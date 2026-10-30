"""
Gemini Antigravity Office Local Daemon
High-performance FastAPI HTTPS server running on localhost:18890.
Provides CORS-free bridge for Office WebView2 Taskpane, MCP Connectors,
Google Gemini Subscription Auth, and Word/PowerPoint Win32 COM engines.
"""

import os
import sys
import json
import re
import argparse
import uvicorn
from typing import Optional, Dict, Any, List
from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, FileResponse, StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

# Relative imports
from .cert_gen import generate_localhost_cert, install_cert_to_windows_store, CERT_FILE, KEY_FILE
from .auth import auth_bridge
from .gemini_client import gemini_client
from .word_engine import word_engine
from .ppt_engine import ppt_engine
from .excel_engine import excel_engine
from .outlook_engine import outlook_engine
from .suite_orchestrator import suite_orchestrator
from .mcp_bridge import mcp_bridge
from .mcp_server import router as mcp_router
from .agent_session import agent_session
from .antigravity_backend import ensure_mcp_registered
from .office.perception import Perception
from .core.agy_process import AgyModelsError, ModelCache, find_agy
import asyncio

perception = Perception()

app = FastAPI(title="Gemini Antigravity Office Daemon", version="1.0.0")
app.include_router(mcp_router)

# Allow all origins for WebView2 and localhost
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def get_base_dir() -> str:
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
        if os.path.exists(os.path.join(exe_dir, "web")):
            return exe_dir
        if hasattr(sys, "_MEIPASS") and os.path.exists(os.path.join(sys._MEIPASS, "web")):
            return sys._MEIPASS
        parent = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        if os.path.exists(os.path.join(parent, "web")):
            return parent
        return exe_dir
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

BASE_DIR = get_base_dir()
WEB_DIR = os.path.join(BASE_DIR, "web")
ASSETS_DIR = os.path.join(BASE_DIR, "assets")

# Mount Assets Directory
if os.path.exists(ASSETS_DIR):
    app.mount("/assets", StaticFiles(directory=ASSETS_DIR), name="assets")

# ---------------------------------------------------------------------------
# Static Web Routes for Office Taskpane
# ---------------------------------------------------------------------------
@app.get("/", response_class=HTMLResponse)
@app.get("/index.html", response_class=HTMLResponse)
async def serve_index():
    index_path = os.path.join(WEB_DIR, "index.html")
    if os.path.exists(index_path):
        return FileResponse(index_path)
    return HTMLResponse("<h2>Gemini Antigravity Office Daemon Running</h2>")

@app.get("/taskpane.css")
async def serve_css():
    return FileResponse(os.path.join(WEB_DIR, "taskpane.css"), media_type="text/css")

@app.get("/taskpane.js")
async def serve_js():
    return FileResponse(os.path.join(WEB_DIR, "taskpane.js"), media_type="application/javascript")

@app.get("/office-bridge.js")
async def serve_bridge_js():
    return FileResponse(os.path.join(WEB_DIR, "office-bridge.js"), media_type="application/javascript")

@app.get("/word-agent.js")
async def serve_word_agent_js():
    return FileResponse(os.path.join(WEB_DIR, "word-agent.js"), media_type="application/javascript")

@app.get("/health")
async def health():
    return {
        "status": "ok",
        "app": "Gemini Antigravity Office Daemon",
        "version": "1.0.0",
        "port": 18890,
        "protocol": "https"
    }

# ---------------------------------------------------------------------------
# Authentication Routes (Direct Google Gemini Subscription)
# ---------------------------------------------------------------------------
@app.get("/api/auth/status")
async def get_auth_status():
    return auth_bridge.get_auth_status()

@app.post("/api/auth/login")
async def login():
    url = auth_bridge.start_oauth_flow()
    return {"status": "started", "auth_url": url}

@app.get("/auth/callback")
async def auth_callback(code: str = "success"):
    auth_bridge.handle_callback(code)
    return HTMLResponse("""
    <html>
      <body style="font-family:sans-serif;text-align:center;padding:50px;">
        <h2 style="color:#1A73E8;">Google Gemini Subscription Connected!</h2>
        <p>You can close this tab and return to Microsoft Office.</p>
        <script>setTimeout(() => window.close(), 1500);</script>
      </body>
    </html>
    """)

@app.get("/api/auth/api_key")
async def get_api_key():
    key = auth_bridge.get_api_key()
    has_key = bool(key)
    preview = f"{key[:6]}...{key[-4:]}" if has_key and len(key) > 10 else ""
    return {"has_key": has_key, "preview": preview}

@app.post("/api/auth/api_key")
async def set_api_key(request: Request):
    data = await request.json()
    key = data.get("api_key", "").strip()
    auth_bridge.set_api_key(key)
    return {"status": "saved", "has_key": bool(key)}

# ---------------------------------------------------------------------------
# MCP Connectors & Plugins Routes
# ---------------------------------------------------------------------------
@app.get("/api/connectors")
async def list_connectors():
    return mcp_bridge.get_connector_list()

@app.post("/api/connectors/sync")
async def sync_connectors():
    connectors = mcp_bridge.discover_connectors()
    return {"status": "synced", "connectors": connectors}

# ---------------------------------------------------------------------------
# AI Chat & Streaming Routes
# ---------------------------------------------------------------------------
@app.on_event("startup")
async def on_startup():
    try:
        ensure_mcp_registered()
    except Exception as e:
        print(f"[Daemon] Startup MCP registration notice: {e}")

@app.post("/api/agent/run")
async def agent_run(request: Request):
    payload = await request.json()
    prompt = payload.get("prompt", "")
    host = payload.get("host", "Word")
    model = payload.get("model", "gemini-3.8-flash-high")

    if agent_session.is_running:
        async def _already_running_stream():
            err_obj = {
                "event": "error",
                "error": "Another agent task is already running. Please wait or cancel it.",
                "code": "already_running",
            }
            yield f"data: {json.dumps(err_obj, ensure_ascii=False)}\n\n"
            yield "data: [DONE]\n\n"
        return StreamingResponse(_already_running_stream(), media_type="text/event-stream")

    perceived = await perception.analyze(payload)

    async def _stream_wrapper():
        async for chunk in agent_session.run(
            prompt=prompt, host=host, model=model,
            document_map=perceived.map_text,
            analysis={**perceived.summary, "cached": perceived.cached},
            snapshot=perceived.snapshot,
            chat_id=payload.get("chat_id"),
            doc_url=payload.get("doc_url"),
            doc_title=payload.get("doc_title"),
            effort=payload.get("effort")
        ):
            if "Another agent task is already running" in chunk and '"code":' not in chunk:
                try:
                    for line in chunk.splitlines():
                        if line.startswith("data: "):
                            d = json.loads(line[6:])
                            if d.get("event") == "error" and not d.get("code"):
                                d["code"] = "already_running"
                                chunk = f"data: {json.dumps(d, ensure_ascii=False)}\n\n"
                except Exception:
                    pass
            yield chunk

    return StreamingResponse(
        _stream_wrapper(),
        media_type="text/event-stream"
    )

_models_cache = ModelCache(ttl=600.0)


def _format_family_name(family: Optional[str]) -> str:
    if not family:
        return "Other"
    known = {
        "gemini-3.8-flash": "Gemini 3.8 Flash",
        "gemini-4-argon": "Gemini 4 Argon",
        "claude-sonnet-5-5": "Claude Sonnet 5.5",
        "claude-opus-5-5": "Claude Opus 5.5",
        "gpt-oss-120b": "GPT-OSS 120B",
    }
    if family in known:
        return known[family]
    s = re.sub(r"(\d)-(\d)", r"\1.\2", family)
    words = s.split("-")
    formatted = []
    for w in words:
        if w.lower() in ("gpt", "oss", "ai", "llm"):
            formatted.append(w.upper())
        elif w:
            formatted.append(w.capitalize())
    return " ".join(formatted)


def _list_models():
    agy = find_agy()
    return _models_cache.get(agy) if agy else []


@app.get("/api/models")
async def list_models_endpoint():
    try:
        raw_models = await asyncio.to_thread(_list_models)
        models = []
        families = {}
        for m in raw_models:
            d = dict(m)
            fam = d.get("family") or d.get("id", "other")
            display = _format_family_name(fam)
            d["family_display"] = display
            d["family_label"] = display
            families[fam] = display
            models.append(d)
        return {"models": models, "families": families}
    except AgyModelsError as e:
        return JSONResponse({"models": [], "families": {}, "error": str(e), "code": e.code}, status_code=200)


@app.get("/api/chats")
async def list_chats_endpoint(host: str = None, doc_url: str = None, limit: int = 50):
    return {"chats": agent_session.store.list_chats(host=host, doc_url=doc_url, limit=limit, only_with_assistant=True)}


@app.get("/api/chats/{chat_id}")
async def get_chat_endpoint(chat_id: str):
    chat = agent_session.store.get_chat(chat_id)
    if chat is None:
        return JSONResponse({"error": "Chat not found"}, status_code=404)
    messages = agent_session.store.list_messages(chat_id)
    tool_calls = agent_session.store.list_tool_calls(chat_id)
    plans = agent_session.store.list_plans(chat_id) if hasattr(agent_session.store, "list_plans") else []
    turns = agent_session.store.list_turns(chat_id) if hasattr(agent_session.store, "list_turns") else []

    turn_map = {}
    for t in turns:
        a_msg = t.get("assistant_message")
        if a_msg and a_msg.get("id"):
            turn_map[a_msg["id"]] = t

    enriched_messages = []
    for msg in messages:
        m = dict(msg)
        if m["role"] == "assistant" and m["id"] in turn_map:
            m["plans"] = turn_map[m["id"]]["plans"]
            m["tool_calls"] = turn_map[m["id"]]["tool_calls"]
        elif m["role"] == "assistant":
            m.setdefault("plans", [])
            m.setdefault("tool_calls", [])
        enriched_messages.append(m)

    return {
        "chat": chat,
        "messages": enriched_messages,
        "tool_calls": tool_calls,
        "plans": plans,
        "turns": turns,
    }


@app.patch("/api/chats/{chat_id}")
async def rename_chat_endpoint(chat_id: str, request: Request):
    title = ((await request.json()) or {}).get("title", "").strip()
    if not title or agent_session.store.get_chat(chat_id) is None:
        return JSONResponse({"ok": False, "error": "Chat not found or empty title"}, status_code=400)
    agent_session.store.rename_chat(chat_id, title)
    return {"ok": True}


@app.delete("/api/chats/{chat_id}")
async def delete_chat_endpoint(chat_id: str):
    agent_session.store.delete_chat(chat_id)
    return {"ok": True}


@app.post("/api/agent/tool_result")
async def agent_tool_result(request: Request):
    data = await request.json()
    call_id = data.get("call_id")
    result = data.get("result", {})
    resolved = agent_session.resolve_tool(call_id, result)
    return {"status": "resolved" if resolved else "ignored"}

@app.post("/api/agent/cancel")
async def agent_cancel():
    await agent_session.cancel()
    return {"status": "cancelled"}

@app.post("/api/agent/reset")
async def agent_reset():
    agent_session.reset_history()
    return {"status": "reset"}

@app.post("/api/chat")
async def chat_endpoint(request: Request):
    payload = await request.json()
    return StreamingResponse(
        gemini_client.stream_chat(payload),
        media_type="text/event-stream"
    )

# ---------------------------------------------------------------------------
# Microsoft Word Automation Endpoints
# ---------------------------------------------------------------------------
@app.post("/api/document/word/read")
async def word_read():
    return word_engine.read_document()

@app.post("/api/document/word/replace")
async def word_replace(request: Request):
    data = await request.json()
    text = data.get("text", "")
    return word_engine.replace_selection(text)

@app.post("/api/document/word/comment")
async def word_comment(request: Request):
    data = await request.json()
    comment = data.get("comment", "")
    return word_engine.insert_comment(comment)

@app.post("/api/document/word/track_changes")
async def word_track_changes(request: Request):
    data = await request.json()
    enable = data.get("enable", True)
    return word_engine.toggle_track_changes(enable)

# ---------------------------------------------------------------------------
# Microsoft PowerPoint Automation Endpoints
# ---------------------------------------------------------------------------
@app.post("/api/document/ppt/read")
async def ppt_read():
    return ppt_engine.read_presentation()

@app.post("/api/document/ppt/create_slide")
async def ppt_create_slide(request: Request):
    data = await request.json()
    title = data.get("title", "Executive Slide")
    bullets = data.get("bullets", [])
    return ppt_engine.create_slide(title, bullets)

@app.post("/api/document/ppt/storylining")
async def ppt_storylining(request: Request):
    data = await request.json()
    slides = data.get("slides", [])
    return ppt_engine.build_storyline(slides)

# ---------------------------------------------------------------------------
# Microsoft Excel Automation Endpoints
# ---------------------------------------------------------------------------
@app.post("/api/document/excel/read")
async def excel_read():
    return excel_engine.read_sheet()

@app.post("/api/document/excel/write")
async def excel_write(request: Request):
    data = await request.json()
    start_cell = data.get("start_cell", "A1")
    matrix = data.get("matrix", [[]])
    format_table = data.get("format_table", True)
    return excel_engine.write_cell_range(start_cell, matrix, format_table)

@app.post("/api/document/excel/financial_model")
async def excel_model(request: Request):
    data = await request.json()
    model_type = data.get("model_type", "saas_revenue")
    currency = data.get("currency", "$")
    return excel_engine.apply_financial_model(model_type, currency)

@app.post("/api/document/excel/chart")
async def excel_chart(request: Request):
    data = await request.json()
    chart_type = data.get("chart_type", "column")
    data_range = data.get("data_range", "A1:E4")
    title = data.get("title", "Executive Chart")
    return excel_engine.create_chart(chart_type, data_range, title)

# ---------------------------------------------------------------------------
# Microsoft Outlook Automation Endpoints
# ---------------------------------------------------------------------------
@app.post("/api/document/outlook/read")
async def outlook_read():
    return outlook_engine.read_current_email()

@app.post("/api/document/outlook/reply")
async def outlook_reply(request: Request):
    data = await request.json()
    reply_body = data.get("reply_html") or data.get("reply_body", "")
    reply_all = data.get("reply_all", False)
    return outlook_engine.draft_reply(reply_body, reply_all)

@app.post("/api/document/outlook/compose")
async def outlook_compose(request: Request):
    data = await request.json()
    to = data.get("to", "")
    subject = data.get("subject", "Executive Memo")
    body_html = data.get("body_html", "")
    return outlook_engine.create_email(to, subject, body_html)

@app.post("/api/document/outlook/meeting")
async def outlook_meeting(request: Request):
    data = await request.json()
    subject = data.get("subject", "Strategic Sync")
    start_time = data.get("start_time", "2026-10-29 10:00")
    duration = data.get("duration_min", 30)
    attendees = data.get("attendees", [])
    return outlook_engine.schedule_meeting(subject, start_time, duration, attendees)

# ---------------------------------------------------------------------------
# Cross-Application Suite Orchestrator Endpoints
# ---------------------------------------------------------------------------
@app.post("/api/suite/doc_to_slides")
async def suite_doc_to_slides():
    return suite_orchestrator.doc_to_slides()

@app.post("/api/suite/excel_to_report")
async def suite_excel_to_report():
    return suite_orchestrator.excel_to_report()

@app.post("/api/suite/email_summary")
async def suite_email_summary(request: Request):
    data = await request.json()
    recipient = data.get("recipient", "team@cloud4job.com")
    return suite_orchestrator.email_suite_summary(recipient)

# ---------------------------------------------------------------------------
# Entrypoint & Daemon Server Launch
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Gemini Antigravity Office Daemon")
    parser.add_argument("--background", action="store_true", help="Run as background service")
    parser.add_argument("--port", type=int, default=18890, help="HTTPS Port (Default: 18890)")
    args = parser.parse_args()

    # Ensure SSL Certs are generated and trusted
    cert_path, key_path = generate_localhost_cert(force=False)
    install_cert_to_windows_store(cert_path)

    print(f"============================================================")
    print(f"  Gemini Antigravity Office Local Daemon")
    print(f"  Serving HTTPS on: https://localhost:{args.port}")
    print(f"  Word & PowerPoint COM Automation: Active")
    print(f"  Google Subscription Auth Bridge: Active")
    print(f"============================================================")

    uvicorn.run(
        app,
        host="127.0.0.1",
        port=args.port,
        ssl_keyfile=key_path,
        ssl_certfile=cert_path,
        log_level="info"
    )

if __name__ == "__main__":
    main()
