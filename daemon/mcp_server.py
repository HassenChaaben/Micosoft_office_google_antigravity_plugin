"""
Internal Model Context Protocol (MCP) Server for Microsoft Word
Implements JSON-RPC 2.0 endpoints for Streamable HTTP MCP communication.
Exposes word-office tools to the local Antigravity AI CLI (`agy`).
"""

import json
from fastapi import APIRouter, Request, Response
from fastapi.responses import JSONResponse

import time
import uuid
from .word_tools import get_tool_definitions, get_tool
from .agent_session import agent_session
from .office.word_com_tools import TOOLS as COM_TOOLS

router = APIRouter()

# Tools executed by the daemon itself (not forwarded as-is to the taskpane).
DAEMON_TOOLS = [
    {
        "name": "office_verify",
        "description": (
            "Verify your work like an Office expert before you report back. Re-reads the whole document "
            "and compares it with the state at the start of this turn: fields filled, fields still empty, "
            "fields accidentally cleared, tables/rows removed, leftover placeholders, heading level skips, "
            "stale table of contents. Call it after your edits; fix every issue it reports, then call it again."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "expect_all_fields_filled": {
                    "type": "boolean",
                    "description": "True (default) when the user asked to fill/complete the document.",
                }
            },
        },
    }
]
BLOCKS_DOC = (
    "List of blocks. Types: heading{level 1-6,text} | title{text} | subtitle{text} | "
    "paragraph{text | runs[{text,bold,italic,underline,color '#RRGGBB',size,highlight,font}], align left|center|right|justify, style} | "
    "bullets{items} | numbered{items} (item = string or {text, children:[items]}, max 3 levels) | "
    "table{columns[{width_pct}], header_rows=1, repeat_header=true, style, borders, header_shading, "
    "rows[[cell]]} where cell = string or {text|runs, colspan, rowspan, shading, bold, align, valign}; do NOT write the cells "
    "covered by a rowspan/colspan | page_break | caption{label,text} | quote{text} | toc{levels:'1-3'}."
)
DAEMON_TOOLS.append({
    "name": "word_insert_content",
    "description": (
        "FASTEST and preferred way to write any multi-part content (report sections, headings + text, lists, "
        "tables with merged cells and shading, captions, TOC). Everything is inserted in ONE atomic native Word "
        "operation with built-in styles (Heading 1-6, Title, Subtitle, Caption...), which work in any Word language. "
        "Returns block_id; the content is wrapped in content control 'ga:<block_id>' so you can target it later "
        "(location 'after:ga:<id>' or 'replace:ga:<id>')."
    ),
    "inputSchema": {
        "type": "object",
        "properties": {
            "blocks": {"type": "array", "items": {"type": "object"}, "description": BLOCKS_DOC},
            "location": {
                "type": "string",
                "description": "cursor (default) | start | end | replace_selection | after:p<index> | after:<heading text> "
                               "| after:ga:<block_id> | replace:ga:<block_id>",
            },
        },
        "required": ["blocks"],
    },
})
DAEMON_TOOLS += [
    {
        "name": "agent_plan",
        "description": "Show the user your plan BEFORE acting on any multi-step task (one short line per step, "
                       "in the user's language). Then execute it and call office_verify at the end.",
        "inputSchema": {"type": "object", "properties": {"steps": {"type": "array", "items": {"type": "string"}}},
                        "required": ["steps"]},
    },
    {
        "name": "skill_load",
        "description": "Load the full instructions of an Office expert skill listed in your prompt (by name).",
        "inputSchema": {"type": "object", "properties": {"name": {"type": "string"}}, "required": ["name"]},
    },
    {
        "name": "memory_save",
        "description": "Remember a durable user preference or project fact for future chats (e.g. 'subtitles centered, "
                       "14pt', 'project code PRJ-2026-014, budget € 75.000,00'). scope 'document' = only this "
                       "document, 'global' = every document.",
        "inputSchema": {"type": "object", "properties": {
            "text": {"type": "string"}, "scope": {"type": "string", "enum": ["document", "global"]}},
            "required": ["text"]},
    },
    {
        "name": "memory_search",
        "description": "Search remembered preferences and project facts (this document + global).",
        "inputSchema": {"type": "object", "properties": {"query": {"type": "string"}}, "required": ["query"]},
    },
]
DAEMON_TOOLS.append({
    "name": "word_fill_fields",
    "description": (
        "Fill form/template fields IN PLACE, keeping each cell's existing formatting. Keys: field paths exactly as in "
        "the Document Map (e.g. 'Cliente › Azienda'), grid cells as '<grid> › #<record> › <column>' "
        "(e.g. 'Minacce › #2 › Causa'), cell coordinates 'T1r2c3', or content-control tags. Fill a whole table "
        "(or the whole document) in ONE call. Unknown keys are reported with suggestions; nothing else is touched."
    ),
    "inputSchema": {"type": "object", "properties": {
        "fields": {"type": "object", "description": "{key: value} — values are plain text."}},
        "required": ["fields"]},
})
COM_TOOL_NAMES = set(COM_TOOLS.keys())
VISUAL_TOOLS = {"word_insert_chart", "word_insert_smartart", "word_insert_diagram"}

for _tname, _tdata in COM_TOOLS.items():
    DAEMON_TOOLS.append({
        "name": _tname,
        "description": _tdata["description"],
        "inputSchema": _tdata["inputSchema"],
        "destructive": _tdata.get("destructive", False),
    })

DAEMON_TOOL_NAMES = {t["name"] for t in DAEMON_TOOLS}


def _memory_scopes(doc_url):
    return ["global"] + ([f"doc:{doc_url}"] if doc_url else [])


async def run_daemon_tool(name: str, args: dict) -> dict:
    if name == "office_verify":
        return await agent_session.verify_document(args)
    if name == "word_insert_content":
        from .office.ooxml_builder import build_ooxml, validate_blocks

        blocks = args.get("blocks") or []
        errors = validate_blocks(blocks)
        if errors:
            return {"ok": False, "error": "Invalid blocks, fix and retry:\n- " + "\n- ".join(errors)}
        block_id = args.get("block_id") or uuid.uuid4().hex[:8]
        ooxml = build_ooxml(blocks, block_id=block_id)
        res = await agent_session.call_tool(
            "word_insert_ooxml", {"ooxml": ooxml, "location": args.get("location") or "cursor", "block_id": block_id})
        return {**res, "block_id": block_id, "blocks": len(blocks)} if res.get("ok") else res
    if name == "word_fill_fields":
        from .office.fill import resolve_fields

        fields = args.get("fields") or {}
        if not isinstance(fields, dict) or not fields:
            return {"ok": False, "error": "Pass 'fields' as an object {key: value}."}
        if agent_session.turn_snapshot is None:
            return {"ok": False, "error": "No document snapshot for this turn; use T{t}r{r}c{c} keys or read the document."}
        plan = resolve_fields(agent_session.turn_snapshot, fields)
        if not plan["cells"] and not plan["content_controls"]:
            return {"ok": False, "error": "No field matched.\n- " + "\n- ".join(plan["errors"])}
        res = await agent_session.call_tool("word_fill_cells", {
            "cells": [{k: c[k] for k in ("table", "row", "cell", "text", "placeholder") if k in c} for c in plan["cells"]],
            "content_controls": plan["content_controls"]})
        if not res.get("ok"):
            return res
        return {"ok": True, "written": res.get("written", len(plan["cells"])) , "failed": res.get("failed", []),
                "not_found": plan["errors"]}
    if name == "agent_plan":
        steps = [str(s) for s in (args.get("steps") or []) if str(s).strip()]
        if not steps:
            return {"ok": False, "error": "Give at least one step."}
        if agent_session._current_chat:
            agent_session.store.add_plan(agent_session._current_chat, None, steps)
        await agent_session.event_queue.put({"event": "plan", "steps": steps})
        return {"ok": True, "steps": len(steps)}
    if name == "skill_load":
        wanted = (args.get("name") or "").strip().lstrip("/")
        skill = next((s for s in agent_session.skills() if s.name == wanted), None)
        if skill is None:
            return {"ok": False, "error": f"No skill '{wanted}'. Available: "
                    + ", ".join(s.name for s in agent_session.skills())}
        return {"ok": True, "name": skill.name, "body": skill.body}
    doc_url = getattr(agent_session, "_current_doc_url", None)
    if name == "memory_save":
        text = (args.get("text") or "").strip()
        if not text:
            return {"ok": False, "error": "Empty memory."}
        scope = f"doc:{doc_url}" if args.get("scope", "document") == "document" and doc_url else "global"
        mem = agent_session.store.add_memory(scope, text)
        return {"ok": True, "id": mem["id"], "scope": scope}
    if name == "memory_search":
        found = agent_session.store.search_memories(args.get("query") or "", _memory_scopes(doc_url), limit=10)
        return {"ok": True, "memories": [{"text": m["text"], "scope": m["scope"]} for m in found]}

    if name in COM_TOOL_NAMES:
        worker = agent_session.com_worker
        tool_entry = COM_TOOLS[name]
        fn = tool_entry["fn"]

        call_args = dict(args)
        # Adapt legacy paragraph_index / comment parameters if provided
        if "paragraph_index" in call_args and "anchor" not in call_args:
            call_args["anchor"] = f"p{call_args['paragraph_index']}"
        if name == "word_add_comment" and "comment" in call_args and "text" not in call_args:
            call_args["text"] = call_args["comment"]
        if name == "word_insert_cross_reference" and call_args.get("kind") == "bookmark":
            target = str(call_args.get("target_anchor", ""))
            if target.startswith("ga:"):
                call_args["target_anchor"] = f"ga_{target[3:].replace('-', '_').replace(':', '_')}"

        started = time.monotonic()

        def _execute_com():
            try:
                doc = worker.get_document(doc_url)
            except Exception as exc:
                return {"ok": False, "error": f"Failed to get Word document: {exc}"}

            try:
                res = fn(doc, **call_args)
            except Exception as exc:
                return {"ok": False, "error": f"COM tool '{name}' failed: {exc}"}

            if not isinstance(res, dict):
                return {"ok": True, "result": res}

            if res.get("ok") and name in VISUAL_TOOLS:
                anchor_id = str(call_args.get("id") or call_args.get("anchor_id") or uuid.uuid4().hex[:8])
                if anchor_id.startswith("ga:"):
                    anchor_id = anchor_id[3:]
                ga_tag = f"ga:{anchor_id}"
                clean_id = anchor_id.replace("-", "_").replace(":", "_")
                bm_name = f"ga_{clean_id}"

                target_range = None
                if name == "word_insert_chart":
                    idx = res.get("inline_shape_index")
                    if idx is not None:
                        try:
                            ish = doc.InlineShapes(idx + 1)
                            target_range = ish.Range
                        except Exception:
                            pass
                elif name == "word_insert_smartart":
                    anchor_str = str(res.get("anchor", ""))
                    if anchor_str.startswith("i") and anchor_str[1:].isdigit():
                        try:
                            idx = int(anchor_str[1:])
                            ish = doc.InlineShapes(idx + 1)
                            target_range = ish.Range
                        except Exception:
                            pass
                    if target_range is None:
                        try:
                            if doc.Shapes.Count > 0:
                                target_range = doc.Shapes(doc.Shapes.Count).Anchor
                        except Exception:
                            pass
                elif name == "word_insert_diagram":
                    shape_name = res.get("name")
                    if shape_name:
                        try:
                            shape = doc.Shapes(shape_name)
                            target_range = shape.Anchor
                        except Exception:
                            pass

                if target_range is not None:
                    try:
                        cc = doc.ContentControls.Add(1, target_range)  # 1 = wdContentControlRichText
                        cc.Tag = anchor_id
                        cc.Title = ga_tag
                    except Exception:
                        pass
                    try:
                        doc.Bookmarks.Add(bm_name, target_range)
                    except Exception:
                        pass
                    if clean_id != bm_name:
                        try:
                            doc.Bookmarks.Add(clean_id, target_range)
                        except Exception:
                            pass

                res["anchor"] = ga_tag
                res["anchor_id"] = anchor_id
                res["bookmark"] = bm_name

            return res

        res = await worker.run(_execute_com)

        agent_session._current_turn_tools.append(f"{name}({json.dumps(args)})")
        agent_session._record_tool(name, args, res, started)

        if res.get("ok"):
            agent_session.invalidate_perception_cache()
            await agent_session.refresh_perception(doc_url)

        return res

    return {"ok": False, "error": f"Unknown daemon tool {name}"}


@router.get("/mcp")
async def mcp_get():
    """Streamable HTTP MCP probe / health check."""
    return Response(content="Word Office MCP Server Active", media_type="text/plain")


@router.post("/mcp")
async def mcp_post(request: Request):
    """
    Handles JSON-RPC 2.0 messages from the Antigravity CLI MCP client.
    Methods: initialize, notifications/initialized, tools/list, tools/call
    """
    try:
        body = await request.json()
    except Exception:
        return JSONResponse(
            {"jsonrpc": "2.0", "id": None, "error": {"code": -32700, "message": "Parse error"}},
            status_code=400
        )

    method = body.get("method")
    msg_id = body.get("id")

    # Notifications do not have an ID and do not return a JSON-RPC body
    if msg_id is None:
        return Response(status_code=202)

    # 1. Initialize
    if method == "initialize":
        client_version = body.get("params", {}).get("protocolVersion", "2025-03-26")
        return JSONResponse({
            "jsonrpc": "2.0",
            "id": msg_id,
            "result": {
                "protocolVersion": client_version,
                "capabilities": {
                    "tools": {}
                },
                "serverInfo": {
                    "name": "word-office",
                    "version": "1.0.0"
                }
            }
        })

    # 2. Tools List
    if method == "tools/list":
        builtin_tools = [t for t in get_tool_definitions() if t["name"] not in DAEMON_TOOL_NAMES]
        return JSONResponse({
            "jsonrpc": "2.0",
            "id": msg_id,
            "result": {
                "tools": builtin_tools + DAEMON_TOOLS
            }
        })

    # 3. Tools Call
    if method == "tools/call":
        params = body.get("params", {})
        tool_name = params.get("name")
        arguments = params.get("arguments", {})

        if tool_name in DAEMON_TOOL_NAMES:
            tool_result = await run_daemon_tool(tool_name, arguments)
            return JSONResponse({
                "jsonrpc": "2.0",
                "id": msg_id,
                "result": {
                    "content": [{"type": "text", "text": json.dumps(tool_result, ensure_ascii=False)}],
                    "isError": tool_result.get("ok") is False and "error" in tool_result,
                },
            })

        tool_def = get_tool(tool_name)
        if not tool_def:
            return JSONResponse({
                "jsonrpc": "2.0",
                "id": msg_id,
                "error": {
                    "code": -32601,
                    "message": f"Unknown tool: '{tool_name}'"
                }
            })

        # Delegate execution to the active Word taskpane session
        tool_result = await agent_session.call_tool(tool_name, arguments)

        return JSONResponse({
            "jsonrpc": "2.0",
            "id": msg_id,
            "result": {
                "content": [
                    {
                        "type": "text",
                        "text": json.dumps(tool_result, ensure_ascii=False)
                    }
                ]
            }
        })

    # Method not found
    return JSONResponse({
        "jsonrpc": "2.0",
        "id": msg_id,
        "error": {
            "code": -32601,
            "message": f"Method not found: {method}"
        }
    })
