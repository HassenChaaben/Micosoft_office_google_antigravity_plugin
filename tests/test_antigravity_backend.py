import asyncio
import json

import pytest

from daemon import antigravity_backend as ab
from daemon import gemini_client as gc_module
from daemon.gemini_client import GeminiClient


def _delta_line(text, state="ACTIVE", step_type="agent_response"):
    return json.dumps({
        "event": "step_update",
        "step_update": {"step_index": 1, "state": state, "step_type": step_type, "text_delta": text},
    })


def test_parse_stream_line_extracts_agent_text_delta():
    assert ab.parse_stream_line(_delta_line("| name |")) == ("delta", "| name |")


def test_parse_stream_line_ignores_user_input_and_init_events():
    assert ab.parse_stream_line(_delta_line("echo", step_type="user_input")) is None
    assert ab.parse_stream_line(json.dumps({"event": "init", "init": {}})) is None


def test_parse_stream_line_reports_result_status():
    ok = json.dumps({"event": "result", "result": {"status": "SUCCESS", "response": "x"}})
    bad = json.dumps({"event": "result", "result": {"status": "ERROR", "error": "quota exceeded"}})
    assert ab.parse_stream_line(ok) == ("done", "")
    kind, msg = ab.parse_stream_line(bad)
    assert kind == "error"
    assert "quota exceeded" in msg


def test_parse_stream_line_skips_non_json_noise():
    assert ab.parse_stream_line("Fetching available models...") is None
    assert ab.parse_stream_line("") is None


def test_build_prompt_contains_request_host_and_document():
    p = ab.build_prompt("three columns product_name, quantity, price", "Word", "Existing doc text")
    assert "three columns product_name, quantity, price" in p
    assert "Microsoft Word" in p
    assert "Existing doc text" in p


def test_build_prompt_truncates_huge_document_context():
    p = ab.build_prompt("summarize", "Word", "a" * 100_000)
    assert len(p) < ab.MAX_PROMPT_CHARS + 2_000


def test_build_command_maps_ui_models_and_uses_stream_json():
    cmd = ab.build_command("agy.exe", "hello", "gemini-pro")
    assert cmd[0] == "agy.exe"
    assert "--output-format" in cmd and cmd[cmd.index("--output-format") + 1] == "stream-json"
    assert cmd[cmd.index("--model") + 1] == ab.MODEL_MAP["gemini-pro"]
    assert cmd[-2:] == ["-p", "hello"]
    assert ab.build_command("agy.exe", "x", "unknown-model")[ab.build_command("agy.exe", "x", "unknown-model").index("--model") + 1] == ab.DEFAULT_MODEL


def _collect(agen):
    async def run():
        return [chunk async for chunk in agen]
    return "".join(asyncio.run(run()))


def test_stream_chat_streams_real_antigravity_deltas(monkeypatch):
    monkeypatch.delenv("GEMINI_OFFICE_BACKEND", raising=False)
    monkeypatch.setattr(ab, "find_agy", lambda: "agy.exe")
    monkeypatch.setattr(ab, "stream_events", lambda *a, **k: iter([
        ("delta", "| Product Name | Quantity | Price |"), ("delta", "\n| A | 1 | $2 |"), ("done", ""),
    ]))
    body = _collect(GeminiClient().stream_chat({"prompt": "3 columns", "host": "Word"}))
    assert "Product Name" in body
    assert "data: [DONE]" in body
    assert "Executive Report" not in body


def test_stream_chat_without_any_backend_returns_honest_error(monkeypatch):
    monkeypatch.delenv("GEMINI_OFFICE_BACKEND", raising=False)
    monkeypatch.setattr(ab, "find_agy", lambda: None)
    monkeypatch.setattr(gc_module.auth_bridge, "get_api_key", lambda: None)
    body = _collect(GeminiClient().stream_chat({"prompt": "create a table that for daily tasks tracking", "host": "Word"}))
    assert "agy" in body
    assert "Daily Tasks Tracking Matrix" not in body
    assert "data: [DONE]" in body


def test_stream_chat_reports_antigravity_error_instead_of_template(monkeypatch):
    monkeypatch.delenv("GEMINI_OFFICE_BACKEND", raising=False)
    monkeypatch.setattr(ab, "find_agy", lambda: "agy.exe")
    monkeypatch.setattr(ab, "stream_events", lambda *a, **k: iter([("error", "not logged in")]))
    monkeypatch.setattr(gc_module.auth_bridge, "get_api_key", lambda: None)
    body = _collect(GeminiClient().stream_chat({"prompt": "create a table that for daily tasks tracking", "host": "Word"}))
    assert "not logged in" in body
    assert "Daily Tasks Tracking Matrix" not in body
