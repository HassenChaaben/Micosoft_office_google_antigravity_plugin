import asyncio
import re
from pathlib import Path

from fastapi.testclient import TestClient

from daemon.agent_session import agent_session
from daemon.main import app
from daemon.office.word_snapshot import build_snapshot

FIX = Path(__file__).parent / "fixtures" / "docs"
CHARTER = (FIX / "project_charter.body.xml").read_text(encoding="utf-8")
client = TestClient(app)


def _rpc(method, params=None, id_=1):
    return client.post("/mcp", json={"jsonrpc": "2.0", "id": id_, "method": method, "params": params or {}}).json()


def test_office_verify_is_listed_as_a_tool():
    names = [t["name"] for t in _rpc("tools/list")["result"]["tools"]]
    assert "office_verify" in names


def test_office_verify_captures_document_and_compares_with_turn_start(monkeypatch):
    filled = re.sub(r"(>Azienda<.*?<w:tc[ >].*?<w:p(?=[ >/])[^>]*?)/>", r"\1><w:r><w:t>Cloud4Job</w:t></w:r></w:p>",
                    CHARTER, count=1, flags=re.S)
    captured = {}

    async def fake_call_tool(name, args, timeout_sec=60.0):
        captured["name"] = name
        return {"ok": True, "document_ooxml": filled, "selection": "", "headers_footers": []}

    monkeypatch.setattr(agent_session, "call_tool", fake_call_tool)
    agent_session.turn_snapshot = build_snapshot(CHARTER)
    res = asyncio.run(agent_session.verify_document({}))
    assert captured["name"] == "__capture_document"
    assert "Cliente › Azienda" in res["changes"]["fields_filled"]
    assert res["summary"]["fields_empty"] == 63


def test_office_verify_without_capture_reports_error(monkeypatch):
    async def fake_call_tool(name, args, timeout_sec=60.0):
        return {"ok": False, "error": "Word is not connected"}

    monkeypatch.setattr(agent_session, "call_tool", fake_call_tool)
    res = asyncio.run(agent_session.verify_document({}))
    assert res["ok"] is False and "Word is not connected" in res["error"]
