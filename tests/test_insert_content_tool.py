import json

from fastapi.testclient import TestClient

from daemon.agent_session import agent_session
from daemon.main import app

client = TestClient(app)


def _call(name, args):
    r = client.post("/mcp", json={"jsonrpc": "2.0", "id": 7, "method": "tools/call",
                                  "params": {"name": name, "arguments": args}}).json()
    return r["result"], json.loads(r["result"]["content"][0]["text"])


def test_insert_content_is_listed():
    r = client.post("/mcp", json={"jsonrpc": "2.0", "id": 1, "method": "tools/list"}).json()
    tool = next(t for t in r["result"]["tools"] if t["name"] == "word_insert_content")
    assert "blocks" in tool["inputSchema"]["required"]


def test_valid_blocks_are_built_daemon_side_and_sent_as_one_ooxml_insert(monkeypatch):
    sent = {}

    async def fake_call_tool(name, args, timeout_sec=60.0):
        sent["name"], sent["args"] = name, args
        return {"ok": True, "inserted": True}

    monkeypatch.setattr(agent_session, "call_tool", fake_call_tool)
    blocks = [{"type": "heading", "level": 1, "text": "Executive summary"},
              {"type": "table", "rows": [[{"text": "Budget", "colspan": 2}], ["CAPEX", "€ 75.000,00"]]}]
    result, payload = _call("word_insert_content", {"blocks": blocks, "location": "end"})
    assert payload["ok"] is True
    assert sent["name"] == "word_insert_ooxml"
    assert sent["args"]["location"] == "end"
    assert "<pkg:package" in sent["args"]["ooxml"] and "gridSpan" in sent["args"]["ooxml"]
    assert sent["args"]["block_id"]  # every agent block is tagged so it can be edited later
    assert payload["block_id"] == sent["args"]["block_id"]


def test_invalid_blocks_return_errors_to_the_model_without_touching_word(monkeypatch):
    called = []

    async def fake_call_tool(name, args, timeout_sec=60.0):
        called.append(name)
        return {"ok": True}

    monkeypatch.setattr(agent_session, "call_tool", fake_call_tool)
    result, payload = _call("word_insert_content", {"blocks": [{"type": "heading", "level": 9, "text": "x"}]})
    assert called == []
    assert payload["ok"] is False and "blocks[0]" in payload["error"]
    assert result["isError"] is True
