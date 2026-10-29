import asyncio
import json

import pytest
from fastapi.testclient import TestClient

from daemon.agent_session import agent_session
from daemon.core.store import Store
from daemon.main import app

client = TestClient(app)


@pytest.fixture(autouse=True)
def fresh_store(tmp_path):
    agent_session.configure(store=Store(tmp_path / "h.db"))
    chat = agent_session.store.create_chat(host="Word", doc_url="C:\\d\\a.docx")
    agent_session._current_chat = chat["id"]
    agent_session._current_doc_url = "C:\\d\\a.docx"
    agent_session.event_queue = asyncio.Queue()
    yield chat


def _call(name, args):
    r = client.post("/mcp", json={"jsonrpc": "2.0", "id": 9, "method": "tools/call",
                                  "params": {"name": name, "arguments": args}}).json()
    return json.loads(r["result"]["content"][0]["text"])


def test_tools_are_listed():
    r = client.post("/mcp", json={"jsonrpc": "2.0", "id": 1, "method": "tools/list"}).json()
    names = {t["name"] for t in r["result"]["tools"]}
    assert {"skill_load", "memory_save", "memory_search", "agent_plan"} <= names


def test_skill_load_returns_full_body():
    res = _call("skill_load", {"name": "smartart-and-diagrams"})
    assert res["ok"] and "## Procedure" in res["body"]
    assert _call("skill_load", {"name": "nope"})["ok"] is False


def test_memory_save_and_search_document_scope():
    assert _call("memory_save", {"text": "Il sottotitolo va centrato, 14pt", "scope": "document"})["ok"]
    assert _call("memory_save", {"text": "Scrivi sempre in italiano", "scope": "global"})["ok"]
    found = _call("memory_search", {"query": "sottotitolo"})["memories"]
    assert found[0]["text"].startswith("Il sottotitolo")
    assert found[0]["scope"] == "doc:C:\\d\\a.docx"


def test_agent_plan_is_stored_and_shown_to_the_user(fresh_store):
    res = _call("agent_plan", {"steps": ["Leggere i campi", "Compilare T0", "Verificare"]})
    assert res["ok"]
    ev = agent_session.event_queue.get_nowait()
    assert ev["event"] == "plan" and len(ev["steps"]) == 3


def test_word_tool_calls_are_recorded_in_history(fresh_store, monkeypatch):
    agent_session.is_running = True

    async def scenario():
        task = asyncio.create_task(agent_session.call_tool("word_search", {"search_text": "Budget"}, timeout_sec=5))
        ev = await agent_session.event_queue.get()
        agent_session.resolve_tool(ev["call_id"], {"ok": True, "matches": 1})
        return await task

    try:
        assert asyncio.run(scenario())["ok"]
    finally:
        agent_session.is_running = False
    calls = agent_session.store.list_tool_calls(fresh_store["id"])
    assert calls[0]["tool"] == "word_search" and calls[0]["ok"] is True
