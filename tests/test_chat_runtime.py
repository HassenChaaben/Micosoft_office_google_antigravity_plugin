import json
import sys
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from daemon.agent_session import agent_session
from daemon.core.agy_process import AgyProcess
from daemon.core.store import Store
from daemon.main import app

ROOT = Path(__file__).parent
FAKE = [sys.executable, str(ROOT / "fake_agy.py")]
CHARTER = (ROOT / "fixtures" / "docs" / "project_charter.body.xml").read_text(encoding="utf-8")


@pytest.fixture
def client(tmp_path, monkeypatch):
    env_holder = {"FAKE_AGY_MODE": "remember", "FAKE_AGY_STATE_DIR": str(tmp_path / "state")}

    def factory(model, effort=None, conversation_id=None):
        import os
        env = dict(os.environ, **env_holder)
        return AgyProcess(None, model, effort=effort, conversation_id=conversation_id,
                          command_prefix=FAKE, cwd=str(tmp_path), env=env)

    agent_session.configure(store=Store(tmp_path / "h.db"), process_factory=factory)
    # one event loop for the whole test, like uvicorn in production (warm processes belong to it)
    with TestClient(app) as c:
        yield c, env_holder
        c.portal.call(agent_session.shutdown)


def _run(c, **body):
    events = []
    with c.stream("POST", "/api/agent/run", json={"host": "Word", "model": "gemini-3.8-flash-low", **body}) as r:
        for line in r.iter_lines():
            if line.startswith("data: ") and line[6:] != "[DONE]":
                events.append(json.loads(line[6:]))
    return events


def _text(events):
    return "".join(e.get("delta", "") for e in events if e["event"] == "text")


def test_run_creates_a_persisted_chat_and_streams_the_answer(client):
    c, _ = client
    ev = _run(c, prompt="Ciao, prepara il documento", document_ooxml=CHARTER, doc_url="C:\\x\\charter.docx")
    chat_id = next(e["chat_id"] for e in ev if e["event"] == "chat")
    assert any(e["event"] == "analysis" for e in ev)
    assert ev[-1]["event"] == "agent_finished"

    chats = c.get("/api/chats", params={"doc_url": "C:\\x\\charter.docx"}).json()["chats"]
    assert chats[0]["id"] == chat_id
    detail = c.get(f"/api/chats/{chat_id}").json()
    assert [m["role"] for m in detail["messages"]] == ["user", "assistant"]
    assert detail["messages"][0]["content"] == "Ciao, prepara il documento"


def test_second_turn_in_same_chat_remembers_the_first(client):
    c, _ = client
    ev1 = _run(c, prompt="Il codice progetto è PRJ-2026-014", document_ooxml=CHARTER)
    chat_id = next(e["chat_id"] for e in ev1 if e["event"] == "chat")
    ev2 = _run(c, prompt="Qual è il codice?", chat_id=chat_id, document_ooxml=CHARTER)
    # fake agy in 'remember' mode answers with the previous user message it received
    assert "PRJ-2026-014" in _text(ev2)


def test_form_request_routes_expert_skills_into_the_first_message(client):
    c, _ = client
    ev1 = _run(c, prompt="compila il documento", document_ooxml=CHARTER)
    chat_id = next(e["chat_id"] for e in ev1 if e["event"] == "chat")
    skills_ev = next(e for e in ev1 if e["event"] == "skills")
    assert "template-form-filling" in skills_ev["names"]
    ev2 = _run(c, prompt="ok", chat_id=chat_id, document_ooxml=CHARTER)
    first_message = _text(ev2)  # the echo of what the model received on turn 1
    assert "Template form filling" in first_message and "DOCUMENT MAP" in first_message


def test_quota_error_is_reported_with_reset_time(client):
    c, env = client
    env["FAKE_AGY_MODE"] = "quota"
    ev = _run(c, prompt="hello", model="gemini-4-argon-medium")
    err = next(e for e in ev if e["event"] == "error")
    assert err["code"] == "quota" and "96h" in (err.get("retry_after") or "")


def test_rename_and_delete_chat(client):
    c, _ = client
    ev = _run(c, prompt="primo messaggio")
    chat_id = next(e["chat_id"] for e in ev if e["event"] == "chat")
    assert c.patch(f"/api/chats/{chat_id}", json={"title": "Charter Cloud4Job"}).json()["ok"]
    assert c.get(f"/api/chats/{chat_id}").json()["chat"]["title"] == "Charter Cloud4Job"
    assert c.delete(f"/api/chats/{chat_id}").json()["ok"]
    assert c.get(f"/api/chats/{chat_id}").status_code == 404


def test_models_endpoint_lists_agy_models(client, monkeypatch):
    c, _ = client
    import daemon.main as main
    monkeypatch.setattr(main, "_list_models", lambda: [{"id": "claude-sonnet-5-5-medium", "label": "Claude Sonnet 5.5 (Medium)"}])
    models = c.get("/api/models").json()["models"]
    assert models[0]["id"] == "claude-sonnet-5-5-medium"
