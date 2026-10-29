"""Tests for daemon.core.store (SQLite chat history)."""

import json
import sqlite3
import time

import pytest

from daemon.core.store import SCHEMA_VERSION, Store, default_db_path


@pytest.fixture
def store(tmp_path):
    s = Store(str(tmp_path / "history.db"))
    yield s
    s.close()


def test_default_path_under_localappdata(monkeypatch, tmp_path):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    assert default_db_path().startswith(str(tmp_path))
    assert default_db_path().endswith("GeminiAntigravityOffice\\history.db") or default_db_path().endswith(
        "GeminiAntigravityOffice/history.db")


def test_wal_and_version(store, tmp_path):
    conn = sqlite3.connect(str(tmp_path / "history.db"))
    assert conn.execute("PRAGMA journal_mode").fetchone()[0].lower() == "wal"
    assert conn.execute("PRAGMA user_version").fetchone()[0] == SCHEMA_VERSION
    conn.close()


def test_migrations_idempotent(tmp_path):
    path = str(tmp_path / "h.db")
    s1 = Store(path)
    cid = s1.create_chat(host="word", doc_url="u", doc_title="Doc", model="m")["id"]
    s1.close()
    s2 = Store(path)
    assert s2.get_chat(cid)["doc_title"] == "Doc"
    s2.close()


def test_create_and_get_chat(store):
    chat = store.create_chat(host="word", doc_url="file:///a.docx", doc_title="A", model="gemini-3.8-flash-low",
                             effort="low")
    assert len(chat["id"]) == 36
    got = store.get_chat(chat["id"])
    assert got["host"] == "word"
    assert got["model"] == "gemini-3.8-flash-low"
    assert got["effort"] == "low"
    assert got["title"] == "New chat"
    assert got["agy_conversation_id"] is None
    assert got["created_at"] and got["updated_at"]
    assert store.get_chat("missing") is None


def test_auto_title_from_first_user_message(store):
    cid = store.create_chat(host="word")["id"]
    store.add_message(cid, "system", "system prompt should not title")
    long = "Crea un registro dei rischi per il progetto di automazione Cloud4Job con dieci righe"
    store.add_message(cid, "user", long)
    store.add_message(cid, "user", "second message")
    title = store.get_chat(cid)["title"]
    assert title.startswith("Crea un registro dei rischi")
    assert len(title) <= 61
    assert "second" not in title


def test_auto_title_keeps_manual_rename(store):
    cid = store.create_chat(host="word")["id"]
    store.rename_chat(cid, "My title")
    store.add_message(cid, "user", "hello")
    assert store.get_chat(cid)["title"] == "My title"


def test_messages_order_and_role_check(store):
    cid = store.create_chat(host="word")["id"]
    m1 = store.add_message(cid, "user", "hi")
    m2 = store.add_message(cid, "assistant", "hello")
    msgs = store.list_messages(cid)
    assert [m["id"] for m in msgs] == [m1["id"], m2["id"]]
    assert [m["role"] for m in msgs] == ["user", "assistant"]
    with pytest.raises(ValueError):
        store.add_message(cid, "robot", "x")


def test_list_chats_order_filter_and_preview(store):
    a = store.create_chat(host="word", doc_url="d1")["id"]
    time.sleep(0.01)
    b = store.create_chat(host="excel", doc_url="d2")["id"]
    time.sleep(0.01)
    store.add_message(a, "user", "first question")
    store.add_message(a, "assistant", "latest answer")
    chats = store.list_chats()
    assert [c["id"] for c in chats] == [a, b]
    assert chats[0]["last_message"] == "latest answer"
    assert chats[1]["last_message"] is None
    assert [c["id"] for c in store.list_chats(host="excel")] == [b]
    assert [c["id"] for c in store.list_chats(doc_url="d1")] == [a]
    assert len(store.list_chats(limit=1)) == 1


def test_rename_and_conversation_id(store):
    cid = store.create_chat(host="word")["id"]
    store.rename_chat(cid, "Renamed")
    store.set_conversation_id(cid, "conv-123")
    chat = store.get_chat(cid)
    assert chat["title"] == "Renamed"
    assert chat["agy_conversation_id"] == "conv-123"


def test_delete_chat_cascades(store, tmp_path):
    cid = store.create_chat(host="word")["id"]
    mid = store.add_message(cid, "assistant", "x")["id"]
    store.add_tool_call(cid, mid, "word_insert_table", {"rows": 2}, {"ok": True}, True, 12)
    store.add_plan(cid, mid, [{"step": 1, "text": "do"}])
    store.delete_chat(cid)
    assert store.get_chat(cid) is None
    assert store.list_messages(cid) == []
    assert store.list_tool_calls(cid) == []
    conn = sqlite3.connect(str(tmp_path / "history.db"))
    assert conn.execute("SELECT COUNT(*) FROM plans").fetchone()[0] == 0
    conn.close()


def test_tool_calls_roundtrip(store):
    cid = store.create_chat(host="word")["id"]
    mid = store.add_message(cid, "assistant", "x")["id"]
    store.add_tool_call(cid, mid, "word_format_text", {"bold": True}, {"ok": False, "error": "e"}, False, 5)
    calls = store.list_tool_calls(cid)
    assert len(calls) == 1
    c = calls[0]
    assert c["tool"] == "word_format_text"
    assert json.loads(c["args_json"]) == {"bold": True}
    assert c["ok"] is False
    assert c["duration_ms"] == 5
    assert c["message_id"] == mid


def test_plan(store):
    cid = store.create_chat(host="word")["id"]
    plan = store.add_plan(cid, None, [{"step": 1}])
    assert plan["id"]
    assert json.loads(plan["steps_json"]) == [{"step": 1}]


def test_snapshot_upsert(store):
    assert store.get_snapshot("doc") is None
    store.save_snapshot("doc", "k1", {"paras": 1}, "map1")
    store.save_snapshot("doc", "k2", {"paras": 2}, "map2")
    snap = store.get_snapshot("doc")
    assert snap["change_key"] == "k2"
    assert snap["snapshot"] == {"paras": 2}
    assert snap["map_text"] == "map2"


def test_memories_search(store):
    store.add_memory("global", "User prefers Italian deliverables")
    time.sleep(0.01)
    store.add_memory("doc:file:///a.docx", "Budget is 75000 euro")
    time.sleep(0.01)
    store.add_memory("global", "Use Italian date format")
    res = store.search_memories("italian", ["global"])
    assert [m["text"] for m in res] == ["Use Italian date format", "User prefers Italian deliverables"]
    assert store.search_memories("budget", ["global"]) == []
    assert len(store.search_memories("budget euro", ["global", "doc:file:///a.docx"])) == 1
    # empty query returns newest in scope
    assert len(store.search_memories("", ["global"])) == 2
    with pytest.raises(ValueError):
        store.add_memory("bad-scope", "x")


def test_delete_all(store):
    cid = store.create_chat(host="word")["id"]
    store.add_message(cid, "user", "x")
    store.add_memory("global", "m")
    store.save_snapshot("d", "k", {}, "")
    store.delete_all()
    assert store.list_chats() == []
    assert store.search_memories("", ["global"]) == []
    assert store.get_snapshot("d") is None


def test_threadsafe_writes(store):
    import threading
    cid = store.create_chat(host="word")["id"]

    def worker(n):
        for i in range(20):
            store.add_message(cid, "user", f"{n}-{i}")

    threads = [threading.Thread(target=worker, args=(n,)) for n in range(5)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    assert len(store.list_messages(cid)) == 100
