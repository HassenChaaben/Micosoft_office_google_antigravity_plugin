"""Local SQLite history: chats, messages, tool calls, plans, document snapshots and memories.

Thread-safe (one connection, check_same_thread=False, guarded by a threading.Lock), WAL mode,
schema versioned with PRAGMA user_version.
"""

from __future__ import annotations

import json
import os
import re
import sqlite3
import threading
import uuid
from datetime import datetime, timezone
from typing import Any, Dict, Iterable, List, Optional

SCHEMA_VERSION = 1
DEFAULT_TITLE = "New chat"
TITLE_CHARS = 60
ROLES = ("user", "assistant", "system")

_MIGRATIONS = {
    1: """
    CREATE TABLE IF NOT EXISTS chats (
        id TEXT PRIMARY KEY,
        host TEXT,
        doc_url TEXT,
        doc_title TEXT,
        title TEXT NOT NULL DEFAULT 'New chat',
        title_locked INTEGER NOT NULL DEFAULT 0,
        model TEXT,
        effort TEXT,
        agy_conversation_id TEXT,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_chats_updated ON chats(updated_at DESC);
    CREATE INDEX IF NOT EXISTS idx_chats_doc ON chats(doc_url);
    CREATE TABLE IF NOT EXISTS messages (
        id TEXT PRIMARY KEY,
        chat_id TEXT NOT NULL REFERENCES chats(id) ON DELETE CASCADE,
        role TEXT NOT NULL CHECK (role IN ('user','assistant','system')),
        content TEXT NOT NULL,
        created_at TEXT NOT NULL,
        seq INTEGER NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_messages_chat ON messages(chat_id, seq);
    CREATE TABLE IF NOT EXISTS tool_calls (
        id TEXT PRIMARY KEY,
        chat_id TEXT NOT NULL REFERENCES chats(id) ON DELETE CASCADE,
        message_id TEXT REFERENCES messages(id) ON DELETE SET NULL,
        tool TEXT NOT NULL,
        args_json TEXT,
        result_json TEXT,
        ok INTEGER,
        duration_ms INTEGER,
        created_at TEXT NOT NULL,
        seq INTEGER NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_tool_calls_chat ON tool_calls(chat_id, seq);
    CREATE TABLE IF NOT EXISTS plans (
        id TEXT PRIMARY KEY,
        chat_id TEXT NOT NULL REFERENCES chats(id) ON DELETE CASCADE,
        message_id TEXT REFERENCES messages(id) ON DELETE SET NULL,
        steps_json TEXT NOT NULL,
        created_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS snapshots (
        doc_url TEXT PRIMARY KEY,
        change_key TEXT,
        snapshot_json TEXT,
        map_text TEXT,
        updated_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS memories (
        id TEXT PRIMARY KEY,
        scope TEXT NOT NULL,
        text TEXT NOT NULL,
        created_at TEXT NOT NULL,
        seq INTEGER NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_memories_scope ON memories(scope, seq);
    """,
}


def default_db_path() -> str:
    base = os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Local")
    return os.path.join(base, "GeminiAntigravityOffice", "history.db")


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="microseconds")


def _auto_title(text: str) -> str:
    one_line = " ".join((text or "").split())
    if len(one_line) <= TITLE_CHARS:
        return one_line or DEFAULT_TITLE
    cut = one_line[:TITLE_CHARS].rstrip()
    return cut + "…"


class Store:
    def __init__(self, path: Optional[str] = None):
        self.path = path or default_db_path()
        if self.path != ":memory:":
            os.makedirs(os.path.dirname(os.path.abspath(self.path)), exist_ok=True)
        self._lock = threading.Lock()
        self._conn = sqlite3.connect(self.path, check_same_thread=False, isolation_level=None)
        self._conn.row_factory = sqlite3.Row
        with self._lock:
            self._conn.execute("PRAGMA journal_mode=WAL")
            self._conn.execute("PRAGMA foreign_keys=ON")
            self._conn.execute("PRAGMA busy_timeout=5000")
            self._migrate()

    # ------------------------------------------------------------ infra

    def _migrate(self) -> None:
        version = self._conn.execute("PRAGMA user_version").fetchone()[0]
        for target in sorted(_MIGRATIONS):
            if target > version:
                self._conn.execute("BEGIN")
                try:
                    for stmt in _MIGRATIONS[target].split(";"):
                        if stmt.strip():
                            self._conn.execute(stmt)
                    self._conn.execute(f"PRAGMA user_version={int(target)}")
                    self._conn.execute("COMMIT")
                except Exception:
                    self._conn.execute("ROLLBACK")
                    raise

    def close(self) -> None:
        with self._lock:
            self._conn.close()

    def _next_seq(self, table: str) -> int:
        return self._conn.execute(f"SELECT COALESCE(MAX(seq), 0) + 1 FROM {table}").fetchone()[0]

    @staticmethod
    def _row(row: Optional[sqlite3.Row]) -> Optional[Dict[str, Any]]:
        return dict(row) if row is not None else None

    # ------------------------------------------------------------ chats

    def create_chat(self, host: Optional[str] = None, doc_url: Optional[str] = None,
                    doc_title: Optional[str] = None, title: Optional[str] = None, model: Optional[str] = None,
                    effort: Optional[str] = None, agy_conversation_id: Optional[str] = None) -> Dict[str, Any]:
        cid = str(uuid.uuid4())
        now = _now()
        with self._lock:
            self._conn.execute(
                "INSERT INTO chats (id, host, doc_url, doc_title, title, title_locked, model, effort,"
                " agy_conversation_id, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                (cid, host, doc_url, doc_title, title or DEFAULT_TITLE, 1 if title else 0, model, effort,
                 agy_conversation_id, now, now))
        return self.get_chat(cid)

    def get_chat(self, chat_id: str) -> Optional[Dict[str, Any]]:
        with self._lock:
            row = self._conn.execute("SELECT * FROM chats WHERE id=?", (chat_id,)).fetchone()
        chat = self._row(row)
        if chat:
            chat.pop("title_locked", None)
        return chat

    def list_chats(self, host: Optional[str] = None, doc_url: Optional[str] = None,
                   limit: int = 50) -> List[Dict[str, Any]]:
        where, params = [], []
        if host is not None:
            where.append("c.host=?")
            params.append(host)
        if doc_url is not None:
            where.append("c.doc_url=?")
            params.append(doc_url)
        sql = ("SELECT c.*, (SELECT m.content FROM messages m WHERE m.chat_id=c.id ORDER BY m.seq DESC LIMIT 1)"
               " AS last_message FROM chats c")
        if where:
            sql += " WHERE " + " AND ".join(where)
        sql += " ORDER BY c.updated_at DESC, c.rowid DESC LIMIT ?"
        params.append(int(limit))
        with self._lock:
            rows = self._conn.execute(sql, params).fetchall()
        out = []
        for r in rows:
            d = dict(r)
            d.pop("title_locked", None)
            if d["last_message"] is not None and len(d["last_message"]) > 200:
                d["last_message"] = d["last_message"][:200] + "…"
            out.append(d)
        return out

    def rename_chat(self, chat_id: str, title: str) -> None:
        with self._lock:
            self._conn.execute("UPDATE chats SET title=?, title_locked=1, updated_at=? WHERE id=?",
                               (title, _now(), chat_id))

    def delete_chat(self, chat_id: str) -> None:
        with self._lock:
            self._conn.execute("DELETE FROM chats WHERE id=?", (chat_id,))

    def set_conversation_id(self, chat_id: str, conversation_id: Optional[str]) -> None:
        with self._lock:
            self._conn.execute("UPDATE chats SET agy_conversation_id=?, updated_at=? WHERE id=?",
                               (conversation_id, _now(), chat_id))

    def update_chat(self, chat_id: str, **fields: Any) -> None:
        """Update model / effort / doc_url / doc_title / host."""
        allowed = {"model", "effort", "doc_url", "doc_title", "host"}
        bad = set(fields) - allowed
        if bad:
            raise ValueError(f"Unknown chat fields: {sorted(bad)}")
        if not fields:
            return
        cols = ", ".join(f"{k}=?" for k in fields)
        with self._lock:
            self._conn.execute(f"UPDATE chats SET {cols}, updated_at=? WHERE id=?",
                               (*fields.values(), _now(), chat_id))

    # ------------------------------------------------------------ messages

    def add_message(self, chat_id: str, role: str, content: str) -> Dict[str, Any]:
        if role not in ROLES:
            raise ValueError(f"role must be one of {ROLES}")
        mid = str(uuid.uuid4())
        now = _now()
        with self._lock:
            self._conn.execute("BEGIN IMMEDIATE")
            try:
                seq = self._next_seq("messages")
                self._conn.execute("INSERT INTO messages (id, chat_id, role, content, created_at, seq)"
                                   " VALUES (?,?,?,?,?,?)", (mid, chat_id, role, content, now, seq))
                if role == "user":
                    self._conn.execute(
                        "UPDATE chats SET title=? WHERE id=? AND title_locked=0 AND title=? AND NOT EXISTS"
                        " (SELECT 1 FROM messages WHERE chat_id=? AND role='user' AND id<>?)",
                        (_auto_title(content), chat_id, DEFAULT_TITLE, chat_id, mid))
                self._conn.execute("UPDATE chats SET updated_at=? WHERE id=?", (now, chat_id))
                self._conn.execute("COMMIT")
            except Exception:
                self._conn.execute("ROLLBACK")
                raise
        return {"id": mid, "chat_id": chat_id, "role": role, "content": content, "created_at": now}

    def list_messages(self, chat_id: str) -> List[Dict[str, Any]]:
        with self._lock:
            rows = self._conn.execute("SELECT id, chat_id, role, content, created_at FROM messages"
                                      " WHERE chat_id=? ORDER BY seq", (chat_id,)).fetchall()
        return [dict(r) for r in rows]

    # ------------------------------------------------------------ tool calls & plans

    def add_tool_call(self, chat_id: str, message_id: Optional[str], tool: str, args: Any = None,
                      result: Any = None, ok: Optional[bool] = None,
                      duration_ms: Optional[int] = None) -> Dict[str, Any]:
        tid = str(uuid.uuid4())
        now = _now()
        args_json = json.dumps(args, ensure_ascii=False, default=str)
        result_json = json.dumps(result, ensure_ascii=False, default=str)
        with self._lock:
            self._conn.execute("BEGIN IMMEDIATE")
            try:
                seq = self._next_seq("tool_calls")
                self._conn.execute(
                    "INSERT INTO tool_calls (id, chat_id, message_id, tool, args_json, result_json, ok,"
                    " duration_ms, created_at, seq) VALUES (?,?,?,?,?,?,?,?,?,?)",
                    (tid, chat_id, message_id, tool, args_json, result_json,
                     None if ok is None else int(bool(ok)), duration_ms, now, seq))
                self._conn.execute("COMMIT")
            except Exception:
                self._conn.execute("ROLLBACK")
                raise
        return {"id": tid, "chat_id": chat_id, "message_id": message_id, "tool": tool, "args_json": args_json,
                "result_json": result_json, "ok": ok, "duration_ms": duration_ms, "created_at": now}

    def list_tool_calls(self, chat_id: str) -> List[Dict[str, Any]]:
        with self._lock:
            rows = self._conn.execute(
                "SELECT id, chat_id, message_id, tool, args_json, result_json, ok, duration_ms, created_at"
                " FROM tool_calls WHERE chat_id=? ORDER BY seq", (chat_id,)).fetchall()
        out = []
        for r in rows:
            d = dict(r)
            d["ok"] = None if d["ok"] is None else bool(d["ok"])
            out.append(d)
        return out

    def add_plan(self, chat_id: str, message_id: Optional[str], steps: Any) -> Dict[str, Any]:
        pid = str(uuid.uuid4())
        now = _now()
        steps_json = steps if isinstance(steps, str) else json.dumps(steps, ensure_ascii=False, default=str)
        with self._lock:
            self._conn.execute("INSERT INTO plans (id, chat_id, message_id, steps_json, created_at)"
                               " VALUES (?,?,?,?,?)", (pid, chat_id, message_id, steps_json, now))
        return {"id": pid, "chat_id": chat_id, "message_id": message_id, "steps_json": steps_json, "created_at": now}

    # ------------------------------------------------------------ snapshots

    def save_snapshot(self, doc_url: str, change_key: Optional[str], snapshot: Any, map_text: str = "") -> None:
        snap_json = snapshot if isinstance(snapshot, str) else json.dumps(snapshot, ensure_ascii=False, default=str)
        with self._lock:
            self._conn.execute(
                "INSERT INTO snapshots (doc_url, change_key, snapshot_json, map_text, updated_at) VALUES (?,?,?,?,?)"
                " ON CONFLICT(doc_url) DO UPDATE SET change_key=excluded.change_key,"
                " snapshot_json=excluded.snapshot_json, map_text=excluded.map_text, updated_at=excluded.updated_at",
                (doc_url, change_key, snap_json, map_text, _now()))

    def get_snapshot(self, doc_url: str) -> Optional[Dict[str, Any]]:
        with self._lock:
            row = self._conn.execute("SELECT * FROM snapshots WHERE doc_url=?", (doc_url,)).fetchone()
        if row is None:
            return None
        d = dict(row)
        try:
            d["snapshot"] = json.loads(d["snapshot_json"]) if d["snapshot_json"] else None
        except json.JSONDecodeError:
            d["snapshot"] = None
        return d

    # ------------------------------------------------------------ memories

    @staticmethod
    def _check_scope(scope: str) -> None:
        if scope != "global" and not (scope.startswith("doc:") and len(scope) > 4):
            raise ValueError("scope must be 'global' or 'doc:<url>'")

    def add_memory(self, scope: str, text: str) -> Dict[str, Any]:
        self._check_scope(scope)
        mid = str(uuid.uuid4())
        now = _now()
        with self._lock:
            self._conn.execute("BEGIN IMMEDIATE")
            try:
                seq = self._next_seq("memories")
                self._conn.execute("INSERT INTO memories (id, scope, text, created_at, seq) VALUES (?,?,?,?,?)",
                                   (mid, scope, text, now, seq))
                self._conn.execute("COMMIT")
            except Exception:
                self._conn.execute("ROLLBACK")
                raise
        return {"id": mid, "scope": scope, "text": text, "created_at": now}

    def search_memories(self, query: str, scopes: Iterable[str], limit: int = 20) -> List[Dict[str, Any]]:
        """Memories in `scopes` whose text contains ANY query word (case-insensitive), newest first.
        An empty query returns the newest memories in scope."""
        scopes = list(scopes)
        if not scopes:
            return []
        words = [w for w in re.findall(r"\w+", (query or "").lower()) if len(w) > 1]
        sql = f"SELECT id, scope, text, created_at FROM memories WHERE scope IN ({','.join('?' * len(scopes))})"
        params: List[Any] = list(scopes)
        if words:
            sql += " AND (" + " OR ".join("lower(text) LIKE ? ESCAPE '\\'" for _ in words) + ")"
            for w in words:
                params.append("%" + w.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_") + "%")
        sql += " ORDER BY seq DESC LIMIT ?"
        params.append(int(limit))
        with self._lock:
            rows = self._conn.execute(sql, params).fetchall()
        return [dict(r) for r in rows]

    def delete_memory(self, memory_id: str) -> None:
        with self._lock:
            self._conn.execute("DELETE FROM memories WHERE id=?", (memory_id,))

    # ------------------------------------------------------------ wipe

    def delete_all(self) -> None:
        with self._lock:
            self._conn.execute("BEGIN IMMEDIATE")
            try:
                for table in ("tool_calls", "plans", "messages", "chats", "snapshots", "memories"):
                    self._conn.execute(f"DELETE FROM {table}")
                self._conn.execute("COMMIT")
            except Exception:
                self._conn.execute("ROLLBACK")
                raise
