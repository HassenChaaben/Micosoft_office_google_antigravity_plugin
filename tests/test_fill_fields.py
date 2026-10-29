import json
from pathlib import Path

from fastapi.testclient import TestClient

from daemon.agent_session import agent_session
from daemon.main import app
from daemon.office.fill import resolve_fields
from daemon.office.word_snapshot import build_snapshot

FIX = Path(__file__).parent / "fixtures" / "docs"
CHARTER = build_snapshot((FIX / "project_charter.body.xml").read_text(encoding="utf-8"))
RISK = build_snapshot((FIX / "risk_register.body.xml").read_text(encoding="utf-8"))


def test_paths_resolve_to_cells():
    res = resolve_fields(CHARTER, {"Cliente › Azienda": "Cloud4Job S.r.l.", "Sponsor › Cognome Nome": "Rossi Mario"})
    assert res["errors"] == []
    by_text = {c["text"]: c for c in res["cells"]}
    az = by_text["Cloud4Job S.r.l."]
    assert (az["table"], az["row"], az["cell"]) == (0, 6, 1)


def test_ascii_separator_case_and_accents_are_tolerated():
    res = resolve_fields(RISK, {"minacce > #1 > causa": "Ritardo fornitore", "Opportunita > #1 > Cod.": "O01"})
    assert res["errors"] == []
    assert len(res["cells"]) == 2


def test_grid_cells_by_record_number_and_coordinates():
    res = resolve_fields(RISK, {"Minacce › #2 › Evento": "Blocco", "T1r2c0": "R01"})
    assert res["errors"] == []
    rows = sorted((c["table"], c["row"], c["cell"]) for c in res["cells"])
    assert rows == [(1, 2, 0), (1, 3, 4)]


def test_unknown_key_gets_suggestions_and_missing_record_hints_clone():
    res = resolve_fields(CHARTER, {"Cliente › Aziende": "x"})
    assert res["cells"] == []
    assert "Cliente › Azienda" in res["errors"][0]
    res = resolve_fields(RISK, {"Minacce › #9 › Causa": "x"})
    assert "word_table_clone_row" in res["errors"][0]


def test_content_control_tags_pass_through():
    snap = dict(CHARTER, content_controls=[{"tag": "cliente", "title": "Cliente", "type": "text", "text": ""}])
    res = resolve_fields(snap, {"cliente": "ACME"})
    assert res["content_controls"] == {"cliente": "ACME"}


def test_mcp_word_fill_fields_sends_one_batched_cell_write(monkeypatch):
    sent = {}

    async def fake_call_tool(name, args, timeout_sec=60.0):
        sent["name"], sent["args"] = name, args
        return {"ok": True, "written": len(args["cells"])}

    monkeypatch.setattr(agent_session, "call_tool", fake_call_tool)
    agent_session.turn_snapshot = CHARTER
    r = TestClient(app).post("/mcp", json={"jsonrpc": "2.0", "id": 3, "method": "tools/call", "params": {
        "name": "word_fill_fields", "arguments": {"fields": {"Cliente › Azienda": "Cloud4Job", "Nope": "x"}}}}).json()
    payload = json.loads(r["result"]["content"][0]["text"])
    assert sent["name"] == "word_fill_cells" and len(sent["args"]["cells"]) == 1
    assert payload["written"] == 1 and payload["not_found"]
