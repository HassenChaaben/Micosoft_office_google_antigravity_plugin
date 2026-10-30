"""
Live automated validation on real PM Templates:
- 2 Project charter.docx
- 23 Registro dei rischi.docx
- 8 WBS e dizionario della WBS.docx
- 19 Reticolo di progetto.docx

Validates:
1. Perception snapshot: paragraph coordinates, table indices, cells, form fields.
2. Form field resolution & grid mapping.
3. Live Word COM manipulation: cell writing, WBS diagram generation, CPM network diagram generation.
"""

import os
import shutil
import tempfile
import zipfile
import pytest

from daemon.office.word_snapshot import build_snapshot
from daemon.office.fill import resolve_fields, fill_cell_com
from daemon.office.ooxml_builder import parse_location
from daemon.office.word_com_tools import (
    tree_layout,
    critical_path,
    layered_layout,
    word_insert_diagram
)

win32com = pytest.importorskip("win32com.client")

PM_DIR = os.path.join(
    os.path.dirname(os.path.dirname(__file__)),
    "practical_tests",
    "Template - Progetto PM-20260227",
    "Template - Progetto PM",
    "Template_Progetto_PM"
)

if not os.path.isdir(PM_DIR):
    pytestmark = pytest.mark.skip(reason="practical_tests local templates not present in repo")

@pytest.fixture(scope="module")
def word():
    try:
        app = win32com.DispatchEx("Word.Application")
    except Exception as exc:
        pytest.skip(f"Word unavailable: {exc}")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        yield app
    finally:
        try:
            for i in range(app.Documents.Count, 0, -1):
                app.Documents(i).Close(False)
        finally:
            app.Quit()


@pytest.fixture
def temp_workspace():
    tmp = tempfile.mkdtemp(prefix="pm_live_")
    yield tmp
    shutil.rmtree(tmp, ignore_errors=True)


def test_project_charter_snapshot_and_live_com(word, temp_workspace):
    src = os.path.join(PM_DIR, "2 Project charter.docx")
    assert os.path.exists(src), f"Template not found: {src}"

    # 1. Snapshot test
    with zipfile.ZipFile(src, "r") as z:
        doc_xml = z.read("word/document.xml").decode("utf-8")

    snap = build_snapshot(doc_xml)
    assert snap["doc_type"] == "form_template"
    assert snap["stats"]["fields_total"] >= 50
    assert snap["stats"]["tables"] >= 2

    # 2. Field resolution test
    requested = {
        "Codice del documento": "DOC-2026-001",
        "Anagrafica progetto > Titolo": "Automazione Microsoft Office con Antigravity",
        "Anagrafica progetto > Descrizione breve": "Senior Office Agent autonomo",
        "Cliente > Azienda": "Cloud4Job",
    }
    resolved = resolve_fields(snap, requested)
    assert len(resolved["cells"]) == 4
    assert len(resolved["errors"]) == 0

    # 3. Live Word COM manipulation on copy
    dest = os.path.join(temp_workspace, "2 Project charter.docx")
    shutil.copy2(src, dest)
    doc = word.Documents.Open(dest)
    try:
        for c in resolved["cells"]:
            ok = fill_cell_com(doc, c["table"], c["row"], c["cell"], c["text"])
            assert ok is True
        doc.Save()
    finally:
        doc.Close(False)


def test_registro_rischi_grid_and_live_com(word, temp_workspace):
    src = os.path.join(PM_DIR, "23 Registro dei rischi.docx")
    assert os.path.exists(src), f"Template not found: {src}"

    with zipfile.ZipFile(src, "r") as z:
        doc_xml = z.read("word/document.xml").decode("utf-8")

    snap = build_snapshot(doc_xml)
    assert snap["doc_type"] == "form_template"
    assert snap["stats"]["tables"] >= 10
    assert snap["stats"]["fields_total"] > 100

    # Live Word COM verify open and cell write
    dest = os.path.join(temp_workspace, "23 Registro dei rischi.docx")
    shutil.copy2(src, dest)
    doc = word.Documents.Open(dest)
    try:
        assert doc.Tables.Count >= 10
        # Write risk item into first data row of table 0
        ok = fill_cell_com(doc, 0, 1, 0, "R-01")
        assert ok is True
        doc.Save()
    finally:
        doc.Close(False)


def test_wbs_template_diagram_and_anchors(word, temp_workspace):
    src = os.path.join(PM_DIR, "8 WBS e dizionario della WBS.docx")
    assert os.path.exists(src), f"Template not found: {src}"

    with zipfile.ZipFile(src, "r") as z:
        doc_xml = z.read("word/document.xml").decode("utf-8")

    snap = build_snapshot(doc_xml)
    assert snap["stats"]["tables"] >= 1

    # Location parsing
    mode, anchor = parse_location("before:p1")
    assert mode == "before"
    assert anchor == "p1"

    # Pure Python WBS tree layout algorithm
    nodes = [
        {"id": "1", "label": "Progetto Antigravity", "parent": None},
        {"id": "1.1", "label": "Daemon Core", "parent": "1"},
        {"id": "1.2", "label": "Taskpane UI", "parent": "1"},
        {"id": "1.3", "label": "COM Automation", "parent": "1"},
    ]
    edges = [("1", "1.1"), ("1", "1.2"), ("1", "1.3")]
    pos, leaves = tree_layout(["1", "1.1", "1.2", "1.3"], edges)
    assert leaves == 3
    assert pos["1"][0] == 0

    # Live Word COM insertion of WBS diagram into document
    dest = os.path.join(temp_workspace, "8 WBS e dizionario della WBS.docx")
    shutil.copy2(src, dest)
    doc = word.Documents.Open(dest)
    try:
        wbs_edges = [{"from": "1", "to": "1.1"}, {"from": "1", "to": "1.2"}, {"from": "1", "to": "1.3"}]
        res = word_insert_diagram(doc, kind="wbs", nodes=nodes, edges=wbs_edges, location="end")
        assert res["ok"] is True
        assert res["kind"] == "wbs"
        assert res["shapes"] >= 4
        assert res["connectors"] == 3
        doc.Save()
    finally:
        doc.Close(False)


def test_reticolo_progetto_cpm_gantt(word, temp_workspace):
    src = os.path.join(PM_DIR, "19 Reticolo di progetto.docx")
    assert os.path.exists(src), f"Template not found: {src}"

    with zipfile.ZipFile(src, "r") as z:
        doc_xml = z.read("word/document.xml").decode("utf-8")

    snap = build_snapshot(doc_xml)
    assert snap["stats"]["tables"] >= 1

    # Pure Python CPM algorithm
    nodes = [
        {"id": "A", "label": "Attivita A", "duration": 5, "critical": True},
        {"id": "B", "label": "Attivita B", "duration": 3, "critical": False},
        {"id": "C", "label": "Attivita C", "duration": 7, "critical": True},
    ]
    edges = [
        {"from": "A", "to": "B", "type": "FS"},
        {"from": "A", "to": "C", "type": "FS"},
    ]
    durations = {"A": 5, "B": 3, "C": 7}
    es, ef, slack, crit = critical_path(["A", "B", "C"], [("A", "B"), ("A", "C")], durations)
    assert "A" in crit
    assert "C" in crit
    assert slack["B"] > 0

    # Live Word COM insertion of Network diagram into document
    dest = os.path.join(temp_workspace, "19 Reticolo di progetto.docx")
    shutil.copy2(src, dest)
    doc = word.Documents.Open(dest)
    try:
        res = word_insert_diagram(doc, kind="network", nodes=nodes, edges=edges, location="end")
        assert res["ok"] is True
        assert res["kind"] == "network"
        assert res["shapes"] >= 3
        assert res["connectors"] == 2
        assert "A" in res["critical"]
        assert "C" in res["critical"]
        doc.Save()
    finally:
        doc.Close(False)
