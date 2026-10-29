import asyncio
from pathlib import Path

from daemon.office.perception import Perception

FIX = Path(__file__).parent / "fixtures" / "docs"


def _payload(name="project_charter", **extra):
    p = {
        "doc_url": "C:\\docs\\charter.docx",
        "doc_title": "2 Project charter.docx",
        "document_ooxml": (FIX / f"{name}.body.xml").read_text(encoding="utf-8"),
        "selection": "",
        "headers_footers": [],
    }
    p.update(extra)
    return p


def test_analyze_returns_map_and_human_summary():
    res = asyncio.run(Perception().analyze(_payload()))
    assert "DOCUMENT MAP" in res.map_text
    assert "Cliente › Azienda" in res.map_text
    assert res.summary["doc_type"] == "form_template"
    assert res.summary["fields_empty"] == 64
    assert "form" in res.summary["headline"].lower()


def test_selection_and_headers_reach_the_map():
    res = asyncio.run(Perception().analyze(_payload(selection="Budget complessivo", headers_footers=["header1: Cloud4Job"])))
    assert "SELECTION: 'Budget complessivo'" in res.map_text
    assert "Cloud4Job" in res.map_text


def test_snapshot_is_cached_for_identical_document():
    p = Perception()
    a = asyncio.run(p.analyze(_payload()))
    b = asyncio.run(p.analyze(_payload()))
    assert a.cached is False and b.cached is True
    assert a.map_text == b.map_text


def test_missing_document_gives_honest_notice_instead_of_failing():
    res = asyncio.run(Perception().analyze({"doc_url": "", "document_ooxml": ""}))
    assert res.snapshot is None
    assert "could not be read" in res.map_text.lower()


def test_map_budget_is_respected():
    res = asyncio.run(Perception(map_budget=5000).analyze(_payload("risk_register")))
    assert len(res.map_text) <= 5000
