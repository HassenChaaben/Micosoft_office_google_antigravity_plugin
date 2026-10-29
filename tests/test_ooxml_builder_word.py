"""Integration: insert builder OOXML into a real, HIDDEN, separate Word instance.

Uses DispatchEx so a new WINWORD process is started; the user's open documents are never touched.
`Range.InsertXML` is the COM equivalent of Office.js `range.insertOoxml`.
"""

import pytest

from daemon.office.ooxml_builder import build_ooxml

win32com = pytest.importorskip("win32com.client")
pythoncom = pytest.importorskip("pythoncom")

pytestmark = pytest.mark.word

WD_STYLE_HEADING1, WD_STYLE_HEADING2, WD_STYLE_TITLE, WD_STYLE_SUBTITLE = -2, -3, -63, -75
WD_STYLE_CAPTION, WD_STYLE_QUOTE = -35, -181
WD_FIELD_SEQUENCE, WD_FIELD_TOC = 12, 13
WD_LIST_BULLET, WD_LIST_SIMPLE_NUMBERING, WD_LIST_OUTLINE_NUMBERING = 2, 3, 4

REPORT = [
    {"type": "title", "text": "Progetto di Automazione Cloud4Job"},
    {"type": "subtitle", "text": "Rapporto di avanzamento"},
    {"type": "toc", "levels": "1-3"},
    {"type": "heading", "level": 1, "text": "Introduzione"},
    {"type": "paragraph", "align": "justify", "runs": [
        {"text": "Questo rapporto è "}, {"text": "importante", "bold": True},
        {"text": " & urgente", "italic": True, "color": "#C00000", "highlight": "yellow", "size": 12}]},
    {"type": "heading", "level": 2, "text": "Obiettivi"},
    {"type": "bullets", "items": ["Ridurre i tempi", {"text": "Automatizzare", "children": [
        "Report", {"text": "Moduli", "children": ["Charter"]}]}]},
    {"type": "numbered", "items": ["Analisi", "Sviluppo", "Collaudo"]},
    {"type": "table", "header_rows": 2, "columns": [{"width_pct": 40}, {"width_pct": 20}, {"width_pct": 20}, {"width_pct": 20}],
     "rows": [
         [{"text": "Attività", "rowspan": 2, "valign": "center"}, {"text": "Trimestre", "colspan": 3, "align": "center"}],
         ["Q1", "Q2", "Q3"],
         ["Analisi", "x", "", ""],
         ["Sviluppo", "", {"text": "x", "shading": "#E2EFDA"}, "x"],
     ]},
    {"type": "caption", "label": "Tabella", "text": "Pianificazione trimestrale"},
    {"type": "page_break"},
    {"type": "quote", "text": "La qualità non è un atto, è un'abitudine."},
]


@pytest.fixture(scope="module")
def word():
    pythoncom.CoInitialize()
    try:
        app = win32com.DispatchEx("Word.Application")
    except Exception as exc:  # Word not installed / COM not registered
        pytest.skip(f"Word not available: {exc}")
    try:
        app.Visible = False
        app.DisplayAlerts = 0
        yield app
    finally:
        try:
            app.Quit(0)
        finally:
            pythoncom.CoUninitialize()


@pytest.fixture
def doc(word):
    d = word.Documents.Add()
    try:
        yield d
    finally:
        d.Close(False)


def _paragraphs(doc):
    return [doc.Paragraphs(i) for i in range(1, doc.Paragraphs.Count + 1)]


def _style_of(doc, text):
    for p in _paragraphs(doc):
        if p.Range.Text.strip("\r\x07 ") == text:
            return p.Style.NameLocal
    raise AssertionError(f"paragraph {text!r} not found")


def test_rich_report_inserts_with_localized_builtin_styles(doc):
    doc.Content.InsertXML(build_ooxml(REPORT))
    builtin = lambda wd: doc.Styles(wd).NameLocal  # noqa: E731
    assert _style_of(doc, "Introduzione") == builtin(WD_STYLE_HEADING1)
    assert _style_of(doc, "Obiettivi") == builtin(WD_STYLE_HEADING2)
    assert _style_of(doc, "Progetto di Automazione Cloud4Job") == builtin(WD_STYLE_TITLE)
    assert _style_of(doc, "Rapporto di avanzamento") == builtin(WD_STYLE_SUBTITLE)
    assert _style_of(doc, "La qualità non è un atto, è un'abitudine.") == builtin(WD_STYLE_QUOTE)
    caption = next(p for p in _paragraphs(doc) if "Pianificazione trimestrale" in p.Range.Text)
    assert caption.Style.NameLocal == builtin(WD_STYLE_CAPTION)
    assert caption.Range.Text.strip() == "Tabella 1: Pianificazione trimestrale"


def test_table_merges_survive_insertion(doc):
    doc.Content.InsertXML(build_ooxml(REPORT))
    assert doc.Tables.Count == 1
    t = doc.Tables(1)
    assert t.Uniform is False or t.Uniform == 0
    # 4 rows x 4 grid cols = 16, minus 2 (colspan 3) minus 1 (rowspan 2 continuation) = 13
    assert t.Range.Cells.Count == 13
    # Word refuses row access when a table has vertically merged cells: proof the rowspan survived
    with pytest.raises(Exception):
        t.Rows(1)
    assert t.Cell(1, 2).Range.Text.strip("\r\x07") == "Trimestre"


def test_lists_and_fields_are_live(doc):
    doc.Content.InsertXML(build_ooxml(REPORT))
    kinds = {}
    for p in _paragraphs(doc):
        txt = p.Range.Text.strip("\r\x07 ")
        if txt in ("Ridurre i tempi", "Charter", "Analisi", "Collaudo") and p.Range.Information(12) is False:
            kinds[txt] = (p.Range.ListFormat.ListType, p.Range.ListFormat.ListLevelNumber)
    assert kinds["Ridurre i tempi"][0] == WD_LIST_BULLET or kinds["Ridurre i tempi"][0] == WD_LIST_OUTLINE_NUMBERING
    assert kinds["Charter"][1] == 3
    assert kinds["Analisi"][0] in (WD_LIST_SIMPLE_NUMBERING, WD_LIST_OUTLINE_NUMBERING)
    assert doc.Paragraphs(1).Range.ListFormat.ListType == 0
    numbered = [p.Range.ListFormat.ListString for p in _paragraphs(doc)
                if p.Range.Text.strip("\r") in ("Analisi", "Sviluppo", "Collaudo") and p.Range.Information(12) is False]
    assert numbered == ["1.", "2.", "3."]
    field_types = [doc.Fields(i).Type for i in range(1, doc.Fields.Count + 1)]
    assert WD_FIELD_SEQUENCE in field_types
    assert WD_FIELD_TOC in field_types


def test_toc_updates_to_real_entries(doc):
    doc.Content.InsertXML(build_ooxml(REPORT))
    doc.TablesOfContents(1).Update()
    toc_text = doc.TablesOfContents(1).Range.Text
    assert "Introduzione" in toc_text and "Obiettivi" in toc_text


def test_content_control_is_findable_by_tag(doc):
    doc.Content.InsertXML(build_ooxml(REPORT[:4], block_id="rep-42"))
    ccs = doc.SelectContentControlsByTag("ga:rep-42")
    assert ccs.Count == 1
    cc = ccs(1)
    assert cc.Title == "Gemini Antigravity"
    assert "Introduzione" in cc.Range.Text
    assert cc.Range.Paragraphs(1).Style.NameLocal == doc.Styles(WD_STYLE_TITLE).NameLocal
