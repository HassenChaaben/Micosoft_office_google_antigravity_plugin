import os
import struct
import tempfile
import zlib
from unittest.mock import MagicMock
import pytest

from daemon.office import word_com_tools as wt
from daemon.office.fill import fill_cell_com, fill_cells_com_fallback, resolve_fields
from daemon.office.ooxml_builder import LOCATION_RE, build_ooxml, parse_location
from daemon.office.word_snapshot import build_snapshot

win32com = pytest.importorskip("win32com.client")
pythoncom = pytest.importorskip("pythoncom")


# ==============================================================================
# 1. OOXML BUILDER & LOCATION PARSING TESTS (Unit)
# ==============================================================================

def test_parse_location_before_after_replace():
    assert parse_location("before:p0") == ("before", "p0")
    assert parse_location("after:T0") == ("after", "T0")
    assert parse_location("replace:heading1") == ("replace", "heading1")
    assert parse_location("before:T0r1c2") == ("before", "T0r1c2")
    assert parse_location("start") == ("before", "start")
    assert parse_location("end") == ("after", "end")
    assert parse_location(None) == ("after", "end")
    assert parse_location("p5") == ("after", "p5")


def test_build_ooxml_with_before_location():
    blocks = [{"type": "paragraph", "text": "Test paragraph"}]
    xml = build_ooxml(blocks, block_id="b1", location="before:p0")
    assert "Test paragraph" in xml
    assert 'w:tag w:val="ga:b1"' in xml
    mode, anchor = parse_location("before:p0")
    assert mode == "before"
    assert anchor == "p0"


# ==============================================================================
# 2. WORD SNAPSHOT NESTED TABLES TESTS (Unit XML)
# ==============================================================================

def test_snapshot_indexes_nested_tables():
    # Package XML with an outer 2x2 table and a nested 1x2 table inside cell (0, 0)
    xml = (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">\n'
        '  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">\n'
        '    <pkg:xmlData>\n'
        '      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
        '        <w:body>\n'
        '          <w:p><w:r><w:t>Documento Principale</w:t></w:r></w:p>\n'
        '          <w:tbl>\n'
        '            <w:tr>\n'
        '              <w:tc>\n'
        '                <w:p><w:r><w:t>Cella Esterna 0,0</w:t></w:r></w:p>\n'
        '                <w:tbl>\n'
        '                  <w:tr>\n'
        '                    <w:tc><w:p><w:r><w:t>Etichetta Annidata</w:t></w:r></w:p></w:tc>\n'
        '                    <w:tc><w:p><w:r><w:t>[Valore Annidato]</w:t></w:r></w:p></w:tc>\n'
        '                  </w:tr>\n'
        '                </w:tbl>\n'
        '              </w:tc>\n'
        '              <w:tc><w:p><w:r><w:t>Cella Esterna 0,1</w:t></w:r></w:p></w:tc>\n'
        '            </w:tr>\n'
        '            <w:tr>\n'
        '              <w:tc><w:p><w:r><w:t>Campo Esterno</w:t></w:r></w:p></w:tc>\n'
        '              <w:tc><w:p><w:r><w:t>[Valore Esterno]</w:t></w:r></w:p></w:tc>\n'
        '            </w:tr>\n'
        '          </w:tbl>\n'
        '        </w:body>\n'
        '      </w:document>\n'
        '    </pkg:xmlData>\n'
        '  </pkg:part>\n'
        '</pkg:package>'
    )
    snap = build_snapshot(xml)
    assert len(snap["tables"]) == 2

    # Top-level table T0
    t0 = snap["tables"][0]
    assert t0.get("nested") is not True

    # Nested table T1
    t1 = snap["tables"][1]
    assert t1.get("nested") is True
    assert t1.get("parent_table") == 0
    assert t1.get("parent_row") == 0
    assert t1.get("parent_cell") == 0

    # Fields in nested table reference table 1
    nested_fields = [f for f in snap["fields"] if f.get("loc", {}).get("table") == 1]
    assert len(nested_fields) >= 1
    assert "Annidat" in nested_fields[0]["path"]


# ==============================================================================
# 3. FILL COM FALLBACK ON MERGED CELLS (Unit)
# ==============================================================================

def test_resolve_fields_includes_paragraph_index():
    fake_snap = {
        "fields": [
            {
                "path": "Sezione › Chiave",
                "kind": "cell",
                "loc": {"table": 0, "row": 1, "col": 1, "cell": 1, "paragraph": 5},
                "placeholder": "[Inserire testo]",
            }
        ],
        "tables": [{"rows": 2, "cols": 2}],
    }
    resolved = resolve_fields(fake_snap, {"Sezione › Chiave": "Valore Test"})
    assert len(resolved["cells"]) == 1
    cell = resolved["cells"][0]
    assert cell["paragraph"] == 5
    assert cell["table"] == 0
    assert cell["row"] == 1
    assert cell["cell"] == 1
    assert cell["text"] == "Valore Test"


def test_fill_cell_com_fallback_to_paragraph():
    # Simulate a Word doc where Cell(r, c) fails (e.g. vertically merged cell)
    mock_doc = MagicMock()
    mock_table = MagicMock()
    mock_doc.Tables.Count = 1
    mock_doc.Tables.return_value = mock_table

    # Direct cell access raises exception
    mock_table.Cell.side_effect = Exception("Cannot access individual rows because table has vertically merged cells")
    mock_table.Range.Cells.Count = 0

    # Paragraph access succeeds
    mock_doc.Paragraphs.Count = 10
    mock_para = MagicMock()
    mock_doc.Paragraphs.return_value = mock_para
    mock_rng = MagicMock()
    mock_rng.Text = "vecchio testo\r\x07"
    mock_para.Range = mock_rng

    success = fill_cell_com(mock_doc, table_index=0, row_index=1, col_index=1, text="Nuovo Testo", paragraph_index=4)
    assert success is True
    assert mock_rng.Text == "Nuovo Testo"


def test_fill_cells_com_fallback_batch():
    mock_doc = MagicMock()
    mock_doc.Tables.Count = 1
    mock_table = MagicMock()
    mock_doc.Tables.return_value = mock_table
    mock_table.Cell.side_effect = Exception("Merged")
    mock_table.Range.Cells.Count = 0

    mock_doc.Paragraphs.Count = 10
    mock_para = MagicMock()
    mock_doc.Paragraphs.return_value = mock_para
    mock_rng = MagicMock()
    mock_rng.Text = ""
    mock_para.Range = mock_rng

    failed_cells = [
        {"table": 0, "row": 1, "cell": 1, "text": "Aggiornato", "paragraph": 2, "key": "T0r1c1"},
    ]
    res = fill_cells_com_fallback(mock_doc, failed_cells)
    assert res["ok"] is True
    assert res["written"] == 1
    assert res["failed"] == []


# ==============================================================================
# 4. WORD COM NATIVE INTEGRATION FIXTURES & TESTS
# ==============================================================================

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
def doc(word):
    d = word.Documents.Add()
    try:
        d.Content.InsertAfter("Introduzione\r")
        d.Content.InsertAfter("Testo introduttivo al capitolo.\r")
        d.Content.InsertAfter("Architettura\r")
        d.Content.InsertAfter("Dettagli tecnici del sistema.\r")
        d.Paragraphs(1).Style = -2  # wdStyleHeading1
        d.Paragraphs(3).Style = -2  # wdStyleHeading1
        yield d
    finally:
        d.Close(False)


def test_word_com_before_anchor_resolution(doc):
    # Insert before paragraph p0 using insertion_range
    rng0 = wt.insertion_range(doc, "before:p0")
    rng0.Text = "Titolo Iniziale\r"
    assert doc.Paragraphs(1).Range.Text.strip() == "Titolo Iniziale"

    # Add a table at the very beginning of the document
    tbl = doc.Tables.Add(doc.Range(0, 0), 2, 2)
    assert doc.Tables.Count >= 1

    # Insert before T0 (table starting at 0, 0)
    rng_tbl = wt.insertion_range(doc, "before:T0")
    rng_tbl.Text = "Testo prima di T0\r"
    assert doc.Paragraphs(1).Range.Text.strip() == "Testo prima di T0"


def test_word_cross_reference_kinds(doc):
    # Add caption to figure
    cap = wt.word_insert_caption(doc, anchor="p1", label="Figure", text="Diagramma Architettura")
    assert cap["ok"] is True

    # 1. Cross reference kind: label_and_number
    res_lbl = wt.word_insert_cross_reference(doc, location="end", target_anchor="Figure 1",
                                             kind="label_and_number", show="label_and_number")
    assert res_lbl["ok"] is True
    assert "field" in res_lbl
    assert res_lbl["result"]

    # Link headings to numbering
    wt.word_link_numbering_to_headings(doc, scheme="1.1")

    # 2. Cross reference kind: heading_number
    res_head = wt.word_insert_cross_reference(doc, location="end", target_anchor="Introduzione",
                                              kind="heading_number")
    assert res_head["ok"] is True
    assert "field" in res_head

    # 3. Cross reference kind: page
    res_page = wt.word_insert_cross_reference(doc, location="end", target_anchor="Introduzione",
                                              kind="page")
    assert res_page["ok"] is True
    assert "field" in res_page


def test_word_diagram_critical_dates_milestone_and_edge_types(doc):
    nodes = [
        {"id": "A", "text": "Fase Iniziale", "start": "01/05/2026", "finish": "05/05/2026", "critical": True},
        {"id": "B", "text": "Implementazione", "start": "2026-05-06", "finish": "2026-05-15", "critical": False},
        {"id": "M1", "text": "Milestone Go-Live", "start": "2026-05-15", "finish": "2026-05-15",
         "milestone": True, "critical": True},
    ]
    edges = [
        {"from": "A", "to": "B", "type": "FS", "critical": True},
        {"from": "B", "to": "M1", "type": "SS", "critical": False},
    ]

    res = wt.word_insert_diagram(doc, kind="gantt", nodes=nodes, edges=edges, location="end")
    assert res["ok"] is True
    assert res["shapes"] > 0
    assert "critical" in res
    assert "A" in res["critical"] and "M1" in res["critical"]

    # Network diagram with FF and SF edge types
    net_nodes = [
        {"id": "N1", "text": "Task 1", "duration": 3, "critical": True},
        {"id": "N2", "text": "Task 2", "duration": 4},
        {"id": "N3", "text": "Task 3", "duration": 2},
    ]
    net_edges = [
        {"from": "N1", "to": "N2", "type": "FF", "critical": True},
        {"from": "N2", "to": "N3", "type": "SF"},
    ]
    net_res = wt.word_insert_diagram(doc, kind="network", nodes=net_nodes, edges=net_edges, location="end")
    assert net_res["ok"] is True
    assert net_res["shapes"] > 0


def test_word_table_of_figures_and_update(doc):
    # Insert caption
    wt.word_insert_caption(doc, anchor="p1", label="Figure", text="Grafico di prova")

    # Insert Table of Figures
    res_tof = wt.word_insert_table_of_figures(doc, location="end", label="Figure")
    assert res_tof["ok"] is True
    assert res_tof["tables_of_figures"] >= 1

    # Update fields
    upd = wt.word_update_fields(doc)
    assert upd["ok"] is True
    assert upd["tables_of_figures"] >= 1


def _create_test_png(path: str):
    width, height = 16, 16
    raw = b"\x00" + b"\xff\x00\x00\xff" * width
    raw = raw * height
    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xffffffff)
    png = b"\x89PNG\r\n\x1a\n"
    png += chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0))
    png += chunk(b"IDAT", zlib.compress(raw))
    png += chunk(b"IEND", b"")
    with open(path, "wb") as f:
        f.write(png)


def test_word_insert_image_with_alt_text_and_dimensions(doc, tmp_path):
    img_path = str(tmp_path / "sample.png")
    _create_test_png(img_path)

    res = wt.word_insert_image(
        doc,
        file_path=img_path,
        location="end",
        alt_text="Descrizione alternativa immagine",
        width_pt=100.0,
        height_pt=75.0,
        caption="Figura di test",
    )
    assert res["ok"] is True
    assert res["alt_text"] == "Descrizione alternativa immagine"
    assert abs(res["width"] - 100.0) < 1.0
    assert abs(res["height"] - 75.0) < 1.0
    assert res["caption"] is not None
    assert res["caption"]["ok"] is True


def test_word_set_header_footer_page_number_styles_and_restart(doc):
    res = wt.word_set_header_footer(
        doc,
        section=1,
        kind="footer",
        text="Pagina",
        page_numbers=True,
        page_number_format="roman_lower",
        page_number_restart=True,
        page_number_start=1,
    )
    assert res["ok"] is True
    pns = doc.Sections(1).Footers(1).PageNumbers
    assert pns.RestartNumberingAtSection is True
    assert pns.StartingNumber == 1
    assert pns.NumberStyle == 2  # wdPageNumberStyleLowercaseRoman


def test_word_set_checkbox_adjacent_label_and_unicode_symbol(doc):
    # 1. Adjacent cell label
    tbl = doc.Tables.Add(doc.Paragraphs(doc.Paragraphs.Count).Range, 1, 2)
    tbl.Cell(1, 1).Range.Text = "Consenso al trattamento"
    cc = doc.ContentControls.Add(8, tbl.Cell(1, 2).Range)  # wdContentControlCheckBox = 8
    cc.Checked = False

    res = wt.word_set_checkbox(doc, anchor="Consenso al trattamento", checked=True)
    assert res["ok"] is True
    assert cc.Checked is True

    # 2. Unicode symbol toggle
    doc.Content.InsertAfter("Stato approvazione: \u2610 da verificare\r")
    res_sym = wt.word_set_checkbox(doc, anchor="Stato approvazione", checked=True)
    assert res_sym["ok"] is True
    assert res_sym["symbol"] == "\u2611"

    res_sym_off = wt.word_set_checkbox(doc, anchor="Stato approvazione", checked=False)
    assert res_sym_off["ok"] is True
    assert res_sym_off["symbol"] == "\u2610"


def test_word_com_nested_table_resolution(doc):
    t0 = doc.Tables.Add(doc.Paragraphs(doc.Paragraphs.Count).Range, 2, 2)
    t1 = doc.Tables.Add(t0.Cell(1, 1).Range, 1, 2)

    assert doc.Tables.Count == 1  # COM doc.Tables only counts top-level tables

    # wt._table must resolve T1 via _all_tables
    resolved_t1 = wt._table(doc, "T1")
    assert resolved_t1 is not None

    # Clone row on T1
    clone_res = wt.word_table_clone_row(doc, "T1", 0, 1)
    assert clone_res["ok"] is True
    assert clone_res["rows"] == 2
