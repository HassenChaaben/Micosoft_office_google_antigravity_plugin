"""Unit tests for daemon.office.ooxml_builder (blocks JSON -> flat-OPC OOXML)."""

import xml.etree.ElementTree as ET

import pytest

from daemon.office.ooxml_builder import build_ooxml, validate_blocks
from daemon.office.word_snapshot import build_snapshot

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
PKG = "{http://schemas.microsoft.com/office/2006/xmlPackage}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def parts(ooxml: str) -> dict:
    root = ET.fromstring(ooxml)
    assert root.tag == PKG + "package"
    out = {}
    for part in root.findall(PKG + "part"):
        data = part.find(PKG + "xmlData")
        out[part.get(PKG + "name")] = data[0]
    return out


def body(ooxml: str) -> ET.Element:
    return parts(ooxml)["/word/document.xml"].find(W + "body")


def pstyles(el: ET.Element) -> list:
    return [p.find(W + "pPr/" + W + "pStyle").get(W + "val")
            for p in el.iter(W + "p") if p.find(W + "pPr/" + W + "pStyle") is not None]


def text_of(el: ET.Element) -> str:
    return "".join(t.text or "" for t in el.iter(W + "t"))


MERGED_TABLE = {
    "type": "table",
    "header_rows": 2,
    "rows": [
        [{"text": "Attività", "rowspan": 2}, {"text": "Trimestre", "colspan": 3}],
        ["Q1", "Q2", "Q3"],
        ["Analisi", "x", "", "x"],
    ],
}

RICH_REPORT = [
    {"type": "title", "text": "Progetto di Automazione"},
    {"type": "subtitle", "text": "Cloud4Job"},
    {"type": "toc", "levels": "1-3"},
    {"type": "heading", "level": 1, "text": "Introduzione"},
    {"type": "heading", "level": 2, "text": "Obiettivi"},
    {"type": "paragraph", "align": "justify", "runs": [
        {"text": "Testo normale, "}, {"text": "grassetto", "bold": True},
        {"text": " e ", "italic": True}, {"text": "rosso", "color": "#C00000", "size": 14,
                                           "highlight": "yellow", "font": "Arial", "underline": True}]},
    {"type": "bullets", "items": ["Uno", {"text": "Due", "children": ["Due.a", {"text": "Due.b", "children": ["Due.b.i"]}]}]},
    {"type": "numbered", "items": ["Primo", "Secondo"]},
    MERGED_TABLE,
    {"type": "caption", "label": "Tabella", "text": "Pianificazione"},
    {"type": "page_break"},
    {"type": "quote", "text": "Fine del rapporto"},
]


# --- package shape ------------------------------------------------------------------

def test_package_has_all_parts_and_relationships():
    p = parts(build_ooxml([{"type": "paragraph", "text": "ciao"}]))
    assert set(p) >= {"/_rels/.rels", "/word/document.xml", "/word/_rels/document.xml.rels",
                      "/word/styles.xml", "/word/numbering.xml"}
    targets = {r.get("Target") for r in p["/word/_rels/document.xml.rels"].findall(REL + "Relationship")}
    assert {"styles.xml", "numbering.xml"} <= targets
    root_targets = {r.get("Target") for r in p["/_rels/.rels"].findall(REL + "Relationship")}
    assert "word/document.xml" in root_targets


def test_rich_report_is_well_formed_and_defines_builtin_styles():
    ooxml = build_ooxml(RICH_REPORT)
    p = parts(ooxml)
    ids = {s.get(W + "styleId"): s.find(W + "name").get(W + "val") for s in p["/word/styles.xml"].findall(W + "style")}
    for sid, name in {"Title": "Title", "Subtitle": "Subtitle", "Heading1": "heading 1", "Heading6": "heading 6",
                      "Caption": "caption", "Quote": "Quote", "TableGrid": "Table Grid",
                      "ListBullet": "List Bullet", "ListNumber": "List Number", "TOCHeading": "TOC Heading"}.items():
        assert ids.get(sid) == name
    used = pstyles(body(ooxml))
    for sid in ("Title", "Subtitle", "Heading1", "Heading2", "Caption", "Quote"):
        assert sid in used
    # every style used by the document is defined in the package
    assert set(used) <= set(ids)


def test_heading_levels_map_to_style_ids():
    b = body(build_ooxml([{"type": "heading", "level": lvl, "text": f"H{lvl}"} for lvl in range(1, 7)]))
    assert pstyles(b) == [f"Heading{i}" for i in range(1, 7)]


# --- runs, escaping, alignment -------------------------------------------------------

def test_runs_formatting_properties():
    b = body(build_ooxml([RICH_REPORT[5]]))
    runs = b.findall(W + "p/" + W + "r")
    red = runs[3].find(W + "rPr")
    assert red.find(W + "color").get(W + "val") == "C00000"
    assert red.find(W + "sz").get(W + "val") == "28"
    assert red.find(W + "highlight").get(W + "val") == "yellow"
    assert red.find(W + "rFonts").get(W + "ascii") == "Arial"
    assert red.find(W + "u").get(W + "val") == "single"
    assert runs[1].find(W + "rPr/" + W + "b") is not None
    assert runs[2].find(W + "rPr/" + W + "i") is not None
    assert b.find(W + "p/" + W + "pPr/" + W + "jc").get(W + "val") == "both"


def test_escaping_and_preserved_spaces_and_line_breaks():
    nasty = 'A & B < C > D "q" \'s\'  doppio\nriga2'
    ooxml = build_ooxml([{"type": "paragraph", "text": nasty}])
    p = body(ooxml).find(W + "p")
    ts = p.findall(W + "r/" + W + "t")
    assert ts[0].text == 'A & B < C > D "q" \'s\'  doppio'
    assert ts[0].get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"
    assert p.find(W + "r/" + W + "br") is not None
    assert ts[-1].text == "riga2"


def test_paragraph_style_by_builtin_name_and_space_after():
    b = body(build_ooxml([{"type": "paragraph", "text": "x", "style": "Intense Quote", "space_after_pt": 12}]))
    assert pstyles(b) == ["IntenseQuote"]
    assert b.find(W + "p/" + W + "pPr/" + W + "spacing").get(W + "after") == "240"


# --- lists ----------------------------------------------------------------------------

def test_bullets_and_numbered_reference_numbering_definitions():
    ooxml = build_ooxml([RICH_REPORT[6], RICH_REPORT[7]])
    p = parts(ooxml)
    nums = {n.get(W + "numId"): n.find(W + "abstractNumId").get(W + "val") for n in p["/word/numbering.xml"].findall(W + "num")}
    fmts = {a.get(W + "abstractNumId"): a.find(W + "lvl/" + W + "numFmt").get(W + "val")
            for a in p["/word/numbering.xml"].findall(W + "abstractNum")}
    items = [(pp.find(W + "pPr/" + W + "numPr/" + W + "numId").get(W + "val"),
              int(pp.find(W + "pPr/" + W + "numPr/" + W + "ilvl").get(W + "val")), text_of(pp))
             for pp in body(ooxml).iter(W + "p")]
    assert [(lvl, t) for _, lvl, t in items] == [(0, "Uno"), (0, "Due"), (1, "Due.a"), (1, "Due.b"), (2, "Due.b.i"),
                                                 (0, "Primo"), (0, "Secondo")]
    bullet_num, numbered_num = items[0][0], items[-1][0]
    assert fmts[nums[bullet_num]] == "bullet"
    assert fmts[nums[numbered_num]] == "decimal"
    assert all(len(a.findall(W + "lvl")) >= 3 for a in p["/word/numbering.xml"].findall(W + "abstractNum"))


def test_each_numbered_list_restarts_with_its_own_num():
    ooxml = build_ooxml([{"type": "numbered", "items": ["a"]}, {"type": "numbered", "items": ["b"]}])
    ids = [pp.find(W + "pPr/" + W + "numPr/" + W + "numId").get(W + "val") for pp in body(ooxml).iter(W + "p")]
    assert ids[0] != ids[1]


# --- tables ---------------------------------------------------------------------------

def _cells(tr):
    out = []
    for tc in tr.findall(W + "tc"):
        pr = tc.find(W + "tcPr")
        gs = pr.find(W + "gridSpan")
        vm = pr.find(W + "vMerge")
        out.append((int(gs.get(W + "val")) if gs is not None else 1,
                    None if vm is None else (vm.get(W + "val") or "continue"), text_of(tc)))
    return out


def test_table_rowspan_and_colspan_emit_vmerge_and_gridspan():
    tbl = body(build_ooxml([MERGED_TABLE])).find(W + "tbl")
    assert len(tbl.findall(W + "tblGrid/" + W + "gridCol")) == 4
    assert tbl.find(W + "tblPr/" + W + "tblW").get(W + "type") == "pct"
    assert tbl.find(W + "tblPr/" + W + "tblW").get(W + "w") == "5000"
    assert tbl.find(W + "tblPr/" + W + "tblStyle").get(W + "val") == "TableGrid"
    rows = tbl.findall(W + "tr")
    assert _cells(rows[0]) == [(1, "restart", "Attività"), (3, None, "Trimestre")]
    assert _cells(rows[1]) == [(1, "continue", ""), (1, None, "Q1"), (1, None, "Q2"), (1, None, "Q3")]
    assert [c[:2] for c in _cells(rows[2])] == [(1, None)] * 4
    # header rows repeat and are shaded
    assert all(r.find(W + "trPr/" + W + "tblHeader") is not None for r in rows[:2])
    assert rows[2].find(W + "trPr/" + W + "tblHeader") is None
    assert rows[0].find(W + "tc/" + W + "tcPr/" + W + "shd").get(W + "fill") == "D9E2F3"
    # every cell carries at least one paragraph
    assert all(tc.find(W + "p") is not None for tc in tbl.iter(W + "tc"))


def test_rowspan_in_middle_column_and_cell_options():
    t = {"type": "table", "header_rows": 0, "columns": [{"width_pct": 20}, {"width_pct": 50}, {"width_pct": 30}],
         "rows": [["a", {"text": "B", "rowspan": 3, "shading": "#FFEE00", "valign": "center", "align": "center",
                         "bold": True}, "c"],
                  ["d", "f"],
                  ["g", "i"]]}
    tbl = body(build_ooxml([t])).find(W + "tbl")
    rows = tbl.findall(W + "tr")
    assert [c[1] for c in _cells(rows[1])] == [None, "continue", None]
    assert [c[2] for c in _cells(rows[2])] == ["g", "", "i"]
    b = rows[0].findall(W + "tc")[1]
    assert b.find(W + "tcPr/" + W + "vAlign").get(W + "val") == "center"
    assert b.find(W + "tcPr/" + W + "shd").get(W + "fill") == "FFEE00"
    assert b.find(W + "p/" + W + "pPr/" + W + "jc").get(W + "val") == "center"
    assert b.find(W + "p/" + W + "r/" + W + "rPr/" + W + "b") is not None
    widths = [int(g.get(W + "w")) for g in tbl.findall(W + "tblGrid/" + W + "gridCol")]
    assert widths[1] > widths[2] > widths[0]


def test_consecutive_tables_are_separated_by_a_paragraph():
    b = body(build_ooxml([{"type": "table", "rows": [["a"]]}, {"type": "table", "rows": [["b"]]}]))
    tags = [c.tag for c in b]
    assert tags[:3] == [W + "tbl", W + "p", W + "tbl"]


def test_borders_false_removes_borders():
    tbl = body(build_ooxml([{"type": "table", "borders": False, "rows": [["a"]]}])).find(W + "tbl")
    borders = tbl.find(W + "tblPr/" + W + "tblBorders")
    assert all(b.get(W + "val") == "nil" for b in borders)


# --- caption, toc, page break, quote ------------------------------------------------

def test_caption_contains_seq_field_and_numbers_per_label():
    b = body(build_ooxml([{"type": "caption", "label": "Tabella", "text": "Uno"},
                          {"type": "caption", "label": "Tabella", "text": "Due"}]))
    ps = b.findall(W + "p")
    fs = ps[0].find(W + "fldSimple")
    assert fs.get(W + "instr") == " SEQ Tabella \\* ARABIC "
    assert pstyles(b) == ["Caption", "Caption"]
    assert text_of(ps[0]) == "Tabella 1: Uno"
    assert text_of(ps[1]) == "Tabella 2: Due"


def test_toc_is_complex_field_with_placeholder():
    p = body(build_ooxml([{"type": "toc", "levels": "1-2"}])).find(W + "p")
    kinds = [f.get(W + "fldCharType") for f in p.iter(W + "fldChar")]
    assert kinds == ["begin", "separate", "end"]
    assert p.find(W + "r/" + W + "instrText").text.strip() == 'TOC \\o "1-2" \\h \\z \\u'
    assert "Aggiorna il sommario (F9)" in text_of(p)


def test_page_break_and_quote():
    b = body(build_ooxml([{"type": "page_break"}, {"type": "quote", "text": "q"}]))
    assert b.find(W + "p/" + W + "r/" + W + "br").get(W + "type") == "page"
    assert pstyles(b) == ["Quote"]


# --- content control -----------------------------------------------------------------

def test_block_id_wraps_body_in_tagged_content_control():
    b = body(build_ooxml([{"type": "paragraph", "text": "x"}, MERGED_TABLE], block_id="rep1"))
    children = list(b)
    assert len(children) == 1 and children[0].tag == W + "sdt"
    pr = children[0].find(W + "sdtPr")
    assert pr.find(W + "tag").get(W + "val") == "ga:rep1"
    assert pr.find(W + "alias").get(W + "val") == "Gemini Antigravity"
    content = children[0].find(W + "sdtContent")
    assert content.find(W + "tbl") is not None


def test_no_block_id_means_no_sdt():
    assert body(build_ooxml([{"type": "paragraph", "text": "x"}])).find(W + "sdt") is None


# --- validation ----------------------------------------------------------------------

def test_valid_report_has_no_errors():
    assert validate_blocks(RICH_REPORT) == []


@pytest.mark.parametrize("blocks, needle", [
    ([{"type": "banana"}], "unknown type"),
    ([{"type": "heading", "level": 7, "text": "x"}], "level"),
    ([{"type": "heading", "level": 1}], "text"),
    ([{"type": "paragraph", "text": "x", "align": "middle"}], "align"),
    ([{"type": "paragraph", "runs": [{"text": "x", "color": "red"}]}], "color"),
    ([{"type": "paragraph", "runs": [{"text": "x", "highlight": "orange"}]}], "highlight"),
    ([{"type": "bullets", "items": []}], "items"),
    ([{"type": "bullets", "items": [{"text": "a", "children": [{"text": "b", "children": [
        {"text": "c", "children": ["d"]}]}]}]}], "3 levels"),
    ([{"type": "table", "rows": [["a", "b"], ["c"]]}], "row 2"),
    ([{"type": "table", "rows": [[{"text": "a", "colspan": 2}], ["b", "c", "d"]]}], "row 2"),
    ([{"type": "table", "rows": [[{"text": "a", "rowspan": 3}, "b"], ["c"]]}], "rowspan"),
    ([{"type": "table", "columns": [{"width_pct": 50}], "rows": [["a", "b"]]}], "columns"),
    ([{"type": "table", "rows": [["a"]], "header_shading": "blue"}], "header_shading"),
    ([{"type": "caption", "text": "x"}], "label"),
    ([{"type": "toc", "levels": "abc"}], "levels"),
    (["not a dict"], "object"),
])
def test_validation_errors(blocks, needle):
    errors = validate_blocks(blocks)
    assert errors, blocks
    assert any(needle in e for e in errors), errors
    with pytest.raises(ValueError) as exc:
        build_ooxml(blocks)
    assert needle in str(exc.value)


def test_rowspan_covered_cells_must_not_be_supplied_by_caller():
    # caller wrongly repeats the covered cell -> row 2 is too wide
    errors = validate_blocks([{"type": "table", "rows": [[{"text": "a", "rowspan": 2}, "b"], ["", "c"]]}])
    assert any("row 2" in e for e in errors)


def test_validation_reports_all_errors_with_block_index():
    errors = validate_blocks([{"type": "heading", "level": 0, "text": "x"}, {"type": "nope"}])
    assert len(errors) >= 2
    assert errors[0].startswith("blocks[0]") and any(e.startswith("blocks[1]") for e in errors)


def test_blocks_must_be_a_list():
    assert validate_blocks({"type": "title"})


# --- round trip with the perception layer -------------------------------------------

def test_round_trip_snapshot_detects_outline_and_merges():
    snap = build_snapshot(build_ooxml(RICH_REPORT))
    assert [(o["level"], o["text"]) for o in snap["outline"]] == [(1, "Introduzione"), (2, "Obiettivi")]
    assert snap["stats"]["tables"] == 1
    t = snap["tables"][0]
    assert t["grid_cols"] == 4 and t["rows"] == 3
    assert t["merged_cells"] == 3  # restart + continue + gridSpan
    assert snap["field_codes"]["SEQ"] == 1
    assert snap["field_codes"]["TOC"] == 1
