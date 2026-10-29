import re
from pathlib import Path

from daemon.office.word_snapshot import build_snapshot
from daemon.office.verify import verify

FIX = Path(__file__).parent / "fixtures" / "docs"
CHARTER = (FIX / "project_charter.body.xml").read_text(encoding="utf-8")


def _fill_first_empty_cell_after(xml: str, label: str, value: str) -> str:
    """Put `value` into the first empty paragraph that follows the label cell (test helper)."""
    i = xml.index(f">{label}<")
    j = min(x for x in (xml.find("<w:tc>", i), xml.find("<w:tc ", i)) if x != -1)
    m = re.compile(r"<w:p(?=[ >/])[^>]*?(/?)>").search(xml, j)
    run = f"<w:r><w:t>{value}</w:t></w:r>"
    if m.group(1):  # self-closing empty paragraph
        return xml[:m.start()] + m.group(0)[:-2] + ">" + run + "</w:p>" + xml[m.end():]
    k = xml.index("</w:p>", m.end())
    return xml[:k] + run + xml[k:]


def test_reports_filled_fields_and_remaining_empty_ones():
    before = build_snapshot(CHARTER)
    after = build_snapshot(_fill_first_empty_cell_after(CHARTER, "Azienda", "Cloud4Job S.r.l."))
    res = verify(before, after)
    assert "Cliente › Azienda" in res["changes"]["fields_filled"]
    assert res["summary"]["fields_empty"] == before["stats"]["fields_empty"] - 1
    assert any(i["type"] == "empty_fields" for i in res["issues"])


def test_structure_change_in_a_form_is_flagged():
    before = build_snapshot(CHARTER)
    broken = CHARTER.replace("<w:tbl>", "<w:p><w:r><w:t>x</w:t></w:r></w:p><w:tbl>", 1)
    i = broken.rindex("<w:tbl>")
    broken = broken[:i] + "<w:p/>" + broken[broken.index("</w:tbl>", i) + len("</w:tbl>"):]
    after = build_snapshot(broken)
    res = verify(before, after)
    assert any(i["type"] == "tables_removed" for i in res["issues"])


def test_placeholders_and_heading_skips_are_flagged():
    W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

    def doc(body):
        return f'<w:document {W}><w:body>{body}</w:body></w:document>'

    def p(text, style=None):
        st = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
        return f"<w:p>{st}<w:r><w:t>{text}</w:t></w:r></w:p>"

    before = build_snapshot(doc(p("Intro")))
    after = build_snapshot(doc(p("Titolo", "Heading1") + p("Dettaglio", "Heading3") + p("[inserire data]")))
    types = {i["type"] for i in verify(before, after)["issues"]}
    assert "heading_level_skip" in types
    assert "placeholders_left" in types


def test_toc_present_but_headings_changed_asks_for_field_update():
    W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    toc = '<w:p><w:r><w:instrText> TOC \\o "1-3" </w:instrText></w:r></w:p>'
    h = '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>{}</w:t></w:r></w:p>'
    before = build_snapshot(f'<w:document {W}><w:body>{toc}{h.format("A")}</w:body></w:document>')
    after = build_snapshot(f'<w:document {W}><w:body>{toc}{h.format("A")}{h.format("B")}</w:body></w:document>')
    res = verify(before, after)
    assert any(i["type"] == "toc_stale" for i in res["issues"])


def test_clean_result_has_no_issues():
    s = build_snapshot(CHARTER.replace("Project charter", "Project charter"))
    res = verify(s, s, expect_all_fields_filled=False)
    assert res["issues"] == []
    assert res["ok"] is True
