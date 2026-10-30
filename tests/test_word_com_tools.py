"""
Native Word COM tools — tested against a SEPARATE hidden Word instance.

Safety: every Word object here comes from DispatchEx (a new, invisible process).
We never call GetActiveObject and never touch the user's open documents.
"""

import asyncio
import threading

import pytest

from daemon.office import word_com_tools as wt
from daemon.office.word_snapshot import build_snapshot

pytestmark = pytest.mark.word

win32com = pytest.importorskip("win32com.client")
pythoncom = pytest.importorskip("pythoncom")


# --- fixtures -------------------------------------------------------------------------

@pytest.fixture(scope="module")
def word():
    try:
        app = win32com.DispatchEx("Word.Application")
    except Exception as exc:  # Word not installed / COM broken
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


HEADINGS = ["Introduzione", "Obiettivi", "Rischi"]


@pytest.fixture
def doc(word):
    d = word.Documents.Add()
    try:
        for h in HEADINGS:
            d.Content.InsertAfter(h + "\r")
            d.Content.InsertAfter(f"Testo del paragrafo {h}.\r")
        for i in range(1, d.Paragraphs.Count + 1):
            p = d.Paragraphs(i)
            if p.Range.Text.strip() in HEADINGS:
                p.Style = -2  # wdStyleHeading1 (works in any UI language)
        yield d
    finally:
        d.Close(False)


def snapshot(doc):
    return build_snapshot(doc.Content.WordOpenXML)


def ok(result):
    assert result.get("ok") is True, result
    return result


# --- anchors --------------------------------------------------------------------------

def test_resolve_anchor_paragraph_is_zero_based(doc):
    assert wt.resolve_anchor(doc, "p0").Text.startswith("Introduzione")
    assert wt.resolve_anchor(doc, "p1").Text.startswith("Testo del paragrafo Introduzione")


def test_resolve_anchor_heading_text_and_start_end(doc):
    assert wt.resolve_anchor(doc, "Obiettivi").Text.startswith("Obiettivi")
    assert wt.resolve_anchor(doc, "start").Start == 0
    assert wt.resolve_anchor(doc, "end").End >= doc.Content.End - 1


def test_resolve_anchor_tables_and_cells(doc):
    t = doc.Tables.Add(wt.resolve_anchor(doc, "end"), 3, 2)
    t.Cell(3, 2).Range.Text = "target"
    assert wt.resolve_anchor(doc, "T0").Start == t.Range.Start
    assert wt.resolve_anchor(doc, "T0r2c1").Text.startswith("target")


def test_resolve_anchor_content_control_tag(doc):
    cc = doc.ContentControls.Add(1, doc.Paragraphs(2).Range)  # wdContentControlText
    cc.Tag = "blk1"
    assert wt.resolve_anchor(doc, "ga:blk1").Start == cc.Range.Start


def test_resolve_anchor_unknown_raises(doc):
    with pytest.raises(ValueError):
        wt.resolve_anchor(doc, "Nonexistent heading xyz")


def test_insert_after_heading_creates_normal_paragraph(doc):
    before = doc.Paragraphs.Count
    rng = wt.insertion_range(doc, "Obiettivi")
    assert doc.Paragraphs.Count == before + 1
    para = rng.Paragraphs(1)
    assert para.Range.Text == "\r"
    assert para.Style.NameLocal == doc.Styles(-1).NameLocal
    assert doc.Paragraphs(3).Range.Text.startswith("Obiettivi")
    assert doc.Paragraphs(4).Range.Text == "\r"


def test_insert_before_and_replace(doc):
    rng = wt.insertion_range(doc, "before:p2")
    assert doc.Paragraphs(3).Range.Text == "\r"
    assert rng.Start == doc.Paragraphs(3).Range.Start
    rng = wt.insertion_range(doc, "replace:p1")
    assert rng.Text == ""
    assert doc.Paragraphs(2).Range.Text == "\r"


# --- charts ---------------------------------------------------------------------------

def test_insert_column_chart_with_data(doc):
    r = ok(wt.word_insert_chart(doc, type="column", categories=["Q1", "Q2", "Q3"],
                                series=[{"name": "Ricavi", "values": [10, 20, 30]},
                                        {"name": "Costi", "values": [5, 8, 12]}],
                                title="Andamento", axes={"x": "Trimestre", "y": "EUR"}, location="Obiettivi"))
    ish = doc.InlineShapes(1)
    assert ish.HasChart
    ch = ish.Chart
    assert ch.SeriesCollection().Count == 2
    assert ch.ChartTitle.Text == "Andamento"
    assert list(ch.SeriesCollection(1).Values) == [10, 20, 30]
    assert r["series"] == 2 and r["points"] == 3
    assert snapshot(doc)["visuals"].get("chart") == 1


def test_insert_pie_chart_has_data_labels(doc):
    ok(wt.word_insert_chart(doc, type="pie", categories=["A", "B"],
                            series=[{"name": "Quota", "values": [60, 40]}], title="Quote", location="end"))
    ch = doc.InlineShapes(1).Chart
    assert ch.ChartType == 5
    assert ch.SeriesCollection(1).HasDataLabels


def test_insert_combo_chart_second_series_is_line(doc):
    ok(wt.word_insert_chart(doc, type="combo", categories=["Gen", "Feb"],
                            series=[{"name": "Vendite", "values": [3, 4]}, {"name": "Trend", "values": [3.5, 4.2]}],
                            location="end"))
    ch = doc.InlineShapes(1).Chart
    assert ch.SeriesCollection(2).ChartType == 4  # xlLine


def test_insert_chart_rejects_bad_input(doc):
    r = wt.word_insert_chart(doc, type="donut3d", categories=["a"], series=[{"name": "s", "values": [1]}])
    assert r["ok"] is False and "type" in r["error"]
    r = wt.word_insert_chart(doc, type="column", categories=["a", "b"], series=[{"name": "s", "values": [1]}])
    assert r["ok"] is False
    assert doc.InlineShapes.Count == 0


# --- SmartArt -------------------------------------------------------------------------

def test_layout_ids_exist_in_this_office(word, doc):
    ids = {word.SmartArtLayouts(i).Id for i in range(1, word.SmartArtLayouts.Count + 1)}
    for kind, urns in wt.SMARTART_LAYOUTS.items():
        assert any(u in ids for u in urns), kind


def test_insert_smartart_process(doc):
    r = ok(wt.word_insert_smartart(doc, layout="process",
                                   nodes=[{"text": "Analisi"}, {"text": "Sviluppo"}, {"text": "Test"}],
                                   location="Obiettivi"))
    sa = doc.InlineShapes(1).SmartArt
    assert [sa.Nodes(i).TextFrame2.TextRange.Text for i in range(1, sa.Nodes.Count + 1)] == ["Analisi", "Sviluppo", "Test"]
    assert r["nodes"] == 3 and r["inline"] is True
    assert snapshot(doc)["visuals"].get("smartart") == 1


def test_insert_smartart_org_chart_levels(doc):
    ok(wt.word_insert_smartart(doc, layout="org_chart", nodes=[
        {"text": "CEO", "level": 1}, {"text": "CTO", "level": 2}, {"text": "Dev", "level": 3},
        {"text": "CFO", "level": 2}], location="end"))
    sa = doc.InlineShapes(1).SmartArt
    got = [(n.TextFrame2.TextRange.Text, n.Level) for n in (sa.AllNodes(i) for i in range(1, sa.AllNodes.Count + 1))]
    assert ("CEO", 1) in got and ("CTO", 2) in got and ("Dev", 3) in got and ("CFO", 2) in got
    assert len(got) == 4


@pytest.mark.parametrize("layout", ["hierarchy", "cycle", "list", "pyramid", "matrix", "timeline"])
def test_insert_smartart_all_layouts(doc, layout):
    ok(wt.word_insert_smartart(doc, layout=layout, nodes=[{"text": t} for t in ("A", "B", "C", "D")], location="end"))
    assert doc.InlineShapes(1).HasSmartArt


# --- native diagrams (canvas + glued connectors) ---------------------------------------

def _canvas(doc, name):
    s = doc.Shapes(name)
    assert s.Type == 20  # msoCanvas
    return s


def test_insert_flowchart_with_decision_and_glued_connectors(doc):
    r = ok(wt.word_insert_diagram(doc, kind="flowchart", nodes=[
        {"id": "s", "text": "Inizio", "type": "start"}, {"id": "a", "text": "Raccolta requisiti"},
        {"id": "d", "text": "Approvato?"}, {"id": "b", "text": "Sviluppo"}, {"id": "e", "text": "Fine", "type": "end"}],
        edges=[{"from": "s", "to": "a"}, {"from": "a", "to": "d"}, {"from": "d", "to": "b", "label": "Sì"},
               {"from": "d", "to": "a", "label": "No"}, {"from": "b", "to": "e"}], location="Obiettivi"))
    cv = _canvas(doc, r["name"])
    items = [cv.CanvasItems(i) for i in range(1, cv.CanvasItems.Count + 1)]
    connectors = [i for i in items if i.Connector]
    assert len(connectors) == 5
    assert r["shapes"] == 5 and r["glued"] == 5
    assert all(c.ConnectorFormat.BeginConnected and c.ConnectorFormat.EndConnected for c in connectors)
    diamond = [i for i in items if not i.Connector and i.AutoShapeType == 63]
    assert len(diamond) == 1
    assert cv.WrapFormat.Type == 7  # inline with text
    assert cv.Anchor.Paragraphs(1).Range.Start == doc.Paragraphs(4).Range.Start
    assert snapshot(doc)["visuals"].get("canvas") == 1


def test_insert_wbs_tree(doc):
    r = ok(wt.word_insert_diagram(doc, kind="wbs", nodes=[
        {"id": "0", "text": "Progetto"}, {"id": "1", "text": "Analisi"}, {"id": "2", "text": "Sviluppo"},
        {"id": "2.1", "text": "Backend"}, {"id": "2.2", "text": "Frontend"}],
        edges=[{"from": "0", "to": "1"}, {"from": "0", "to": "2"}, {"from": "2", "to": "2.1"}, {"from": "2", "to": "2.2"}],
        location="end"))
    cv = _canvas(doc, r["name"])
    assert r["shapes"] == 5 and r["connectors"] == 4
    tops = {cv.CanvasItems(i).TextFrame.TextRange.Text.strip(): cv.CanvasItems(i).Top
            for i in range(1, cv.CanvasItems.Count + 1) if not cv.CanvasItems(i).Connector}
    assert tops["Progetto"] < tops["Analisi"] < tops["Backend"]
    assert tops["Analisi"] == tops["Sviluppo"]


def test_insert_network_marks_critical_path_red(doc):
    r = ok(wt.word_insert_diagram(doc, kind="network", nodes=[
        {"id": "A", "text": "A", "duration": 3}, {"id": "B", "text": "B", "duration": 2},
        {"id": "C", "text": "C", "duration": 5}, {"id": "D", "text": "D", "duration": 1}],
        edges=[{"from": "A", "to": "B"}, {"from": "A", "to": "C"}, {"from": "B", "to": "D"}, {"from": "C", "to": "D"}],
        location="end"))
    assert set(r["critical"]) == {"A", "C", "D"}
    cv = _canvas(doc, r["name"])
    red = [cv.CanvasItems(i) for i in range(1, cv.CanvasItems.Count + 1)
           if cv.CanvasItems(i).Line.ForeColor.RGB == wt.RED]
    assert len(red) >= 5  # 3 critical nodes + 2 critical edges


def test_insert_gantt_bars(doc):
    r = ok(wt.word_insert_diagram(doc, kind="gantt", nodes=[
        {"id": "a", "text": "Analisi", "duration": 5}, {"id": "b", "text": "Sviluppo", "duration": 10},
        {"id": "c", "text": "Test", "duration": 3, "critical": True}],
        edges=[{"from": "a", "to": "b"}, {"from": "b", "to": "c"}], location="end"))
    cv = _canvas(doc, r["name"])
    assert r["bars"] == 3
    bars = {cv.CanvasItems(i).Name: cv.CanvasItems(i) for i in range(1, cv.CanvasItems.Count + 1)
            if cv.CanvasItems(i).Name.startswith("bar_")}
    assert bars["bar_1"].Left > bars["bar_0"].Left
    assert bars["bar_1"].Width > bars["bar_0"].Width > bars["bar_2"].Width


def test_diagram_rejects_unknown_edge(doc):
    r = wt.word_insert_diagram(doc, kind="flowchart", nodes=[{"id": "a", "text": "A"}], edges=[{"from": "a", "to": "zz"}])
    assert r["ok"] is False and "zz" in r["error"]


def test_layered_layout_breaks_cycles_and_layers_longest_path():
    layer, order, back = wt.layered_layout(["s", "a", "d", "b"], [("s", "a"), ("a", "d"), ("d", "a"), ("d", "b")])
    assert back == {("d", "a")}
    assert [layer[i] for i in ("s", "a", "d", "b")] == [0, 1, 2, 3]


def test_critical_path_matches_hand_computed_cpm():
    es, ef, slack, crit = wt.critical_path(["A", "B", "C", "D"], [("A", "B"), ("A", "C"), ("B", "D"), ("C", "D")],
                                           {"A": 3, "B": 2, "C": 5, "D": 1})
    assert crit == ["A", "C", "D"] and slack["B"] == 3 and ef["D"] == 9


def test_tree_layout_centres_parents_over_children():
    pos, leaves = wt.tree_layout(["0", "1", "2", "2.1", "2.2"], [("0", "1"), ("0", "2"), ("2", "2.1"), ("2", "2.2")])
    assert leaves == 3 and pos["2"] == (1, 1.5) and pos["0"][0] == 0


def test_glue_connectors_xml_adds_connection_and_geometry():
    xml = ('<wps:cNvPr id="11" name="n_0"/><wps:cNvPr id="12" name="n_1"/>'
           '<wps:cNvPr id="13" name="e_0"/><wps:cNvCnPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/>'
           '<a:ext cx="1" cy="1"/></a:xfrm><a:prstGeom prst="bentConnector3">')
    out, glued = wt.glue_connectors_xml(xml, [("e_0", "n_0", 2, "n_1", 0, ("bentConnector2", 1, 2, 3, 4, 0, False, True))])
    assert glued == 1
    assert '<a:stCxn id="11" idx="2"/><a:endCxn id="12" idx="0"/>' in out
    assert 'flipV="1"' in out and 'prst="bentConnector2"' in out and '<a:off x="12700" y="25400"/>' in out


# --- fields: TOC, captions, cross references --------------------------------------------

def test_insert_toc_and_update_fields(doc):
    ok(wt.word_insert_toc(doc, location="start", levels="1-2"))
    assert doc.TablesOfContents.Count == 1
    toc = doc.TablesOfContents(1)
    assert toc.UpperHeadingLevel == 1 and toc.LowerHeadingLevel == 2
    assert "Obiettivi" in toc.Range.Text
    assert snapshot(doc)["field_codes"].get("TOC") == 1
    r = ok(wt.word_update_fields(doc))
    assert r["tocs"] == 1


def test_caption_and_cross_reference(doc):
    ok(wt.word_insert_chart(doc, type="bar", categories=["a", "b"], series=[{"name": "s", "values": [1, 2]}],
                            location="Rischi"))
    r = ok(wt.word_insert_caption(doc, anchor="i0", label="Figure", text="Rischi per area", position="below"))
    assert "Rischi per area" in r["text"]
    codes = [doc.Fields(i).Code.Text.strip().split()[0] for i in range(1, doc.Fields.Count + 1)]
    assert "SEQ" in codes
    ok(wt.word_insert_cross_reference(doc, location="p1", target_anchor="Figure 1", kind="caption", prefix=" (vedi ",
                                      suffix=")"))
    codes = [doc.Fields(i).Code.Text.strip().split()[0] for i in range(1, doc.Fields.Count + 1)]
    assert "REF" in codes
    assert "(vedi " in doc.Paragraphs(2).Range.Text
    snap = snapshot(doc)
    assert snap["field_codes"].get("SEQ") == 1 and snap["field_codes"].get("REF") == 1


def test_caption_for_table_above(doc):
    doc.Tables.Add(wt.insertion_range(doc, "end"), 2, 2)
    r = ok(wt.word_insert_caption(doc, anchor="T0", label="Table", text="Budget", position="above"))
    assert r["label"] == doc.Application.CaptionLabels(-2).Name


def test_cross_reference_heading(doc):
    ok(wt.word_insert_cross_reference(doc, location="p1", target_anchor="Rischi", kind="heading"))
    assert "Rischi" in doc.Paragraphs(2).Range.Text
    assert snapshot(doc)["field_codes"].get("REF") == 1


# --- layout: sections, headers/footers ---------------------------------------------------

def test_section_break_landscape(doc):
    r = ok(wt.word_insert_section_break(doc, location="Obiettivi", type="next_page", orientation="landscape"))
    assert doc.Sections.Count == 2
    assert doc.Sections(2).PageSetup.Orientation == 1
    assert doc.Sections(1).PageSetup.Orientation == 0
    assert r["section"] == 2
    assert doc.Sections(2).Range.Paragraphs(1).Range.Text.startswith("Testo del paragrafo Obiettivi")


def test_header_footer_with_page_numbers(doc):
    ok(wt.word_set_header_footer(doc, section=1, kind="header", text="Progetto Cloud4Job", alignment="right"))
    ok(wt.word_set_header_footer(doc, section=1, kind="footer", text="Pagina", page_numbers=True,
                                 different_first_page=True, alignment="center"))
    sec = doc.Sections(1)
    assert sec.Headers(1).Range.Text.strip() == "Progetto Cloud4Job"
    assert sec.Headers(1).Range.ParagraphFormat.Alignment == 2
    assert sec.PageSetup.DifferentFirstPageHeaderFooter
    footer = sec.Footers(1).Range
    assert any(footer.Fields(i).Type == 33 for i in range(1, footer.Fields.Count + 1))  # wdFieldPage


# --- review ---------------------------------------------------------------------------

def test_track_changes_and_comment(doc):
    ok(wt.word_track_changes(doc, on=True))
    assert doc.TrackRevisions
    ok(wt.word_track_changes(doc, on=False))
    assert not doc.TrackRevisions
    ok(wt.word_add_comment(doc, anchor="Rischi", text="Aggiungere matrice dei rischi"))
    assert doc.Comments.Count == 1
    assert doc.Comments(1).Range.Text == "Aggiungere matrice dei rischi"


# --- styles & numbering ------------------------------------------------------------------

def test_modify_builtin_style_by_english_name(doc):
    ok(wt.word_modify_style(doc, style="Heading 1", font="Arial", size_pt=18, color="#1F4E79", bold=True,
                            space_before_pt=12, space_after_pt=6, align="center"))
    st = doc.Styles(-2)
    assert st.Font.Name == "Arial" and st.Font.Size == 18 and st.Font.Bold
    assert st.Font.Color == 0x1F + 0x4E * 256 + 0x79 * 65536
    assert st.ParagraphFormat.SpaceBefore == 12 and st.ParagraphFormat.Alignment == 1


def test_modify_style_unknown_name_fails(doc):
    r = wt.word_modify_style(doc, style="No Such Style 42", bold=True)
    assert r["ok"] is False


def test_apply_style_localised_alias(doc):
    ok(wt.word_apply_style(doc, anchor="p1", style="Titolo 2"))
    assert doc.Paragraphs(2).Style.NameLocal == doc.Styles(-3).NameLocal
    ok(wt.word_apply_style(doc, anchor="p3", style="Quote"))
    assert doc.Paragraphs(4).Style.NameLocal == doc.Styles(-181).NameLocal


def test_link_numbering_to_headings(doc):
    doc.Paragraphs(2).Style = -3
    ok(wt.word_link_numbering_to_headings(doc, scheme="1.1.1"))
    assert doc.Paragraphs(1).Range.ListFormat.ListString == "1"
    assert doc.Paragraphs(2).Range.ListFormat.ListString == "1.1"
    assert doc.Paragraphs(3).Range.ListFormat.ListString == "2"


# --- forms & tables -----------------------------------------------------------------------

def test_set_checkbox(doc):
    cc = doc.ContentControls.Add(8, doc.Paragraphs(2).Range)  # wdContentControlCheckBox
    cc.Tag = "approved"
    ok(wt.word_set_checkbox(doc, anchor="ga:approved", checked=True))
    assert cc.Checked
    ok(wt.word_set_checkbox(doc, anchor="ga:approved", checked=False))
    assert not cc.Checked
    r = wt.word_set_checkbox(doc, anchor="p0", checked=True)
    assert r["ok"] is False


def test_table_clone_row_keeps_formatting_and_order(doc):
    t = doc.Tables.Add(wt.insertion_range(doc, "end"), 3, 2)
    t.Cell(1, 1).Range.Text = "Header"
    t.Cell(2, 1).Range.Text = "Riga 1"
    t.Rows(2).Shading.BackgroundPatternColor = 255
    t.Cell(3, 1).Range.Text = "Totale"
    r = ok(wt.word_table_clone_row(doc, table="T0", row=1, count=2))
    assert t.Rows.Count == 5
    assert r["new_rows"] == [2, 3]
    texts = [t.Cell(i, 1).Range.Text[:-2] for i in range(1, 6)]
    assert texts == ["Header", "Riga 1", "", "", "Totale"]
    assert all(t.Rows(i).Shading.BackgroundPatternColor == 255 for i in (2, 3, 4))


def test_delete_column_and_row(doc):
    t = doc.Tables.Add(wt.insertion_range(doc, "end"), 3, 3)
    t.Cell(1, 1).Range.Text = "A1"
    t.Cell(1, 2).Range.Text = "B1"
    t.Cell(1, 3).Range.Text = "C1"
    t.Cell(2, 1).Range.Text = "A2"
    t.Cell(3, 1).Range.Text = "A3"

    # Delete column 1 (0-based, "B1")
    r_col = ok(wt.word_delete_column(doc, table="T0", col=1))
    assert r_col["remaining_columns"] == 2
    assert t.Columns.Count == 2
    assert t.Cell(1, 1).Range.Text[:-2] == "A1"
    assert t.Cell(1, 2).Range.Text[:-2] == "C1"

    # Delete row 1 (0-based, "A2")
    r_row = ok(wt.word_delete_row(doc, table="T0", row=1))
    assert r_row["remaining_rows"] == 2
    assert t.Rows.Count == 2
    assert t.Cell(1, 1).Range.Text[:-2] == "A1"
    assert t.Cell(2, 1).Range.Text[:-2] == "A3"


# --- registry ---------------------------------------------------------------------------

def test_registry_contract():
    expected = {"word_insert_chart", "word_insert_smartart", "word_insert_diagram", "word_insert_toc",
                "word_insert_caption", "word_insert_cross_reference", "word_update_fields",
                "word_insert_section_break", "word_set_header_footer", "word_track_changes", "word_add_comment",
                "word_modify_style", "word_apply_style", "word_link_numbering_to_headings", "word_set_checkbox",
                "word_table_clone_row", "word_delete_column", "word_delete_row"}
    assert set(wt.TOOLS) == expected
    for name, spec in wt.TOOLS.items():
        assert callable(spec["fn"]) and len(spec["description"]) > 40
        assert spec["inputSchema"]["type"] == "object"
        assert isinstance(spec["destructive"], bool)


def test_call_tool_reports_bad_arguments(doc):
    r = wt.call_tool(doc, "word_track_changes", {"on": True, "bogus": 1})
    assert r["ok"] is True and r.get("ignored_args") == ["bogus"]
    r = wt.call_tool(doc, "word_no_such_tool", {})
    assert r["ok"] is False


# --- COM worker -----------------------------------------------------------------------

def test_com_worker_runs_jobs_on_one_sta_thread_and_finds_documents(tmp_path):
    from daemon.office.com_worker import ComWorker

    created = {}

    def factory():
        app = win32com.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        created["app"] = app
        return app

    worker = ComWorker(app_factory=factory)
    worker.start()
    try:
        path = tmp_path / "Rapporto.docx"

        def make():
            app = worker.app()
            d = app.Documents.Add()
            d.Content.Text = "ciao"
            d.SaveAs2(str(path))
            app.Documents.Add()  # becomes the active document
            return threading.get_ident()

        async def main():
            t1 = await worker.run(make)
            t2 = await worker.run(threading.get_ident)
            by_path = await worker.run(lambda: worker.get_document(str(path).upper()).Name)
            by_name = await worker.run(lambda: worker.get_document("rapporto.docx").Name)
            by_url = await worker.run(
                lambda: worker.get_document("https://d.docs.live.net/abc/Documenti/Rapporto.docx").Name)
            fallback = await worker.run(lambda: worker.get_document(None).Name)
            return t1, t2, by_path, by_name, by_url, fallback

        t1, t2, by_path, by_name, by_url, fallback = asyncio.run(main())
        assert t1 == t2 != threading.get_ident()
        assert by_path == by_name == by_url == "Rapporto.docx"
        assert fallback != "Rapporto.docx"

        with pytest.raises(ZeroDivisionError):
            asyncio.run(worker.run(lambda: 1 / 0))
    finally:
        def shutdown():
            app = created.get("app")
            if app is not None:
                for i in range(app.Documents.Count, 0, -1):
                    app.Documents(i).Close(False)
                app.Quit()
        worker.call(shutdown)
        worker.stop()
