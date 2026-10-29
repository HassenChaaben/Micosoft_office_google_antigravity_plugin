from pathlib import Path

import pytest

from daemon.office.word_snapshot import build_snapshot, render_map

FIXTURES = Path(__file__).parent / "fixtures" / "docs"


def load(name: str) -> str:
    return (FIXTURES / f"{name}.body.xml").read_text(encoding="utf-8")


@pytest.fixture(scope="module")
def charter():
    return build_snapshot(load("project_charter"))


@pytest.fixture(scope="module")
def risk():
    return build_snapshot(load("risk_register"))


def _fields(snap, **match):
    return [f for f in snap["fields"] if all(f.get(k) == v for k, v in match.items())]


# --- identity -----------------------------------------------------------------

def test_pm_templates_are_classified_as_form_templates(charter, risk):
    assert charter["doc_type"] == "form_template"
    assert risk["doc_type"] == "form_template"


def test_language_is_detected_as_italian(charter):
    assert charter["language"] == "it"


def test_blank_document_is_classified_blank():
    xml = (
        '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
        '<pkg:part pkg:name="/word/document.xml"><pkg:xmlData>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:body><w:p/></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>"
    )
    snap = build_snapshot(xml)
    assert snap["doc_type"] == "blank"
    assert snap["fields"] == []


# --- tables ---------------------------------------------------------------------

def test_tables_report_grid_size_and_merges(charter):
    t0 = charter["tables"][0]
    assert t0["rows"] == 28
    assert t0["grid_cols"] >= 9
    assert t0["merged_cells"] > 0
    assert t0["title"] == "Project charter"


def test_record_grid_table_is_detected_with_its_columns(risk):
    minacce = [t for t in risk["tables"] if t["title"] == "Minacce"][0]
    assert minacce["kind"] == "record_grid"
    assert minacce["columns"][:3] == ["Cod.", "Categoria", "WBS"]
    assert minacce["empty_records"] == 6


def test_two_level_headers_are_joined(risk):
    cols = [t for t in risk["tables"] if "qualitativa" in t["title"].lower()][0]["columns"]
    assert "Probabilità accadimento › Val. num." in cols
    assert "Strategia" in cols


# --- fields ---------------------------------------------------------------------

def test_form_fields_get_section_and_label_path(charter):
    azienda = _fields(charter, path="Cliente › Azienda")
    assert len(azienda) == 1
    f = azienda[0]
    assert f["empty"] is True
    assert {"table", "row", "cell"} <= set(f["loc"])


def test_same_label_in_different_sections_is_disambiguated(charter):
    paths = {f["path"] for f in charter["fields"]}
    assert "Cliente › Cognome Nome" in paths
    assert "Sponsor › Cognome Nome" in paths


def test_inline_label_value_pairs_in_one_row(charter):
    paths = {f["path"] for f in charter["fields"]}
    assert "Anagrafica progetto › Codice" in paths
    assert "Anagrafica progetto › Titolo" in paths


def test_record_grid_fields_have_record_number_and_column(risk):
    f = _fields(risk, path="Minacce › #1 › Causa")
    assert len(f) == 1


def test_labels_and_bands_are_never_reported_as_fields(charter):
    for f in charter["fields"]:
        assert f["path"].split(" › ")[-1] not in ("", "Cliente", "Sponsor")


def test_vertical_merge_continuations_are_not_fields(risk):
    # the header continuation cells under 'Cod.' (vMerge) must not appear as empty inputs
    for f in risk["fields"]:
        assert "Cod. › " not in f["path"] or "#" in f["path"]


def test_field_counts(charter, risk):
    assert charter["stats"]["fields_total"] >= 30
    assert charter["stats"]["fields_empty"] == charter["stats"]["fields_total"]
    assert risk["stats"]["fields_total"] > 100


# --- formatting awareness -------------------------------------------------------

def test_direct_formatting_is_reported_when_only_normal_style_is_used(charter):
    assert charter["styles"]["direct_formatting"] is True
    assert "Normal" in charter["styles"]["used"]


# --- map ------------------------------------------------------------------------

def test_map_mentions_type_fields_and_paths(charter):
    text = render_map(charter)
    assert "form_template" in text
    assert "Cliente › Azienda" in text
    assert "Fields to fill" in text


def test_map_respects_budget(risk):
    text = render_map(risk, budget_chars=4000)
    assert len(text) <= 4000
    assert "Fields to fill" in text  # structure is kept, detail is shortened


def test_tables_with_both_form_rows_and_record_rows_are_mixed(charter):
    assert charter["tables"][0]["kind"] == "mixed"


def test_record_fields_carry_group_record_and_column(risk):
    f = _fields(risk, path="Minacce › #1 › Causa")[0]
    assert f["kind"] == "record"
    assert (f["group"], f["record"], f["column"]) == ("Minacce", 1, "Causa")


def test_map_groups_record_grids_instead_of_listing_every_cell(risk):
    text = render_map(risk)
    assert "Minacce › records #1-#6" in text
    assert "Minacce › #1 › Causa" not in text
    assert len(text) < 6000
