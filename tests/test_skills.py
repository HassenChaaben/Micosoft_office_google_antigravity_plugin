"""Tests for the Word expertise skill files and the skill loader/router (daemon/core/skills.py)."""
from pathlib import Path

import pytest

from daemon.core import skills as sk

ROOT = Path(__file__).resolve().parents[1]
BUILTIN = ROOT / "daemon" / "skills"

EXPECTED = {
    "office-document-analysis",
    "template-form-filling",
    "pm-documentation-suite",
    "long-report-architecture",
    "styles-and-numbering",
    "sections-and-page-layout",
    "fields-captions-references",
    "advanced-tables",
    "native-charts",
    "smartart-and-diagrams",
    "review-and-track-changes",
    "document-quality-check",
    "research-for-content",
}

# Tool names allowed by the tool contract; skills must not invent others.
CONTRACT_TOOLS = {
    "word_read_section", "word_read_table", "word_find", "word_get_formatting", "office_verify",
    "agent_plan", "memory_save", "memory_search", "skill_load",
    "word_insert_content", "word_update_block",
    "word_fill_fields", "word_table_clone_row", "word_set_checkbox",
    "word_apply_style", "word_modify_style", "word_link_numbering_to_headings",
    "word_insert_section_break", "word_set_header_footer", "word_insert_toc",
    "word_insert_caption", "word_insert_cross_reference", "word_update_fields",
    "word_track_changes", "word_add_comment",
    "word_insert_chart", "word_insert_smartart", "word_insert_diagram",
    "word_replace_text", "word_set_paragraph_text", "word_delete_paragraph",
}


@pytest.fixture(scope="module")
def builtin():
    return sk.load_skills([BUILTIN])


def by_name(skills):
    return {s.name: s for s in skills}


# ---------------------------------------------------------------- loading

def test_all_13_skills_load_with_valid_frontmatter(builtin):
    names = by_name(builtin)
    assert EXPECTED <= set(names)
    for name in EXPECTED:
        s = names[name]
        assert s.description and "\n" not in s.description
        assert s.hosts == ["word"]
        assert isinstance(s.always_on, bool)
        assert isinstance(s.doc_types, list) and isinstance(s.keywords, list)
        assert s.keywords, f"{name} has no keywords"
        assert s.path.name == "SKILL.md"
        for section in ("## Office mental model", "## Procedure", "## Tool recipes",
                        "## Pitfalls", "## Verification checklist"):
            assert section in s.body, f"{name} missing {section}"
        assert 60 <= len(s.body.splitlines()) <= 260, f"{name} body length off"


def test_only_document_analysis_is_always_on(builtin):
    on = {s.name for s in builtin if s.always_on}
    assert on == {"office-document-analysis"}


def test_skills_reference_only_contract_tools(builtin):
    import re
    for s in builtin:
        used = set(re.findall(r"\b((?:word|office|agent|memory|skill)_[a-z_]+)\(", s.body))
        unknown = used - CONTRACT_TOOLS
        assert not unknown, f"{s.name} references non-contract tools {unknown}"


def test_keywords_are_bilingual(builtin):
    names = by_name(builtin)
    kw = {k.lower() for k in names["template-form-filling"].keywords}
    assert {"compila", "fill"} <= kw
    kw = {k.lower() for k in names["smartart-and-diagrams"].keywords}
    assert {"diagramma", "wbs", "reticolo"} <= kw


# ---------------------------------------------------------------- frontmatter parser

def test_frontmatter_parser_handles_inline_and_block_lists():
    text = (
        "---\n"
        "name: demo\n"
        "description: \"A demo: with colon\"\n"
        "hosts: [word]\n"
        "always_on: false\n"
        "triggers:\n"
        "  doc_types: [report, blank]\n"
        "  keywords:\n"
        "    - relazione\n"
        "    - 'report'\n"
        "---\n"
        "Body here\n"
    )
    meta, body = sk.parse_frontmatter(text)
    assert meta["name"] == "demo"
    assert meta["description"] == "A demo: with colon"
    assert meta["hosts"] == ["word"]
    assert meta["always_on"] is False
    assert meta["triggers"]["doc_types"] == ["report", "blank"]
    assert meta["triggers"]["keywords"] == ["relazione", "report"]
    assert body.strip() == "Body here"


# ---------------------------------------------------------------- routing

def test_form_template_compila_routes_form_and_pm(builtin):
    routed = sk.route(builtin, host="word", doc_type="form_template",
                      request="compila il documento")
    full = [s.name for s in routed["full"]]
    assert full[0] == "office-document-analysis"
    assert "template-form-filling" in full
    assert "pm-documentation-suite" in full
    listed = {s.name for s in routed["listed"]}
    assert not (set(full) & listed)
    assert set(full) | listed == {s.name for s in builtin}


def test_blank_report_request_routes_long_report(builtin):
    routed = sk.route(builtin, host="word", doc_type="blank",
                      request="scrivi una relazione con indice e grafico")
    full = [s.name for s in routed["full"]]
    assert "office-document-analysis" in full
    assert "long-report-architecture" in full
    assert len(full) <= 1 + 3


def test_accent_insensitive_keywords(builtin):
    routed = sk.route(builtin, host="word", doc_type="document",
                      request="Inserisci la matrice probabilità impatto")
    assert "pm-documentation-suite" in [s.name for s in routed["full"]]
    routed2 = sk.route(builtin, host="word", doc_type="document",
                       request="inserisci la matrice probabilita impatto")
    assert [s.name for s in routed["full"]] == [s.name for s in routed2["full"]]


def test_forced_skill_via_slash(builtin):
    forced, cleaned = sk.parse_forced("/smartart-and-diagrams fai il reticolo")
    assert forced == ["smartart-and-diagrams"]
    assert cleaned == "fai il reticolo"
    routed = sk.route(builtin, host="word", doc_type="blank", request=cleaned, forced=forced)
    full = [s.name for s in routed["full"]]
    assert full[0] == "office-document-analysis"
    assert full[1] == "smartart-and-diagrams"


def test_parse_forced_known_filter_keeps_unknown_commands():
    forced, cleaned = sk.parse_forced("/check-doc please", known={"native-charts"})
    assert forced == []
    assert cleaned == "/check-doc please"
    forced, cleaned = sk.parse_forced("no slash here")
    assert forced == [] and cleaned == "no slash here"


def test_route_filters_by_host(builtin):
    routed = sk.route(builtin, host="excel", doc_type="blank", request="grafico")
    assert routed["full"] == [] and routed["listed"] == []


# ---------------------------------------------------------------- user overrides

def test_user_dir_overrides_builtin_by_name(tmp_path):
    d = tmp_path / "native-charts"
    d.mkdir()
    (d / "SKILL.md").write_text(
        "---\nname: native-charts\ndescription: My custom charts\nhosts: [word]\n"
        "always_on: false\ntriggers:\n  doc_types: []\n  keywords: [torta]\n---\nCUSTOM BODY\n",
        encoding="utf-8",
    )
    extra = tmp_path / "my-extra"
    extra.mkdir()
    (extra / "SKILL.md").write_text(
        "---\nname: my-extra\ndescription: Extra\nhosts: [word]\n---\nX\n", encoding="utf-8")
    skills = by_name(sk.load_skills([BUILTIN, tmp_path]))
    assert skills["native-charts"].description == "My custom charts"
    assert skills["native-charts"].body.strip() == "CUSTOM BODY"
    assert "my-extra" in skills
    assert len(skills) == len(EXPECTED) + 1


def test_default_dirs_include_localappdata(monkeypatch, tmp_path):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    user = tmp_path / "GeminiAntigravityOffice" / "skills" / "zz-user"
    user.mkdir(parents=True)
    (user / "SKILL.md").write_text("---\nname: zz-user\ndescription: U\nhosts: [word]\n---\nB\n",
                                   encoding="utf-8")
    names = {s.name for s in sk.load_skills()}
    assert "zz-user" in names and EXPECTED <= names


def test_broken_skill_file_is_skipped(tmp_path):
    bad = tmp_path / "bad"
    bad.mkdir()
    (bad / "SKILL.md").write_text("no frontmatter at all", encoding="utf-8")
    assert sk.load_skills([tmp_path]) == []


# ---------------------------------------------------------------- rendering

def test_render_for_prompt_contains_full_bodies_and_listed_names(builtin):
    routed = sk.route(builtin, host="word", doc_type="form_template", request="compila")
    text = sk.render_for_prompt(routed)
    assert "skill_load" in text
    for s in routed["full"]:
        assert s.name in text
        assert s.body.strip().splitlines()[0] in text
    for s in routed["listed"]:
        assert f"{s.name} — {s.description}" in text
