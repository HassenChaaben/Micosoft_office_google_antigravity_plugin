"""
Word document perception.

Turns the flat-OPC OOXML of a Word document (Office.js `body.getOoxml()` or COM
`Range.WordOpenXML`) into a structured snapshot, and renders a compact Document Map
that is given to the model before every turn, so the agent is never blind.

The snapshot understands how real templates are built: shaded full-width rows are
section bands, shaded cells are labels, white empty cells are input fields, and a
shaded header row followed by empty rows is a record grid.
"""

import re
import xml.etree.ElementTree as ET
from collections import Counter
from typing import Any, Dict, List, Optional

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
PKG = "{http://schemas.microsoft.com/office/2006/xmlPackage}"
SEP = " › "

PLACEHOLDER_RE = re.compile(
    r"^\s*(\[[^\]]*\]|<[^>]*>|x{3,}|_{3,}|\.{4,}|gg/mm/aaaa|dd/mm/yyyy|inserire.*|insert .*|tbd|n\.d\.)\s*$",
    re.IGNORECASE,
)
STOPWORDS = {
    "it": {"il", "di", "che", "della", "del", "per", "con", "una", "dei", "delle", "sono", "nel", "alla", "progetto"},
    "en": {"the", "of", "and", "to", "in", "is", "for", "with", "that", "this", "are", "project"},
    "fr": {"le", "la", "les", "des", "du", "et", "pour", "avec", "dans", "est", "une", "projet"},
}
DIRECT_RUN_PROPS = ("b", "sz", "color", "rFonts", "i", "u", "highlight")
GRAPHIC_KINDS = {
    "chart": "chart",
    "diagram": "smartart",
    "picture": "picture",
    "wordprocessingShape": "shape",
    "wordprocessingCanvas": "canvas",
    "wordprocessingGroup": "group",
}


# --- package helpers ----------------------------------------------------------------

def _parts(ooxml: str) -> Dict[str, ET.Element]:
    root = ET.fromstring(ooxml)
    if root.tag == W + "document":
        return {"/word/document.xml": root}
    parts = {}
    for part in root.iter(PKG + "part"):
        data = part.find(PKG + "xmlData")
        if data is not None and len(data):
            parts[part.get(PKG + "name")] = data[0]
    return parts


def _text(el: ET.Element) -> str:
    out = []
    for node in el.iter():
        if node.tag == W + "t" and node.text:
            out.append(node.text)
        elif node.tag in (W + "tab",):
            out.append("\t")
        elif node.tag in (W + "br", W + "cr"):
            out.append("\n")
    return "".join(out)


def _val(el: Optional[ET.Element], tag: str) -> Optional[str]:
    if el is None:
        return None
    child = el.find(W + tag)
    return None if child is None else child.get(W + "val")


def _fill(tc_pr: Optional[ET.Element]) -> Optional[str]:
    if tc_pr is None:
        return None
    shd = tc_pr.find(W + "shd")
    fill = None if shd is None else shd.get(W + "fill")
    if not fill or fill.lower() in ("auto", "ffffff"):
        return None
    return fill.upper()


# --- styles -------------------------------------------------------------------------

def _style_names(parts: Dict[str, ET.Element]) -> Dict[str, str]:
    styles = parts.get("/word/styles.xml")
    names = {}
    if styles is not None:
        for st in styles.findall(W + "style"):
            sid = st.get(W + "styleId")
            name = _val(st, "name")
            if sid and name:
                names[sid] = name
    return names


def _heading_level(style_name: str, ppr: Optional[ET.Element]) -> Optional[int]:
    # styles.xml stores built-in names in English ("heading 1") even in Italian Word
    m = re.match(r"^(heading|titolo|titre|überschrift)\s*(\d)$", (style_name or "").strip(), re.IGNORECASE)
    if m:
        return int(m.group(2))
    lvl = _val(ppr, "outlineLvl")
    if lvl is not None and lvl.isdigit() and int(lvl) < 9:
        return int(lvl) + 1
    return None


# --- tables -------------------------------------------------------------------------

def _row_cells(tr: ET.Element, ptr: List[int]) -> List[Dict[str, Any]]:
    cells, gc = [], 0
    trpr = tr.find(W + "trPr")
    before = _val(trpr, "gridBefore")
    if before and before.isdigit():
        gc = int(before)
    for ci, tc in enumerate(tr.findall(W + "tc")):
        pr = tc.find(W + "tcPr")
        span = int(_val(pr, "gridSpan") or 1)
        vm_el = None if pr is None else pr.find(W + "vMerge")
        vmerge = None if vm_el is None else (vm_el.get(W + "val") or "continue")
        text = _text(tc).strip()
        bold = any(r.find(W + "rPr/" + W + "b") is not None for r in tc.iter(W + "r"))
        first_p = ptr[0]
        ptr[0] += len(list(tc.iter(W + "p")))
        cells.append({
            "ci": ci, "gc": gc, "span": span, "vmerge": vmerge, "fill": _fill(pr),
            "text": text, "bold": bold, "p": first_p,
        })
        gc += span
    return cells


def _is_band(cells, grid_cols) -> bool:
    live = [c for c in cells if c["vmerge"] != "continue"]
    return len(live) == 1 and live[0]["text"] and (live[0]["fill"] or live[0]["span"] >= max(2, grid_cols - 1))


def _is_spacer(cells) -> bool:
    return all(not c["text"] and not c["fill"] for c in cells) and len(cells) == 1


def _is_header(cells) -> bool:
    live = [c for c in cells if c["vmerge"] != "continue"]
    if len(cells) < 2 or not live:
        return False
    return all(c["fill"] and c["text"] for c in live)


def _covering(cells, gc):
    for c in cells:
        if c["gc"] <= gc < c["gc"] + c["span"]:
            return c
    return None


def _header_label(header_rows, gc, span) -> str:
    parts = []
    for row in header_rows:
        c = _covering(row, gc)
        if c is None:
            continue
        if c["vmerge"] == "continue":
            continue
        if c["text"] and c["text"] not in parts:
            # a sub-header narrower than the data cell is not the cell's own label
            if c["span"] < span and row is not header_rows[0]:
                continue
            parts.append(c["text"])
    return SEP.join(parts)


def _analyze_table(tbl: ET.Element, t_index: int, ptr: List[int], prev: Optional[Dict]) -> Dict[str, Any]:
    grid = tbl.find(W + "tblGrid")
    grid_cols = len(grid.findall(W + "gridCol")) if grid is not None else 0
    rows = [_row_cells(tr, ptr) for tr in tbl.findall(W + "tr")]
    grid_cols = max([grid_cols] + [sum(c["span"] for c in r) for r in rows] or [0])

    title = ""
    if rows:
        title = next((c["text"] for c in rows[0] if c["fill"] and c["text"]), "") or (rows[0][0]["text"] if rows[0] else "")

    leading_bands = 0
    for r in rows:
        if _is_band(r, grid_cols):
            leading_bands += 1
        else:
            break
    prefix: List[str] = [rows[0][0]["text"]] if leading_bands >= 2 else []

    fields: List[Dict[str, Any]] = []
    section: List[str] = []
    last_was_band = False
    header_group: List[List[Dict]] = []
    header_done = False
    record_no = 0
    empty_records = 0
    vm_origin: Dict[int, str] = {}
    columns: List[str] = []
    merged = 0
    kinds = Counter()

    for r_i, cells in enumerate(rows):
        merged += sum(1 for c in cells if c["span"] > 1 or c["vmerge"])
        for c in cells:
            if c["vmerge"] == "restart":
                vm_origin[c["gc"]] = c["text"]
        if _is_spacer(cells):
            last_was_band = False
            continue
        if _is_band(cells, grid_cols):
            text = cells[0]["text"]
            if prefix and r_i == 0:
                last_was_band = True
                continue
            section = (section + [text]) if last_was_band and section else [text]
            last_was_band = True
            header_group, header_done, record_no = [], False, 0
            continue
        last_was_band = False

        if _is_header(cells):
            if header_done:
                header_group, record_no = [], 0
            header_group.append(cells)
            continue
        if header_group:
            header_done = True

        has_shaded_label = any(c["fill"] and (c["text"] or c["vmerge"] == "continue") for c in cells)
        empties = [c for c in cells if not c["fill"] and c["vmerge"] != "continue"
                   and (not c["text"] or PLACEHOLDER_RE.match(c["text"]))]
        base = prefix + section

        if header_group and not has_shaded_label:
            first = cells[0]
            row_label = None
            if first["text"] and not PLACEHOLDER_RE.match(first["text"]) and first not in empties:
                row_label = first["text"]
            else:
                record_no += 1
                row_label = f"#{record_no}"
            if not columns:
                columns = [_header_label(header_group, c["gc"], c["span"]) for c in cells]
            if row_label.startswith("#") and all(not c["text"] for c in cells):
                empty_records += 1
            for c in empties:
                col = _header_label(header_group, c["gc"], c["span"]) or f"col {c['ci'] + 1}"
                f = _field(base + [row_label, col], t_index, r_i, c)
                f.update(kind="record", group=SEP.join(base), record=record_no if row_label.startswith("#") else row_label, column=col)
                fields.append(f)
                kinds["record"] += 1
            continue

        # form row: label is the nearest text cell on the left (or the vertical-merge origin)
        for c in empties:
            label = ""
            for left in reversed([x for x in cells if x["gc"] < c["gc"]]):
                if left["text"] and not PLACEHOLDER_RE.match(left["text"]):
                    label = left["text"]
                    break
                if left["vmerge"] == "continue" and vm_origin.get(left["gc"]):
                    label = vm_origin[left["gc"]]
                    break
            if not label and r_i > 0:
                above = _covering(rows[r_i - 1], c["gc"])
                if above is not None and above["text"] and above["fill"]:
                    label = above["text"]
            if not label:
                continue
            f = _field(base + [label], t_index, r_i, c)
            f.update(kind="form", group=SEP.join(base), record=None, column=label)
            fields.append(f)
            kinds["form"] += 1

    if prev and prefix == [] and prev.get("prefix") and leading_bands == 1 and columns and prev.get("columns"):
        a, b = set(columns), set(prev["columns"])
        if len(a & b) >= 0.7 * max(len(a), 1):
            prefix = prev["prefix"]
            for f in fields:
                f["path"] = SEP.join(prefix + f["path"].split(SEP))
                f["group"] = SEP.join(prefix + ([f["group"]] if f["group"] else []))

    if not fields:
        kind = "layout" if all(not c["text"] for r in rows for c in r) else "data"
    else:
        kind = "mixed" if kinds["record"] and kinds["form"] else ("record_grid" if kinds["record"] else "form")

    return {
        "index": t_index, "title": title, "kind": kind, "rows": len(rows), "grid_cols": grid_cols,
        "merged_cells": merged, "columns": columns, "empty_records": empty_records,
        "sections": sorted({f["path"].split(SEP)[len(prefix)] for f in fields if SEP in f["path"]}, key=str),
        "prefix": prefix, "fields": fields,
    }


def _field(path_parts: List[str], t: int, r: int, cell: Dict[str, Any]) -> Dict[str, Any]:
    text = cell["text"]
    return {
        "path": SEP.join(p for p in path_parts if p),
        "loc": {"table": t, "row": r, "cell": cell["ci"], "paragraph": cell["p"]},
        "empty": not text,
        "placeholder": text or None,
    }


# --- snapshot -----------------------------------------------------------------------

def build_snapshot(ooxml: str, extras: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Parse the document OOXML into a structured snapshot (see module docstring)."""
    parts = _parts(ooxml)
    doc = parts.get("/word/document.xml")
    body = doc.find(W + "body") if doc is not None else None
    style_names = _style_names(parts)
    extras = extras or {}

    paragraphs, tables, outline = [], [], []
    style_counter, runs_total, runs_direct = Counter(), 0, 0
    ptr = [0]
    prev_table = None
    all_text = []

    def visit_paragraph(p: ET.Element):
        ppr = p.find(W + "pPr")
        sid = _val(ppr, "pStyle") or "Normal"
        name = style_names.get(sid, sid)
        text = _text(p).strip()
        idx = ptr[0]
        ptr[0] += 1
        if text:
            all_text.append(text)
        level = _heading_level(name, ppr)
        item = {"index": idx, "text": text, "style": _display_style(name), "align": _val(ppr, "jc"),
                "numbered": ppr is not None and ppr.find(W + "numPr") is not None}
        if level:
            item["heading"] = level
            if text:
                outline.append({"level": level, "text": text, "paragraph": idx})
        paragraphs.append(item)

    if body is not None:
        for child in body:
            if child.tag == W + "p":
                visit_paragraph(child)
            elif child.tag == W + "tbl":
                start = ptr[0]
                info = _analyze_table(child, len(tables), ptr, prev_table)
                info["first_paragraph"] = start
                tables.append(info)
                prev_table = info
                for tc_text in (_text(tc).strip() for tc in child.iter(W + "tc")):
                    if tc_text:
                        all_text.append(tc_text)
            elif child.tag == W + "sdt":
                content = child.find(W + "sdtContent")
                for p in (content.findall(W + "p") if content is not None else []):
                    visit_paragraph(p)

    fields = _dedupe([f for t in tables for f in t.pop("fields")])
    for t in tables:
        t.pop("prefix", None)

    root = doc if doc is not None else ET.Element("x")
    for p in root.iter(W + "p"):
        if _text(p).strip():
            sid = _val(p.find(W + "pPr"), "pStyle") or "Normal"
            style_counter[_display_style(style_names.get(sid, sid))] += 1
    for r in root.iter(W + "r"):
        if r.find(W + "t") is None:
            continue
        runs_total += 1
        rpr = r.find(W + "rPr")
        if rpr is not None and any(rpr.find(W + t) is not None for t in DIRECT_RUN_PROPS):
            runs_direct += 1
    content_controls = []
    for sdt in root.iter(W + "sdt"):
        pr = sdt.find(W + "sdtPr")
        kind = next((k for k in ("date", "dropDownList", "comboBox", "picture", "checkbox", "richText", "text")
                     if pr is not None and any(ch.tag.endswith("}" + k) for ch in pr)), "richText")
        content_controls.append({
            "tag": _val(pr, "tag"), "title": _val(pr, "alias"), "type": kind,
            "text": _text(sdt.find(W + "sdtContent") if sdt.find(W + "sdtContent") is not None else sdt).strip()[:120],
        })

    field_codes = Counter()
    for instr in root.iter(W + "instrText"):
        word = (instr.text or "").strip().split(" ")[0].upper()
        if word:
            field_codes[word] += 1
    for fs in root.iter(W + "fldSimple"):
        word = (fs.get(W + "instr") or "").strip().split(" ")[0].upper()
        if word:
            field_codes[word] += 1

    visuals = Counter()
    for gd in root.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}graphicData"):
        uri = gd.get("uri") or ""
        visuals[next((v for k, v in GRAPHIC_KINDS.items() if uri.rstrip("/").endswith(k) or k in uri), "other")] += 1
    for _ in root.iter("{urn:schemas-microsoft-com:vml}shape"):
        visuals["legacy_shape"] += 1

    settings = parts.get("/word/settings.xml")
    review = {
        "track_changes": settings is not None and settings.find(W + "trackRevisions") is not None,
        "insertions": sum(1 for _ in root.iter(W + "ins")),
        "deletions": sum(1 for _ in root.iter(W + "del")),
        "comments": sum(1 for _ in root.iter(W + "commentRangeStart")),
    }

    sections = []
    for sect in root.iter(W + "sectPr"):
        pg = sect.find(W + "pgSz")
        mar = sect.find(W + "pgMar")
        sections.append({
            "orientation": "landscape" if pg is not None and pg.get(W + "orient") == "landscape" else "portrait",
            "margins_cm": None if mar is None else {
                k: round(int(mar.get(W + k, "0")) / 567, 2) for k in ("top", "bottom", "left", "right")},
            "title_page": sect.find(W + "titlePg") is not None,
        })

    joined = " ".join(all_text)
    snapshot = {
        "doc_type": "blank",
        "language": _language(joined),
        "stats": {
            "paragraphs": len(paragraphs),
            "non_empty_paragraphs": sum(1 for p in paragraphs if p["text"]),
            "tables": len(tables),
            "words": len(joined.split()),
            "fields_total": len(fields),
            "fields_empty": sum(1 for f in fields if f["empty"]),
        },
        "styles": {
            "used": dict(style_counter.most_common()),
            "direct_formatting": runs_total > 0 and runs_direct / runs_total >= 0.3,
        },
        "outline": outline,
        "paragraphs": paragraphs,
        "tables": tables,
        "fields": fields,
        "content_controls": content_controls,
        "field_codes": dict(field_codes),
        "visuals": dict(visuals),
        "review": review,
        "sections": sections,
        "placeholders": [p["text"] for p in paragraphs if p["text"] and PLACEHOLDER_RE.match(p["text"])][:30],
        "selection": extras.get("selection"),
        "headers_footers": extras.get("headers_footers") or [],
        "document": extras.get("document") or {},
    }
    snapshot["doc_type"] = _classify(snapshot, joined)
    return snapshot


def _display_style(name: str) -> str:
    return {"normal": "Normal"}.get((name or "").lower(), name)


def _dedupe(fields: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    seen = Counter()
    for f in fields:
        seen[f["path"]] += 1
        if seen[f["path"]] > 1:
            f["path"] = f"{f['path']} #{seen[f['path']]}"
    return fields


def _language(text: str) -> str:
    words = re.findall(r"[a-zàèéìòù]+", text.lower())
    scores = {lang: sum(1 for w in words if w in sw) for lang, sw in STOPWORDS.items()}
    best = max(scores, key=scores.get)
    return best if scores[best] > 0 else "unknown"


def _classify(s: Dict[str, Any], text: str) -> str:
    st = s["stats"]
    if st["words"] == 0 and not st["tables"]:
        return "blank"
    if st["fields_empty"] >= 5 or len(s["content_controls"]) >= 3 or len(s["placeholders"]) >= 3:
        return "form_template"
    low = text.lower()
    if re.search(r"\b(contratto|tra le parti|articolo \d|agreement|whereas|clause)\b", low):
        return "contract"
    if re.search(r"\b(gentile|egregio|cordiali saluti|dear |sincerely|distinti saluti)\b", low):
        return "letter"
    if re.search(r"\b(verbale|ordine del giorno|minutes|agenda|partecipanti)\b", low):
        return "minutes"
    if len(s["outline"]) >= 3 or st["words"] > 800:
        return "report"
    return "document"


# --- Document Map -------------------------------------------------------------------

def render_map(s: Dict[str, Any], budget_chars: int = 24000) -> str:
    """Compact text view of the snapshot for the model. Structure is kept; detail shrinks to fit."""
    st = s["stats"]
    head = [
        "=== DOCUMENT MAP (auto-analysed before your turn) ===",
        f"Type: {s['doc_type']} · Language: {s['language']} · Words: {st['words']} · "
        f"Paragraphs: {st['paragraphs']} · Tables: {st['tables']}",
    ]
    doc = s.get("document") or {}
    if doc.get("name"):
        head.append(f"File: {doc['name']}")
    used = list(s["styles"]["used"])[:8]
    style_line = f"Styles in use: {', '.join(used) or 'none'}"
    if s["styles"]["direct_formatting"]:
        style_line += " · DIRECT FORMATTING is used: preserve fonts, bold, sizes and shading when editing"
    head.append(style_line)
    for i, sec in enumerate(s["sections"]):
        head.append(f"Section {i + 1}: {sec['orientation']}" + (" · different first page" if sec["title_page"] else ""))
    for hf in s.get("headers_footers") or []:
        head.append(f"Header/footer: {hf}")
    if s["review"]["track_changes"] or s["review"]["insertions"] or s["review"]["comments"]:
        r = s["review"]
        head.append(f"Review: track changes {'ON' if r['track_changes'] else 'off'} · "
                    f"{r['insertions']} insertions · {r['deletions']} deletions · {r['comments']} comments")
    if s["field_codes"]:
        head.append("Field codes: " + ", ".join(f"{k}×{v}" for k, v in s["field_codes"].items()))
    if s["visuals"]:
        head.append("Visuals: " + ", ".join(f"{k}×{v}" for k, v in s["visuals"].items()))
    if s["content_controls"]:
        head.append("Content controls: " + "; ".join(
            f"{c['title'] or c['tag'] or c['type']}={c['text'] or '(empty)'}" for c in s["content_controls"][:20]))
    if s.get("selection"):
        sel = s["selection"]
        head.append(f"SELECTION: {sel if isinstance(sel, str) else sel.get('text', '')!r}"[:600])

    head.append(f"Fields to fill: {st['fields_total']} ({st['fields_empty']} empty)")
    tables = ["Tables:"] + [
        f"  T{t['index']} \"{t['title'][:50]}\" {t['kind']} {t['rows']}x{t['grid_cols']} merges={t['merged_cells']}"
        + (f" columns=[{', '.join(c for c in t['columns'] if c)[:160]}]" if t["columns"] and t["kind"] == "data" else "")
        + (f" empty_records={t['empty_records']}" if t["empty_records"] else "")
        for t in s["tables"]
    ] if s["tables"] else []

    outline = ["Outline:"] + [f"  {'  ' * (o['level'] - 1)}H{o['level']} {o['text'][:90]} [p{o['paragraph']}]"
                              for o in s["outline"]] if s["outline"] else []

    field_lines: List[str] = []
    grids: Dict[tuple, List[Dict[str, Any]]] = {}
    for f in s["fields"]:
        if f.get("kind") == "record" and isinstance(f.get("record"), int):
            key = (f["loc"]["table"], f["group"])
            if key not in grids:
                grids[key] = []
                field_lines.append(key)  # placeholder, replaced below
            grids[key].append(f)
        else:
            field_lines.append(
                f"  - {f['path']} [T{f['loc']['table']} r{f['loc']['row']} c{f['loc']['cell']}]"
                + ("" if f["empty"] else f" = {f['placeholder']!r}"))
    for i, line in enumerate(field_lines):
        if isinstance(line, tuple):
            fs = grids[line]
            recs = sorted({f["record"] for f in fs})
            rows = sorted({f["loc"]["row"] for f in fs})
            cols = list(dict.fromkeys(f["column"] for f in fs))
            field_lines[i] = (f"  - {line[1] + SEP if line[1] else ''}records #{recs[0]}-#{recs[-1]} "
                              f"× [{', '.join(cols)}] [T{line[0]} rows r{rows[0]}-r{rows[-1]}, one cell per column]")
    body_lines = [f"  [p{p['index']}]{'(' + p['style'] + ')' if p['style'] != 'Normal' else ''} {p['text'][:200]}"
                  for p in s["paragraphs"] if p["text"]]

    fixed = "\n".join(head + tables + outline)
    tail_note = "(use word_read_* tools for anything not shown)"
    remaining = budget_chars - len(fixed) - len(tail_note) - 64
    if remaining < 0:
        return (fixed[: budget_chars - len(tail_note) - 2] + "\n" + tail_note)[:budget_chars]

    def fit(title: str, lines: List[str], room: int) -> List[str]:
        out, used_chars = [], len(title) + 1
        for i, line in enumerate(lines):
            if used_chars + len(line) + 1 > room - 40:
                out.append(f"  ... and {len(lines) - i} more")
                break
            out.append(line)
            used_chars += len(line) + 1
        return [title] + out if lines else []

    fields_part = fit("Field paths (fill these; loc = table/row/cell):", field_lines, int(remaining * 0.65))
    left = remaining - sum(len(x) + 1 for x in fields_part)
    body_part = fit("Body text:", body_lines, left)
    text = "\n".join([fixed] + fields_part + body_part + [tail_note])
    return text[:budget_chars]
