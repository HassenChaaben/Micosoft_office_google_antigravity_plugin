"""
Blocks -> OOXML builder.

Turns a JSON list of "blocks" (title, headings, paragraphs with runs, nested lists,
tables with merged cells, captions, TOC, quotes, page breaks) into ONE flat-OPC
OOXML package string. The taskpane inserts it with a single Office.js call
`range.insertOoxml(ooxml, location)` (COM: `Range.InsertXML`).

Styles are referenced by their language-independent built-in style IDs (Heading1,
Title, Caption, ...). The package carries a minimal /word/styles.xml that defines
them with their English built-in names ("heading 1", ...), so Word maps them onto
the localized built-ins of the target document (e.g. Italian "Titolo 1").
"""

import re
import zlib
from typing import Any, Dict, List, Optional, Tuple

BLOCK_TYPES = ("heading", "title", "subtitle", "paragraph", "bullets", "numbered", "table",
               "page_break", "caption", "quote", "toc")
ALIGN = {"left": "left", "center": "center", "right": "right", "justify": "both"}
VALIGN = {"top": "top", "center": "center", "bottom": "bottom"}
HIGHLIGHTS = {"yellow", "green", "cyan", "magenta", "blue", "red", "darkBlue", "darkCyan", "darkGreen",
              "darkMagenta", "darkRed", "darkYellow", "darkGray", "lightGray", "black", "white", "none"}
HEX_RE = re.compile(r"^#?[0-9A-Fa-f]{6}$")
LEVELS_RE = re.compile(r"^([1-9])-([1-9])$")
INVALID_XML_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")
MAX_LIST_DEPTH = 3
LOCATION_RE = re.compile(r"^(after|before|replace):(.*)$", re.I | re.S)
TABLE_WIDTH_DXA = 9638  # A4 with 2 cm margins; the table itself is 100 % (pct 5000)
SDT_ALIAS = "Gemini Antigravity"
TOC_PLACEHOLDER = "Aggiorna il sommario (F9)"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_NS = "http://schemas.microsoft.com/office/2006/xmlPackage"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# built-in style names (as written by the model or by humans) -> style IDs
STYLE_IDS = {
    "normal": "Normal", "title": "Title", "subtitle": "Subtitle", "caption": "Caption", "quote": "Quote",
    "intense quote": "IntenseQuote", "list paragraph": "ListParagraph", "list bullet": "ListBullet",
    "list number": "ListNumber", "no spacing": "NoSpacing", "toc heading": "TOCHeading",
    "table grid": "TableGrid",
    **{f"heading {i}": f"Heading{i}" for i in range(1, 7)},
    **{f"toc {i}": f"TOC{i}" for i in range(1, 4)},
}


# --- XML helpers ----------------------------------------------------------------------

def _clean(s: Any) -> str:
    return INVALID_XML_RE.sub("", str(s))


def _esc(s: Any) -> str:
    return _clean(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _attr(s: Any) -> str:
    return _esc(s).replace('"', "&quot;")


def _hex(color: str) -> str:
    return color.lstrip("#").upper()


def _style_id(name: str) -> str:
    return STYLE_IDS.get(str(name).strip().lower(), str(name).strip())


# --- validation ---------------------------------------------------------------------

def _is_int(v: Any) -> bool:
    return isinstance(v, int) and not isinstance(v, bool)


def _check_inline(obj: Dict[str, Any], where: str, errors: List[str]) -> None:
    """An inline container needs `text` (string) or `runs` (list of run objects)."""
    text, runs = obj.get("text"), obj.get("runs")
    if runs is None:
        if not isinstance(text, (str, int, float)) or isinstance(text, bool):
            errors.append(f"{where}: needs 'text' (string) or 'runs' (list of {{text, bold, ...}})")
        return
    if not isinstance(runs, list) or not runs:
        errors.append(f"{where}: 'runs' must be a non-empty list of run objects")
        return
    for i, run in enumerate(runs):
        rw = f"{where}.runs[{i}]"
        if not isinstance(run, dict):
            errors.append(f"{rw}: a run must be an object like {{\"text\": \"...\", \"bold\": true}}")
            continue
        if not isinstance(run.get("text"), (str, int, float)) or isinstance(run.get("text"), bool):
            errors.append(f"{rw}: run needs a 'text' string")
        if run.get("color") is not None and not (isinstance(run["color"], str) and HEX_RE.match(run["color"])):
            errors.append(f"{rw}: color must be \"#RRGGBB\" (got {run['color']!r})")
        if run.get("size") is not None and (not isinstance(run["size"], (int, float)) or isinstance(run["size"], bool)
                                            or not 1 <= run["size"] <= 1638):
            errors.append(f"{rw}: size must be a number of points between 1 and 1638")
        if run.get("highlight") is not None and run["highlight"] not in HIGHLIGHTS:
            errors.append(f"{rw}: highlight must be one of {sorted(HIGHLIGHTS)} (got {run['highlight']!r})")
        if run.get("font") is not None and not isinstance(run["font"], str):
            errors.append(f"{rw}: font must be a font name string")


def _check_align(obj: Dict[str, Any], where: str, errors: List[str], key: str = "align", allowed=ALIGN) -> None:
    if obj.get(key) is not None and obj[key] not in allowed:
        errors.append(f"{where}: {key} must be one of {list(allowed)} (got {obj[key]!r})")


def _check_items(items: Any, where: str, depth: int, errors: List[str]) -> None:
    if not isinstance(items, list) or not items:
        errors.append(f"{where}: 'items' must be a non-empty list of strings or {{text, children}} objects")
        return
    for i, item in enumerate(items):
        iw = f"{where}[{i}]"
        if isinstance(item, str):
            continue
        if not isinstance(item, dict):
            errors.append(f"{iw}: list item must be a string or an object {{text|runs, children}}")
            continue
        _check_inline(item, iw, errors)
        children = item.get("children")
        if children:
            if depth >= MAX_LIST_DEPTH:
                errors.append(f"{iw}: lists support at most {MAX_LIST_DEPTH} levels of nesting")
            else:
                _check_items(children, f"{iw}.children", depth + 1, errors)


def _layout_table(rows: List[List[Any]], ncols: int, where: str, errors: List[str]):
    """Place caller cells on the grid. Returns per-row lists of
    ("cell", cell, col, colspan, rowspan) and ("cont", origin, col, colspan) entries.
    Cells covered by a rowspan from above are inserted by the builder."""
    covered: Dict[Tuple[int, int], Tuple[Dict[str, Any], int, int]] = {}  # (r, c) -> (origin, origin_col, span)
    layout = []
    for r, row in enumerate(rows):
        entries, col = [], 0
        rw = f"{where} row {r + 1}"

        def flush_covered(limit: Optional[int] = None):
            nonlocal col
            while col < ncols and (r, col) in covered and (limit is None or col < limit):
                origin, ocol, span = covered[(r, col)]
                if ocol == col:
                    entries.append(("cont", origin, col, span))
                col += 1

        for ci, cell in enumerate(row):
            flush_covered()
            cell = cell if isinstance(cell, dict) else {"text": "" if cell is None else cell}
            cs = cell.get("colspan", 1) or 1
            rs = cell.get("rowspan", 1) or 1
            if col + cs > ncols:
                errors.append(f"{rw}: spans {col + cs} grid columns but the table has {ncols} "
                              f"(count colspan; cells covered by a rowspan from above are added "
                              f"automatically, do not repeat them)")
                return None
            if any((r, c) in covered for c in range(col, col + cs)):
                errors.append(f"{rw}, cell {ci + 1}: colspan overlaps a cell covered by a rowspan from above")
                return None
            if r + rs > len(rows):
                errors.append(f"{rw}, cell {ci + 1}: rowspan {rs} goes past the last row of the table")
                return None
            for rr in range(r + 1, r + rs):
                for c in range(col, col + cs):
                    covered[(rr, c)] = (cell, col, cs)
            entries.append(("cell", cell, col, cs, rs))
            col += cs
        flush_covered()
        if col != ncols:
            errors.append(f"{rw}: covers {col} grid columns but the table has {ncols} "
                          f"(check colspan/rowspan; cells covered by a rowspan from above are added automatically)")
            return None
        layout.append(entries)
    return layout


def _row_width(row: List[Any]) -> int:
    return sum((c.get("colspan", 1) or 1) if isinstance(c, dict) else 1 for c in row)


def _check_table(block: Dict[str, Any], where: str, errors: List[str]) -> None:
    rows = block.get("rows")
    if not isinstance(rows, list) or not rows or not all(isinstance(r, list) and r for r in rows):
        errors.append(f"{where}: 'rows' must be a non-empty list of non-empty rows (lists of cells)")
        return
    n_err = len(errors)
    for r, row in enumerate(rows):
        for ci, cell in enumerate(row):
            cw = f"{where} row {r + 1}, cell {ci + 1}"
            if cell is None or isinstance(cell, (str, int, float)) and not isinstance(cell, bool):
                continue
            if not isinstance(cell, dict):
                errors.append(f"{cw}: a cell must be a string or an object {{text|runs, colspan, rowspan, ...}}")
                continue
            if "text" in cell or "runs" in cell:  # a cell with neither is simply empty
                _check_inline(cell, cw, errors)
            for key in ("colspan", "rowspan"):
                if cell.get(key) is not None and (not _is_int(cell[key]) or cell[key] < 1):
                    errors.append(f"{cw}: {key} must be an integer >= 1")
            if cell.get("shading") is not None and not (isinstance(cell["shading"], str) and HEX_RE.match(cell["shading"])):
                errors.append(f"{cw}: shading must be \"#RRGGBB\"")
            _check_align(cell, cw, errors)
            _check_align(cell, cw, errors, "valign", VALIGN)
    hr = block.get("header_rows", 1)
    if not _is_int(hr) or not 0 <= hr <= len(rows):
        errors.append(f"{where}: header_rows must be an integer between 0 and the number of rows ({len(rows)})")
    if block.get("header_shading") is not None and not (
            isinstance(block["header_shading"], str) and HEX_RE.match(block["header_shading"])):
        errors.append(f"{where}: header_shading must be \"#RRGGBB\" (got {block['header_shading']!r})")
    columns = block.get("columns")
    if len(errors) > n_err:
        return
    ncols = _row_width(rows[0])
    if columns is not None:
        if not isinstance(columns, list) or not columns:
            errors.append(f"{where}: columns must be a list of {{width_pct}} objects")
            return
        for i, c in enumerate(columns):
            if not isinstance(c, dict) or not isinstance(c.get("width_pct", 1), (int, float)) or c.get("width_pct", 1) <= 0:
                errors.append(f"{where}: columns[{i}] must be {{\"width_pct\": number > 0}}")
        if len(columns) != ncols:
            errors.append(f"{where}: {len(columns)} columns given but row 1 spans {ncols} grid columns")
            return
    _layout_table(rows, ncols, where, errors)


def validate_blocks(blocks: Any) -> List[str]:
    """Return human-readable errors (empty list when the blocks are valid)."""
    if not isinstance(blocks, list):
        return ["blocks must be a list of block objects like {\"type\": \"paragraph\", \"text\": \"...\"}"]
    errors: List[str] = []
    for i, block in enumerate(blocks):
        where = f"blocks[{i}]"
        if not isinstance(block, dict):
            errors.append(f"{where}: each block must be an object with a 'type'")
            continue
        t = block.get("type")
        if t not in BLOCK_TYPES:
            errors.append(f"{where}: unknown type {t!r}; valid types are {', '.join(BLOCK_TYPES)}")
            continue
        where = f"{where} ({t})"
        if t == "heading":
            lvl = block.get("level")
            if not _is_int(lvl) or not 1 <= lvl <= 6:
                errors.append(f"{where}: level must be an integer 1-6 (got {lvl!r})")
            _check_inline(block, where, errors)
        elif t in ("title", "subtitle", "quote"):
            _check_inline(block, where, errors)
            _check_align(block, where, errors)
        elif t == "paragraph":
            _check_inline(block, where, errors)
            _check_align(block, where, errors)
            if block.get("style") is not None and not isinstance(block["style"], str):
                errors.append(f"{where}: style must be a style name or styleId string")
            sa = block.get("space_after_pt")
            if sa is not None and (not isinstance(sa, (int, float)) or isinstance(sa, bool) or sa < 0):
                errors.append(f"{where}: space_after_pt must be a number >= 0")
        elif t in ("bullets", "numbered"):
            _check_items(block.get("items"), f"{where}.items", 1, errors)
        elif t == "table":
            _check_table(block, where, errors)
        elif t == "caption":
            label = block.get("label")
            if not isinstance(label, str) or not label.strip():
                errors.append(f"{where}: label is required (e.g. \"Tabella\" or \"Figura\")")
            if not isinstance(block.get("text"), str):
                errors.append(f"{where}: text is required")
        elif t == "toc":
            m = LEVELS_RE.match(str(block.get("levels", "1-3")))
            if not m or int(m.group(1)) > int(m.group(2)):
                errors.append(f"{where}: levels must look like \"1-3\" (got {block.get('levels')!r})")
    return errors


# --- runs and paragraphs ------------------------------------------------------------

def _text_runs(text: Any, rpr: str = "") -> str:
    """Runs for a text that may contain \\n (line break) and \\t (tab)."""
    out = []
    for li, line in enumerate(_clean(text).split("\n")):
        if li:
            out.append(f"<w:r>{rpr}<w:br/></w:r>")
        for ti, chunk in enumerate(line.split("\t")):
            if ti:
                out.append(f"<w:r>{rpr}<w:tab/></w:r>")
            if chunk:
                out.append(f'<w:r>{rpr}<w:t xml:space="preserve">{_esc(chunk)}</w:t></w:r>')
    return "".join(out)


def _rpr(run: Dict[str, Any], force_bold: bool = False) -> str:
    parts = []
    if run.get("font"):
        f = _attr(run["font"])
        parts.append(f'<w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>')
    if run.get("bold") or force_bold:
        parts.append("<w:b/><w:bCs/>")
    if run.get("italic"):
        parts.append("<w:i/><w:iCs/>")
    if run.get("color"):
        parts.append(f'<w:color w:val="{_hex(run["color"])}"/>')
    if run.get("size"):
        hp = int(round(float(run["size"]) * 2))
        parts.append(f'<w:sz w:val="{hp}"/><w:szCs w:val="{hp}"/>')
    if run.get("highlight"):
        parts.append(f'<w:highlight w:val="{run["highlight"]}"/>')
    if run.get("underline"):
        parts.append('<w:u w:val="single"/>')
    return f"<w:rPr>{''.join(parts)}</w:rPr>" if parts else ""


def _inline(obj: Dict[str, Any], force_bold: bool = False) -> str:
    if obj.get("runs"):
        return "".join(_text_runs(r.get("text", ""), _rpr(r, force_bold)) for r in obj["runs"])
    return _text_runs(obj.get("text", ""), _rpr({}, force_bold))


def _ppr(style: Optional[str] = None, num: Optional[Tuple[str, int]] = None, align: Optional[str] = None,
         space_after_pt: Optional[float] = None) -> str:
    parts = []
    if style:
        parts.append(f'<w:pStyle w:val="{_attr(style)}"/>')
    if num:
        parts.append(f'<w:numPr><w:ilvl w:val="{num[1]}"/><w:numId w:val="{num[0]}"/></w:numPr>')
    if space_after_pt is not None:
        parts.append(f'<w:spacing w:after="{int(round(float(space_after_pt) * 20))}"/>')
    if align:
        parts.append(f'<w:jc w:val="{ALIGN[align]}"/>')
    return f"<w:pPr>{''.join(parts)}</w:pPr>" if parts else ""


def _para(obj: Dict[str, Any], style: Optional[str] = None, **kw) -> str:
    return f"<w:p>{_ppr(style, **kw)}{_inline(obj)}</w:p>"


# --- tables -----------------------------------------------------------------------------

def _borders(on: bool) -> str:
    val = 'w:val="single" w:sz="4" w:space="0" w:color="auto"' if on else 'w:val="nil"'
    return "<w:tblBorders>" + "".join(
        f"<w:{side} {val}/>" for side in ("top", "left", "bottom", "right", "insideH", "insideV")) + "</w:tblBorders>"


def _table(block: Dict[str, Any]) -> str:
    rows = block["rows"]
    ncols = _row_width(rows[0])
    layout = _layout_table(rows, ncols, "table", [])
    columns = block.get("columns") or [{"width_pct": 1}] * ncols
    weights = [float(c.get("width_pct", 1)) for c in columns]
    total = sum(weights)
    widths = [int(TABLE_WIDTH_DXA * w / total) for w in weights]
    header_rows = block.get("header_rows", 1)
    repeat = block.get("repeat_header", True)
    header_fill = _hex(block.get("header_shading") or "D9E2F3")
    style = _style_id(block.get("style") or "TableGrid")

    out = [f'<w:tbl><w:tblPr><w:tblStyle w:val="{_attr(style)}"/><w:tblW w:w="5000" w:type="pct"/>',
           _borders(block.get("borders", True)),
           '<w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" '
           'w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>',
           "".join(f'<w:gridCol w:w="{w}"/>' for w in widths), "</w:tblGrid>"]
    for r, entries in enumerate(layout):
        is_header = r < header_rows
        out.append("<w:tr>")
        if is_header and repeat:
            out.append("<w:trPr><w:tblHeader/></w:trPr>")
        for entry in entries:
            cell, col, span = entry[1], entry[2], entry[3]
            cont = entry[0] == "cont"
            rs = 1 if cont else entry[4]
            pr = [f'<w:tcW w:w="{sum(widths[col:col + span])}" w:type="dxa"/>']
            if span > 1:
                pr.append(f'<w:gridSpan w:val="{span}"/>')
            if cont:
                pr.append("<w:vMerge/>")
            elif rs > 1:
                pr.append('<w:vMerge w:val="restart"/>')
            fill = _hex(cell["shading"]) if cell.get("shading") else (header_fill if is_header else None)
            if fill:
                pr.append(f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>')
            if cell.get("valign"):
                pr.append(f'<w:vAlign w:val="{VALIGN[cell["valign"]]}"/>')
            bold = cell.get("bold", is_header)
            content = "" if cont else _inline(cell, force_bold=bool(bold))
            out.append(f"<w:tc><w:tcPr>{''.join(pr)}</w:tcPr><w:p>{_ppr(align=cell.get('align'))}{content}</w:p></w:tc>")
        out.append("</w:tr>")
    out.append("</w:tbl>")
    return "".join(out)


# --- numbering ------------------------------------------------------------------------

BULLET_ABSTRACT, DECIMAL_ABSTRACT = 1, 2
BULLET_CHARS = ("•", "o", "▪")
BULLET_FONTS = ("Symbol", "Courier New", "Wingdings")
DECIMAL_FORMATS = (("decimal", "%1."), ("lowerLetter", "%2)"), ("lowerRoman", "%3."))


def _numbering_xml(nums: List[Tuple[int, int]]) -> str:
    def lvl(i: int, fmt: str, text: str, font: Optional[str] = None) -> str:
        left = 720 * (i + 1)
        rpr = f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:hint="default"/></w:rPr>' if font else ""
        return (f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
                f'<w:lvlText w:val="{_attr(text)}"/><w:lvlJc w:val="left"/>'
                f'<w:pPr><w:ind w:left="{left}" w:hanging="360"/></w:pPr>{rpr}</w:lvl>')

    # bullet glyphs: "•" in Symbol is rendered as ; use plain Unicode with a Unicode font instead
    bullets = "".join(lvl(i, "bullet", BULLET_CHARS[i], ("Arial", "Courier New", "Arial")[i]) for i in range(3))
    decimals = "".join(lvl(i, fmt, text) for i, (fmt, text) in enumerate(DECIMAL_FORMATS))
    num_xml = "".join(
        f'<w:num w:numId="{nid}"><w:abstractNumId w:val="{aid}"/>'
        + ("".join(f'<w:lvlOverride w:ilvl="{i}"><w:startOverride w:val="1"/></w:lvlOverride>' for i in range(3))
           if aid == DECIMAL_ABSTRACT else "")
        + "</w:num>" for nid, aid in nums)
    return (f'<w:numbering xmlns:w="{W_NS}">'
            f'<w:abstractNum w:abstractNumId="{BULLET_ABSTRACT}"><w:multiLevelType w:val="hybridMultilevel"/>{bullets}</w:abstractNum>'
            f'<w:abstractNum w:abstractNumId="{DECIMAL_ABSTRACT}"><w:multiLevelType w:val="hybridMultilevel"/>{decimals}</w:abstractNum>'
            f"{num_xml}</w:numbering>")


def _list(items: List[Any], num_id: int, style: str, level: int = 0) -> str:
    out = []
    for item in items:
        obj = item if isinstance(item, dict) else {"text": item}
        out.append(_para(obj, style, num=(str(num_id), level)))
        if isinstance(item, dict) and item.get("children"):
            out.append(_list(item["children"], num_id, style, level + 1))
    return "".join(out)


# --- styles ---------------------------------------------------------------------------

def _pstyle(sid: str, name: str, ppr: str = "", rpr: str = "", based: Optional[str] = "Normal",
            prio: int = 99, link_next: bool = True) -> str:
    return (f'<w:style w:type="paragraph" w:styleId="{sid}"><w:name w:val="{name}"/>'
            + (f'<w:basedOn w:val="{based}"/>' if based else "")
            + ('<w:next w:val="Normal"/>' if link_next else "")
            + f'<w:uiPriority w:val="{prio}"/><w:qFormat/>'
            + (f"<w:pPr>{ppr}</w:pPr>" if ppr else "") + (f"<w:rPr>{rpr}</w:rPr>" if rpr else "")
            + "</w:style>")


HEADING_SIZES = (32, 26, 24, 22, 22, 22)


def _styles_xml() -> str:
    styles = ['<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>']
    for i in range(1, 7):
        styles.append(_pstyle(
            f"Heading{i}", f"heading {i}",
            f'<w:keepNext/><w:keepLines/><w:spacing w:before="{240 if i == 1 else 120}" w:after="60"/>'
            f'<w:outlineLvl w:val="{i - 1}"/>',
            f'<w:b/><w:bCs/><w:sz w:val="{HEADING_SIZES[i - 1]}"/><w:szCs w:val="{HEADING_SIZES[i - 1]}"/>', prio=9))
    styles += [
        _pstyle("Title", "Title", '<w:spacing w:after="120"/><w:contextualSpacing/>',
                '<w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/>', prio=10),
        _pstyle("Subtitle", "Subtitle", '<w:spacing w:after="160"/>',
                '<w:color w:val="5A5A5A"/><w:sz w:val="30"/><w:szCs w:val="30"/>', prio=11),
        _pstyle("Caption", "caption", '<w:spacing w:after="200"/>',
                '<w:i/><w:iCs/><w:sz w:val="18"/><w:szCs w:val="18"/>', prio=35),
        _pstyle("Quote", "Quote", '<w:spacing w:before="200" w:after="160"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/>',
                '<w:i/><w:iCs/><w:color w:val="404040"/>', prio=29),
        _pstyle("IntenseQuote", "Intense Quote",
                '<w:pBdr><w:top w:val="single" w:sz="4" w:space="10" w:color="4472C4"/>'
                '<w:bottom w:val="single" w:sz="4" w:space="10" w:color="4472C4"/></w:pBdr>'
                '<w:spacing w:before="360" w:after="360"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/>',
                '<w:i/><w:iCs/><w:color w:val="4472C4"/>', prio=30),
        _pstyle("NoSpacing", "No Spacing", '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>', prio=1),
        _pstyle("ListParagraph", "List Paragraph", '<w:ind w:left="720"/><w:contextualSpacing/>', prio=34),
        _pstyle("ListBullet", "List Bullet", '<w:contextualSpacing/>', prio=99),
        _pstyle("ListNumber", "List Number", '<w:contextualSpacing/>', prio=99),
        _pstyle("TOCHeading", "TOC Heading", '<w:outlineLvl w:val="9"/>', based="Heading1", prio=39),
    ]
    for i in range(1, 4):
        styles.append(_pstyle(f"TOC{i}", f"toc {i}", f'<w:spacing w:after="100"/><w:ind w:left="{(i - 1) * 220}"/>', prio=39))
    styles.append(
        '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/>'
        '<w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
        f'<w:tblPr>{_borders(True)}</w:tblPr></w:style>')
    styles.append(
        '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>'
        '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>'
        '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>'
        '<w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>')
    return f'<w:styles xmlns:w="{W_NS}">{"".join(styles)}</w:styles>'


# --- fields -----------------------------------------------------------------------------

def _seq_name(label: str) -> str:
    return re.sub(r"\W+", "_", label.strip()) or "Seq"


def _caption(block: Dict[str, Any], counters: Dict[str, int]) -> str:
    label = block["label"].strip()
    seq = _seq_name(label)
    counters[seq] = counters.get(seq, 0) + 1
    return (f'<w:p>{_ppr("Caption")}<w:r><w:t xml:space="preserve">{_esc(label)} </w:t></w:r>'
            f'<w:fldSimple w:instr=" SEQ {seq} \\* ARABIC "><w:r><w:rPr><w:noProof/></w:rPr>'
            f"<w:t>{counters[seq]}</w:t></w:r></w:fldSimple>"
            f'{_text_runs(": " + block["text"])}</w:p>')


def _toc(block: Dict[str, Any]) -> str:
    levels = str(block.get("levels", "1-3"))
    return ('<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:instrText xml:space="preserve"> TOC \\o "{levels}" \\h \\z \\u </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            f"<w:r><w:t>{_esc(TOC_PLACEHOLDER)}</w:t></w:r>"
            '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')


# --- package -----------------------------------------------------------------------------

def _part(name: str, content_type: str, xml: str) -> str:
    return (f'<pkg:part pkg:name="{name}" pkg:contentType="{content_type}">'
            f"<pkg:xmlData>{xml}</pkg:xmlData></pkg:part>")


def parse_location(location: Optional[str]) -> Tuple[str, str]:
    """Parse a location string into (mode, anchor).

    Modes: 'after', 'before', 'replace'.
    Supports 'before:<anchor>', 'after:<anchor>', 'replace:<anchor>',
    as well as plain anchors ('start', 'end', 'cursor', 'p12', 'T0', etc.).
    """
    loc = "end" if location is None else str(location).strip()
    mode = "after"
    m = LOCATION_RE.match(loc)
    if m:
        return m.group(1).lower(), m.group(2).strip() or "cursor"
    if loc.lower() == "start":
        return "before", "start"
    if loc.lower() == "replace_selection":
        return "replace", "cursor"
    return mode, loc


def build_ooxml(blocks: List[Dict[str, Any]], block_id: Optional[str] = None,
                location: Optional[str] = None) -> str:
    """Build a flat-OPC package for `range.insertOoxml`. Raises ValueError listing every problem."""
    errors = validate_blocks(blocks)
    if errors:
        raise ValueError("Invalid blocks:\n- " + "\n- ".join(errors))

    body: List[str] = []
    nums: List[Tuple[int, int]] = []
    counters: Dict[str, int] = {}
    prev_type = None
    for block in blocks:
        t = block["type"]
        if t == "table" and prev_type == "table":
            body.append("<w:p/>")  # two adjacent tables would merge into one
        if t == "heading":
            body.append(_para(block, f"Heading{block['level']}"))
        elif t == "title":
            body.append(_para(block, "Title", align=block.get("align")))
        elif t == "subtitle":
            body.append(_para(block, "Subtitle", align=block.get("align")))
        elif t == "quote":
            body.append(_para(block, "Quote", align=block.get("align")))
        elif t == "paragraph":
            style = _style_id(block["style"]) if block.get("style") else None
            body.append(_para(block, style, align=block.get("align"), space_after_pt=block.get("space_after_pt")))
        elif t in ("bullets", "numbered"):
            num_id = len(nums) + 1
            nums.append((num_id, BULLET_ABSTRACT if t == "bullets" else DECIMAL_ABSTRACT))
            body.append(_list(block["items"], num_id, "ListParagraph"))
        elif t == "table":
            body.append(_table(block))
        elif t == "page_break":
            body.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
        elif t == "caption":
            body.append(_caption(block, counters))
        elif t == "toc":
            body.append(_toc(block))
        prev_type = t

    content = "".join(body)
    if block_id:
        content = (f'<w:sdt><w:sdtPr><w:alias w:val="{_attr(SDT_ALIAS)}"/>'
                   f'<w:tag w:val="ga:{_attr(block_id)}"/><w:id w:val="{zlib.crc32(block_id.encode('utf-8')) % 2_000_000_000}"/>'
                   f"</w:sdtPr><w:sdtContent>{content}</w:sdtContent></w:sdt>")

    document = f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body>{content}</w:body></w:document>'
    root_rels = (f'<Relationships xmlns="{RELS_NS}"><Relationship Id="rId1" Type="{REL_BASE}/officeDocument" '
                 f'Target="word/document.xml"/></Relationships>')
    doc_rels = (f'<Relationships xmlns="{RELS_NS}">'
                f'<Relationship Id="rId1" Type="{REL_BASE}/styles" Target="styles.xml"/>'
                f'<Relationship Id="rId2" Type="{REL_BASE}/numbering" Target="numbering.xml"/></Relationships>')
    rels_ct = "application/vnd.openxmlformats-package.relationships+xml"
    wml = "application/vnd.openxmlformats-officedocument.wordprocessingml"
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<?mso-application progid="Word.Document"?>'
            f'<pkg:package xmlns:pkg="{PKG_NS}">'
            + _part("/_rels/.rels", rels_ct, root_rels)
            + _part("/word/document.xml", f"{wml}.document.main+xml", document)
            + _part("/word/_rels/document.xml.rels", rels_ct, doc_rels)
            + _part("/word/styles.xml", f"{wml}.styles+xml", _styles_xml())
            + _part("/word/numbering.xml", f"{wml}.numbering+xml", _numbering_xml(nums))
            + "</pkg:package>")
