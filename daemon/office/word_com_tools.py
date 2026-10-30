"""
Native Word tools over COM — the "senior Office expert" toolbox.

Every tool takes the target `doc` (a Word.Document COM object, obtained on the COM worker
thread) plus keyword arguments, and returns {"ok": True, ...details} or
{"ok": False, "error": "..."}. Tools never pretend: if Word refuses, the error is returned.

Everything produced is NATIVE Word content: real charts (embedded chart workbook), real
SmartArt, drawing canvases with native shapes and glued connectors, TOC/SEQ/REF fields,
sections, header/footer fields and built-in styles (resolved language-independently).

Anchors (0-based, matching the Document Map from word_snapshot):
    "p12"        paragraph 12            -> doc.Paragraphs(13).Range
    "T3"         table 3                 -> doc.Tables(4).Range
    "T3r2c1"     table 3, row 2, cell 1  -> doc.Tables(4).Cell(3, 2).Range
    "i0"         inline picture/chart 0  -> doc.InlineShapes(1).Range
    "ga:<id>"    content control with Tag <id>
    "<text>"     first paragraph starting with this text (headings preferred)
    "cursor"     the selection in the document's window (also None)
    "start" / "end"
Locations for inserting block content: "<anchor>" or "after:<anchor>" (new paragraph after
the anchor), "before:<anchor>", "replace:<anchor>" (replaces the anchor's content).
"""

import inspect
import re
import uuid
from collections import defaultdict
from datetime import date
from functools import wraps
from typing import Any, Dict, List, Optional, Tuple

# --- constants ---------------------------------------------------------------------------

RED = 0x0000FF  # Word colours are BGR integers
DARK_TEXT = 0x262626
BLUE_FILL = 222 + 235 * 256 + 247 * 65536  # #DEEBF7
BLUE_LINE = 0x1F + 0x4E * 256 + 0x79 * 65536  # #1F4E79
BAR_FILL = 0x44 + 0x72 * 256 + 0xC4 * 65536  # #4472C4
RED_FILL = 0xFC + 0xE4 * 256 + 0xD6 * 65536  # #FCE4D6
PAGE_WIDTH_PT = 468.0  # 6.5 inch text column

CHART_TYPES = {"column": 51, "bar": 57, "line": 4, "pie": 5, "area": 1, "scatter": -4169, "combo": 51}

_URN = "urn:microsoft.com/office/officeart/2005/8/layout/"
SMARTART_LAYOUTS = {
    "process": [_URN + "process1"],
    "hierarchy": [_URN + "hierarchy1"],
    "org_chart": [_URN + "orgChart1", _URN + "hierarchy1"],
    "cycle": [_URN + "cycle2", _URN + "cycle1"],
    "list": [_URN + "default", _URN + "vList2"],
    "pyramid": [_URN + "pyramid1", _URN + "pyramid2"],
    "matrix": [_URN + "matrix3", _URN + "matrix1"],
    "timeline": [_URN + "hProcess11", "urn:microsoft.com/office/officeart/2008/layout/CircleAccentTimeline"],
}

BUILTIN_STYLES = {
    "normal": -1, "title": -63, "subtitle": -75, "caption": -35, "quote": -181, "intense quote": -182,
    "list paragraph": -180, "no spacing": -158, "header": -32, "footer": -33, "hyperlink": -86,
    "strong": -88, "emphasis": -89, "list bullet": -49, "list number": -50,
    # common localized names, so "Titolo" works on French Word and vice versa
    "normale": -1, "titolo": -63, "titre": -63, "sottotitolo": -75, "sous-titre": -75,
    "didascalia": -35, "légende": -35, "citazione": -181, "citation": -181,
    "paragrafo elenco": -180, "paragraphe de liste": -180, "nessuna spaziatura": -158, "sans interligne": -158,
}
_HEADING_RE = re.compile(r"^(heading|titolo|titre|t[ií]tulo|überschrift|kop)\s*(\d)$", re.I)
_TOC_RE = re.compile(r"^(toc|sommario|tdm)\s*(\d)$", re.I)

CAPTION_LABELS = {"figure": -1, "figura": -1, "abbildung": -1, "table": -2, "tabella": -2, "tableau": -2,
                  "tabla": -2, "tabelle": -2, "equation": -3, "equazione": -3, "équation": -3, "ecuación": -3}

ALIGN = {"left": 0, "center": 1, "centre": 1, "right": 2, "justify": 3, "justified": 3}
LIST_STYLES = {"1": 0, "I": 1, "i": 2, "A": 3, "a": 4}  # wdListNumberStyle*
PAGE_NUMBER_STYLES = {
    "arabic": 0, "1": 0, "0": 0,
    "roman_upper": 1, "i_upper": 1, "I": 1,
    "roman_lower": 2, "roman": 2, "i": 2,
    "letter_upper": 3, "a_upper": 3, "A": 3,
    "letter_lower": 4, "letter": 4, "a": 4,
}

# --- errors & tool wrapper -----------------------------------------------------------------


def _err(exc: BaseException) -> str:
    args = getattr(exc, "args", ())
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) > 2 and args[2][2]:
        return f"Word: {args[2][2]}"  # pywintypes.com_error excepinfo description
    if len(args) >= 2 and isinstance(args[0], int) and isinstance(args[1], str):
        return f"Word COM error {args[0]}: {args[1]}"
    return f"{type(exc).__name__}: {exc}"


def _tool(fn):
    """Catch every failure as {"ok": False}, and group the edits into one Undo step."""

    @wraps(fn)
    def wrapper(doc, *args, **kwargs):
        undo = None
        try:
            try:
                undo = doc.Application.UndoRecord
                if undo.IsRecordingCustomRecord:
                    undo = None
                else:
                    undo.StartCustomRecord("AI: " + fn.__name__.replace("word_", "").replace("_", " "))
            except Exception:
                undo = None
            return fn(doc, *args, **kwargs)
        except Exception as exc:
            return {"ok": False, "error": _err(exc)}
        finally:
            if undo is not None:
                try:
                    undo.EndCustomRecord()
                except Exception:
                    pass

    return wrapper


# --- anchors ----------------------------------------------------------------------------

_P_RE = re.compile(r"^p(\d+)$", re.I)
_T_RE = re.compile(r"^T(\d+)(?:r(\d+)c(\d+))?$", re.I)
_I_RE = re.compile(r"^i(\d+)$", re.I)
_MODE_RE = re.compile(r"^(after|before|replace):(.*)$", re.I | re.S)


def _find_paragraph(doc, text: str):
    """First paragraph whose text starts with `text`; headings win over body text."""
    wanted = text.strip()
    if not wanted:
        raise ValueError("Empty anchor text")
    rng = doc.Content
    find = rng.Find
    find.ClearFormatting()
    fallback = None
    for _ in range(200):
        if not find.Execute(FindText=wanted[:250].replace("^", "^^"), MatchCase=False, Forward=True, Wrap=0,
                            MatchWildcards=False):
            break
        para = rng.Paragraphs(1)
        if rng.Start == para.Range.Start:
            if para.OutlineLevel != 10:  # wdOutlineLevelBodyText
                return para.Range
            fallback = fallback or para.Range
        rng.Collapse(0)
    if fallback is not None:
        return fallback
    raise ValueError(f"No paragraph starts with '{wanted}'. Use a 'pN' anchor from the Document Map.")


def _all_tables(doc):
    """Return a flat list of all tables in the document in document order, including nested tables."""
    tables = []
    def _collect(tbl):
        tables.append(tbl)
        try:
            for i in range(1, tbl.Range.Cells.Count + 1):
                c = tbl.Range.Cells(i)
                if c.Tables.Count > 0:
                    for nt in range(1, c.Tables.Count + 1):
                        _collect(c.Tables(nt))
        except Exception:
            pass
    for i in range(1, doc.Tables.Count + 1):
        _collect(doc.Tables(i))
    return tables


def _table(doc, ref):
    if isinstance(ref, int):
        idx = ref
    else:
        m = _T_RE.match(str(ref).strip())
        if not m:
            raise ValueError(f"Bad table reference '{ref}' (use 'T0', 'T1', ...)")
        idx = int(m.group(1))
    if idx < 0:
        raise ValueError(f"Bad table index {idx}")
    if idx < doc.Tables.Count:
        try:
            all_tbls = _all_tables(doc)
            if len(all_tbls) > doc.Tables.Count:
                if idx < len(all_tbls):
                    return all_tbls[idx]
        except Exception:
            pass
        return doc.Tables(idx + 1)
    all_tbls = _all_tables(doc)
    if 0 <= idx < len(all_tbls):
        return all_tbls[idx]
    raise ValueError(f"Table T{idx} does not exist (document has {len(all_tbls)} tables)")


def resolve_anchor(doc, anchor: Optional[str]):
    """Turn an anchor string (see module docstring) into a Word Range."""
    a = "cursor" if anchor is None else str(anchor).strip()
    low = a.lower()
    if low in ("", "cursor", "selection"):
        return doc.ActiveWindow.Selection.Range
    if low == "start":
        return doc.Range(0, 0)
    if low == "end":
        return doc.Paragraphs(doc.Paragraphs.Count).Range
    m = _P_RE.match(a)
    if m:
        i = int(m.group(1))
        if i >= doc.Paragraphs.Count:
            raise ValueError(f"Paragraph p{i} does not exist (document has {doc.Paragraphs.Count} paragraphs)")
        return doc.Paragraphs(i + 1).Range
    m = _T_RE.match(a)
    if m:
        tbl = _table(doc, int(m.group(1)))
        if m.group(2) is None:
            return tbl.Range
        return tbl.Cell(int(m.group(2)) + 1, int(m.group(3)) + 1).Range
    m = _I_RE.match(a)
    if m:
        i = int(m.group(1))
        if i >= doc.InlineShapes.Count:
            raise ValueError(f"Inline shape i{i} does not exist (document has {doc.InlineShapes.Count})")
        return doc.InlineShapes(i + 1).Range
    if low.startswith("ga:"):
        ccs = doc.SelectContentControlsByTag(a[3:])
        if ccs.Count == 0:
            raise ValueError(f"No content control with tag '{a[3:]}'")
        return ccs(1).Range
    return _find_paragraph(doc, a)


def _is_whole_table(doc, rng) -> bool:
    try:
        if rng.Tables.Count == 0:
            return False
        t = rng.Tables(1).Range
        return rng.Start <= t.Start and rng.End >= t.End
    except Exception:
        return False


def _as_normal(doc, pos: int):
    r = doc.Range(pos, pos)
    try:
        r.Paragraphs(1).Style = -1
        r.ListFormat.RemoveNumbers()
    except Exception:
        pass
    return doc.Range(pos, pos)


def _insert_point(doc, location: Optional[str]) -> Tuple[Any, bool]:
    """Collapsed range for new block content, and whether we created a new paragraph for it."""
    loc = "end" if location is None else str(location).strip()
    mode = "after"
    m = _MODE_RE.match(loc)
    if m:
        mode, loc = m.group(1).lower(), m.group(2).strip() or "cursor"
    if loc.lower() == "start" and mode == "after":
        mode = "before"
    rng = resolve_anchor(doc, loc)
    whole_table = _is_whole_table(doc, rng)

    if mode == "replace":
        if whole_table:
            start = rng.Tables(1).Range.Start
            rng.Tables(1).Delete()
            return _as_normal(doc, start), False
        r = rng.Duplicate
        if r.End > r.Start and r.Text.endswith("\r"):
            r.MoveEnd(1, -1)
        r.Text = ""
        return doc.Range(r.Start, r.Start), False

    if mode == "before":
        start = rng.Tables(1).Range.Start if whole_table else rng.Paragraphs(1).Range.Start
        if whole_table:
            if start == 0:
                doc.Range(0, 0).InsertParagraphBefore()
                return _as_normal(doc, 0), True
            prev = doc.Range(start - 1, start - 1).Paragraphs(1).Range
            end = prev.End
            prev.InsertParagraphAfter()
            return _as_normal(doc, end), True
        doc.Range(start, start).InsertParagraphBefore()
        return _as_normal(doc, start), True

    if whole_table:
        end = rng.Tables(1).Range.End
        nxt = doc.Range(end, end).Paragraphs(1).Range
        if nxt.Text == "\r" and nxt.End >= doc.Content.End:
            return _as_normal(doc, end), False
        doc.Range(end, end).InsertParagraphBefore()
        return _as_normal(doc, end), True
    para = rng.Paragraphs(rng.Paragraphs.Count).Range
    if para.Text in ("\r", "\r\x07"):
        return _as_normal(doc, para.Start), False
    end = para.End
    para.InsertParagraphAfter()
    return _as_normal(doc, end), True


def insertion_range(doc, location: Optional[str]):
    """Collapsed Range in an empty Normal paragraph where block content (chart, diagram...) goes."""
    return _insert_point(doc, location)[0]


def _drop_if_empty(doc, rng, created: bool) -> None:
    if not created:
        return
    try:
        para = rng.Paragraphs(1).Range
        if para.Text == "\r" and para.End < doc.Content.End:
            para.Delete()
    except Exception:
        pass


def _inline_point(doc, location: Optional[str]):
    """Collapsed range for inline content: end of the anchor's text, or its start with before:."""
    loc = "cursor" if location is None else str(location).strip()
    mode = "after"
    m = _MODE_RE.match(loc)
    if m:
        mode, loc = m.group(1).lower(), m.group(2).strip() or "cursor"
    rng = resolve_anchor(doc, loc)
    if mode == "before":
        return doc.Range(rng.Start, rng.Start)
    end = rng.End
    text = rng.Text or ""
    if text.endswith("\r\x07"):
        end -= 2
    elif text.endswith(("\r", "\x07")):
        end -= 1
    end = max(end, rng.Start)
    return doc.Range(end, end)


def _text_range(rng):
    """The anchor range without its trailing paragraph / cell mark (for comments, styles...)."""
    r = rng.Duplicate
    while r.End > r.Start and (r.Text or "").endswith(("\r", "\x07")):
        r.MoveEnd(1, -1)
    return r


def _shape_index(doc, rng) -> int:
    """0-based index of the inline shape at `rng` (for the i<N> anchor)."""
    return doc.Range(0, rng.Start).InlineShapes.Count


# --- charts -----------------------------------------------------------------------------

def _col(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


@_tool
def word_insert_chart(doc, type: str = "column", categories: Optional[List[Any]] = None,
                      series: Optional[List[Dict[str, Any]]] = None, title: Optional[str] = None,
                      axes: Optional[Dict[str, str]] = None, location: Optional[str] = "end",
                      legend: Optional[bool] = None) -> Dict[str, Any]:
    kind = str(type or "").lower()
    if kind not in CHART_TYPES:
        return {"ok": False, "error": f"Unknown chart type '{type}'. Use one of {sorted(CHART_TYPES)}"}
    categories = list(categories or [])
    series = list(series or [])
    if not categories or not series:
        return {"ok": False, "error": "A chart needs non-empty 'categories' and 'series'"}
    for s in series:
        vals = s.get("values") if isinstance(s, dict) else None
        if not isinstance(vals, (list, tuple)) or len(vals) != len(categories):
            return {"ok": False, "error": f"Series '{(s or {}).get('name')}' must have exactly {len(categories)} values"}
        if not all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in vals):
            return {"ok": False, "error": f"Series '{s.get('name')}' values must be numbers"}
    if kind == "pie" and len(series) != 1:
        return {"ok": False, "error": "A pie chart takes exactly one series"}
    if kind == "combo" and len(series) < 2:
        return {"ok": False, "error": "A combo chart needs at least two series (first = columns, others = lines)"}
    if kind == "scatter" and not all(isinstance(c, (int, float)) for c in categories):
        return {"ok": False, "error": "For a scatter chart 'categories' are the numeric X values"}

    rng, created = _insert_point(doc, location)
    try:
        ish = doc.InlineShapes.AddChart2(-1, CHART_TYPES[kind], rng)
    except Exception:
        _drop_if_empty(doc, rng, created)
        raise
    try:
        ch = ish.Chart
        cd = ch.ChartData
        try:
            cd.ActivateChartDataWindow()
        except Exception:
            cd.Activate()
        wb = cd.Workbook
        try:
            try:
                wb.Application.Visible = False
            except Exception:
                pass
            ws = wb.Worksheets(1)
            try:
                while ws.ListObjects.Count:
                    ws.ListObjects(1).Unlist()
            except Exception:
                pass
            ws.Cells.ClearContents()
            rows, cols = len(categories) + 1, len(series) + 1
            data = [[""] + [str(s.get("name") or f"Series {i + 1}") for i, s in enumerate(series)]]
            for r, cat in enumerate(categories):
                data.append([cat if kind == "scatter" else str(cat)] + [s["values"][r] for s in series])
            ws.Range(ws.Cells(1, 1), ws.Cells(rows, cols)).Value = data
            ch.SetSourceData(f"='{ws.Name}'!$A$1:${_col(cols)}${rows}", 2)  # xlColumns
        finally:
            wb.Close()
        if kind == "combo":
            for k in range(2, ch.SeriesCollection().Count + 1):
                ch.SeriesCollection(k).ChartType = 4  # xlLine
        ch.HasTitle = bool(title)
        if title:
            ch.ChartTitle.Text = str(title)
        ch.HasLegend = legend if legend is not None else (kind == "pie" or len(series) > 1)
        if kind == "pie":
            s1 = ch.SeriesCollection(1)
            s1.HasDataLabels = True
            try:
                s1.DataLabels().ShowPercentage = True
            except Exception:
                pass
        elif axes:
            for key, idx in (("x", 1), ("y", 2)):
                if axes.get(key):
                    ax = ch.Axes(idx)
                    ax.HasTitle = True
                    ax.AxisTitle.Text = str(axes[key])
        ish.Range.ParagraphFormat.Alignment = 1
        n_series = ch.SeriesCollection().Count
    except Exception:
        ish.Delete()
        _drop_if_empty(doc, rng, created)
        raise
    idx = _shape_index(doc, ish.Range)
    return {"ok": True, "chart_type": kind, "series": n_series, "points": len(categories),
            "anchor": f"i{idx}", "inline_shape_index": idx}


# --- SmartArt ---------------------------------------------------------------------------

_LAYOUT_INDEX: Dict[str, int] = {}


def _smartart_layout(app, layout: str):
    if not _LAYOUT_INDEX:
        for i in range(1, app.SmartArtLayouts.Count + 1):
            _LAYOUT_INDEX[app.SmartArtLayouts(i).Id] = i
    wanted = SMARTART_LAYOUTS.get(str(layout).lower())
    if wanted is None:  # raw URN or suffix such as "chevron1"
        wanted = [u for u in _LAYOUT_INDEX if u == layout or u.lower().endswith("/" + str(layout).lower())]
    for urn in wanted:
        if urn in _LAYOUT_INDEX:
            lay = app.SmartArtLayouts(_LAYOUT_INDEX[urn])
            if lay.Id == urn:
                return lay
            _LAYOUT_INDEX.clear()  # stale cache (another Office build): rebuild once
            return _smartart_layout(app, layout)
    raise ValueError(f"SmartArt layout '{layout}' not available. Use one of {sorted(SMARTART_LAYOUTS)}")


@_tool
def word_insert_smartart(doc, layout: str = "process", nodes: Optional[List[Dict[str, Any]]] = None,
                         location: Optional[str] = "end", width_pt: Optional[float] = None,
                         height_pt: Optional[float] = None) -> Dict[str, Any]:
    nodes = [n if isinstance(n, dict) else {"text": str(n)} for n in (nodes or [])]
    if not nodes:
        return {"ok": False, "error": "SmartArt needs at least one node"}
    app = doc.Application
    lay = _smartart_layout(app, layout)
    rng, created = _insert_point(doc, location)
    inline = True
    try:
        try:
            ish = doc.InlineShapes.AddSmartArt(lay, rng)
            holder = ish
        except Exception:
            holder = doc.Shapes.AddSmartArt(lay, Anchor=rng)
            try:
                holder = holder.ConvertToInlineShape()
            except Exception:
                inline = False
    except Exception:
        _drop_if_empty(doc, rng, created)
        raise
    try:
        sa = holder.SmartArt
        # Keep the layout's own top-level nodes as spares (fixed layouts such as the 2x2 matrix
        # refuse AddNode), drop their default children, delete unused spares at the end.
        for i in range(sa.AllNodes.Count, 0, -1):
            if sa.AllNodes(i).Level > 1:
                sa.AllNodes(i).Delete()
        spares = [sa.Nodes(i) for i in range(1, sa.Nodes.Count + 1)]
        stack: List[Any] = []
        for n in nodes:
            level = max(1, min(int(n.get("level") or 1), len(stack) + 1))
            if level == 1:
                if spares:
                    node = spares.pop(0)
                elif stack:
                    node = stack[0].AddNode(2)  # msoSmartArtNodeAfter
                else:
                    node = sa.Nodes.Add()
            elif len(stack) >= level:
                node = stack[level - 1].AddNode(2)  # next sibling under the same parent
            else:
                node = stack[level - 2].AddNode(5)  # msoSmartArtNodeBelow = first child
            node.TextFrame2.TextRange.Text = str(n.get("text", ""))
            stack = stack[:level - 1] + [node]
        for spare in reversed(spares):
            try:
                spare.Delete()
            except Exception:  # fixed-size layout: leave the slot blank
                spare.TextFrame2.TextRange.Text = ""
        if width_pt:
            holder.Width = float(width_pt)
        if height_pt:
            holder.Height = float(height_pt)
        if inline:
            holder.Range.ParagraphFormat.Alignment = 1
        count = sa.AllNodes.Count
    except Exception:
        holder.Delete()
        _drop_if_empty(doc, rng, created)
        raise
    return {"ok": True, "layout": lay.Id, "nodes": count, "inline": inline,
            **({"anchor": f"i{_shape_index(doc, holder.Range)}"} if inline else {})}


# --- diagram layout (pure Python) ---------------------------------------------------------

def _graph(nodes: List[Dict[str, Any]], edges: List[Dict[str, Any]]):
    ids = [str(n.get("id", i)) for i, n in enumerate(nodes)]
    if len(set(ids)) != len(ids):
        raise ValueError("Node ids must be unique")
    known = set(ids)
    pairs = []
    for e in edges:
        u, v = str(e.get("from")), str(e.get("to"))
        for x in (u, v):
            if x not in known:
                raise ValueError(f"Edge {u}->{v} references unknown node '{x}'")
        pairs.append((u, v))
    return ids, pairs


def layered_layout(ids: List[str], pairs: List[Tuple[str, str]]):
    """Longest-path layers (cycles broken at DFS back edges) + one barycenter ordering pass.
    Returns (layer: {id: int}, order: {id: position in layer}, back_edges: set)."""
    succ = defaultdict(list)
    for u, v in pairs:
        succ[u].append(v)
    state: Dict[str, int] = {}
    back = set()

    def dfs(u):
        state[u] = 1
        for v in succ[u]:
            if state.get(v) == 1:
                back.add((u, v))
            elif v not in state:
                dfs(v)
        state[u] = 2

    for i in ids:
        if i not in state:
            dfs(i)
    dag = [(u, v) for u, v in pairs if (u, v) not in back and u != v]
    preds, dsucc = defaultdict(list), defaultdict(list)
    for u, v in dag:
        preds[v].append(u)
        dsucc[u].append(v)
    indeg = {i: len(preds[i]) for i in ids}
    ready = [i for i in ids if indeg[i] == 0]
    topo = []
    while ready:
        u = ready.pop(0)
        topo.append(u)
        for v in dsucc[u]:
            indeg[v] -= 1
            if indeg[v] == 0:
                ready.append(v)
    layer: Dict[str, int] = {}
    for u in topo:
        layer[u] = max((layer[p] + 1 for p in preds[u]), default=0)
    rows = defaultdict(list)
    for u in topo:
        rows[layer[u]].append(u)
    order: Dict[str, float] = {}
    for lvl in sorted(rows):
        members = rows[lvl]
        if lvl > 0:
            def bary(u):
                ps = [order[p] for p in preds[u] if p in order]
                return sum(ps) / len(ps) if ps else float(members.index(u))
            members.sort(key=bary)
        for pos, u in enumerate(members):
            order[u] = pos
    return layer, order, back


def tree_layout(ids: List[str], pairs: List[Tuple[str, str]]):
    """Tidy tree from parent->child edges: leaves get consecutive slots, parents are centred.
    Returns {id: (depth, slot)} and the number of leaf slots."""
    children = defaultdict(list)
    has_parent = set()
    for u, v in pairs:
        if v not in has_parent and v != u:
            children[u].append(v)
            has_parent.add(v)
    pos: Dict[str, Tuple[int, float]] = {}
    nxt = [0.0]

    def place(u, depth, seen):
        seen.add(u)
        kids = [c for c in children[u] if c not in seen]
        if not kids:
            pos[u] = (depth, nxt[0])
            nxt[0] += 1
            return
        for c in kids:
            place(c, depth + 1, seen)
        pos[u] = (depth, (pos[kids[0]][1] + pos[kids[-1]][1]) / 2)

    seen: set = set()
    for r in [i for i in ids if i not in has_parent] + ids:
        if r not in seen:
            place(r, 0, seen)
    return pos, max(int(nxt[0]), 1)


def critical_path(ids: List[str], pairs: List[Tuple[str, str]], durations: Dict[str, float]):
    """CPM: returns (es, ef, slack, critical_ids)."""
    layer, _, back = layered_layout(ids, pairs)
    dag = [(u, v) for u, v in pairs if (u, v) not in back and u != v]
    topo = sorted(ids, key=lambda i: layer[i])
    es, ef = {}, {}
    for u in topo:
        es[u] = max((ef[p] for p, v in dag if v == u), default=0.0)
        ef[u] = es[u] + durations[u]
    finish = max(ef.values(), default=0.0)
    lf, ls = {}, {}
    for u in reversed(topo):
        lf[u] = min((ls[v] for p, v in dag if p == u), default=finish)
        ls[u] = lf[u] - durations[u]
    slack = {u: ls[u] - es[u] for u in ids}
    return es, ef, slack, [u for u in ids if abs(slack[u]) < 1e-9]


# --- diagram drawing (COM canvas + glued connectors) -----------------------------------------

FLOW_SHAPES = {"start": 69, "end": 69, "terminator": 69, "decision": 63, "io": 64, "data": 64,
               "document": 67, "process": 61}
SITE = {"top": 1, "left": 2, "bottom": 3, "right": 4}  # COM 1-based; DrawingML idx = site - 1


def _site_point(box, site):
    x, y, w, h = box
    return {1: (x + w / 2, y), 2: (x, y + h / 2), 3: (x + w / 2, y + h), 4: (x + w, y + h / 2)}[site]


def _style_box(shape, text, critical=False, fill=BLUE_FILL, size=9):
    shape.Fill.ForeColor.RGB = RED_FILL if critical else fill
    shape.Line.ForeColor.RGB = RED if critical else BLUE_LINE
    shape.Line.Weight = 2.25 if critical else 1.0
    if text is not None:
        tf = shape.TextFrame
        for side in ("MarginLeft", "MarginRight", "MarginTop", "MarginBottom"):
            setattr(tf, side, 2)
        try:
            tf.VerticalAnchor = 3  # msoAnchorMiddle
        except Exception:
            pass
        tr = tf.TextRange
        tr.Text = str(text)
        tr.Font.Size = size
        tr.Font.Color = DARK_TEXT
        tr.ParagraphFormat.Alignment = 1
        tr.ParagraphFormat.SpaceAfter = 0
        tr.ParagraphFormat.SpaceBefore = 0


def _label(canvas, text, x, y, w=60, h=16, size=8, align=1, name=None):
    tb = canvas.CanvasItems.AddTextbox(1, x, y, w, h)
    if name:
        tb.Name = name
    tb.Line.Visible = 0
    tb.Fill.Visible = 0
    tf = tb.TextFrame
    for side in ("MarginLeft", "MarginRight", "MarginTop", "MarginBottom"):
        setattr(tf, side, 0)
    tr = tf.TextRange
    tr.Text = str(text)
    tr.Font.Size = size
    tr.Font.Color = DARK_TEXT
    tr.ParagraphFormat.Alignment = align
    tr.ParagraphFormat.SpaceAfter = 0
    return tb


class _Canvas:
    """Collects shapes and connectors; connectors are glued afterwards through OOXML."""

    def __init__(self, doc, anchor, width, height, name):
        self.doc = doc
        self.shape = doc.Shapes.AddCanvas(0, 0, width, height, anchor)
        self.shape.Name = name
        self.boxes: Dict[str, Tuple[float, float, float, float]] = {}
        self.links: List[Tuple[str, str, int, str, int]] = []
        self.n_shapes = 0

    def box(self, name, autoshape, x, y, w, h, text, critical=False, fill=BLUE_FILL, size=9):
        s = self.shape.CanvasItems.AddShape(autoshape, x, y, w, h)
        s.Name = name
        _style_box(s, text, critical, fill, size)
        self.boxes[name] = (x, y, w, h)
        self.n_shapes += 1
        return s

    def connect(self, a, a_site, b, b_site, critical=False, elbow=None, label=None, route=None):
        """Connector from site a_site of box a to site b_site of box b (sites: 1 top, 2 left,
        3 bottom, 4 right). route: None (Word's horizontal-first elbow or straight line),
        "vfirst" (down-across-down) or "hv" (one bend: across then down)."""
        bx, by = _site_point(self.boxes[a], a_site)
        ex, ey = _site_point(self.boxes[b], b_site)
        dx, dy = ex - bx, ey - by
        if elbow is None:
            elbow = abs(dx) > 0.5 and abs(dy) > 0.5
        if route is None and elbow and a_site == 3 and b_site == 1 and abs(dx) > 0.5 and dy > 0.5:
            route = "vfirst"
        geometry = None
        if route == "vfirst":
            # bentConnector3 is horizontal-first; rotate its (dy x |dx|) box by 90 degrees about the
            # centre (flipV when running right) to get down-across-down.
            cx, cy = (bx + ex) / 2, (by + ey) / 2
            geometry = ("bentConnector3", cx - dy / 2, cy - abs(dx) / 2, dy, abs(dx), 5400000, False, dx > 0)
            elbow = True
        elif route == "hv" and dx > 0.5 and dy > 0.5:
            geometry = ("bentConnector2", bx, by, dx, dy, 0, False, False)
            elbow = True
        c = self.shape.CanvasItems.AddConnector(2 if elbow else 1, bx, by, ex, ey)
        name = f"e_{len(self.links)}"
        c.Name = name
        c.Line.ForeColor.RGB = RED if critical else BLUE_LINE
        c.Line.Weight = 2.0 if critical else 1.0
        c.Line.EndArrowheadStyle = 2  # msoArrowheadTriangle
        self.links.append((name, a, a_site - 1, b, b_site - 1, geometry))
        if label:
            _label(self.shape, label, (bx + ex) / 2 + 2, (by + ey) / 2 - 8, 40, 14, 8, 0)
        return c


_CNV_RE = r'<wps:cNvPr id="(\d+)" name="([^"]*)"'
EMU = 12700  # per point


def _xfrm(prst, x, y, w, h, rot, flip_h, flip_v):
    attrs = (f' rot="{rot}"' if rot else "") + (' flipH="1"' if flip_h else "") + (' flipV="1"' if flip_v else "")
    return (f'<a:xfrm{attrs}><a:off x="{round(x * EMU)}" y="{round(y * EMU)}"/>'
            f'<a:ext cx="{round(w * EMU)}" cy="{round(h * EMU)}"/></a:xfrm>'), prst


def glue_connectors_xml(xml: str, links) -> Tuple[str, int]:
    """Glue connectors (a:stCxn/a:endCxn) named in `links` and apply their custom geometry.
    links: (connector_name, from_shape, from_idx, to_shape, to_idx[, geometry]) -> (xml, glued)."""
    ids = {name: sid for sid, name in re.findall(_CNV_RE, xml)}
    glued = 0
    for link in links:
        cname, a, ai, b, bi = link[:5]
        geometry = link[5] if len(link) > 5 else None
        m = re.search(r'<wps:cNvPr id="\d+" name="%s"\s*/>' % re.escape(cname), xml)
        if not m:
            continue
        if a in ids and b in ids:
            pat = re.compile(r"<wps:cNvCnPr\s*/>")
            repl = '<wps:cNvCnPr><a:stCxn id="%s" idx="%d"/><a:endCxn id="%s" idx="%d"/></wps:cNvCnPr>' % (
                ids[a], ai, ids[b], bi)
            m2 = pat.match(xml, m.end())
            if m2:
                xml = xml[:m2.start()] + repl + xml[m2.end():]
                glued += 1
        if geometry:
            new_xfrm, prst = _xfrm(*geometry)
            x0 = xml.find("<a:xfrm", m.end())
            x1 = xml.find("</a:xfrm>", x0) + len("</a:xfrm>")
            if x0 > 0 and x1 > x0:
                xml = xml[:x0] + new_xfrm + xml[x1:]
                p0 = xml.find('prst="', x0)
                if p0 > 0:
                    p1 = xml.find('"', p0 + 6)
                    xml = xml[:p0 + 6] + prst + xml[p1:]
    return xml, glued


def canvas_run_xml(pkg_xml: str, name: str) -> str:
    """Return a flat-OPC package whose body is one clean paragraph holding only the drawing run
    whose wp:docPr is `name` (the source paragraph may also hold heading text)."""
    m = re.search(r'<wp:docPr\b[^>]*\bname="%s"' % re.escape(name), pkg_xml)
    if not m:
        raise ValueError(f"Drawing '{name}' not found in the paragraph XML")
    starts = list(re.finditer(r"<w:r[ >]", pkg_xml[:m.start()]))
    if not starts:
        raise ValueError("Drawing run not found")
    i = starts[-1].start()
    depth = 0
    for tok in re.finditer(r"<w:r[ >]|</w:r>", pkg_xml[i:]):
        depth += 1 if tok.group(0) != "</w:r>" else -1
        if depth == 0:
            run = pkg_xml[i:i + tok.end()]
            break
    else:
        raise ValueError("Unbalanced drawing run")
    b0 = pkg_xml.index("<w:body>") + len("<w:body>")
    b1 = pkg_xml.find("<w:sectPr", b0)
    if b1 < 0:
        b1 = pkg_xml.index("</w:body>", b0)
    return pkg_xml[:b0] + "<w:p>" + run + "</w:p>" + pkg_xml[b1:]


def _finish_canvas(doc, cv: _Canvas, ip) -> Dict[str, Any]:
    """Inline the canvas at `ip` with glued connectors.

    COM's ConnectorFormat.BeginConnect cannot be called from Python (its [in,out] Shape**
    argument is rejected with DISP_E_TYPEMISMATCH under both late and early binding), so the
    canvas is built with COM, then its OOXML gets a:stCxn/a:endCxn and is re-inserted."""
    name = cv.shape.Name
    src = cv.shape.ConvertToInlineShape().Range
    xml = canvas_run_xml(src.Paragraphs(1).Range.WordOpenXML, name)
    patched, glued = glue_connectors_xml(xml, cv.links)
    # Delete Word's copy first: if it already sits at `ip`, inserting at that boundary would
    # extend `src` over the new canvas and a later src.Delete() would remove both.
    src.Delete()
    try:
        ip.InsertXML(patched)
    except Exception:
        glued = 0
        ip.InsertXML(xml)  # unglued, but still native shapes and connectors
    shape = doc.Shapes(name)
    para = shape.Anchor.Paragraphs(1)
    try:
        para.Style = -1
        para.Alignment = 1
    except Exception:
        pass
    return {"name": name, "glued": glued, "canvas_items": shape.CanvasItems.Count}


def _draw_layered(cv_factory, ids, texts, kinds, pairs, critical, crit_edges, labels, horizontal):
    layer, order, back = layered_layout(ids, pairs)
    groups = defaultdict(list)
    for u in ids:
        groups[layer[u]].append(u)
    for g in groups.values():
        g.sort(key=lambda u: order[u])
    n_layers = max(groups) + 1
    widest = max(len(g) for g in groups.values())
    gap, vgap = 14.0, 30.0
    if horizontal:
        w = min(110.0, (PAGE_WIDTH_PT - gap * (n_layers + 1)) / n_layers)
        h = 46.0
        hgap = (PAGE_WIDTH_PT - n_layers * w) / (n_layers + 1)
        height = widest * (h + 16) + 16
        width = PAGE_WIDTH_PT
    else:
        w = min(130.0, (PAGE_WIDTH_PT - gap * (widest + 1)) / widest)
        h = 40.0
        height = n_layers * (h + vgap) + vgap / 2
        width = PAGE_WIDTH_PT
    cv = cv_factory(width, height)
    for lvl, members in groups.items():
        k = len(members)
        for j, u in enumerate(members):
            if horizontal:
                x = hgap + lvl * (w + hgap)
                y = (height - k * (h + 16) + 16) / 2 + j * (h + 16)
            else:
                x = (width - (k * w + (k - 1) * gap)) / 2 + j * (w + gap)
                y = vgap / 2 + lvl * (h + vgap)
            cv.box(f"n_{ids.index(u)}", kinds[u], x, y, w, h, texts[u], u in critical)
    for (u, v), lab in zip(pairs, labels):
        a, b = f"n_{ids.index(u)}", f"n_{ids.index(v)}"
        crit = (u, v) in crit_edges
        if u == v:
            continue
        if (u, v) in back or layer[v] <= layer[u]:
            if horizontal:
                cv.connect(a, SITE["bottom"], b, SITE["bottom"], crit, True, lab)
            else:
                cv.connect(a, SITE["right"], b, SITE["right"], crit, True, lab)
        elif horizontal:
            cv.connect(a, SITE["right"], b, SITE["left"], crit, None, lab)
        else:
            cv.connect(a, SITE["bottom"], b, SITE["top"], crit, None, lab)
    return cv


def _parse_date(val: Any) -> Optional[date]:
    if isinstance(val, date):
        return val
    if not isinstance(val, str) or not val.strip():
        return None
    s = val.strip()
    try:
        return date.fromisoformat(s[:10])
    except ValueError:
        pass
    m = re.match(r"^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$", s)
    if m:
        day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            return date(year, month, day)
        except ValueError:
            pass
    raise ValueError(f"Date '{val}' is neither 'gg/mm/aaaa' nor an ISO date (YYYY-MM-DD)")


def _to_days(values: List[Any]) -> Tuple[List[Optional[float]], Optional[date]]:
    """Numbers stay numbers; gg/mm/aaaa or ISO dates become day offsets from the earliest date."""
    dates = {}
    for i, v in enumerate(values):
        if isinstance(v, str):
            try:
                dates[i] = _parse_date(v)
            except ValueError as e:
                raise ValueError(f"start '{v}' is neither a number, gg/mm/aaaa nor an ISO date (YYYY-MM-DD)") from e
    origin = min(dates.values()) if dates else None
    out: List[Optional[float]] = []
    for i, v in enumerate(values):
        if i in dates:
            out.append(float((dates[i] - origin).days))
        elif isinstance(v, (int, float)) and not isinstance(v, bool):
            out.append(float(v))
        else:
            out.append(None)
    return out, origin


def _nice_step(total: float, scale: float = 0.0, min_px: float = 34.0) -> float:
    """Smallest 'nice' tick step giving <= 10 ticks and at least `min_px` points between ticks."""
    for step in (1, 2, 5, 7, 10, 14, 20, 30, 50, 60, 90, 100, 180, 365):
        if total / step <= 10 and step * scale >= min_px:
            return float(step)
    return float(max(int(total / 10) + 1, 1))


@_tool
def word_insert_diagram(doc, kind: str = "flowchart", nodes: Optional[List[Dict[str, Any]]] = None,
                        edges: Optional[List[Dict[str, Any]]] = None, location: Optional[str] = "end"
                        ) -> Dict[str, Any]:
    kind = str(kind or "").lower()
    if kind not in ("flowchart", "wbs", "network", "gantt"):
        return {"ok": False, "error": f"Unknown diagram kind '{kind}'. Use flowchart, wbs, network or gantt"}
    nodes = list(nodes or [])
    edges = list(edges or [])
    if not nodes:
        return {"ok": False, "error": "A diagram needs at least one node"}
    ids, pairs = _graph(nodes, edges)
    labels = [e.get("label") or (e.get("type") if e.get("type") and e.get("type") != "FS" else None) for e in edges]
    by_id = dict(zip(ids, nodes))
    texts = {i: str(by_id[i].get("text", i)) for i in ids}

    durations = {}
    if kind in ("network", "gantt"):
        for i in ids:
            node = by_id[i]
            is_milestone = bool(node.get("milestone"))
            start_val = node.get("start")
            finish_val = node.get("finish")
            d = node.get("duration")
            if d is not None and not isinstance(d, (int, float)):
                raise ValueError(f"Node '{i}': duration must be a number")
            if is_milestone:
                durations[i] = 0.0
            elif d is not None:
                durations[i] = float(d)
            elif start_val and finish_val:
                sd = _parse_date(start_val)
                fd = _parse_date(finish_val)
                if sd and fd:
                    days_diff = (fd - sd).days
                    durations[i] = float(max(days_diff + 1 if days_diff >= 0 else 1, 1))
                else:
                    durations[i] = 1.0 if kind == "gantt" else 0.0
            else:
                durations[i] = 1.0 if kind == "gantt" else 0.0

    flagged = {i for i in ids if by_id[i].get("critical")}
    es: Dict[str, float] = {}
    if flagged:
        critical = set(flagged)
    elif kind in ("network", "gantt") and len(durations) == len(ids) and all(durations[i] > 0 for i in ids):
        es, _, _, crit_list = critical_path(ids, pairs, durations)
        critical = set(crit_list)
    else:
        critical = set()
    if kind in ("network", "gantt") and len(durations) == len(ids) and all(durations[i] > 0 for i in ids) and not es:
        es = critical_path(ids, pairs, durations)[0]

    explicit_crit_edges = {(e.get("from"), e.get("to")) for e in edges if e.get("critical")}
    crit_edges = {(u, v) for u, v in pairs if ((u, v) in explicit_crit_edges or
                  (u in critical and v in critical and (not es or abs(es.get(u, 0) + durations.get(u, 0) - es.get(v, 0)) < 1e-9)))}
    for u, v in explicit_crit_edges:
        if u in ids:
            critical.add(u)
        if v in ids:
            critical.add(v)

    rng, created = _insert_point(doc, location)
    name = f"diagram_{kind}_{uuid.uuid4().hex[:6]}"
    holder: Dict[str, _Canvas] = {}

    def factory(width, height):
        holder["cv"] = _Canvas(doc, rng, width, height, name)
        return holder["cv"]

    try:
        extra: Dict[str, Any] = {}
        if kind == "flowchart":
            kinds = {}
            for i in ids:
                t = str(by_id[i].get("type") or "").lower()
                if not t and texts[i].rstrip().endswith("?"):
                    t = "decision"
                kinds[i] = FLOW_SHAPES.get(t, 61)
            cv = _draw_layered(factory, ids, texts, kinds, pairs, critical, crit_edges, labels, horizontal=False)
        elif kind == "network":
            if es:
                for i in ids:
                    texts[i] = f"{texts[i]}\n{durations[i]:g}d · ES {es[i]:g} / EF {es[i] + durations[i]:g}"
            kinds = {i: 5 for i in ids}  # msoShapeRoundedRectangle
            cv = _draw_layered(factory, ids, texts, kinds, pairs, critical, crit_edges, labels, horizontal=True)
        elif kind == "wbs":
            pos, leaves = tree_layout(ids, pairs)
            depth = max(d for d, _ in pos.values()) + 1
            slot = PAGE_WIDTH_PT / leaves
            w, h, vgap = min(120.0, slot - 8), 40.0, 30.0
            cv = factory(PAGE_WIDTH_PT, depth * (h + vgap) + vgap / 2)
            for i in ids:
                d, s = pos[i]
                cv.box(f"n_{ids.index(i)}", 5 if d else 1, slot * (s + 0.5) - w / 2, vgap / 2 + d * (h + vgap), w, h,
                       texts[i], i in critical, size=9 if w > 60 else 7)
            for u, v in pairs:
                if pos[v][0] == pos[u][0] + 1:
                    cv.connect(f"n_{ids.index(u)}", SITE["bottom"], f"n_{ids.index(v)}", SITE["top"],
                               (u, v) in crit_edges)
        else:  # gantt
            raw_starts, origin = _to_days([by_id[i].get("start") for i in ids])
            starts = {i: (raw_starts[k] if raw_starts[k] is not None else es.get(i, 0.0)) for k, i in enumerate(ids)}
            total = max(starts[i] + max(durations[i], 1.0) for i in ids) or 1.0
            label_w, row_h, bar_h, top = 120.0, 22.0, 14.0, 30.0
            area = PAGE_WIDTH_PT - label_w - 10
            scale = area / total
            cv = factory(PAGE_WIDTH_PT, top + len(ids) * row_h + 8)
            step = _nice_step(total, scale)
            t = 0.0
            while t <= total + 1e-9:
                x = label_w + t * scale
                text = (date.fromordinal(origin.toordinal() + int(t)).strftime("%d/%m") if origin else f"{t:g}")
                lw = max(min(40.0, step * scale - 2), 16.0)
                _label(cv.shape, text, min(max(x - lw / 2, 0.0), PAGE_WIDTH_PT - lw), 4, lw, 14,
                       7 if lw >= 30 else 6)
                t += step
            axis = cv.shape.CanvasItems.AddLine(label_w, top - 6, label_w + area, top - 6)
            axis.Line.ForeColor.RGB = BLUE_LINE
            for k, i in enumerate(ids):
                y = top + k * row_h
                _label(cv.shape, texts[i], 2, y + 3, label_w - 6, 16, 8, 0, name=f"lbl_{k}")
                is_milestone = bool(by_id[i].get("milestone"))
                is_crit = i in critical
                if is_milestone:
                    m_size = bar_h
                    center_x = label_w + starts[i] * scale
                    bar = cv.box(f"bar_{k}", 63, center_x - m_size / 2, y + (row_h - m_size) / 2,
                                 m_size, m_size, None, is_crit, fill=RED_FILL if is_crit else BAR_FILL)
                else:
                    bar_w = max(durations[i] * scale, 2.0)
                    bar = cv.box(f"bar_{k}", 1, label_w + starts[i] * scale, y + (row_h - bar_h) / 2,
                                 bar_w, bar_h, None, is_crit, fill=BAR_FILL)
                if is_crit:
                    bar.Fill.ForeColor.RGB = RED
            edge_by_pair = {(e.get("from"), e.get("to")): e for e in edges}
            for u, v in pairs:
                a, b = f"bar_{ids.index(u)}", f"bar_{ids.index(v)}"
                edge_info = edge_by_pair.get((u, v), {})
                dep_type = str(edge_info.get("type") or "FS").upper()
                is_crit = (u, v) in crit_edges
                if dep_type == "SS":
                    cv.connect(a, SITE["left"], b, SITE["left"], is_crit, True, route="hv")
                elif dep_type == "FF":
                    cv.connect(a, SITE["right"], b, SITE["right"], is_crit, True, route="hv")
                elif dep_type == "SF":
                    cv.connect(a, SITE["left"], b, SITE["right"], is_crit, True, route="hv")
                else:
                    gap_x = cv.boxes[b][0] - (cv.boxes[a][0] + cv.boxes[a][2])
                    if gap_x > 6:  # finish -> start with room for an elbow
                        cv.connect(a, SITE["right"], b, SITE["left"], is_crit, True)
                    elif cv.boxes[b][0] + cv.boxes[b][2] / 2 > cv.boxes[a][0] + cv.boxes[a][2] + 3:
                        # bars touch: from the end of u across, then down onto the top of v
                        cv.connect(a, SITE["right"], b, SITE["top"], is_crit, True, route="hv")
                    else:  # overlapping tasks
                        cv.connect(a, SITE["bottom"], b, SITE["top"], is_crit, True)
            extra["bars"] = len(ids)
            extra["span"] = total
        info = _finish_canvas(doc, cv, rng)
    except Exception:
        cv_obj = holder.get("cv")
        if cv_obj is not None:
            try:
                doc.Shapes(name).Delete()
            except Exception:
                pass
        _drop_if_empty(doc, rng, created)
        raise
    return {"ok": True, "kind": kind, "name": info["name"], "shapes": cv.n_shapes,
            "connectors": len(cv.links), "glued": info["glued"], "canvas_items": info["canvas_items"],
            "critical": sorted(critical), **extra}


# --- fields: TOC, captions, cross references --------------------------------------------------

@_tool
def word_insert_toc(doc, location: Optional[str] = "start", levels: str = "1-3",
                    label: Optional[str] = None) -> Dict[str, Any]:
    if label:
        return word_insert_table_of_figures(doc, location=location, label=label)
    m = re.match(r"^\s*(\d)\s*-\s*(\d)\s*$", str(levels))
    if not m or not 1 <= int(m.group(1)) <= int(m.group(2)) <= 9:
        return {"ok": False, "error": "levels must look like '1-3'"}
    upper, lower = int(m.group(1)), int(m.group(2))
    rng, created = _insert_point(doc, location)
    try:
        toc = doc.TablesOfContents.Add(Range=rng, UseHeadingStyles=True, UpperHeadingLevel=upper,
                                       LowerHeadingLevel=lower)
        toc.Update()
    except Exception:
        _drop_if_empty(doc, rng, created)
        raise
    return {"ok": True, "levels": f"{upper}-{lower}", "entries": max(toc.Range.Paragraphs.Count - 1, 0),
            "tocs": doc.TablesOfContents.Count}


@_tool
def word_insert_table_of_figures(doc, location: Optional[str] = "end", label: str = "Figure") -> Dict[str, Any]:
    app = doc.Application
    name = caption_label(app, label)
    rng, created = _insert_point(doc, location)
    try:
        tof = doc.TablesOfFigures.Add(Range=rng, Caption=name)
        tof.Update()
    except Exception:
        _drop_if_empty(doc, rng, created)
        raise
    return {"ok": True, "label": name, "entries": max(tof.Range.Paragraphs.Count - 1, 0),
            "tables_of_figures": doc.TablesOfFigures.Count}


def caption_label(app, label: Optional[str]) -> str:
    """Localized caption label name ('Figure' -> 'Figura' on Italian Word); creates custom labels."""
    lab = str(label or "Figure").strip()
    builtin = CAPTION_LABELS.get(lab.lower())
    if builtin is not None:
        return app.CaptionLabels(builtin).Name
    for i in range(1, app.CaptionLabels.Count + 1):
        if app.CaptionLabels(i).Name.lower() == lab.lower():
            return app.CaptionLabels(i).Name
    app.CaptionLabels.Add(lab)
    return lab


@_tool
def word_insert_caption(doc, anchor: str, label: str = "Figure", text: str = "",
                        position: str = "below") -> Dict[str, Any]:
    pos = str(position or "below").lower()
    if pos not in ("below", "above"):
        return {"ok": False, "error": "position must be 'below' or 'above'"}
    app = doc.Application
    name = caption_label(app, label)
    rng = resolve_anchor(doc, anchor)
    if _is_whole_table(doc, rng):
        rng = rng.Tables(1).Range
        start, end = rng.Start, rng.End
    else:
        start = rng.Paragraphs(1).Range.Start
        end = rng.Paragraphs(rng.Paragraphs.Count).Range.End
    title = f": {text}" if text and not str(text).startswith((":", " ", ".", "-", "–")) else str(text or "")
    rng.InsertCaption(Label=name, Title=title, Position=1 if pos == "below" else 0)
    cap = doc.Range(end, end).Paragraphs(1) if pos == "below" else doc.Range(start, start).Paragraphs(1)
    return {"ok": True, "label": name, "text": cap.Range.Text.strip(),
            "paragraph": doc.Range(0, cap.Range.Start).Paragraphs.Count}


@_tool
def word_insert_image(doc, file_path: str, location: Optional[str] = "end",
                      alt_text: Optional[str] = None, width_pt: Optional[float] = None,
                      height_pt: Optional[float] = None, caption: Optional[str] = None) -> Dict[str, Any]:
    import os
    abs_path = os.path.abspath(str(file_path).strip())
    if not os.path.exists(abs_path):
        return {"ok": False, "error": f"Image file not found: '{file_path}'"}
    rng, created = _insert_point(doc, location)
    try:
        shape = doc.InlineShapes.AddPicture(FileName=abs_path, LinkToFile=False,
                                            SaveWithDocument=True, Range=rng)
        if alt_text:
            try:
                shape.AlternativeText = str(alt_text)
            except Exception:
                pass
            try:
                shape.Title = str(alt_text)
            except Exception:
                pass
        if width_pt and float(width_pt) > 0:
            shape.Width = float(width_pt)
        if height_pt and float(height_pt) > 0:
            shape.Height = float(height_pt)
        idx = doc.InlineShapes.Count - 1
        cap_res = None
        if caption:
            cap_res = word_insert_caption(doc, anchor=f"i{idx}", label="Figure", text=str(caption), position="below")
        return {"ok": True, "anchor": f"i{idx}", "width": shape.Width, "height": shape.Height,
                "alt_text": getattr(shape, "AlternativeText", alt_text), "caption": cap_res}
    except Exception:
        _drop_if_empty(doc, rng, created)
        raise


def _items(doc, ref_type) -> List[str]:
    items = doc.GetCrossReferenceItems(ref_type)
    return [str(x) for x in (items or ())]


@_tool
def word_insert_cross_reference(doc, location: Optional[str] = "cursor", target_anchor: str = "",
                                kind: str = "caption", label: Optional[str] = None, show: str = "label_number",
                                prefix: str = "", suffix: str = "", hyperlink: bool = True) -> Dict[str, Any]:
    kind = str(kind or "").lower()
    target = str(target_anchor or "").strip()
    if not target:
        return {"ok": False, "error": "target_anchor is required"}
    app = doc.Application
    k = kind.lower()
    if k in ("caption", "label_and_number"):
        m = re.match(r"^(.+?)\s+(\d+)$", target)
        ref_type = caption_label(app, m.group(1) if m else label)
        items = _items(doc, ref_type)
        if m:
            item = int(m.group(2))
        else:
            item = next((idx + 1 for idx, t in enumerate(items) if target.lower() in t.lower()), 0)
        show_map = {"label_number": 3, "label_and_number": 3, "number": 3, "text": 4, "full": 2, "page": 7}
        ref_kind = show_map.get(show, 3)
    elif k in ("heading", "heading_number"):
        ref_type = 1  # wdRefTypeHeading
        items = _items(doc, 1)
        tgt = target.strip().lower()
        item = next((idx + 1 for idx, t in enumerate(items)
                     if t.strip().lower().startswith(tgt) or tgt in t.lower()
                     or re.sub(r"^[\d.\s]+", "", t.strip().lower()).startswith(tgt)), 0)
        default_show = -4 if k == "heading_number" else -1
        show_map = {"text": -1, "number": -4, "heading_number": -4, "page": 7, "full": -2}
        ref_kind = show_map.get(show, default_show)
    elif k in ("page",):
        ref_kind = 7  # wdPageNumber
        m = re.match(r"^(.+?)\s+(\d+)$", target)
        candidate_label = caption_label(app, m.group(1)) if m else None
        caption_items = _items(doc, candidate_label) if candidate_label else []
        if m and candidate_label and 1 <= int(m.group(2)) <= len(caption_items):
            ref_type = candidate_label
            items = caption_items
            item = int(m.group(2))
        else:
            heading_items = _items(doc, 1)
            tgt = target.strip().lower()
            h_match = next((idx + 1 for idx, t in enumerate(heading_items)
                            if t.strip().lower().startswith(tgt) or tgt in t.lower()
                            or re.sub(r"^[\d.\s]+", "", t.strip().lower()).startswith(tgt)), 0)
            if h_match:
                ref_type = 1
                items = heading_items
                item = h_match
            elif doc.Bookmarks.Exists(target):
                ref_type = 2
                items = _items(doc, 2)
                item = target
            else:
                ref_type = 1
                items = heading_items
                item = 0
    elif k == "bookmark":
        ref_type = 2  # wdRefTypeBookmark
        items = _items(doc, 2)
        item = target if doc.Bookmarks.Exists(target) else 0
        ref_kind = {"page": 7}.get(show, -1)
    else:
        return {"ok": False, "error": "kind must be label_and_number, heading_number, page, caption, heading or bookmark"}
    if not item or (isinstance(item, int) and item > len(items)):
        return {"ok": False, "error": f"Target '{target}' not found. Available: {[t.strip() for t in items][:20]}"}
    pt = _inline_point(doc, location)
    start = pt.Start
    if prefix or suffix:
        pt.InsertAfter(f"{prefix}{suffix}")
    at = doc.Range(start + len(prefix), start + len(prefix))
    at.InsertCrossReference(ReferenceType=ref_type, ReferenceKind=ref_kind, ReferenceItem=item,
                            InsertAsHyperlink=bool(hyperlink))
    field = doc.Range(start, doc.Content.End).Fields(1)
    return {"ok": True, "kind": kind, "target": items[item - 1].strip() if isinstance(item, int) else item,
            "field": field.Code.Text.strip(), "result": field.Result.Text}


@_tool
def word_update_fields(doc) -> Dict[str, Any]:
    errors = doc.Fields.Update()
    tocs = 0
    for i in range(1, doc.TablesOfContents.Count + 1):
        doc.TablesOfContents(i).Update()
        tocs += 1
    tofs = 0
    for i in range(1, doc.TablesOfFigures.Count + 1):
        doc.TablesOfFigures(i).Update()
        tofs += 1
    for s in range(1, doc.Sections.Count + 1):
        sec = doc.Sections(s)
        for coll in (sec.Headers, sec.Footers):
            for k in (1, 2, 3):
                try:
                    coll(k).Range.Fields.Update()
                except Exception:
                    pass
    out = {"ok": True, "fields": doc.Fields.Count, "tocs": tocs, "tables_of_figures": tofs}
    if errors:
        out["first_error_field"] = errors
    return out


# --- layout: sections, headers/footers --------------------------------------------------------

@_tool
def word_insert_section_break(doc, location: Optional[str] = "end", type: str = "next_page",
                              orientation: Optional[str] = None) -> Dict[str, Any]:
    kinds = {"next_page": 2, "continuous": 3, "even_page": 4, "odd_page": 5}
    if type not in kinds:
        return {"ok": False, "error": f"type must be one of {sorted(kinds)}"}
    if orientation not in (None, "", "portrait", "landscape"):
        return {"ok": False, "error": "orientation must be portrait or landscape"}
    loc = "end" if location is None else str(location).strip()
    mode = "after"
    m = _MODE_RE.match(loc)
    if m:
        mode, loc = m.group(1).lower(), m.group(2).strip() or "cursor"
    rng = resolve_anchor(doc, loc)
    if mode == "before":
        pos = rng.Paragraphs(1).Range.Start
    else:
        para = rng.Paragraphs(rng.Paragraphs.Count).Range
        pos = para.End if para.End < doc.Content.End else para.End - 1
    doc.Range(pos, pos).InsertBreak(kinds[type])
    new_idx = doc.Range(pos + 1, pos + 1).Sections(1).Index
    sec = doc.Sections(new_idx)
    if orientation:
        want = 1 if orientation == "landscape" else 0
        if sec.PageSetup.Orientation != want:
            sec.PageSetup.Orientation = want
    return {"ok": True, "section": new_idx, "sections": doc.Sections.Count,
            "orientation": "landscape" if sec.PageSetup.Orientation == 1 else "portrait"}


_HF_FIELD_RE = re.compile(r"(\{PAGE\}|\{NUMPAGES\})")


@_tool
def word_set_header_footer(doc, section: int = 1, kind: str = "header", text: str = "",
                           page_numbers: bool = False, different_first_page: bool = False,
                           alignment: str = "center",
                           page_number_format: Optional[str] = None,
                           page_number_restart: Optional[bool] = None,
                           page_number_start: Optional[int] = None) -> Dict[str, Any]:
    if kind not in ("header", "footer"):
        return {"ok": False, "error": "kind must be header or footer"}
    if not 1 <= int(section) <= doc.Sections.Count:
        return {"ok": False, "error": f"section must be 1..{doc.Sections.Count}"}
    if alignment not in ALIGN:
        return {"ok": False, "error": f"alignment must be one of {sorted(ALIGN)}"}
    sec = doc.Sections(int(section))
    sec.PageSetup.DifferentFirstPageHeaderFooter = bool(different_first_page)
    hf = (sec.Headers if kind == "header" else sec.Footers)(1)  # wdHeaderFooterPrimary

    pns = hf.PageNumbers
    if page_number_restart is not None:
        pns.RestartNumberingAtSection = bool(page_number_restart)
    if page_number_start is not None:
        pns.StartingNumber = int(page_number_start)
        pns.RestartNumberingAtSection = True
    if page_number_format:
        fmt = PAGE_NUMBER_STYLES.get(str(page_number_format).lower())
        if fmt is not None:
            pns.NumberStyle = fmt

    text = str(text or "")
    if page_numbers and "{PAGE}" not in text:
        text = f"{text} {{PAGE}}" if text else "{PAGE}"
    hf.Range.Text = ""
    fields = 0
    for piece in _HF_FIELD_RE.split(text):
        if not piece:
            continue
        r = hf.Range
        r.MoveEnd(1, -1) if r.Text.endswith("\r") and r.End > r.Start else None
        r.Collapse(0)
        if piece in ("{PAGE}", "{NUMPAGES}"):
            hf.Range.Fields.Add(r, 33 if piece == "{PAGE}" else 26)  # wdFieldPage / wdFieldNumPages
            fields += 1
        else:
            r.InsertAfter(piece)
    hf.Range.ParagraphFormat.Alignment = ALIGN[alignment]
    return {"ok": True, "section": int(section), "kind": kind, "text": hf.Range.Text.strip(), "fields": fields,
            "different_first_page": bool(sec.PageSetup.DifferentFirstPageHeaderFooter),
            "restart_page_number": bool(pns.RestartNumberingAtSection),
            "starting_page_number": int(pns.StartingNumber)}


@_tool
def word_set_page_numbering(doc, section: int = 1, format: str = "arabic",
                            restart: bool = True, start: int = 1) -> Dict[str, Any]:
    if not 1 <= int(section) <= doc.Sections.Count:
        return {"ok": False, "error": f"section must be 1..{doc.Sections.Count}"}
    sec = doc.Sections(int(section))
    pns = sec.Footers(1).PageNumbers
    pns.RestartNumberingAtSection = bool(restart)
    pns.StartingNumber = int(start)
    fmt = PAGE_NUMBER_STYLES.get(str(format).lower(), 0)
    pns.NumberStyle = fmt
    return {"ok": True, "section": int(section), "format": format, "restart": restart, "start": start}


# --- review -----------------------------------------------------------------------------------

@_tool
def word_track_changes(doc, on: bool = True) -> Dict[str, Any]:
    doc.TrackRevisions = bool(on)
    return {"ok": True, "track_changes": bool(doc.TrackRevisions)}


@_tool
def word_add_comment(doc, anchor: Optional[str], text: str) -> Dict[str, Any]:
    if not str(text or "").strip():
        return {"ok": False, "error": "Comment text is empty"}
    rng = _text_range(resolve_anchor(doc, anchor))
    c = doc.Comments.Add(Range=rng, Text=str(text))
    return {"ok": True, "comments": doc.Comments.Count, "on_text": rng.Text[:80], "index": c.Index}


# --- styles & numbering -------------------------------------------------------------------------

def resolve_style(doc, name: str):
    """Style object from an English built-in name, a localized name or a common alias."""
    key = str(name or "").strip()
    low = key.lower()
    m = _HEADING_RE.match(key)
    const = -(int(m.group(2)) + 1) if m and 1 <= int(m.group(2)) <= 9 else BUILTIN_STYLES.get(low)
    if const is None:
        m = _TOC_RE.match(key)
        if m and 1 <= int(m.group(2)) <= 9:
            const = -(19 + int(m.group(2)))
    if const is not None:
        return doc.Styles(const)
    try:
        return doc.Styles(key)
    except Exception:
        raise ValueError(f"Style '{name}' not found. Use a built-in English name like 'Heading 1' or the exact "
                         "style name shown in the Document Map")


def _bgr(color: str) -> int:
    m = re.match(r"^#?([0-9a-fA-F]{6})$", str(color).strip())
    if not m:
        raise ValueError(f"Colour '{color}' must be #RRGGBB")
    v = m.group(1)
    return int(v[0:2], 16) + int(v[2:4], 16) * 256 + int(v[4:6], 16) * 65536


@_tool
def word_modify_style(doc, style: str, font: Optional[str] = None, size_pt: Optional[float] = None,
                      color: Optional[str] = None, bold: Optional[bool] = None, italic: Optional[bool] = None,
                      space_before_pt: Optional[float] = None, space_after_pt: Optional[float] = None,
                      align: Optional[str] = None) -> Dict[str, Any]:
    if align is not None and align not in ALIGN:
        return {"ok": False, "error": f"align must be one of {sorted(ALIGN)}"}
    rgb = _bgr(color) if color else None
    st = resolve_style(doc, style)
    f = st.Font
    if font:
        f.Name = font
    if size_pt:
        f.Size = float(size_pt)
    if rgb is not None:
        f.Color = rgb
    if bold is not None:
        f.Bold = bool(bold)
    if italic is not None:
        f.Italic = bool(italic)
    pf = None
    try:
        pf = st.ParagraphFormat
    except Exception:
        pass
    if pf is not None:
        if space_before_pt is not None:
            pf.SpaceBefore = float(space_before_pt)
        if space_after_pt is not None:
            pf.SpaceAfter = float(space_after_pt)
        if align is not None:
            pf.Alignment = ALIGN[align]
    return {"ok": True, "style": st.NameLocal, "font": f.Name, "size_pt": f.Size}


@_tool
def word_apply_style(doc, anchor: Optional[str], style: str) -> Dict[str, Any]:
    st = resolve_style(doc, style)
    rng = resolve_anchor(doc, anchor)
    rng.Style = st.NameLocal
    return {"ok": True, "style": st.NameLocal, "paragraphs": rng.Paragraphs.Count}


@_tool
def word_link_numbering_to_headings(doc, scheme: str = "1.1.1") -> Dict[str, Any]:
    tokens = [t for t in str(scheme or "").strip().strip(".").split(".")]
    if not tokens or len(tokens) > 9 or any(t not in LIST_STYLES for t in tokens):
        return {"ok": False, "error": "scheme must look like '1.1.1' (tokens 1, a, A, i, I; up to 9 levels)"}
    lt = doc.ListTemplates.Add(True)  # OutlineNumbered
    for i, tok in enumerate(tokens, start=1):
        lv = lt.ListLevels(i)
        lv.NumberFormat = ".".join(f"%{k}" for k in range(1, i + 1))
        lv.NumberStyle = LIST_STYLES[tok]
        lv.TrailingCharacter = 0  # tab
        lv.Alignment = 0
        lv.NumberPosition = 0
        lv.TextPosition = (0.75 + 0.25 * (i - 1)) * 28.3465  # cm -> pt (CentimetersToPoints fails on some builds)
        lv.TabPosition = lv.TextPosition
        lv.StartAt = 1
        lv.ResetOnHigher = i - 1
        lv.LinkedStyle = doc.Styles(-(i + 1)).NameLocal
    for i in range(1, len(tokens) + 1):
        doc.Styles(-(i + 1)).LinkToListTemplate(lt, i)
    return {"ok": True, "levels": len(tokens), "scheme": ".".join(tokens),
            "styles": [doc.Styles(-(i + 1)).NameLocal for i in range(1, len(tokens) + 1)]}


# --- forms & tables -------------------------------------------------------------------------------

@_tool
def word_set_checkbox(doc, anchor: str, checked: bool = True) -> Dict[str, Any]:
    a = str(anchor or "").strip()
    cands = []
    rng = None
    if a.lower().startswith("ga:"):
        ccs = doc.SelectContentControlsByTag(a[3:])
        cands = [ccs(i) for i in range(1, ccs.Count + 1)]
    else:
        try:
            rng = resolve_anchor(doc, a)
        except Exception:
            rng = None
        is_p_coord = bool(_P_RE.match(a))
        if rng is not None:
            ccs = rng.ContentControls
            cands = [ccs(i) for i in range(1, ccs.Count + 1)]
            if not is_p_coord and not any(getattr(c, "Type", None) == 8 for c in cands):
                try:
                    if rng.Tables.Count > 0 and rng.Cells.Count > 0:
                        row = rng.Cells(1).Row
                        for c_idx in range(1, row.Cells.Count + 1):
                            cell_rng = row.Cells(c_idx).Range
                            cell_ccs = cell_rng.ContentControls
                            for k in range(1, cell_ccs.Count + 1):
                                cands.append(cell_ccs(k))
                except Exception:
                    pass
            if not is_p_coord and not any(getattr(c, "Type", None) == 8 for c in cands):
                try:
                    nxt = rng.Paragraphs(rng.Paragraphs.Count).Next()
                    if nxt is not None:
                        for k in range(1, nxt.Range.ContentControls.Count + 1):
                            cands.append(nxt.Range.ContentControls(k))
                except Exception:
                    pass
        if not is_p_coord and not any(getattr(c, "Type", None) == 8 for c in cands):
            try:
                for i in range(1, doc.ContentControls.Count + 1):
                    cc = doc.ContentControls(i)
                    if cc.Type == 8:
                        tag = (cc.Tag or "").lower()
                        title = (cc.Title or "").lower()
                        if a.lower() in tag or a.lower() in title:
                            cands.append(cc)
            except Exception:
                pass

    boxes = [c for c in cands if getattr(c, "Type", None) == 8]
    if boxes:
        boxes[0].Checked = bool(checked)
        return {"ok": True, "checked": bool(boxes[0].Checked), "tag": boxes[0].Tag}

    # Legacy FormFields
    if rng is not None:
        try:
            ffs = []
            if rng.FormFields.Count > 0:
                ffs = [rng.FormFields(i) for i in range(1, rng.FormFields.Count + 1)]
            elif rng.Tables.Count > 0 and rng.Cells.Count > 0:
                row = rng.Cells(1).Row
                for c_idx in range(1, row.Cells.Count + 1):
                    c_ffs = row.Cells(c_idx).Range.FormFields
                    ffs.extend([c_ffs(k) for k in range(1, c_ffs.Count + 1)])
            ff_boxes = [f for f in ffs if f.Type == 71]
            if ff_boxes:
                ff_boxes[0].CheckBox.Value = bool(checked)
                return {"ok": True, "checked": bool(ff_boxes[0].CheckBox.Value), "name": ff_boxes[0].Name}
        except Exception:
            pass

    # Unicode checkbox symbols: ☐ (\u2610), ☑ (\u2611), ☒ (\u2612)
    if rng is not None:
        try:
            check_ranges = [rng]
            if rng.Tables.Count > 0 and rng.Cells.Count > 0:
                row = rng.Cells(1).Row
                check_ranges.extend([row.Cells(i).Range for i in range(1, row.Cells.Count + 1)])
            for cr in check_ranges:
                t = cr.Text
                for sym in ("\u2610", "\u2611", "\u2612"):
                    if sym in t:
                        find = cr.Find
                        find.ClearFormatting()
                        if find.Execute(FindText=sym, MatchCase=True, Forward=True):
                            new_sym = "\u2611" if checked else "\u2610"
                            cr.Text = new_sym
                            return {"ok": True, "checked": bool(checked), "symbol": new_sym}
        except Exception:
            pass

    return {"ok": False, "error": f"No checkbox found at or adjacent to '{anchor}'"}


@_tool
def word_table_clone_row(doc, table: Any, row: int, count: int = 1) -> Dict[str, Any]:
    tbl = _table(doc, table)
    n = tbl.Rows.Count
    row, count = int(row), int(count)
    if not 0 <= row < n:
        return {"ok": False, "error": f"row must be 0..{n - 1}"}
    if not 1 <= count <= 200:
        return {"ok": False, "error": "count must be 1..200"}
    R = row + 1
    src = tbl.Rows(R)  # fails on vertically merged rows -> reported as error
    for _ in range(count):
        tbl.Rows.Add(tbl.Rows(R))  # inserted before R, with R's formatting
    orig, dst = tbl.Rows(R + count), tbl.Rows(R)
    for c in range(1, orig.Cells.Count + 1):
        s = orig.Cells(c).Range
        s.MoveEnd(1, -1)
        if s.Text:
            d = dst.Cells(c).Range
            d.MoveEnd(1, -1)
            d.FormattedText = s.FormattedText
            s.Text = ""
    del src
    return {"ok": True, "rows": tbl.Rows.Count,
            "new_rows": list(range(row + 1, row + 1 + count))}


@_tool
def word_delete_column(doc, table: Any, col: int) -> Dict[str, Any]:
    tbl = _table(doc, table)
    col = int(col)
    n = tbl.Columns.Count
    if not 0 <= col < n:
        return {"ok": False, "error": f"col must be 0..{n - 1} (table has {n} columns)"}
    if n <= 1:
        return {"ok": False, "error": "Cannot delete the last remaining column of a table"}
    tbl.Columns(col + 1).Delete()
    return {"ok": True, "deleted_column_index": col, "remaining_columns": tbl.Columns.Count}


@_tool
def word_delete_row(doc, table: Any, row: int) -> Dict[str, Any]:
    tbl = _table(doc, table)
    row = int(row)
    n = tbl.Rows.Count
    if not 0 <= row < n:
        return {"ok": False, "error": f"row must be 0..{n - 1} (table has {n} rows)"}
    if n <= 1:
        return {"ok": False, "error": "Cannot delete the last remaining row of a table"}
    tbl.Rows(row + 1).Delete()
    return {"ok": True, "deleted_row_index": row, "remaining_rows": tbl.Rows.Count}


# --- registry ------------------------------------------------------------------------------------

_LOC = {"type": "string", "description": "Where to insert: anchor from the Document Map ('p12', 'T0', 'i0', "
        "'ga:<id>', heading text, 'cursor', 'start', 'end'). Default = new paragraph AFTER the anchor; "
        "prefix with 'before:' or 'replace:' to change that."}
_ANCHOR = {"type": "string", "description": "Anchor: 'p12' paragraph, 'T0' table, 'T0r1c2' cell, 'i0' inline "
           "chart/picture, 'ga:<id>' content control, heading text, or 'cursor'."}


def _schema(props: Dict[str, Any], required: List[str]) -> Dict[str, Any]:
    return {"type": "object", "properties": props, "required": required, "additionalProperties": False}


TOOLS: Dict[str, Dict[str, Any]] = {
    "word_insert_chart": {
        "fn": word_insert_chart, "destructive": False,
        "description": "Insert a NATIVE, editable Word chart (embedded chart workbook, not an image). Use for any "
                       "numeric comparison or trend: column/bar for categories, line/area for time, pie for "
                       "shares of one total, scatter for X-Y numbers, combo for columns+line. Returns anchor "
                       "'iN' to caption it with word_insert_caption.",
        "inputSchema": _schema({
            "type": {"type": "string", "enum": sorted(CHART_TYPES)},
            "categories": {"type": "array", "items": {"type": ["string", "number"]},
                           "description": "X-axis labels (numbers for scatter)."},
            "series": {"type": "array", "items": {"type": "object", "properties": {
                "name": {"type": "string"}, "values": {"type": "array", "items": {"type": "number"}}},
                "required": ["name", "values"]}, "description": "One value per category in each series."},
            "title": {"type": "string"},
            "axes": {"type": "object", "properties": {"x": {"type": "string"}, "y": {"type": "string"}}},
            "legend": {"type": "boolean"},
            "location": _LOC}, ["type", "categories", "series"]),
    },
    "word_insert_smartart": {
        "fn": word_insert_smartart, "destructive": False,
        "description": "Insert real Word SmartArt. Pick the layout by meaning: process (sequential steps), "
                       "timeline (dated milestones), cycle (recurring loop), hierarchy/org_chart (reporting "
                       "lines; use node level 1..n), list (grouped points), pyramid (layers/priorities), "
                       "matrix (2x2 quadrants, give 4 nodes).",
        "inputSchema": _schema({
            "layout": {"type": "string", "enum": sorted(SMARTART_LAYOUTS)},
            "nodes": {"type": "array", "items": {"type": "object", "properties": {
                "text": {"type": "string"}, "level": {"type": "integer", "minimum": 1}},
                "required": ["text"]}},
            "width_pt": {"type": "number"}, "height_pt": {"type": "number"},
            "location": _LOC}, ["layout", "nodes"]),
    },
    "word_insert_diagram": {
        "fn": word_insert_diagram, "destructive": False,
        "description": "Draw a native diagram on a Word drawing canvas: shapes with text joined by glued "
                       "connectors (stay attached when moved). flowchart = steps top-down (type start/end/"
                       "decision/io; text ending '?' is a decision); wbs = work breakdown tree (edges parent->"
                       "child); network = PERT/activity-on-node left->right, durations give ES/EF and the "
                       "critical path in red; gantt = bars on a time axis (duration, optional start number or "
                       "ISO date, finish date, milestone diamond; edges = dependencies with type FS/SS/FF/SF). "
                       "Mark nodes or edges critical:true to force red.",
        "inputSchema": _schema({
            "kind": {"type": "string", "enum": ["flowchart", "wbs", "network", "gantt"]},
            "nodes": {"type": "array", "items": {"type": "object", "properties": {
                "id": {"type": "string"}, "text": {"type": "string"},
                "type": {"type": "string", "enum": ["start", "end", "process", "decision", "io", "document"]},
                "duration": {"type": "number"}, "start": {"type": ["number", "string"]},
                "finish": {"type": ["number", "string"]}, "milestone": {"type": "boolean"},
                "critical": {"type": "boolean"}}, "required": ["id", "text"]}},
            "edges": {"type": "array", "items": {"type": "object", "properties": {
                "from": {"type": "string"}, "to": {"type": "string"}, "label": {"type": "string"},
                "type": {"type": "string", "enum": ["FS", "SS", "FF", "SF"]},
                "critical": {"type": "boolean"}},
                "required": ["from", "to"]}},
            "location": _LOC}, ["kind", "nodes"]),
    },
    "word_insert_toc": {
        "fn": word_insert_toc, "destructive": False,
        "description": "Insert a real Table of Contents field built from Heading styles (default at the start). "
                       "Headings must use Heading 1..n styles (see word_apply_style).",
        "inputSchema": _schema({"location": _LOC, "levels": {"type": "string", "pattern": "^[1-9]-[1-9]$",
                                                             "description": "e.g. '1-3'"}}, []),
    },
    "word_insert_caption": {
        "fn": word_insert_caption, "destructive": False,
        "description": "Add a numbered caption (SEQ field) to a chart/picture ('i0'), table ('T0') or paragraph. "
                       "Labels Figure/Table/Equation are mapped to the document language automatically. Tables "
                       "usually get captions above, figures below.",
        "inputSchema": _schema({"anchor": _ANCHOR, "label": {"type": "string", "default": "Figure"},
                                "text": {"type": "string"}, "position": {"type": "string", "enum": ["below", "above"]}},
                               ["anchor"]),
    },
    "word_insert_cross_reference": {
        "fn": word_insert_cross_reference, "destructive": False,
        "description": "Insert a live cross-reference (REF field) at the end of the anchor's text: to a caption "
                       "('Figure 1'), a heading (its text) or a bookmark name. Use prefix/suffix for words "
                       "around it, e.g. prefix ' (see ', suffix ')'.",
        "inputSchema": _schema({
            "location": _ANCHOR, "target_anchor": {"type": "string"},
            "kind": {"type": "string", "enum": ["caption", "heading", "bookmark", "label_and_number", "heading_number", "page"]},
            "label": {"type": "string"}, "show": {"type": "string", "enum": ["label_number", "label_and_number", "text", "full", "number", "heading_number", "page"]},
            "prefix": {"type": "string"}, "suffix": {"type": "string"}, "hyperlink": {"type": "boolean"}},
            ["target_anchor", "kind"]),
    },
    "word_update_fields": {
        "fn": word_update_fields, "destructive": False,
        "description": "Update every field: TOCs, tables of figures, caption numbers, cross-references, page "
                       "numbers in headers/footers. Run after inserting or moving captions/headings.",
        "inputSchema": _schema({}, []),
    },
    "word_insert_section_break": {
        "fn": word_insert_section_break, "destructive": False,
        "description": "Start a new section after the anchor paragraph (next_page or continuous) and optionally "
                       "make it landscape - e.g. for a wide table or Gantt. Add a second break after the wide "
                       "content and set it back to portrait.",
        "inputSchema": _schema({"location": _LOC,
                                "type": {"type": "string", "enum": ["next_page", "continuous", "even_page", "odd_page"]},
                                "orientation": {"type": "string", "enum": ["portrait", "landscape"]}}, []),
    },
    "word_set_header_footer": {
        "fn": word_set_header_footer, "destructive": True,
        "description": "Set the primary header or footer text of a section (replaces it). Put {PAGE} and "
                       "{NUMPAGES} in the text for live page fields, or page_numbers:true to append the page "
                       "number. different_first_page hides it on a title page. Supports page_number_format "
                       "(arabic, roman_lower, roman_upper, etc.) and page_number_restart.",
        "inputSchema": _schema({"section": {"type": "integer", "minimum": 1},
                                "kind": {"type": "string", "enum": ["header", "footer"]}, "text": {"type": "string"},
                                "page_numbers": {"type": "boolean"}, "different_first_page": {"type": "boolean"},
                                "alignment": {"type": "string", "enum": ["left", "center", "right", "justify"]},
                                "page_number_format": {"type": "string", "enum": ["arabic", "roman_lower", "roman_upper", "letter_lower", "letter_upper"]},
                                "page_number_restart": {"type": "boolean"},
                                "page_number_start": {"type": "integer", "minimum": 1}},
                               ["kind"]),
    },
    "word_track_changes": {
        "fn": word_track_changes, "destructive": False,
        "description": "Turn Track Changes on or off. Turn it on before editing a document someone must review.",
        "inputSchema": _schema({"on": {"type": "boolean"}}, ["on"]),
    },
    "word_add_comment": {
        "fn": word_add_comment, "destructive": False,
        "description": "Add a review comment (balloon) on the anchor's text without changing the text.",
        "inputSchema": _schema({"anchor": _ANCHOR, "text": {"type": "string"}}, ["anchor", "text"]),
    },
    "word_modify_style": {
        "fn": word_modify_style, "destructive": True,
        "description": "Change a style definition so every paragraph using it updates (the professional way to "
                       "restyle a document). Use English built-in names ('Heading 1', 'Title', 'Normal', "
                       "'Caption', 'Quote') - they work in any Word language.",
        "inputSchema": _schema({"style": {"type": "string"}, "font": {"type": "string"}, "size_pt": {"type": "number"},
                                "color": {"type": "string", "pattern": "^#?[0-9A-Fa-f]{6}$"},
                                "bold": {"type": "boolean"}, "italic": {"type": "boolean"},
                                "space_before_pt": {"type": "number"}, "space_after_pt": {"type": "number"},
                                "align": {"type": "string", "enum": ["left", "center", "right", "justify"]}},
                               ["style"]),
    },
    "word_apply_style": {
        "fn": word_apply_style, "destructive": True,
        "description": "Apply a style to the anchor paragraph(s), e.g. make a line 'Heading 2' so it appears in the "
                       "TOC. Built-in English names work in any Word language.",
        "inputSchema": _schema({"anchor": _ANCHOR, "style": {"type": "string"}}, ["anchor", "style"]),
    },
    "word_link_numbering_to_headings": {
        "fn": word_link_numbering_to_headings, "destructive": True,
        "description": "Number all headings automatically (1, 1.1, 1.1.1 ...) with a multilevel list linked to "
                       "Heading 1..n styles - never type numbers by hand.",
        "inputSchema": _schema({"scheme": {"type": "string", "description": "e.g. '1.1.1' or '1.a.i'"}}, []),
    },
    "word_set_checkbox": {
        "fn": word_set_checkbox, "destructive": False,
        "description": "Tick or untick a checkbox content control in a form ('ga:<tag>' or the anchor holding it).",
        "inputSchema": _schema({"anchor": _ANCHOR, "checked": {"type": "boolean"}}, ["anchor", "checked"]),
    },
    "word_table_clone_row": {
        "fn": word_table_clone_row, "destructive": False,
        "description": "Add empty rows after row N of a table, copying that row's formatting (shading, borders, "
                       "heights) - use to grow a record grid in a template. Rows are 0-based like the Document Map.",
        "inputSchema": _schema({"table": {"type": "string", "description": "'T0'"},
                                "row": {"type": "integer", "minimum": 0},
                                "count": {"type": "integer", "minimum": 1, "maximum": 200}}, ["table", "row"]),
    },
    "word_delete_column": {
        "fn": word_delete_column, "destructive": True,
        "description": "Delete a column by 0-based index from a table (e.g. 'T0'). Preserves the rest of the table.",
        "inputSchema": _schema({"table": {"type": "string", "description": "'T0'"},
                                "col": {"type": "integer", "minimum": 0, "description": "0-based column index"}},
                               ["table", "col"]),
    },
    "word_delete_row": {
        "fn": word_delete_row, "destructive": True,
        "description": "Delete a row by 0-based index from a table (e.g. 'T0'). Preserves the rest of the table.",
        "inputSchema": _schema({"table": {"type": "string", "description": "'T0'"},
                                "row": {"type": "integer", "minimum": 0, "description": "0-based row index"}},
                               ["table", "row"]),
    },
}

EXTRA_TOOLS: Dict[str, Dict[str, Any]] = {
    "word_insert_table_of_figures": {
        "fn": word_insert_table_of_figures, "destructive": False,
        "description": "Insert a Table of Figures (built from SEQ caption fields like Figure or Table).",
        "inputSchema": _schema({"location": _LOC, "label": {"type": "string", "default": "Figure"}}, []),
    },
    "word_insert_image": {
        "fn": word_insert_image, "destructive": False,
        "description": "Insert an image/picture with optional alt text, dimensions and caption into the document.",
        "inputSchema": _schema({"file_path": {"type": "string"}, "location": _LOC,
                                "alt_text": {"type": "string"}, "width_pt": {"type": "number"},
                                "height_pt": {"type": "number"}, "caption": {"type": "string"}},
                               ["file_path"]),
    },
    "word_set_page_numbering": {
        "fn": word_set_page_numbering, "destructive": True,
        "description": "Configure page numbering format (arabic, roman_lower, etc.) and restart for a section.",
        "inputSchema": _schema({"section": {"type": "integer", "minimum": 1},
                                "format": {"type": "string", "enum": ["arabic", "roman_lower", "roman_upper", "letter_lower", "letter_upper"]},
                                "restart": {"type": "boolean"}, "start": {"type": "integer", "minimum": 1}},
                               []),
    },
}


def call_tool(doc, name: str, args: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Run a registered tool by name, dropping (and reporting) arguments it does not take."""
    spec = TOOLS.get(name) or EXTRA_TOOLS.get(name)
    if spec is None:
        return {"ok": False, "error": f"Unknown tool '{name}'"}
    fn = spec["fn"]
    params = inspect.signature(getattr(fn, "__wrapped__", fn)).parameters
    args = dict(args or {})
    ignored = sorted(k for k in args if k not in params or k == "doc")
    result = fn(doc, **{k: v for k, v in args.items() if k not in ignored})
    if ignored:
        result = {**result, "ignored_args": ignored}
    return result


__all__ = ["TOOLS", "EXTRA_TOOLS", "call_tool", "resolve_anchor", "insertion_range", "resolve_style", "caption_label",
           "layered_layout", "tree_layout", "critical_path", "glue_connectors_xml", "canvas_run_xml",
           "word_insert_table_of_figures", "word_insert_image", "word_set_page_numbering"] + [n for n in TOOLS]
