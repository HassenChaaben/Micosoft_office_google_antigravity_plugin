"""
word_fill_fields: resolve the keys the model uses (field paths from the Document Map,
"T{t}r{r}c{c}" coordinates, content-control tags) to exact table cells.
"""

import difflib
import re
import unicodedata
from typing import Any, Dict, List, Optional

COORD_RE = re.compile(r"^\s*T(\d+)\s*r(\d+)\s*c(\d+)\s*$", re.IGNORECASE)
RECORD_RE = re.compile(r"#(\d+)")


def _norm(key: str) -> str:
    key = unicodedata.normalize("NFKD", key)
    key = "".join(ch for ch in key if not unicodedata.combining(ch))
    key = re.sub(r"\s*(›|>|»|/)\s*", " > ", key)
    return re.sub(r"\s+", " ", key).strip().lower()


def resolve_fields(snapshot: Dict[str, Any], fields: Dict[str, Any]) -> Dict[str, Any]:
    by_path = {_norm(f["path"]): f for f in snapshot.get("fields", [])}
    tags = {(c.get("tag") or "").lower(): c.get("tag") for c in snapshot.get("content_controls", []) if c.get("tag")}
    tags.update({(c.get("title") or "").lower(): c.get("tag") or c.get("title")
                 for c in snapshot.get("content_controls", []) if c.get("title")})

    cells: List[Dict[str, Any]] = []
    content_controls: Dict[str, str] = {}
    errors: List[str] = []

    for key, value in fields.items():
        text = "" if value is None else str(value)
        m = COORD_RE.match(key)
        if m:
            t, r, c = (int(x) for x in m.groups())
            matched_f = next((f for f in snapshot.get("fields", [])
                              if f.get("loc", {}).get("table") == t
                              and f.get("loc", {}).get("row") == r
                              and f.get("loc", {}).get("cell") == c), None)
            p_idx = matched_f["loc"].get("paragraph") if matched_f else None
            cells.append({"table": t, "row": r, "cell": c, "paragraph": p_idx, "text": text, "key": key})
            continue
        f = by_path.get(_norm(key))
        if f is not None:
            loc = f["loc"]
            cells.append({"table": loc["table"], "row": loc["row"], "cell": loc["cell"],
                          "paragraph": loc.get("paragraph"), "text": text,
                          "key": f["path"], "placeholder": f.get("placeholder")})
            continue
        if key.lower() in tags:
            content_controls[tags[key.lower()]] = text
            continue

        rec = RECORD_RE.search(key)
        if rec:
            group = _norm(key[:rec.start()]).rstrip(" >")
            n = int(rec.group(1))
            known = [f for f in snapshot.get("fields", []) if f.get("kind") == "record"
                     and _norm(f.get("group") or "") == group]
            if known and n > max(f["record"] for f in known if isinstance(f["record"], int)):
                errors.append(f"'{key}': the grid has only {max(f['record'] for f in known if isinstance(f['record'], int))} "
                              f"empty record rows. Call word_table_clone_row first to add rows, then refill.")
                continue
        close = difflib.get_close_matches(_norm(key), list(by_path), n=3, cutoff=0.6)
        hint = ", ".join(by_path[c]["path"] for c in close) or "use a path from the Document Map or T{t}r{r}c{c}"
        errors.append(f"'{key}' not found. Did you mean: {hint}")

    return {"cells": cells, "content_controls": content_controls, "errors": errors}


def fill_cell_com(doc, table_index: int, row_index: int, col_index: int, text: str,
                  paragraph_index: Optional[int] = None) -> bool:
    """Write text into a table cell via Word COM, with fallbacks for vertically merged rows.

    1. Try doc.Tables(table_index + 1).Cell(row_index + 1, col_index + 1).
    2. If that fails (e.g. vertically merged rows in Word COM), fallback to
       doc.Paragraphs(paragraph_index + 1).Range if paragraph_index is provided.
    3. Fallback to searching the table's Range.Cells collection.
    """
    if doc is None:
        return False
    try:
        if 0 <= table_index < doc.Tables.Count:
            tbl = doc.Tables(table_index + 1)
        else:
            from .word_com_tools import _all_tables
            all_tbls = _all_tables(doc)
            if 0 <= table_index < len(all_tbls):
                tbl = all_tbls[table_index]
            else:
                return False
    except Exception:
        return False

    # 1. Direct Cell access
    try:
        cell = tbl.Cell(row_index + 1, col_index + 1)
        rng = cell.Range
        rng.MoveEnd(1, -1)
        rng.Text = str(text)
        return True
    except Exception:
        pass

    # 2. Fallback to paragraph index from snapshot
    if paragraph_index is not None and paragraph_index >= 0:
        try:
            if paragraph_index < doc.Paragraphs.Count:
                p = doc.Paragraphs(paragraph_index + 1)
                rng = p.Range
                if rng.Text.endswith("\r\x07"):
                    rng.MoveEnd(1, -2)
                elif rng.Text.endswith(("\r", "\x07")):
                    rng.MoveEnd(1, -1)
                rng.Text = str(text)
                return True
        except Exception:
            pass

    # 3. Fallback to searching Cells collection in the table
    try:
        for idx in range(1, tbl.Range.Cells.Count + 1):
            c = tbl.Range.Cells(idx)
            try:
                if c.RowIndex == row_index + 1 and c.ColumnIndex == col_index + 1:
                    rng = c.Range
                    rng.MoveEnd(1, -1)
                    rng.Text = str(text)
                    return True
            except Exception:
                continue
    except Exception:
        pass

    return False


def fill_cells_com_fallback(doc, failed_cells: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Retry a list of failed cells via Word COM.

    Takes a list of cell dicts (with 'table', 'row', 'cell', 'text', and optional 'paragraph')
    and attempts to write each using Word COM.
    Returns {"ok": bool, "written": int, "failed": List[str]}.
    """
    written = 0
    remaining_failed = []
    for c in failed_cells:
        if isinstance(c, str):
            m = COORD_RE.match(c)
            if m:
                t, r, col = (int(x) for x in m.groups())
                ok = fill_cell_com(doc, t, r, col, "")
                if ok:
                    written += 1
                else:
                    remaining_failed.append(c)
            else:
                remaining_failed.append(c)
            continue

        t = c.get("table")
        r = c.get("row")
        col = c.get("cell")
        text = c.get("text", "")
        p = c.get("paragraph")
        label = c.get("key") or f"T{t}r{r}c{col}"
        if t is not None and r is not None and col is not None:
            ok = fill_cell_com(doc, t, r, col, text, paragraph_index=p)
            if ok:
                written += 1
            else:
                remaining_failed.append(label)
        else:
            remaining_failed.append(label)
    return {"ok": len(remaining_failed) == 0, "written": written, "failed": remaining_failed}
