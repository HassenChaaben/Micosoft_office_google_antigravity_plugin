"""
word_fill_fields: resolve the keys the model uses (field paths from the Document Map,
"T{t}r{r}c{c}" coordinates, content-control tags) to exact table cells.
"""

import difflib
import re
import unicodedata
from typing import Any, Dict, List

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
            cells.append({"table": t, "row": r, "cell": c, "text": text, "key": key})
            continue
        f = by_path.get(_norm(key))
        if f is not None:
            loc = f["loc"]
            cells.append({"table": loc["table"], "row": loc["row"], "cell": loc["cell"], "text": text,
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
