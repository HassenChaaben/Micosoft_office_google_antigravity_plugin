"""
office_verify: compare the document before and after the agent acted, and list quality
issues an Office expert would check before saying "done".
"""

from typing import Any, Dict, List, Optional


def verify(before: Optional[Dict[str, Any]], after: Dict[str, Any], expect_all_fields_filled: bool = True) -> Dict[str, Any]:
    before = before or {"fields": [], "tables": [], "outline": [], "stats": {"tables": 0}}
    issues: List[Dict[str, Any]] = []

    b_fields = {f["path"]: f for f in before["fields"]}
    a_fields = {f["path"]: f for f in after["fields"]}
    filled = [p for p, f in b_fields.items() if f["empty"] and (p not in a_fields or not a_fields[p]["empty"])]
    emptied = [p for p, f in b_fields.items() if not f["empty"] and p in a_fields and a_fields[p]["empty"]]
    still_empty = [p for p, f in a_fields.items() if f["empty"]]

    if expect_all_fields_filled and still_empty:
        issues.append({
            "type": "empty_fields", "severity": "warning", "count": len(still_empty),
            "message": f"{len(still_empty)} fields are still empty. Fill them or write '[da definire]' if the value is unknown.",
            "examples": still_empty[:15],
        })
    if emptied:
        issues.append({"type": "fields_cleared", "severity": "error", "examples": emptied[:15],
                       "message": "Some fields that had a value are now empty. Check you did not overwrite content."})

    if before["stats"]["tables"] and after["stats"]["tables"] < before["stats"]["tables"]:
        issues.append({"type": "tables_removed", "severity": "error",
                       "message": f"Tables went from {before['stats']['tables']} to {after['stats']['tables']}. "
                                  "A template's structure must be preserved."})
    b_rows = {t["index"]: t["rows"] for t in before["tables"]}
    for t in after["tables"]:
        if t["index"] in b_rows and t["rows"] < b_rows[t["index"]]:
            issues.append({"type": "rows_removed", "severity": "error", "table": t["index"],
                           "message": f"Table T{t['index']} lost rows ({b_rows[t['index']]} → {t['rows']})."})

    placeholders = list(after.get("placeholders") or []) + [
        f["path"] for f in after["fields"] if not f["empty"] and f.get("placeholder") and f["placeholder"].startswith("[")
        and "da definire" not in f["placeholder"].lower()]
    if placeholders:
        issues.append({"type": "placeholders_left", "severity": "warning", "examples": placeholders[:10],
                       "message": "Template placeholders are still in the document."})

    prev = 0
    for h in after["outline"]:
        if prev and h["level"] > prev + 1:
            issues.append({"type": "heading_level_skip", "severity": "warning", "paragraph": h["paragraph"],
                           "message": f"Heading '{h['text'][:60]}' jumps from level {prev} to {h['level']}. "
                                      "Use the next level so numbering and the TOC stay correct."})
        prev = h["level"]

    if after.get("field_codes", {}).get("TOC") and [o["text"] for o in before["outline"]] != [o["text"] for o in after["outline"]]:
        issues.append({"type": "toc_stale", "severity": "warning",
                       "message": "Headings changed and the document has a table of contents: call word_update_fields."})

    if after.get("field_codes", {}).get("SEQ") and before.get("field_codes", {}).get("SEQ") != after["field_codes"]["SEQ"]:
        issues.append({"type": "fields_need_update", "severity": "info",
                       "message": "Captions changed: call word_update_fields so numbers and cross-references are correct."})

    changes = {
        "fields_filled": filled,
        "fields_cleared": emptied,
        "tables": [before["stats"]["tables"], after["stats"]["tables"]],
        "headings_added": [o["text"] for o in after["outline"] if o["text"] not in {x["text"] for x in before["outline"]}],
        "visuals": after.get("visuals", {}),
    }
    return {
        "ok": not any(i["severity"] == "error" for i in issues) and not issues,
        "summary": {"fields_total": after["stats"]["fields_total"], "fields_empty": after["stats"]["fields_empty"],
                    "fields_filled_now": len(filled)},
        "changes": changes,
        "issues": issues,
    }
