"""
Runs document perception before every agent turn: the taskpane sends the document OOXML,
selection and headers/footers; we build the snapshot off the event loop, cache it per
document, and return the Document Map for the model plus a short summary for the UI.
"""

import asyncio
import hashlib
from dataclasses import dataclass
from typing import Any, Dict, Optional

from .word_snapshot import build_snapshot, render_map

TYPE_LABELS = {
    "form_template": "Form / template to fill",
    "report": "Report",
    "contract": "Contract",
    "letter": "Letter",
    "minutes": "Meeting minutes",
    "blank": "Blank document",
    "document": "Document",
}


@dataclass
class PerceptionResult:
    snapshot: Optional[Dict[str, Any]]
    map_text: str
    summary: Dict[str, Any]
    cached: bool = False


class Perception:
    def __init__(self, map_budget: int = 12000, cache_size: int = 16):
        self.map_budget = map_budget
        self.cache_size = cache_size
        self._cache: Dict[str, Dict[str, Any]] = {}

    async def analyze(self, payload: Dict[str, Any]) -> PerceptionResult:
        ooxml = payload.get("document_ooxml") or ""
        if not ooxml.strip():
            return PerceptionResult(
                None,
                "=== DOCUMENT MAP ===\nThe document could not be read before this turn. "
                "Call word_get_document_outline / word_read_* tools before editing.",
                {"doc_type": "unknown", "headline": "Document could not be read", "fields_empty": 0},
            )

        extras = {
            "selection": payload.get("selection") or None,
            "headers_footers": payload.get("headers_footers") or [],
            "document": {"name": payload.get("doc_title") or "", "url": payload.get("doc_url") or ""},
        }
        key = hashlib.sha1(
            (ooxml + repr(sorted((k, str(v)) for k, v in extras.items()))).encode("utf-8", "replace")
        ).hexdigest()
        hit = self._cache.get(key)
        if hit:
            return PerceptionResult(hit["snapshot"], hit["map"], hit["summary"], cached=True)

        snapshot = await asyncio.to_thread(build_snapshot, ooxml, extras)
        map_text = render_map(snapshot, budget_chars=self.map_budget)
        summary = summarize(snapshot)
        if len(self._cache) >= self.cache_size:
            self._cache.pop(next(iter(self._cache)))
        self._cache[key] = {"snapshot": snapshot, "map": map_text, "summary": summary}
        return PerceptionResult(snapshot, map_text, summary)


def summarize(s: Dict[str, Any]) -> Dict[str, Any]:
    st = s["stats"]
    parts = [TYPE_LABELS.get(s["doc_type"], s["doc_type"])]
    if st["fields_total"]:
        parts.append(f"{st['fields_total']} fields ({st['fields_empty']} empty)")
    if st["tables"]:
        parts.append(f"{st['tables']} tables")
    if s["outline"]:
        parts.append(f"{len(s['outline'])} headings")
    if s["visuals"]:
        parts.append(", ".join(f"{v} {k}" for k, v in s["visuals"].items()))
    if s["language"] != "unknown":
        parts.append({"it": "Italian", "en": "English", "fr": "French"}.get(s["language"], s["language"]))
    if s["styles"]["direct_formatting"]:
        parts.append("direct formatting")
    return {
        "doc_type": s["doc_type"],
        "language": s["language"],
        "fields_total": st["fields_total"],
        "fields_empty": st["fields_empty"],
        "tables": st["tables"],
        "words": st["words"],
        "headline": " · ".join(parts),
    }
