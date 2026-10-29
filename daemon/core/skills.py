"""Skill loader and router.

Skills are folders containing a ``SKILL.md`` with a small YAML frontmatter::

    ---
    name: template-form-filling
    description: One line saying when to use it
    hosts: [word]
    always_on: false
    triggers:
      doc_types: [form_template]
      keywords: [compila, fill]
    ---
    Markdown body ...

Built-in skills live in ``daemon/skills``; user skills in
``%LOCALAPPDATA%\\GeminiAntigravityOffice\\skills`` override built-ins by name.
No PyYAML dependency: a minimal parser supports exactly the fields above.
"""
from __future__ import annotations

import logging
import os
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

log = logging.getLogger(__name__)

BUILTIN_DIR = Path(__file__).resolve().parents[1] / "skills"

DOC_TYPE_SCORE = 2
KEYWORD_SCORE = 3


@dataclass
class Skill:
    name: str
    description: str
    hosts: List[str] = field(default_factory=list)
    always_on: bool = False
    doc_types: List[str] = field(default_factory=list)
    keywords: List[str] = field(default_factory=list)
    body: str = ""
    path: Path = Path()


# ------------------------------------------------------------------ frontmatter

def _scalar(raw: str) -> Any:
    v = raw.strip()
    if not v:
        return ""
    if v.startswith("[") and v.endswith("]"):
        inner = v[1:-1].strip()
        return [_scalar(p) for p in _split_inline(inner)] if inner else []
    if len(v) >= 2 and v[0] == v[-1] and v[0] in "\"'":
        return v[1:-1]
    low = v.lower()
    if low in ("true", "yes"):
        return True
    if low in ("false", "no"):
        return False
    return v


def _split_inline(s: str) -> List[str]:
    """Split an inline list on commas, respecting quotes."""
    out, cur, q = [], "", None
    for ch in s:
        if q:
            cur += ch
            if ch == q:
                q = None
        elif ch in "\"'":
            q = ch
            cur += ch
        elif ch == ",":
            out.append(cur.strip())
            cur = ""
        else:
            cur += ch
    if cur.strip():
        out.append(cur.strip())
    return out


def _parse_block(lines: List[Tuple[int, str]]) -> Any:
    """Parse indented (indent, text) lines into a dict or a list."""
    if not lines:
        return {}
    if lines[0][1].startswith("- "):
        return [_scalar(t[2:]) for _, t in lines if t.startswith("- ")]
    result: Dict[str, Any] = {}
    i = 0
    base = lines[0][0]
    while i < len(lines):
        ind, text = lines[i]
        if ind != base or ":" not in text:
            i += 1
            continue
        key, _, rest = text.partition(":")
        key = key.strip()
        if rest.strip():
            result[key] = _scalar(rest)
            i += 1
            continue
        j = i + 1
        child = []
        while j < len(lines) and lines[j][0] > base:
            child.append(lines[j])
            j += 1
        result[key] = _parse_block(child) if child else ""
        i = j
    return result


def parse_frontmatter(text: str) -> Tuple[Dict[str, Any], str]:
    """Return (metadata, body). Raises ValueError when no frontmatter is present."""
    text = text.lstrip("﻿")
    m = re.match(r"^---[ \t]*\r?\n(.*?)\r?\n---[ \t]*(?:\r?\n|$)", text, re.S)
    if not m:
        raise ValueError("missing YAML frontmatter")
    lines = []
    for raw in m.group(1).splitlines():
        if not raw.strip() or raw.lstrip().startswith("#"):
            continue
        lines.append((len(raw) - len(raw.lstrip(" ")), raw.strip()))
    meta = _parse_block(lines)
    if not isinstance(meta, dict):
        raise ValueError("frontmatter must be a mapping")
    return meta, text[m.end():]


def _as_list(v: Any) -> List[str]:
    if v in (None, ""):
        return []
    if isinstance(v, list):
        return [str(x) for x in v]
    return [str(v)]


def _skill_from_file(path: Path) -> Skill:
    meta, body = parse_frontmatter(path.read_text(encoding="utf-8"))
    name = str(meta.get("name") or path.parent.name).strip()
    triggers = meta.get("triggers") if isinstance(meta.get("triggers"), dict) else {}
    return Skill(
        name=name,
        description=str(meta.get("description", "")).strip(),
        hosts=_as_list(meta.get("hosts")),
        always_on=meta.get("always_on") is True,
        doc_types=_as_list(triggers.get("doc_types")),
        keywords=_as_list(triggers.get("keywords")),
        body=body.strip("\n") + "\n",
        path=path,
    )


# ------------------------------------------------------------------ loading

def user_skills_dir() -> Optional[Path]:
    base = os.environ.get("LOCALAPPDATA")
    return Path(base) / "GeminiAntigravityOffice" / "skills" if base else None


def default_dirs() -> List[Path]:
    dirs = [BUILTIN_DIR]
    u = user_skills_dir()
    if u:
        dirs.append(u)
    return dirs


def load_skills(dirs: Optional[List[Path]] = None) -> List[Skill]:
    """Load skills from ``dirs`` in order; later directories override earlier ones by name."""
    found: Dict[str, Skill] = {}
    for d in dirs if dirs is not None else default_dirs():
        d = Path(d)
        if not d.is_dir():
            continue
        for f in sorted(d.glob("*/SKILL.md")):
            try:
                s = _skill_from_file(f)
            except (OSError, ValueError, UnicodeDecodeError) as exc:
                log.warning("skipping skill %s: %s", f, exc)
                continue
            found[s.name] = s
    return sorted(found.values(), key=lambda s: s.name)


# ------------------------------------------------------------------ routing

def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    return "".join(c for c in s if not unicodedata.combining(c)).lower()


def _kw_hit(kw: str, text: str) -> bool:
    k = _norm(kw).strip()
    if not k:
        return False
    pre = r"\b" if k[0].isalnum() else ""
    post = r"\b" if k[-1].isalnum() else ""
    return re.search(pre + re.escape(k) + post, text) is not None


def score(skill: Skill, doc_type: Optional[str], request: str) -> int:
    text = _norm(request)
    s = sum(KEYWORD_SCORE for k in skill.keywords if _kw_hit(k, text))
    if doc_type and _norm(doc_type) in {_norm(d) for d in skill.doc_types}:
        s += DOC_TYPE_SCORE
    return s


_SLASH = re.compile(r"^\s*/([A-Za-z0-9][A-Za-z0-9_-]*)\b\s*")


def parse_forced(request: str, known: Optional[Iterable[str]] = None) -> Tuple[List[str], str]:
    """Extract leading ``/skill-name`` tokens. With ``known``, only those names are stripped."""
    known_set = {k.lower() for k in known} if known is not None else None
    forced: List[str] = []
    rest = request or ""
    while True:
        m = _SLASH.match(rest)
        if not m:
            break
        name = m.group(1).lower()
        if known_set is not None and name not in known_set:
            break
        if name not in forced:
            forced.append(name)
        rest = rest[m.end():]
    return forced, (rest.strip() if forced else request)


def route(skills: List[Skill], host: str, doc_type: Optional[str], request: str,
          forced: Optional[List[str]] = None, max_full: int = 3) -> Dict[str, List[Skill]]:
    """Choose skills whose full body goes into the prompt; list the rest by name.

    Order: always_on skills, then forced skills, then up to ``max_full`` scored skills
    (doc_type match + keyword hits, accent/case-insensitive). Forced skills count
    toward ``max_full`` but are never dropped.
    """
    pool = [s for s in skills if not s.hosts or host.lower() in [h.lower() for h in s.hosts]]
    by = {s.name: s for s in pool}
    full: List[Skill] = [s for s in pool if s.always_on]
    chosen = {s.name for s in full}
    extra = 0
    for name in forced or []:
        s = by.get(name)
        if s and s.name not in chosen:
            full.append(s)
            chosen.add(s.name)
            extra += 1
    ranked = sorted(((score(s, doc_type, request), s.name) for s in pool if s.name not in chosen),
                    key=lambda t: (-t[0], t[1]))
    for sc, name in ranked:
        if extra >= max_full or sc <= 0:
            break
        full.append(by[name])
        chosen.add(name)
        extra += 1
    listed = [s for s in pool if s.name not in chosen]
    return {"full": full, "listed": listed}


def render_for_prompt(routed: Dict[str, List[Skill]]) -> str:
    parts = ["=== ACTIVE SKILLS (expert procedures for this turn) ==="]
    for s in routed.get("full", []):
        parts.append(f"\n### Skill: {s.name}\n{s.body.strip()}\n")
    listed = routed.get("listed", [])
    if listed:
        parts.append("=== OTHER SKILLS (call skill_load(name) to read one before using it) ===")
        parts.extend(f"- {s.name} — {s.description}" for s in listed)
    return "\n".join(parts) + "\n"
