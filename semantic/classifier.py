"""Deterministic heading classification for DOCX paragraphs.

Consumes the resolved style metadata (StyleRegistry) and assigns a heading
level + a single authoritative heading id to each Paragraph.

Design rules (preventing the old TOC false-positive problem):
  * A paragraph is a heading ONLY by DOCX structural metadata:
      1. explicit outline level on the paragraph (w:pPr/w:outlineLvl)
      2. the resolved paragraph style's outline level / heading semantics
      3. style inheritance (based on a heading style)
  * NO visual inference (font size, bold, color, numbering text).
  * Styles whose id/name contain 'toc'/'header'/'footer' are NEVER headings,
    so a Word TOC field or page furniture cannot corrupt heading detection.
  * The TOC is generated separately from the resolved headings.

The assigned heading_id is the ONE authoritative id: it is used both as the
HTML element id and as the TOC navigation target, so links can never drift.
"""

import re
from typing import List, Optional

# Styles that must never be promoted to headings even if their name happens to
# contain a heading-like token. Covers Word TOC fields and page furniture.
_BLOCKED = ("toc", "header", "footer", "footnote", "endnote")


def _is_blocked(style_id: str, name: str) -> bool:
    blob = (style_id + " " + name).lower()
    return any(b in blob for b in _BLOCKED)


def _heading_text(para) -> str:
    return "".join(r.text for r in para.runs)


def _slug(text: str) -> str:
    text = (text or "").lower().strip()
    text = re.sub(r'[^a-z0-9]+', '-', text)
    return text.strip('-') or "section"


def classify_paragraphs(paragraphs: List, registry, used_ids: Optional[dict] = None) -> List:
    """Resolve heading levels + ids in place and return the same list.

    Two-pass model:
      Pass 1 (done by StyleRegistry): resolve every style's metadata.
      Pass 2 (here): classify each paragraph from its resolved style.
    """
    if used_ids is None:
        used_ids = {}
    for para in paragraphs:
        sid = para.style_name
        resolved = registry.resolve(sid) if registry is not None else None
        name = resolved.name if resolved else ""
        level = None

        # Priority 1: explicit outline level on the paragraph itself.
        if para.outline_level is not None and 0 <= para.outline_level <= 8:
            level = para.outline_level + 1

        # Priority 2/3: resolved style (outline or inheritance or name match).
        if level is None and resolved is not None and resolved.is_heading:
            level = resolved.heading_level

        # Never promote blocked styles (TOC / header / footer).
        if level is not None and _is_blocked(sid, name):
            level = None

        if level is None:
            para.heading_level = None
            para.heading_id = None
            continue

        # Clamp to a sane range; HTML only renders h1-h6 but we keep the
        # semantic level for hierarchy/nesting.
        level = max(1, min(level, 9))
        para.heading_level = level

        base = "h%d-%s" % (level, _slug(_heading_text(para)))
        hid = base
        n = 2
        while hid in used_ids:
            hid = "%s-%d" % (base, n)
            n += 1
        used_ids[hid] = True
        para.heading_id = hid
    return paragraphs
