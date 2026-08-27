"""Table-of-contents generation from resolved headings.

The TOC is GENERATED from the resolved heading hierarchy - it does NOT parse
the Word TOC field (which is handled separately and must never corrupt heading
detection). Every TOC entry links to the SAME authoritative heading id that
was assigned to the heading element, so navigation can never drift.

Adapted from x2 docflow's semantic/toc.py: the anomaly-validation concept
(duplicate / out-of-order numbered entries) is reused as detect_heading_anomalies
for QA. The x2 TOC-text-parsing (regex on dot leaders / page numbers) is NOT
copied - we have no need to reverse-engineer TOC text because headings are
already resolved structurally.
"""

import re
from typing import List, Optional

from semantic.hierarchy import HeadingNode
from core.model import format_numbering_label


def build_toc(nodes: List[HeadingNode]) -> List[dict]:
    """Build a nested TOC mirroring the heading hierarchy.

    Each entry carries: level, text, id (authoritative heading id), target
    (navigation target = same id), number (legacy text-derived outline number,
    used only by internal anomaly QA), numbering (resolved path), and
    numbering_label (the VISIBLE label rendered in the TOC, derived from the
    SAME paragraph fields the heading renderer uses - one source of truth).
    """
    entries = []
    for n in nodes:
        p = n.paragraph
        entries.append({
            "level": n.level,
            "text": n.text,
            "id": n.heading_id,
            "target": n.heading_id,
            "number": extract_number(n.text),
            "numbering": p.numbering_path,
            "numbering_label": format_numbering_label(
                p.numbering_path, p.numbering_level_formats, p.numbering_text_pattern),
            "children": build_toc(n.children),
        })
    return entries


def flatten_toc(entries: List[dict]) -> List[dict]:
    """Depth-first flat list of TOC entries (document order)."""
    out: List[dict] = []
    for e in entries:
        out.append(e)
        out.extend(flatten_toc(e.get("children", [])))
    return out


def extract_number(text: str) -> Optional[str]:
    """Extract a leading outline number like '1', '1.2', '3.4.1' if present."""
    m = re.match(r'^\s*(\d+(?:\.\d+)*)\b', text or "")
    return m.group(1) if m else None


def detect_heading_anomalies(entries: List[dict]) -> List[dict]:
    """Reused x2 concept: flag duplicate / out-of-order numbered headings.

    Only headings with an explicit leading number are checked (unnumbered
    headings are perfectly valid and skipped). Returns a list of anomalies.
    """
    numbered = [e for e in entries if e.get("number")]
    anomalies: List[dict] = []
    seen = {}
    for e in numbered:
        num = e["number"]
        if num in seen:
            anomalies.append({"type": "duplicate_number", "number": num, "text": e["text"]})
        else:
            seen[num] = True
    # out-of-order: a child number lexically smaller than an earlier sibling
    prev_by_parent = {}
    for e in numbered:
        parts = tuple(int(p) for p in e["number"].split("."))
        parent = parts[:-1]
        if parent in prev_by_parent and parts < prev_by_parent[parent]:
            anomalies.append({"type": "out_of_order", "number": e["number"], "text": e["text"]})
        prev_by_parent[parent] = parts
    return anomalies
