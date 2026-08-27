"""Style resolution for DOCX OOXML.

Adapted concept from x2 docflow's semantic/style_resolver.py:
  - Two-pass resolution (resolve metadata first, then classify)
  - BasedOn inheritance chain with a cycle guard (depth limit)
but rebuilt for DOCX/OOXML instead of IDML:

  * x2 parsed IDML ParagraphStyle attributes (PointSize/FontStyle/SpaceBefore).
    We instead read OOXML <w:style> (styleId/name/basedOn/w:outlineLvl).
  * x2 used font-size/bold/space SCORING as a heading fallback. That is a
    VISUAL heuristic and is intentionally NOT used here - DOCX gives us
    structural metadata (w:outlineLvl and Heading style semantics) which is
    authoritative and avoids the old TOC false-positive problem.

Resolution priority for a heading level:
  1. explicit outline level (w:outlineLvl, own or inherited via BasedOn)
  2. heading style semantics (styleId/name match)
  3. style inheritance (a style based on a heading style inherits its level)
"""

import re
from dataclasses import dataclass
from typing import Dict, List, Optional

# Heading-name patterns, applied only as a structural fallback for CUSTOM
# heading styles that lack an outline level and are not based on a heading.
_HEADING_NAME_RES = [
    re.compile(r'heading\s*(\d+)', re.IGNORECASE),
    re.compile(r'\bh(\d)\b', re.IGNORECASE),
    re.compile(r'hdr\s*(\d+)', re.IGNORECASE),
    re.compile(r'headline\s*(\d+)', re.IGNORECASE),
]


@dataclass
class ResolvedStyle:
    style_id: str
    name: str = ""
    type: str = "paragraph"
    outline_level: Optional[int] = None   # 0-8 resolved (own or inherited)
    heading_level: Optional[int] = None   # 1-9 if heading else None
    is_heading: bool = False
    source: str = ""                       # outline | name_pattern | none
    num_id: Optional[str] = None           # effective w:numPr/w:numId (own or inherited)
    num_ilvl: Optional[int] = None         # effective w:numPr/w:ilvl (own or inherited)
    # Paragraph layout (w:pPr)
    alignment: Optional[str] = None
    indent_left: Optional[int] = None
    indent_right: Optional[int] = None
    indent_first_line: Optional[int] = None
    indent_hanging: Optional[int] = None
    spacing_before: Optional[int] = None
    spacing_after: Optional[int] = None
    line_spacing: Optional[int] = None
    line_spacing_rule: Optional[str] = None
    # Run typography (w:rPr)
    font_family: Optional[str] = None
    font_size: Optional[int] = None
    font_color: Optional[str] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[str] = None
    superscript: Optional[bool] = None
    subscript: Optional[bool] = None


class StyleRegistry:
    """Resolves inherited style properties from raw StyleDef objects.

    Builds a fully-resolved map of every style. Outline levels (and therefore
    heading levels) are propagated down BasedOn chains with a cycle guard.
    """

    def __init__(self, styles: List):
        self._raw: Dict[str, object] = {}
        for s in styles:
            self._raw[s.style_id] = s
        self._resolved: Dict[str, ResolvedStyle] = {}
        for sid in list(self._raw.keys()):
            self._resolve(sid)

    def _resolve(self, sid: str, depth: int = 0, visiting=None) -> Optional[ResolvedStyle]:
        if sid in self._resolved:
            return self._resolved[sid]
        if depth > 10 or sid not in self._raw:
            return None
        if visiting is None:
            visiting = set()
        if sid in visiting:
            return None  # BasedOn cycle -> treat as unresolved
        visiting.add(sid)

        s = self._raw[sid]
        parent = None
        if s.based_on and s.based_on != sid and s.based_on in self._raw:
            parent = self._resolve(s.based_on, depth + 1, visiting)
        visiting.discard(sid)

        # Outline level: own value wins; otherwise inherit from parent.
        outline = s.outline_level
        if outline is None and parent is not None:
            outline = parent.outline_level

        num_id = s.num_id
        if num_id is None and parent is not None:
            num_id = parent.num_id
        num_ilvl = s.num_ilvl
        if num_ilvl is None and parent is not None:
            num_ilvl = parent.num_ilvl

        # Paragraph layout inheritance
        def _inherit(attr):
            v = getattr(s, attr)
            if v is None and parent is not None:
                return getattr(parent, attr)
            return v

        alignment = _inherit("alignment")
        indent_left = _inherit("indent_left")
        indent_right = _inherit("indent_right")
        indent_first_line = _inherit("indent_first_line")
        indent_hanging = _inherit("indent_hanging")
        spacing_before = _inherit("spacing_before")
        spacing_after = _inherit("spacing_after")
        line_spacing = _inherit("line_spacing")
        line_spacing_rule = _inherit("line_spacing_rule")
        font_family = _inherit("font_family")
        font_size = _inherit("font_size")
        font_color = _inherit("font_color")
        bold = _inherit("bold")
        italic = _inherit("italic")
        underline = _inherit("underline")
        superscript = _inherit("superscript")
        subscript = _inherit("subscript")

        level = None
        source = ""
        if outline is not None and 0 <= outline <= 8:
            level = outline + 1
            source = "outline"
        if level is None:
            nm = _match_heading_name(s.style_id, s.name)
            if nm is not None:
                level = nm
                source = "name_pattern"

        res = ResolvedStyle(
            style_id=sid,
            name=s.name,
            type=s.style_type,
            outline_level=outline,
            heading_level=level,
            is_heading=level is not None,
            source=source,
            num_id=num_id,
            num_ilvl=num_ilvl,
            alignment=alignment,
            indent_left=indent_left,
            indent_right=indent_right,
            indent_first_line=indent_first_line,
            indent_hanging=indent_hanging,
            spacing_before=spacing_before,
            spacing_after=spacing_after,
            line_spacing=line_spacing,
            line_spacing_rule=line_spacing_rule,
            font_family=font_family,
            font_size=font_size,
            font_color=font_color,
            bold=bold,
            italic=italic,
            underline=underline,
            superscript=superscript,
            subscript=subscript,
        )
        self._resolved[sid] = res
        return res

    def resolve(self, style_id: str) -> Optional[ResolvedStyle]:
        return self._resolved.get(style_id)


def _match_heading_name(style_id: str, name: str) -> Optional[int]:
    """Fallback: derive a heading level from a style name/styleId.

    Only used when no outline level is available. Returns the level (1-9) or
    None. Deliberately conservative: requires an explicit 'heading'/'hdr'/'hN'
    signal so arbitrary styles are never promoted.
    """
    for blob in (style_id, name):
        for pat in _HEADING_NAME_RES:
            m = pat.search(blob or "")
            if m:
                try:
                    return int(m.group(1))
                except (IndexError, ValueError):
                    return 1
    low = (name or "").lower()
    if "heading" in low:
        return 1
    return None


def resolve_style(style_id: str, registry=None):
    """Convenience entry point.

    With a registry, returns a resolved dict for the style. Without one
    (legacy / standalone use) it performs a best-effort name-only
    classification so callers still receive a non-None result for heading
    style ids such as "Heading1".
    """
    if registry is not None:
        r = registry.resolve(style_id)
        if r is None:
            return None
        return {
            "name": r.name,
            "type": r.type,
            "level": r.heading_level,
            "is_heading": r.is_heading,
            "source": r.source,
        }
    lvl = _match_heading_name(style_id, style_id)
    return {
        "name": style_id,
        "type": "paragraph",
        "level": lvl,
        "is_heading": lvl is not None,
        "source": "name_pattern" if lvl is not None else "none",
    }
