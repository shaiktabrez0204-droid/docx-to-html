"""Structured numbering: OOXML metadata -> numbering path -> cross-validation.

This module is a SECOND, INDEPENDENT structural signal. It does NOT classify
headings and never mutates a heading's style-derived level.

Pipeline:
  word/numbering.xml  -> NumberingModel (numId -> abstractNumId -> lvl defs)
  paragraph w:numPr   -> (numId, ilvl) reference stored on the Paragraph
  document walk       -> numbering_path [1], [1,1], [1,1,1] ... (stateful)
  cross-validate      -> heading level vs numbering level/path
  validate hierarchy  -> numbering nesting vs heading nesting

All numbering comes from OOXML metadata. Visible text (e.g. "1.1") is never the
source of truth; it is, at most, a fallback/cross-check elsewhere (toc.extract_number).
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional

from core.model import NumberingModel, Paragraph
from semantic.hierarchy import HeadingNode
from semantic.style_resolver import StyleRegistry


@dataclass
class NumberingValidation:
    """Per-heading result of heading-vs-numbering cross-validation."""
    heading_id: Optional[str]
    text: str
    has_numbering: bool
    numbering_level: Optional[int]   # num_ilvl (0-based)
    heading_level: Optional[int]
    numbering_path: Optional[List[int]]
    consistent: Optional[bool]
    issue: Optional[str] = None


class NumberingResolver:
    """Resolves (numId, ilvl) references into paths using numbering.xml metadata."""

    def __init__(self, model: NumberingModel):
        self.model = model

    def resolve(self, paragraphs: List[Paragraph], registry: Optional[StyleRegistry] = None) -> List[Paragraph]:
        """Assign numbering_format/text_pattern/path to each paragraph in place.

        Maintains per-numId counter state across the document walk so that the
        resolved path reflects real OOXML numbering sequence (including valid
        chapter resets, which are just deeper-level resets on a shallower change).
        Effective (numId, ilvl) = direct paragraph w:numPr OR the resolved
        style's w:numPr (inherited through BasedOn).
        """
        counters: Dict[str, Dict[int, int]] = {}

        for p in paragraphs:
            num_id = p.num_id
            num_ilvl = p.num_ilvl
            if registry is not None:
                res = registry.resolve(p.style_name)
                if res is not None:
                    if num_id is None:
                        num_id = res.num_id
                    if num_ilvl is None:
                        num_ilvl = res.num_ilvl

            # Single-level list styles (e.g. ListNumber/ListBullet) carry numId
            # in the style but no explicit ilvl in styles.xml or document.xml.
            # OOXML singleLevel meaning is ilvl 0 – defaulting avoids silent loss.
            if num_id is not None and num_ilvl is None:
                num_ilvl = 0

            p.num_id = num_id
            p.num_ilvl = num_ilvl

            if num_id is None or num_ilvl is None:
                p.numbering_path = None
                continue

            lvl = self.model.resolve_level(num_id, num_ilvl)
            if lvl is not None:
                p.numbering_format = lvl.num_fmt
                p.numbering_text_pattern = lvl.lvl_text
                start = lvl.start
                # Per-level numFmt for each position of the path (0..num_ilvl)
                # so multi-level lvlText ("%1.%2.%3") can format every level
                # with its own numFmt. Reuses the already-built NumberingModel.
                formats = []
                for k in range(0, num_ilvl + 1):
                    lk = self.model.resolve_level(num_id, k)
                    formats.append(lk.num_fmt if lk is not None else "decimal")
                p.numbering_level_formats = formats
            else:
                p.numbering_format = None
                p.numbering_text_pattern = None
                p.numbering_level_formats = None
                start = 1

            if num_id not in counters:
                counters[num_id] = {}
            c = counters[num_id]

            # A shallower level appearing resets all deeper counters (Word semantics).
            for k in [k for k in c if k > num_ilvl]:
                del c[k]

            c[num_ilvl] = c.get(num_ilvl, start - 1) + 1
            # skipped parent levels render as 0 in OOXML lvlText
            p.numbering_path = [c.get(k, 0) for k in range(0, num_ilvl + 1)]

        return paragraphs


def cross_validate(paragraphs: List[Paragraph]) -> List[NumberingValidation]:
    """Compare style-derived heading level against structured numbering.

    Sets paragraph.numbering_consistent. Unnumbered headings are valid (consistent
    stays None). Never changes heading_level.
    """
    results: List[NumberingValidation] = []
    for p in paragraphs:
        if p.heading_level is None:
            continue
        text = "".join(r.text for r in p.runs)
        if p.num_ilvl is None or p.numbering_path is None:
            p.numbering_consistent = None
            results.append(NumberingValidation(
                p.heading_id, text, False, p.num_ilvl, p.heading_level, None, None))
            continue

        expected_ilvl = p.heading_level - 1
        depth_ok = len(p.numbering_path) == p.heading_level
        ilvl_ok = p.num_ilvl == expected_ilvl
        ok = depth_ok and ilvl_ok
        p.numbering_consistent = ok

        issue = None
        if not ok:
            if not ilvl_ok:
                issue = "numbering_ilvl_%s_mismatch_heading_%d" % (p.num_ilvl, p.heading_level)
            else:
                issue = "numbering_path_depth_%d_mismatch_heading_%d" % (len(p.numbering_path), p.heading_level)

        results.append(NumberingValidation(
            p.heading_id, text, True, p.num_ilvl, p.heading_level, p.numbering_path, ok, issue))
    return results


def validate_hierarchy(nodes: List[HeadingNode]) -> List[dict]:
    """Cross-check heading hierarchy against numbering nesting.

    Flags: numbering path not nested under its parent, and duplicate numbering
    paths. Does not mutate the hierarchy; only reports.
    """
    issues: List[dict] = []
    seen: Dict[tuple, str] = {}

    def walk(node: HeadingNode, parent_path: Optional[List[int]]):
        p = node.paragraph
        path = p.numbering_path
        if path is not None:
            if parent_path is not None and path[:len(parent_path)] != parent_path:
                issues.append({
                    "type": "numbering_not_nested",
                    "heading_id": node.heading_id,
                    "text": node.text,
                    "detail": "numbering path %s is not nested under parent path %s" % (path, parent_path),
                })
            key = tuple(path)
            if key in seen and seen[key] != node.heading_id:
                issues.append({
                    "type": "duplicate_numbering_path",
                    "heading_id": node.heading_id,
                    "text": node.text,
                    "detail": "numbering path %s used by multiple headings" % (path,),
                })
            else:
                seen[key] = node.heading_id
        child_parent = path if path is not None else parent_path
        for c in node.children:
            walk(c, child_parent)

    for n in nodes:
        walk(n, None)
    return issues
