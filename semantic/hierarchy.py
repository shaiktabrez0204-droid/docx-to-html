"""Heading hierarchy construction (stack-based, arbitrary depth).

Adapted concept from x2 docflow's semantic/hierarchy.py: a stack walk that
attaches each heading to its nearest ancestor. Reused because it naturally
PRESERVES skipped levels (e.g. H1 -> H3 -> H2 stays H1 > H3, H1 > H2) without
inventing missing levels.

Deliberately NOT copied from x2:
  * build_nav() hard-coded depth=2 limit  -> we support H1-H6+.
  * page/coordinate/sidebar/TOC-alignment machinery -> irrelevant to DOCX,
    where heading level comes from resolved style metadata, not geometry.
"""

from dataclasses import dataclass, field
from typing import List, Optional

from core.model import Paragraph


@dataclass
class HeadingNode:
    paragraph: Paragraph
    level: int
    text: str
    heading_id: str
    order: int
    children: List['HeadingNode'] = field(default_factory=list)


def build_hierarchy(paragraphs: List[Paragraph]) -> List[HeadingNode]:
    """Return the top-level HeadingNode list (document order).

    Each node also carries its descendant chain so the TOC can be rendered
    as a nested navigation tree.
    """
    headings = [p for p in paragraphs if p.heading_level is not None]
    roots: List[HeadingNode] = []
    stack: List[HeadingNode] = []
    order = 0

    for p in headings:
        node = HeadingNode(
            paragraph=p,
            level=p.heading_level,
            text="".join(r.text for r in p.runs),
            heading_id=p.heading_id,
            order=order,
        )
        order += 1

        # Pop until the top of the stack is a strict ancestor (level < node.level).
        while stack and stack[-1].level >= node.level:
            stack.pop()

        parent = stack[-1] if stack else None
        if parent is None:
            roots.append(node)
        else:
            parent.children.append(node)
        stack.append(node)

    return roots


def flatten_hierarchy(nodes: List[HeadingNode]) -> List[HeadingNode]:
    """Depth-first flat list of all heading nodes (document order)."""
    out: List[HeadingNode] = []
    for n in nodes:
        out.append(n)
        out.extend(flatten_hierarchy(n.children))
    return out
