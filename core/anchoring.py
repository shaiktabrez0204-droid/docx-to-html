"""Nearest-block association for floating (wp:anchor) images.

A wp:anchor lives inside a w:r inside a w:p, so the OOXML already places the
drawing in a specific paragraph. This pass turns that structural fact into an
explicit, queryable association: each floating Image gets a ``nearest_block_id``
(referencing a Paragraph.block_id) and a ``nearest_block_confidence`` in 0..1.

Confidence policy
-----------------
* The anchor's own paragraph carries text -> the image is unambiguously anchored
  to that paragraph (confidence 0.95).
* The anchor sits in an EMPTY paragraph (only the drawing, no runs) -> the
  genuine anchor context is ambiguous, so we associate with the nearest
  non-empty neighbour and LOWER the confidence (0.6). We never fabricate a
  precise relationship; we record the uncertainty instead.
* No neighbour at all -> confidence 0.3 (kept, but flagged low).

The dataclass fields are the single source of truth; nothing here rewrites the
image extraction pipeline. Adversarial/unsupported geometry is preserved on the
Image (see adapter.ooxml_parser) and only the block association is computed here.
"""

from typing import List, Optional

from core.model import Paragraph, Image, Run


def _paragraph_has_text(p: Paragraph) -> bool:
    """True when a paragraph carries any visible run text."""
    if any(r.text for r in p.runs):
        return True
    return any(isinstance(c, Run) and c.text for c in p.content)


def _floating_images(paragraphs: List[Paragraph]):
    for p in paragraphs:
        if p.block_id is None:
            continue
        for c in p.content:
            if isinstance(c, Image) and c.wrap_type == "anchor":
                yield p, c


def associate_floating_images(paragraphs: List[Paragraph]) -> List[Paragraph]:
    """Assign block_id to every paragraph, then associate each floating image.

    Mutates ``Paragraph.block_id`` and the floating ``Image``'s
    ``nearest_block_id`` / ``nearest_block_confidence`` in place. Returns the
    same paragraph list for call chaining.
    """
    for i, p in enumerate(paragraphs):
        if p.block_id is None:
            p.block_id = "blk-%d" % i

    last_nonempty_id: Optional[str] = None
    for i, p in enumerate(paragraphs):
        has_text = _paragraph_has_text(p)
        if has_text:
            last_nonempty_id = p.block_id

        for c in p.content:
            if not (isinstance(c, Image) and c.wrap_type == "anchor"):
                continue
            # The containing paragraph is the structural anchor context.
            if has_text:
                c.nearest_block_id = p.block_id
                c.nearest_block_confidence = 0.95
                continue

            # Anchor in an empty paragraph: associate with nearest non-empty
            # block (prefer the preceding one, then look ahead). Low confidence
            # by design - we do NOT invent a precise relationship.
            target_id = last_nonempty_id
            if target_id is None:
                for j in range(i + 1, len(paragraphs)):
                    if _paragraph_has_text(paragraphs[j]):
                        target_id = paragraphs[j].block_id
                        break
            c.nearest_block_id = target_id
            c.nearest_block_confidence = 0.6 if target_id else 0.3

    return paragraphs


def low_confidence_associations(paragraphs: List[Paragraph]):
    """Yield (Image, nearest_block_id, confidence) for any weak association.

    Lets tests/reports surface ambiguous placements instead of hiding them.
    """
    for _p, img in _floating_images(paragraphs):
        if img.nearest_block_confidence is not None and img.nearest_block_confidence < 0.9:
            yield img, img.nearest_block_id, img.nearest_block_confidence
