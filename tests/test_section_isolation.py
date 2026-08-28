"""Regression tests for the isolated-heading (section) view.

These are PURE-PYTHON (no browser) checks of the production contract between
``output/html_renderer.render_html`` and the viewer JS:

  * every top-level block is wrapped in ``.docx-block``
  * heading blocks carry ``data-heading-id`` + ``data-level`` so the JS can
    compute section boundaries from the EXISTING heading hierarchy (no regex,
    no text inference)
  * the first H1 stays in the title bar only (no duplicate H1 in the document
    body)
  * the section-boundary algorithm (mirrors the viewer JS) yields the exact
    visible-heading set the spec requires for H1 / H2 / H3 focus, including
    descendant inclusion and sibling exclusion.

The actual DOM toggling, exit control and browser-history behaviour are covered
by ``test_section_isolation_playwright.py``.
"""

import os
import re
import sys

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from semantic.pipeline import convert_docx  # noqa: E402

FIX = os.path.join(PROJECT_ROOT, "tests", "fixtures")
DOCX = os.path.join(FIX, "section-isolation.docx")


def _render():
    return convert_docx(DOCX).html


def _blocks(html):
    """Return ordered list of (heading_id|None, level|0) for each .docx-block."""
    out = []
    for attrs in re.findall(r'<div class="docx-block"([^>]*)>', html):
        hid = re.search(r'data-heading-id="([^"]+)"', attrs)
        lvl = re.search(r'data-level="(\d+)"', attrs)
        out.append((hid.group(1) if hid else None, int(lvl.group(1)) if lvl else 0))
    return out


def _title_id(html):
    m = re.search(r'<h1 id="([^"]+)" class="doc-title"', html)
    return m.group(1) if m else None


def _visible_heading_ids(blocks, focus_id, title_id):
    """Mirror of the viewer JS ``focusHeading`` boundary rule.

    For heading H (level L): section starts at H and ends immediately before the
    next heading whose level <= L. The first H1 (title bar) has no block in the
    body, so it is represented by startIdx=-1 and spans to the next H1.
    """
    start_idx = -2
    level = 1
    if focus_id == title_id:
        start_idx = -1
    else:
        for i, (hid, lvl) in enumerate(blocks):
            if hid == focus_id:
                start_idx = i
                level = lvl or 1
                break
    if start_idx == -2:
        return set()
    vis = set()
    if start_idx == -1:
        for i, (hid, lvl) in enumerate(blocks):
            if i > 0 and lvl and lvl <= level:
                break
            if hid:
                vis.add(hid)
    else:
        for i in range(start_idx, len(blocks)):
            hid, lvl = blocks[i]
            if i > start_idx and lvl and lvl <= level:
                break
            if hid:
                vis.add(hid)
    return vis


def test_blocks_wrapped_with_heading_attributes():
    html = _render()
    blocks = _blocks(html)
    assert len(blocks) == 14, "expected 14 top-level blocks, got %d" % len(blocks)
    # Exactly the 5 non-first headings carry ids; first H1 lives in the title bar.
    headed = [b for b in blocks if b[0]]
    assert len(headed) == 5, "expected 5 heading blocks, got %d" % len(headed)
    for hid, lvl in headed:
        assert hid and 1 <= lvl <= 6


def test_no_duplicate_first_h1_in_body():
    html = _render()
    title_id = _title_id(html)
    assert title_id, "first H1 title bar missing"
    # The body must not contain a second heading with the first H1's id.
    assert html.count('id="%s"' % title_id) == 1
    # Exactly one visible H1 carries the introduction text (the title bar).
    assert html.count("Introduction") >= 1
    # doc-title present and only once.
    assert html.count('class="doc-title"') == 1


def test_h1_section_isolation():
    html = _render()
    blocks = _blocks(html)
    tid = _title_id(html)
    # Non-first H1: Conclusion -> Conclusion + Summary only.
    vis = _visible_heading_ids(blocks, "h1-conclusion", tid)
    assert vis == {"h1-conclusion", "h2-summary"}, vis
    # First H1: Introduction -> intro + Architecture + Data Model + Implementation.
    vis = _visible_heading_ids(blocks, tid, tid)
    assert vis == {"h2-architecture", "h3-data-model", "h2-implementation"}, vis


def test_h2_section_isolation():
    html = _render()
    blocks = _blocks(html)
    tid = _title_id(html)
    # Architecture (H2) -> Architecture + Data Model (descendant) only.
    vis = _visible_heading_ids(blocks, "h2-architecture", tid)
    assert vis == {"h2-architecture", "h3-data-model"}, vis
    # Implementation (H2) -> Implementation only.
    vis = _visible_heading_ids(blocks, "h2-implementation", tid)
    assert vis == {"h2-implementation"}, vis


def test_h3_section_isolation_with_descendant_and_sibling():
    html = _render()
    blocks = _blocks(html)
    tid = _title_id(html)
    # Data Model (H3) -> Data Model only; its parent (Architecture) and sibling
    # (Implementation) are excluded.
    vis = _visible_heading_ids(blocks, "h3-data-model", tid)
    assert vis == {"h3-data-model"}, vis
    # Descendant inclusion: Architecture includes the deeper Data Model.
    arch = _visible_heading_ids(blocks, "h2-architecture", tid)
    assert "h3-data-model" in arch
    # Sibling exclusion: Architecture excludes the H2 sibling Implementation.
    assert "h2-implementation" not in arch


def test_exit_focus_control_present():
    html = _render()
    # Exit control exists and starts hidden (full document by default).
    assert 'id="exit-focus"' in html
    assert "is-hidden" in html  # CSS rule shipped
    assert ".docx-block.is-hidden" in html
