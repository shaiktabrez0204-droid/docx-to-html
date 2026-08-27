"""Browser verification of VISIBLE numbering rendering (real Chromium).

Loads a converted numbered DOCX in a real browser and proves the numbering is
actually visible in the rendered DOM (not just source), that heading format is
correct, that the TOC shows the SAME number, and that heading ids / TOC links
are preserved.
"""

import os
import sys
import re

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from core.model import format_numbering_label
from semantic.pipeline import convert_docx

import pytest

FIXTURE = os.path.join(PROJECT_ROOT, "tests", "fixtures", "num-h1-h2-h3.docx")
OUT = os.path.join(PROJECT_ROOT, "tests", "test-output", "visible_numbering.html")


def _render():
    res = convert_docx(FIXTURE)
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(res.html)
    return res


def test_visible_numbering_in_browser():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        pytest.skip("playwright not installed")

    res = _render()
    by_text = {"".join(x.text for x in p.runs): p for p in res.paragraphs if p.heading_level}

    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch()
            page = browser.new_page()
            page.goto("file://" + OUT)
            page.wait_for_selector("nav.toc")

            # 1) Every numbered heading exposes a visible .docx-number span whose
            #    text equals the model-resolved label.
            for text, p in by_text.items():
                if p.numbering_path is None:
                    continue
                expected = format_numbering_label(
                    p.numbering_path, p.numbering_level_formats, p.numbering_text_pattern)
                sel = 'h1#%s .docx-number, h2#%s .docx-number, h3#%s .docx-number' % (
                    p.heading_id, p.heading_id, p.heading_id)
                node = page.query_selector(sel)
                assert node is not None, "no visible number for heading %r" % text
                assert node.inner_text().strip() == expected, \
                    "heading %r number %r != expected %r" % (text, node.inner_text(), expected)

            # 2) Heading hierarchy preserved (h1 > h2 > h3).
            tags = [page.eval_on_selector("#" + p.heading_id, "el => el.tagName.toLowerCase()")
                    for p in res.paragraphs if p.heading_id]
            assert "h1" in tags and "h2" in tags and "h3" in tags

            # 3) TOC shows the SAME resolved number and links resolve.
            anchors = page.eval_on_selector_all(
                "nav.toc a", "els => els.map(e => ({href: e.getAttribute('href'), num: e.querySelector('.docx-number') ? e.querySelector('.docx-number').textContent.trim() : null}))")
            assert anchors, "no TOC anchors"
            for a in anchors:
                assert a["href"].startswith("#")
                tid = a["href"][1:]
                assert page.query_selector("#" + tid) is not None, "broken TOC link %s" % a["href"]
                # TOC number must match the heading's resolved number.
                head = by_text.get(_heading_text_for(res, tid))
                if head and head.numbering_path is not None:
                    exp = format_numbering_label(
                        head.numbering_path, head.numbering_level_formats, head.numbering_text_pattern)
                    assert a["num"] == exp, "TOC number %r != heading %r" % (a["num"], exp)

            # 4) Heading ids unchanged (anchors == heading ids).
            hids = set(p.heading_id for p in res.paragraphs if p.heading_id)
            aids = set(a["href"][1:] for a in anchors)
            assert aids == hids

            # 5) No duplicate numbers across distinct headings.
            visible = [a["num"] for a in anchors if a["num"]]
            assert len(visible) == len(set(visible)), "duplicate visible numbers: %s" % visible

            browser.close()
    except Exception as e:
        if "executable doesn't exist" in str(e) or "launch" in str(e).lower():
            pytest.skip("chromium browser binary not installed: %s" % e)
        raise


def _heading_text_for(res, heading_id):
    for p in res.paragraphs:
        if p.heading_id == heading_id:
            return "".join(r.text for r in p.runs)
    return None


if __name__ == "__main__":
    test_visible_numbering_in_browser()
    print("BROWSER NUMBERING CHECK PASSED")
