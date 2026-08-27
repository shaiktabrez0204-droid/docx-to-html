"""Browser verification of TOC navigation (real Chromium via Playwright).

Loads a converted DOCX in a real browser and proves that every TOC anchor
targets an actual heading element and that clicking a TOC entry navigates to
it. Replaces the old stub test_toc_playwright.py with a real check.
"""

import os
import sys

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from semantic.pipeline import convert_docx

import pytest

FIXTURE = os.path.join(PROJECT_ROOT, "tests", "fixtures", "mixed-document.docx")
OUT = os.path.join(PROJECT_ROOT, "tests", "test-output", "toc_nav.html")


def _render():
    res = convert_docx(FIXTURE)
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(res.html)
    return res


def test_toc_navigation_in_browser():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        pytest.skip("playwright not installed")

    res = _render()
    expected_ids = [p.heading_id for p in res.paragraphs if p.heading_id]
    assert expected_ids, "no headings produced"

    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch()
            page = browser.new_page()
            page.goto("file://" + OUT)
            page.wait_for_selector("nav.toc")

            # 1) Every TOC anchor references a real heading element.
            anchors = page.eval_on_selector_all(
                "nav.toc a", "els => els.map(e => e.getAttribute('href'))")
            assert anchors, "no TOC anchors found"
            for href in anchors:
                assert href.startswith("#"), href
                target_id = href[1:]
                assert page.query_selector("#" + target_id) is not None, \
                    "anchor %s has no matching heading" % href

            # 2) Clicking a TOC entry updates the location hash to the target.
            first = anchors[0][1:]
            page.click("nav.toc a[href='#%s']" % first)
            page.wait_for_timeout(150)
            assert first in page.evaluate("() => location.hash"), \
                "clicking TOC did not navigate to %s" % first

            # 3) The target is a real heading element with matching id.
            tag = page.eval_on_selector("#" + first, "el => el.tagName")
            assert tag.lower().startswith("h"), "target %s is not a heading" % first

            browser.close()
    except Exception as e:  # browser binary missing -> skip rather than red
        if "executable doesn't exist" in str(e) or "launch" in str(e).lower():
            pytest.skip("chromium browser binary not installed: %s" % e)
        raise
