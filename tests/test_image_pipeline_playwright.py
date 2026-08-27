"""Browser verification of IMAGE rendering (real Chromium).

Loads converted image-bearing DOCX in a real browser and proves the images
actually render (naturalWidth > 0), that <img> count/order/alt/dimensions match
the resolved model, and that no network request or console error occurs. Images
are inlined as data URLs, so they load from file:// with no external server.
"""

import os
import sys

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from semantic.pipeline import convert_docx
from core.units import emu_to_px

import pytest

# fixture name -> expected number of <img> placements in the DOM
FIXTURES = {
    "img-inline-png": 1,
    "img-inline-jpeg": 1,
    "img-multiple": 4,
    "img-alt-text": 1,
    "img-no-alt": 1,
    "img-explicit-dims": 1,
    "img-floating": 1,
}

OUT_DIR = os.path.join(PROJECT_ROOT, "tests", "test-output")


def test_images_render_in_browser():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        pytest.skip("playwright not installed")

    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch()
            page = browser.new_page()
            failed_reqs = []
            console_errors = []
            page.on("requestfailed", lambda r: failed_reqs.append(r.url))
            page.on("console", lambda m: console_errors.append(m.text)
                    if m.type == "error" else None)

            for name, expected_count in FIXTURES.items():
                path = os.path.join(PROJECT_ROOT, "tests", "fixtures", name + ".docx")
                res = convert_docx(path)
                out = os.path.join(OUT_DIR, name + ".html")
                os.makedirs(OUT_DIR, exist_ok=True)
                with open(out, "w", encoding="utf-8") as f:
                    f.write(res.html)

                model_imgs = [i for p in res.paragraphs for i in p.images]
                failed_reqs.clear()
                console_errors.clear()

                page.goto("file://" + out)
                if expected_count:
                    page.wait_for_selector("img", timeout=5000)
                nodes = page.query_selector_all("img")
                assert len(nodes) == expected_count, \
                    "%s: %d <img> in DOM, expected %d" % (name, len(nodes), expected_count)

                for node, img in zip(nodes, model_imgs):
                    # naturalWidth > 0 is the only real proof the bytes decoded
                    # and the image is visible, not just present in source.
                    nw = node.evaluate("el => el.naturalWidth")
                    assert nw and nw > 0, \
                        "%s: image not actually rendered (naturalWidth=%s)" % (name, nw)
                    assert node.get_attribute("alt") == img.alt_text, \
                        "%s: alt %r != model %r" % (name, node.get_attribute("alt"), img.alt_text)
                    if img.width:
                        if img.wrap_type == "anchor":
                            # Floats carry their size in the inline CSS width
                            # (px); Image.width is already in px (spec mandates
                            # style-based sizing), so compare to the model px.
                            px = node.evaluate(
                                "el => parseInt(getComputedStyle(el).width)")
                            assert abs(px - img.width) <= 2, \
                                "%s: float width %dpx != model %dpx" % (
                                    name, px, img.width)
                        else:
                            assert int(node.get_attribute("width")) == img.width, \
                                "%s: width attr mismatch" % name
                    if img.height:
                        if img.wrap_type == "anchor":
                            px = node.evaluate(
                                "el => parseInt(getComputedStyle(el).height)")
                            assert abs(px - img.height) <= 2, \
                                "%s: float height %dpx != model %dpx" % (
                                    name, px, img.height)
                        else:
                            assert int(node.get_attribute("height")) == img.height, \
                                "%s: height attr mismatch" % name

                assert not failed_reqs, "%s: failed requests: %s" % (name, failed_reqs)
                assert not console_errors, "%s: console errors: %s" % (name, console_errors)

            browser.close()
    except Exception as e:
        if "executable doesn't exist" in str(e) or "launch" in str(e).lower():
            pytest.skip("chromium browser binary not installed: %s" % e)
        raise


if __name__ == "__main__":
    test_images_render_in_browser()
    print("BROWSER IMAGE CHECK PASSED")
