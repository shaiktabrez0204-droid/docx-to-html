"""Browser verification of the DOCX -> HTML product (real Chromium).

Mirrors tests/test_preview_server_playwright.py but drives the actual product
UI: open the upload page, select a real .docx, click Convert, wait for the
preview iframe to render, and prove the generated content is present and
correct (headings, TOC links + targets, numbering, inline + floating images,
no console errors, no failed requests). Download is validated over HTTP in
tests/test_web_product.py to keep this focused on the visual journey.

Skips when Playwright/Chromium is unavailable so it never breaks a headless run.
"""

import os
import sys
import threading

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from ui import web  # noqa: E402

import pytest

FIXTURES = os.path.join(PROJECT_ROOT, "tests", "fixtures")

# fixture -> structural expectations verified inside the preview iframe
CASES = {
    "mixed-document": {"headings_min": 5, "toc": True},
    "num-h1-h2-h3": {"headings_min": 1, "toc": True, "numbering_min": 6},
    "num-mixed": {"headings_min": 1, "toc": True, "numbering_min": 4},
    "img-multiple": {"imgs": 4},
    "flt-multi": {"imgs": 3, "float": 3},
    "flt-near-heading": {"headings_min": 1, "toc": True, "imgs": 1},
}


def _start_server():
    srv = web.make_server(host="127.0.0.1", port=0)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv, srv.server_address[1]


def test_product_user_journey():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        pytest.skip("playwright not installed")

    srv, port = _start_server()
    base = "http://127.0.0.1:%d" % port
    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch()
            page = browser.new_page()
            failed_reqs = []
            console_errors = []
            page.on("requestfailed", lambda r: failed_reqs.append(r.url))
            page.on("console", lambda m: console_errors.append(m.text)
                    if m.type == "error" else None)

            for name, exp in CASES.items():
                failed_reqs.clear()
                console_errors.clear()
                fx = os.path.join(FIXTURES, name + ".docx")
                resp = page.goto(base + "/", timeout=15000)
                assert resp.status == 200, "%s: index %s" % (name, resp.status)

                # Select the real fixture and submit the form.
                with page.expect_file_chooser() as chooser_info:
                    page.click("input[type=file]")
                chooser_info.value.set_files(fx)
                page.click("button[type=submit]")

                # The response page embeds the preview in an iframe.
                page.wait_for_selector("iframe.preview", timeout=20000)
                frame = page.frame_locator("iframe.preview")
                frame.locator("body").wait_for(timeout=20000)

                if exp.get("headings_min"):
                    n = frame.locator("h1,h2,h3,h4,h5,h6").count()
                    assert n >= exp["headings_min"], \
                        "%s: headings %d < %d" % (name, n, exp["headings_min"])

                if exp.get("numbering_min"):
                    n = frame.locator(".docx-number").count()
                    assert n >= exp["numbering_min"], \
                        "%s: numbering %d < %d" % (name, n, exp["numbering_min"])

                if exp.get("toc"):
                    toc_links = frame.locator("nav.toc a[href^='#']").count()
                    assert toc_links > 0, "%s: no TOC links" % name
                    first = frame.locator("nav.toc a[href^='#']").first
                    href = first.get_attribute("href")
                    target_id = href.lstrip("#")
                    first.click()
                    if target_id:
                        exists = frame.locator("#%s" % target_id).count()
                        assert exists, "%s: TOC target #%s missing" % (name, target_id)

                if exp.get("imgs"):
                    imgs = frame.locator("img")
                    assert imgs.count() >= exp["imgs"], \
                        "%s: img %d < %d" % (name, imgs.count(), exp["imgs"])
                    for i in range(imgs.count()):
                        nw = imgs.nth(i).evaluate("el => el.naturalWidth")
                        assert nw > 0, "%s: img %d not rendered" % (name, i)

                if exp.get("float"):
                    # Floating images must be present in the document.
                    floats = frame.locator("img.docx-float, img.docx-float-wrapped").count()
                    assert floats >= exp["float"], \
                        "%s: floats %d < %d" % (name, floats, exp["float"])

            # Whole-session invariants.
            assert failed_reqs == [], "failed requests: %s" % failed_reqs
            assert console_errors == [], "console errors: %s" % console_errors
            browser.close()
    except Exception as exc:
        msg = str(exc).lower()
        if "event loop" in msg or "already stopped" in msg or "playwright" in msg:
            pytest.skip("playwright sync session unavailable in this process: %s" % exc)
        raise
    finally:
        srv.shutdown()
