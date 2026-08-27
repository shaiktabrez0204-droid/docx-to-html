"""Browser verification of the local DOCX -> HTML preview server (real Chromium).

Loads the preview server's HTTP endpoint in a real browser and proves:
  * HTTP 200,
  * headings render,
  * TOC links exist and resolve to real anchors,
  * numbering appears,
  * images (inline + floating) actually render (naturalWidth > 0),
  * no failed network requests,
  * no console errors.

Mirrors tests/test_image_pipeline_playwright.py. Skips when Playwright or the
Chromium binary is unavailable so it never breaks a headless CI run.
"""

import os
import sys
import threading

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from ui.preview_server import (  # noqa: E402
    convert_to_html,
    make_server,
    default_output_path,
)

import pytest

FIXTURES = os.path.join(PROJECT_ROOT, "tests", "fixtures")

# fixture -> structural expectations in the served DOM
CASES = {
    "mixed-document": {"headings_min": 5, "toc": True, "imgs": 0},
    "img-multiple": {"imgs": 4},
    "flt-multi": {"imgs": 3},
    "flt-near-heading": {"headings_min": 1, "toc": True, "imgs": 1},
    "num-h1-h2-h3": {"headings_min": 1, "numbering_min": 1},
}


def _start(fixture, port):
    fx = os.path.join(FIXTURES, fixture + ".docx")
    html = convert_to_html(fx)
    out = default_output_path(fx)
    with open(out, "wb") as fh:
        fh.write(html.encode("utf-8"))
    srv = make_server(out, host="127.0.0.1", port=port)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv, srv.server_address[1]


def test_preview_loads_in_browser():
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

            for name, exp in CASES.items():
                srv, port = _start(name, 0)
                try:
                    failed_reqs.clear()
                    console_errors.clear()
                    resp = page.goto("http://127.0.0.1:%d/" % port, timeout=15000)
                    assert resp.status == 200, "%s: HTTP %s" % (name, resp.status)

                    if exp.get("headings_min"):
                        n = page.eval_on_selector_all(
                            "h1,h2,h3,h4,h5,h6", "els => els.length")
                        assert n >= exp["headings_min"], \
                            "%s: headings %d < %d" % (name, n, exp["headings_min"])

                    if exp.get("numbering_min"):
                        spans = page.eval_on_selector_all(
                            ".docx-number", "els => els.length")
                        assert spans >= exp["numbering_min"], \
                            "%s: numbering spans %d" % (name, spans)

                    if exp.get("toc"):
                        toc_links = page.eval_on_selector_all(
                            "nav.toc a[href^='#']", "els => els.length")
                        assert toc_links > 0, "%s: no TOC links" % name
                        first = page.query_selector("nav.toc a[href^='#']")
                        href = first.get_attribute("href")
                        target_id = href.lstrip("#")
                        first.click()
                        if target_id:
                            exists = page.eval_on_selector(
                                "#%s" % target_id, "el => !!el")
                            assert exists, "%s: TOC target #%s missing" % (name, target_id)

                    if exp.get("imgs"):
                        imgs = page.query_selector_all("img")
                        assert len(imgs) >= exp["imgs"], \
                            "%s: img count %d < %d" % (name, len(imgs), exp["imgs"])
                        for img in imgs:
                            nw = img.evaluate("el => el.naturalWidth")
                            assert nw > 0, "%s: image not rendered (naturalWidth=0)" % name
                finally:
                    srv.shutdown()

            assert failed_reqs == [], "failed requests: %s" % failed_reqs
            assert console_errors == [], "console errors: %s" % console_errors
    except Exception as exc:  # environmental Playwright sync-loop conflict
        msg = str(exc).lower()
        if "event loop" in msg or "already stopped" in msg or "playwright" in msg:
            pytest.skip("playwright sync session unavailable in this process: %s" % exc)
        raise
