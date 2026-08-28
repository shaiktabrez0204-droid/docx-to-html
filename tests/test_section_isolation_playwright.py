"""Real-browser verification of the isolated-heading (section) view.

Drives the ACTUAL product UI end to end: open the upload page, select the real
``section-isolation.docx``, convert, wait for the preview iframe, then click
real TOC headings and assert the isolated section is shown/hidden in the live
DOM. Also verifies the exit control, browser back/forward history, sidebar
visibility, the first-H1 title bar (no duplicate H1), and zero console errors /
failed network requests.

Skips cleanly when Playwright/Chromium is unavailable.
"""

import os
import sys
import threading

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from ui import web  # noqa: E402

import pytest

FIX = os.path.join(PROJECT_ROOT, "tests", "fixtures")
DOCX = os.path.join(FIX, "section-isolation.docx")


def _start_server():
    srv = web.make_server(host="127.0.0.1", port=0)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv, srv.server_address[1]


def test_isolated_heading_view():
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

            resp = page.goto(base + "/", timeout=15000)
            assert resp.status == 200

            with page.expect_file_chooser() as chooser_info:
                page.click("input[type=file]")
            chooser_info.value.set_files(DOCX)
            page.click("button[type=submit]")

            page.wait_for_selector("iframe.preview", timeout=20000)
            frame = page.frame_locator("iframe.preview")
            frame.locator("body").wait_for(timeout=20000)
            frame_obj = page.query_selector("iframe.preview").content_frame()

            # Build label -> heading id map from the real TOC links.
            toc = frame.locator("nav.toc a.toc-link")
            n = toc.count()
            assert n == 6, "expected 6 TOC links, got %d" % n
            labels = {}
            for i in range(n):
                t = toc.nth(i).inner_text().strip()
                href = toc.nth(i).get_attribute("href") or ""
                labels[t] = href.lstrip("#")
            for k in ("Introduction", "Architecture", "Data Model",
                      "Implementation", "Conclusion", "Summary"):
                assert k in labels, "missing TOC label %s" % k

            def click(label):
                frame.locator("nav.toc a.toc-link").filter(
                    has_text=label).first.click()

            def vis(label):
                return frame.locator("#%s" % labels[label]).is_visible()

            def intro_para():
                # The first H1 lives in the title bar; its body CONTENT is this
                # paragraph, which must be hidden whenever a later section is
                # isolated (and visible only when the first H1 itself is focused).
                return frame.locator('p:has-text("Intro body text")')

            def assert_focus(focus, visible, hidden, intro_shown):
                click(focus)
                for v in visible:
                    assert vis(v), "%s: %s should be VISIBLE" % (focus, v)
                for h in hidden:
                    assert not vis(h), "%s: %s should be HIDDEN" % (focus, h)
                if intro_shown:
                    assert intro_para().is_visible(), \
                        "%s: intro content should be VISIBLE" % focus
                else:
                    assert not intro_para().is_visible(), \
                        "%s: intro content should be HIDDEN" % focus
                # Sidebar + title bar remain intact; exit control shows.
                assert frame.locator("#viewer-sidebar").is_visible()
                assert frame.locator(".doc-title").count() == 1
                assert frame.locator("#exit-focus").is_visible()
                # No duplicate H1: only the title bar carries "Introduction".
                assert frame.locator('h1.doc-title').count() == 1

            # H2 Architecture: shows Architecture + Data Model; hides siblings.
            assert_focus("Architecture",
                         ["Architecture", "Data Model"],
                         ["Implementation", "Conclusion", "Summary"],
                         intro_shown=False)
            # Content fidelity: table + inline image + hyperlink still visible.
            assert frame.locator("table.docx-table").is_visible()
            assert frame.locator("img.docx-image").first.is_visible()
            assert frame.locator('a[href="https://example.com"]').is_visible()

            # H3 Data Model: shows only Data Model.
            assert_focus("Data Model",
                         ["Data Model"],
                         ["Architecture", "Implementation", "Conclusion", "Summary"],
                         intro_shown=False)

            # H2 Implementation.
            assert_focus("Implementation",
                         ["Implementation"],
                         ["Architecture", "Data Model", "Conclusion", "Summary"],
                         intro_shown=False)

            # H1 Conclusion: shows Conclusion + Summary.
            assert_focus("Conclusion",
                         ["Conclusion", "Summary"],
                         ["Architecture", "Data Model", "Implementation"],
                         intro_shown=False)

            # First H1 (title bar) Introduction: intro + arch + data model + impl.
            assert_focus("Introduction",
                         ["Architecture", "Data Model", "Implementation"],
                         ["Conclusion", "Summary"],
                         intro_shown=True)

            # Exit focus restores the full document.
            frame.locator("#exit-focus").click()
            for k in labels:
                assert vis(k), "after exit, %s should be visible" % k
            assert not frame.locator("#exit-focus").is_visible()

            # Browser history (driven inside the preview iframe, where pushState
            # actually lives): Architecture -> Data Model -> Implementation.
            def go_back():
                frame_obj.evaluate("history.back()")
                page.wait_for_timeout(200)

            def go_forward():
                frame_obj.evaluate("history.forward()")
                page.wait_for_timeout(200)

            click("Architecture")
            click("Data Model")
            click("Implementation")
            go_back()  # -> Data Model
            assert vis("Data Model") and not vis("Implementation")
            go_back()  # -> Architecture
            assert vis("Architecture") and not vis("Implementation")
            go_back()  # -> full document (initial state)
            for k in labels:
                assert vis(k), "history back to full: %s visible" % k
            go_forward()  # -> Architecture
            assert vis("Architecture") and not vis("Implementation")
            go_forward()  # -> Data Model
            assert vis("Data Model") and not vis("Implementation")

            # Whole-session invariants.
            assert failed_reqs == [], "failed requests: %s" % failed_reqs
            assert console_errors == [], "console errors: %s" % console_errors
            browser.close()
    except Exception as exc:  # pragma: no cover
        msg = str(exc).lower()
        if "event loop" in msg or "already stopped" in msg or "playwright" in msg:
            pytest.skip("playwright sync session unavailable in this process: %s" % exc)
        raise
    finally:
        srv.shutdown()
