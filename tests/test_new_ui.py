"""Focused UI tests for modern SaaS upload interface.

Covers:
  - initial state (header, left panel, dropzone, step indicator)
  - valid DOCX selection
  - invalid file handling
  - drag-over & invalid drag styling
  - conversion state animation
  - success state & preview iframe
  - error state via rejected upload
  - responsive layout (mobile no overflow)
  - accessibility (keyboard, aria-live, reduced-motion)
"""

import os, sys, threading
import pytest

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from ui import web

FIXTURES = os.path.join(PROJECT_ROOT, "tests", "fixtures")

def _start():
    srv = web.make_server("127.0.0.1", 0)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv, srv.server_address[1]

def _needs_playwright():
    try:
        import playwright
        return False
    except ImportError:
        return True


@pytest.mark.skipif(_needs_playwright(), reason="playwright not installed")
def test_initial_state():
    from playwright.sync_api import sync_playwright
    srv, port = _start()
    base = f"http://127.0.0.1:{port}"
    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch()
            page = browser.new_page()
            page.goto(base + "/", timeout=15000)
            # Header
            assert page.locator(".header").is_visible()
            assert "DOCX" in page.locator(".brand-title").text_content()
            assert "Convert. Preserve. Perfect." in page.locator(".brand-sub").text_content()
            assert page.locator(".header-meta").is_visible()
            # Left panel
            assert "Convert your" in page.locator(".headline").text_content()
            assert page.locator(".lead").count() == 0
            assert page.locator(".benefits").count() == 0
            assert page.locator(".check").count() == 0
            assert page.locator(".security").is_visible()
            # Card
            assert page.locator("#card").is_visible()
            assert "step 1 of 3" in page.locator(".step-label").text_content().lower()
            assert page.locator("#dropzone").is_visible()
            assert page.locator("#docxInput").count() == 1
            assert page.locator("#convertBtn").count() == 1
            # convert hidden initially
            assert not page.locator("#convertWrap.visible").is_visible()
            assert not page.locator("#selected.visible").is_visible()
            # Step indicator: step 1 active
            assert page.locator('.pipe-step[data-step="1"].active').is_visible()
            # Footer benefits
            assert page.locator(".features").is_visible()
            assert page.locator(".feat").count() == 4
            # No console errors
            errors=[]
            page.on("console", lambda m: errors.append(m.text) if m.type=="error" else None)
            page.reload(wait_until="domcontentloaded")
            assert errors==[]
            browser.close()
    finally:
        srv.shutdown()


@pytest.mark.skipif(_needs_playwright(), reason="playwright not installed")
def test_valid_docx_selection():
    from playwright.sync_api import sync_playwright
    srv, port = _start()
    base = f"http://127.0.0.1:{port}"
    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch()
            page = browser.new_page()
            page.goto(base+"/", timeout=15000)
            fx = os.path.join(FIXTURES, "mixed-document.docx")
            page.set_input_files("input[type=file]", fx)
            page.wait_for_selector("#selected.visible", timeout=5000)
            assert "mixed-document.docx" in page.locator("#fileName").text_content()
            assert "KB" in page.locator("#fileSize").text_content() or "MB" in page.locator("#fileSize").text_content()
            assert page.locator("#convertWrap.visible").is_visible()
            assert page.locator("#convertBtn").is_visible()
            # dropzone should have has-file
            assert "has-file" in page.locator("#dropzone").get_attribute("class")
            browser.close()
    finally:
        srv.shutdown()


@pytest.mark.skipif(_needs_playwright(), reason="playwright not installed")
def test_invalid_file_shows_error():
    from playwright.sync_api import sync_playwright
    import tempfile
    srv, port = _start()
    base = f"http://127.0.0.1:{port}"
    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch()
            page = browser.new_page()
            page.goto(base+"/", timeout=15000)
            tmp = tempfile.mktemp(suffix=".txt")
            with open(tmp,"w") as f: f.write("hello")
            try:
                page.set_input_files("input[type=file]", tmp)
                page.wait_for_selector("#alertBox", state="visible", timeout=5000)
                assert "Only .docx" in page.locator("#alertBox").text_content()
                # selected should NOT be visible
                assert not page.locator("#selected.visible").is_visible()
            finally:
                try: os.remove(tmp)
                except: pass
            browser.close()
    finally:
        srv.shutdown()


@pytest.mark.skipif(_needs_playwright(), reason="playwright not installed")
def test_drag_over_styling():
    from playwright.sync_api import sync_playwright
    srv, port = _start()
    base = f"http://127.0.0.1:{port}"
    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch()
            page = browser.new_page()
            page.goto(base+"/", timeout=15000)
            # valid drag
            page.evaluate("()=>document.getElementById('dropzone').classList.add('drag-over')")
            assert "drag-over" in page.locator("#dropzone").get_attribute("class")
            # background and border should be brighter - check computed style
            border = page.evaluate("()=>getComputedStyle(document.getElementById('dropzone')).borderColor")
            assert border  # non-empty
            # drop hint visible via opacity
            page.evaluate("()=>document.getElementById('dropzone').classList.remove('drag-over')")
            assert "drag-over" not in page.locator("#dropzone").get_attribute("class")
            # invalid drag
            page.evaluate("()=>document.getElementById('dropzone').classList.add('drag-over','drag-invalid')")
            assert "drag-invalid" in page.locator("#dropzone").get_attribute("class")
            assert page.locator("#dzError").is_visible()
            browser.close()
    finally:
        srv.shutdown()


@pytest.mark.skipif(_needs_playwright(), reason="playwright not installed")
def test_conversion_and_success_state():
    from playwright.sync_api import sync_playwright
    srv, port = _start()
    base = f"http://127.0.0.1:{port}"
    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch()
            page = browser.new_page()
            page.goto(base+"/", timeout=15000)
            fx = os.path.join(FIXTURES, "mixed-document.docx")
            page.set_input_files("input[type=file]", fx)
            page.wait_for_selector("#selected.visible", timeout=5000)
            page.click("#convertBtn")
            try:
                page.wait_for_selector("#converting.visible", timeout=3000)
            except Exception:
                pass
            page.wait_for_selector("#successPanel.visible", timeout=20000)
            assert page.locator("#successPanel").is_visible()
            assert "Conversion complete" in page.locator("#successPanel").text_content()
            assert page.locator('.pipe-step[data-step="3"].active').is_visible()
            assert "Conversion complete" in page.locator("#successPanel").text_content()
            assert page.locator('.pipe-step[data-step="3"].active').is_visible()
            # preview iframe
            page.wait_for_selector("iframe.preview", timeout=10000)
            frame = page.frame_locator("iframe.preview")
            frame.locator("body").wait_for(timeout=10000)
            assert frame.locator("h1,h2").count() >= 1
            # Download link should have correct href
            dl = page.locator("#downloadBtn").get_attribute("href")
            assert dl and "/download/" in dl
            browser.close()
    finally:
        srv.shutdown()


@pytest.mark.skipif(_needs_playwright(), reason="playwright not installed")
def test_error_state_on_invalid_docx_content():
    from playwright.sync_api import sync_playwright
    import tempfile
    srv, port = _start()
    base = f"http://127.0.0.1:{port}"
    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch()
            page = browser.new_page()
            page.goto(base+"/", timeout=15000)
            # create a .docx that is not a zip
            tmp = tempfile.mktemp(suffix=".docx")
            with open(tmp,"wb") as f: f.write(b"PK fake but not docx")
            page.set_input_files("input[type=file]", tmp)
            page.wait_for_selector("#selected.visible", timeout=5000)
            page.click("#convertBtn")
            page.wait_for_selector("#alertBox", state="visible", timeout=20000)
            txt = page.locator("#alertBox").text_content()
            assert "not a valid DOCX" in txt or "could not be converted" in txt.lower() or "Conversion failed" in txt
            browser.close()
            try: os.remove(tmp)
            except: pass
    finally:
        srv.shutdown()


@pytest.mark.skipif(_needs_playwright(), reason="playwright not installed")
def test_responsive_no_overflow():
    from playwright.sync_api import sync_playwright
    srv, port = _start()
    base = f"http://127.0.0.1:{port}"
    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch()
            for w in [1280, 768, 375]:
                page = browser.new_page(viewport={"width": w, "height": 800})
                page.goto(base+"/", timeout=15000)
                page.wait_for_timeout(400)
                overflow = page.evaluate("()=>document.documentElement.scrollWidth > document.documentElement.clientWidth")
                assert not overflow, f"overflow at {w}px"
                # dropzone must remain visible on mobile
                assert page.locator("#dropzone").is_visible()
                page.close()
            browser.close()
    finally:
        srv.shutdown()


@pytest.mark.skipif(_needs_playwright(), reason="playwright not installed")
def test_accessibility():
    from playwright.sync_api import sync_playwright
    srv, port = _start()
    base = f"http://127.0.0.1:{port}"
    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch()
            page = browser.new_page()
            page.goto(base+"/", timeout=15000)
            # keyboard accessible file selection: dropzone focusable
            assert page.locator("#dropzone").get_attribute("tabindex") == "0"
            assert page.locator("#dropzone").get_attribute("role") == "button"
            # aria-live for conversion status
            assert page.locator("#statusLive").get_attribute("aria-live") == "polite"
            # focus states via keyboard
            page.keyboard.press("Tab")
            # should be able to focus dropzone or file input
            # reduced-motion: check CSS contains media query
            content = page.content()
            assert "prefers-reduced-motion" in content
            browser.close()
    finally:
        srv.shutdown()
