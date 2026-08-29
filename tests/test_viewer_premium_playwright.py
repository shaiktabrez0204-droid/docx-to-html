import os, sys, threading
import pytest

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from ui import web
FIXTURES = os.path.join(PROJECT_ROOT, "tests", "fixtures")

def _start():
    srv=web.make_server("127.0.0.1",0)
    threading.Thread(target=srv.serve_forever,daemon=True).start()
    return srv, srv.server_address[1]

def _needs_pw():
    try:
        import playwright
        return False
    except ImportError:
        return True

@pytest.mark.skipif(_needs_pw(), reason="playwright not installed")
def test_viewer_header_and_sidebar_toggle():
    from playwright.sync_api import sync_playwright
    srv,port=_start()
    base=f"http://127.0.0.1:{port}"
    try:
        with sync_playwright() as pw:
            browser=pw.chromium.launch()
            page=browser.new_page()
            fx=os.path.join(FIXTURES,"mixed-document.docx")
            page.goto(base+"/", timeout=15000)
            page.set_input_files("input[type=file]", fx)
            page.wait_for_selector("#selected.visible", timeout=5000)
            page.click("#convertBtn")
            page.wait_for_selector("#successPanel.visible", timeout=20000)
            href=page.locator("#openPreviewBtn").get_attribute("href")
            full=base+href
            page2=browser.new_page()
            page2.goto(full, timeout=15000)
            page2.wait_for_selector(".viewer-header", timeout=10000)
            assert page2.locator(".viewer-header").is_visible()
            assert page2.locator("#viewer-docname").is_visible()
            assert page2.locator("#header-search").is_visible()
            assert page2.locator("#viewer-download").is_visible()
            assert page2.locator("#viewer-sidebar").is_visible()
            page2.click("#sidebar-collapse")
            page2.wait_for_timeout(400)
            assert page2.evaluate("()=>document.querySelector('.viewer').classList.contains('viewer--sidebar-collapsed')")
            page2.click("#header-toc")
            page2.wait_for_timeout(400)
            assert not page2.evaluate("()=>document.querySelector('.viewer').classList.contains('viewer--sidebar-collapsed')")
            browser.close()
    finally:
        srv.shutdown()

@pytest.mark.skipif(_needs_pw(), reason="playwright not installed")
def test_toc_search_and_focus_flow():
    from playwright.sync_api import sync_playwright
    srv,port=_start()
    base=f"http://127.0.0.1:{port}"
    try:
        with sync_playwright() as pw:
            browser=pw.chromium.launch()
            page=browser.new_page()
            fx=os.path.join(FIXTURES,"mixed-document.docx")
            page.goto(base+"/", timeout=15000)
            page.set_input_files("input[type=file]", fx)
            page.wait_for_selector("#selected.visible", timeout=5000)
            page.click("#convertBtn")
            page.wait_for_selector("#successPanel.visible", timeout=20000)
            href=page.locator("#openPreviewBtn").get_attribute("href")
            full=base+href
            page2=browser.new_page()
            page2.goto(full, timeout=15000)
            page2.wait_for_selector("#toc-search", timeout=10000)
            page2.fill("#toc-search", "Arch")
            page2.wait_for_timeout(400)
            assert page2.locator(".toc-mark").count() >= 1
            assert page2.locator("#toc-search-clear.visible").is_visible()
            page2.fill("#toc-search", "zzzznotfound")
            page2.wait_for_timeout(400)
            assert page2.locator("#toc-no-results.visible").is_visible()
            page2.fill("#toc-search", "")
            page2.wait_for_timeout(400)
            assert not page2.locator("#toc-no-results.visible").is_visible()
            # Focus heading
            first=page2.locator("a.toc-link").nth(1)
            first.click()
            page2.wait_for_timeout(500)
            assert page2.locator("#focus-banner.visible").is_visible()
            assert "Show Full Document" in page2.locator("#focus-banner").text_content()
            hidden=page2.evaluate("()=>document.querySelectorAll('.docx-block.is-hidden').length")
            assert hidden > 0
            # child
            child=page2.locator("a.toc-link").nth(2)
            child.click()
            page2.wait_for_timeout(500)
            assert page2.locator("#focus-banner.visible").is_visible()
            # clear
            page2.click("#focus-banner-clear")
            page2.wait_for_timeout(400)
            hidden2=page2.evaluate("()=>document.querySelectorAll('.docx-block.is-hidden').length")
            assert hidden2 == 0
            assert not page2.locator("#focus-banner.visible").is_visible()
            browser.close()
    finally:
        srv.shutdown()

@pytest.mark.skipif(_needs_pw(), reason="playwright not installed")
def test_history_and_responsive_drawer():
    from playwright.sync_api import sync_playwright
    srv,port=_start()
    base=f"http://127.0.0.1:{port}"
    try:
        with sync_playwright() as pw:
            browser=pw.chromium.launch()
            page=browser.new_page()
            fx=os.path.join(FIXTURES,"mixed-document.docx")
            page.goto(base+"/", timeout=15000)
            page.set_input_files("input[type=file]", fx)
            page.wait_for_selector("#selected.visible", timeout=5000)
            page.click("#convertBtn")
            page.wait_for_selector("#successPanel.visible", timeout=20000)
            href=page.locator("#openPreviewBtn").get_attribute("href")
            full=base+href
            page2=browser.new_page()
            page2.goto(full, timeout=15000)
            page2.wait_for_selector("a.toc-link", timeout=10000)
            page2.locator("a.toc-link").nth(1).click()
            page2.wait_for_timeout(400)
            assert page2.locator("#focus-banner.visible").is_visible()
            page2.evaluate("()=>history.back()")
            page2.wait_for_timeout(400)
            assert not page2.locator("#focus-banner.visible").is_visible()
            page2.evaluate("()=>history.forward()")
            page2.wait_for_timeout(400)
            assert page2.locator("#focus-banner.visible").is_visible()
            # responsive drawer
            page3=browser.new_page(viewport={"width":375,"height":800})
            page3.goto(full, timeout=15000)
            page3.wait_for_selector(".viewer", timeout=10000)
            page3.wait_for_timeout(400)
            assert not page3.evaluate("()=>document.querySelector('.viewer').classList.contains('viewer--mobile-open')")
            page3.click("#header-toc")
            page3.wait_for_timeout(400)
            assert page3.evaluate("()=>document.querySelector('.viewer').classList.contains('viewer--mobile-open')")
            page3.locator("a.toc-link").nth(1).click()
            page3.wait_for_timeout(500)
            assert not page3.evaluate("()=>document.querySelector('.viewer').classList.contains('viewer--mobile-open')")
            # no overflow
            overflow=page3.evaluate("()=>document.documentElement.scrollWidth > document.documentElement.clientWidth")
            assert not overflow
            overflow2=page2.evaluate("()=>document.documentElement.scrollWidth > document.documentElement.clientWidth")
            assert not overflow2
            browser.close()
    finally:
        srv.shutdown()

@pytest.mark.skipif(_needs_pw(), reason="playwright not installed")
def test_download_no_errors():
    from playwright.sync_api import sync_playwright
    srv,port=_start()
    base=f"http://127.0.0.1:{port}"
    try:
        with sync_playwright() as pw:
            browser=pw.chromium.launch()
            page=browser.new_page()
            failed=[]
            console=[]
            page.on("requestfailed", lambda r: failed.append(r.url))
            page.on("console", lambda m: console.append(m.text) if m.type=="error" else None)
            fx=os.path.join(FIXTURES,"mixed-document.docx")
            page.goto(base+"/", timeout=15000)
            page.set_input_files("input[type=file]", fx)
            page.wait_for_selector("#selected.visible", timeout=5000)
            page.click("#convertBtn")
            page.wait_for_selector("#successPanel.visible", timeout=20000)
            href=page.locator("#openPreviewBtn").get_attribute("href")
            full=base+href
            page2=browser.new_page()
            page2.on("requestfailed", lambda r: failed.append(r.url))
            page2.on("console", lambda m: console.append(m.text) if m.type=="error" else None)
            page2.goto(full, timeout=15000)
            page2.wait_for_selector("#viewer-download", timeout=10000)
            # download handler should not throw
            page2.click("#viewer-download")
            page2.wait_for_timeout(500)
            assert failed == []
            assert console == []
            browser.close()
    finally:
        srv.shutdown()
