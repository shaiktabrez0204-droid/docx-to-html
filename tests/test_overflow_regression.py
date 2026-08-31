import os, re, time, threading
from semantic.pipeline import convert_docx
from ui.web import make_server

def test_no_horizontal_overflow_multicol():
    server = make_server(host="127.0.0.1", port=8899)
    thr = threading.Thread(target=server.serve_forever, daemon=True)
    thr.start()
    time.sleep(0.5)
    docx_path = os.path.join(os.path.dirname(__file__), "fixtures", "benchmark-multicol.docx")
    try:
        import requests
    except ImportError:
        server.shutdown()
        return
    with open(docx_path, 'rb') as f:
        files = {'docx': ('benchmark-multicol.docx', f, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')}
        r = requests.post('http://127.0.0.1:8899/upload', files=files, headers={'Accept': 'text/html'})
        assert r.status_code == 200
        m = re.search(r'/preview/([a-f0-9]+)', r.text)
        assert m, "preview id not found"
        preview = f"http://127.0.0.1:8899/preview/{m.group(1)}"
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page(viewport={"width": 1280, "height": 800})
        page.goto(preview)
        page.wait_for_timeout(1000)
        metrics = page.evaluate("""() => {
            const de=document.documentElement;
            const docMain=document.getElementById('doc-main');
            const docPage=document.querySelector('.docx-page');
            return {
                de_sw: de.scrollWidth, de_cw: de.clientWidth,
                dm_sw: docMain?docMain.scrollWidth:0, dm_cw: docMain?docMain.clientWidth:0,
                page_sw: docPage?docPage.scrollWidth:0, page_cw: docPage?docPage.clientWidth:0,
                hasH: de.scrollWidth > de.clientWidth + 1,
                dmHasH: docMain ? docMain.scrollWidth > docMain.clientWidth + 1 : false,
                over: (()=>{let c=0; for(const el of document.querySelectorAll('*')){const r=el.getBoundingClientRect(); if(r.right > de.clientWidth+0.5 && r.width>0) c++;} return c;})()
            };
        }""")
        assert metrics['de_sw'] == metrics['de_cw'], f"documentElement scrollWidth {metrics['de_sw']} != clientWidth {metrics['de_cw']} indicates genuine horizontal overflow"
        assert metrics['dm_sw'] == metrics['dm_cw'], f"docMain scrollWidth {metrics['dm_sw']} != clientWidth {metrics['dm_cw']}"
        assert metrics['page_sw'] == metrics['page_cw'], f"docx-page scrollWidth != clientWidth"
        assert not metrics['hasH'], "unexpected horizontal scrollbar at 1280"
        assert metrics['over'] == 0, f"overflowing elements {metrics['over']} at 1280 implies genuine overflow"
        # Expected layout width difference is not overflow: de 1280 vs docMain 1000 vs page 816 is centred layout, not scroll
        # At narrow viewport, page shrinks via max-width, no scrollbar should appear
        page2 = browser.new_page(viewport={"width": 800, "height": 800})
        page2.goto(preview)
        page2.wait_for_timeout(800)
        m2 = page2.evaluate("""() => {
            const de=document.documentElement;
            const docMain=document.getElementById('doc-main');
            return {de_sw: de.scrollWidth, de_cw: de.clientWidth, hasH: de.scrollWidth > de.clientWidth + 1, dm_sw: docMain.scrollWidth, dm_cw: docMain.clientWidth};
        }""")
        assert not m2['hasH'], f"narrow viewport should not have horizontal scrollbar, got {m2}"
        assert m2['de_sw'] == m2['de_cw']
        browser.close()
    server.shutdown()

if __name__ == "__main__":
    test_no_horizontal_overflow_multicol()
    print("overflow regression passed")
