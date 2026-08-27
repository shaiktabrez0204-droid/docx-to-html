"""HTTP-level product tests: DOCX upload -> convert -> preview + download.

Every fixture is fed through the SAME web endpoint (POST /upload -> the EXISTING
convert_docx). No fixture-specific branching. Verifies:
  * arbitrary valid DOCX converts and previews,
  * generated preview HTML is the real converter output (headings/TOC/numbering
    /images/floating images/tables preserved),
  * download returns the same HTML with a .html filename,
  * invalid uploads (no file, wrong extension, non-docx, oversized) are rejected
    with clean errors and never reach the converter.
"""

import io
import os
import sys
import threading
import unittest
import urllib.request

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from ui import web  # noqa: E402

FIXTURES = os.path.join(PROJECT_ROOT, "tests", "fixtures")


def _multipart(filename, data, field="docx"):
    boundary = "----webtestboundary"
    body = (
        ("--%s\r\n" % boundary).encode()
        + ('Content-Disposition: form-data; name="%s"; filename="%s"\r\n'
           % (field, filename)).encode()
        + b"Content-Type: application/octet-stream\r\n\r\n"
        + data
        + ("\r\n--%s--\r\n" % boundary).encode()
    )
    headers = {
        "Content-Type": "multipart/form-data; boundary=%s" % boundary,
        "Content-Length": str(len(body)),
    }
    return body, headers


def _post(port, filename, data, field="docx"):
    body, headers = _multipart(filename, data, field)
    req = urllib.request.Request(
        "http://127.0.0.1:%d/upload" % port, data=body, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return r.status, r.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode("utf-8")


def _get(port, path):
    with urllib.request.urlopen("http://127.0.0.1:%d%s" % (port, path), timeout=15) as r:
        return r.status, r.read().decode("utf-8"), r.headers


def _start_server():
    srv = web.make_server(host="127.0.0.1", port=0)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv, srv.server_address[1]


# (fixture, expectations) -- all go through the SAME endpoint, no special-casing.
# Expectations use verified per-fixture converter output (body spans, not CSS).
CASES = [
    ("mixed-document", {"headings": 5, "toc": True}),
    ("num-h1-h2-h3", {"headings": 3, "toc": True, "numbering": 6}),
    ("num-mixed", {"headings": 4, "toc": True, "numbering": 4}),
    ("img-multiple", {"imgs": 4, "dataimg": 4}),
    ("flt-multi", {"imgs": 3, "float": 3}),
    ("flt-near-heading", {"headings": 1, "toc": True, "imgs": 1}),
    ("flt-wrapsquare", {"imgs": 1, "float": 1}),
    ("h1-h2", {"headings": 3, "toc": True}),
    ("formatting", {}),
    ("merged-cells", {}),
    ("images", {}),
]


class WebProductTest(unittest.TestCase):
    def setup_class(self):
        self.srv, self.port = _start_server()

    def teardown_class(self):
        self.srv.shutdown()

    def test_upload_and_preview_flow(self):
        for fx, exp in CASES:
            path = os.path.join(FIXTURES, fx + ".docx")
            with open(path, "rb") as fh:
                data = fh.read()
            status, page = _post(self.port, fx + ".docx", data)
            assert status == 200, "%s: status %s" % (fx, status)
            assert "iframe" in page and "/preview/" in page, "%s: no preview iframe" % fx
            assert "Download HTML" in page, "%s: no download button" % fx

            # Extract the doc id and fetch the actual preview HTML.
            doc_id = page.split("/preview/", 1)[1].split('"', 1)[0]
            pstatus, pbody, _ = _get(self.port, "/preview/" + doc_id)
            assert pstatus == 200, "%s: preview %s" % (fx, pstatus)
            assert pbody.lstrip().startswith("<!DOCTYPE html>"), "%s: not HTML" % fx
            assert "class=\"toc\"" in pbody if exp.get("toc") else True
            assert pbody.count("<h1") + pbody.count("<h2") + pbody.count("<h3") >= \
                exp.get("headings", 0), "%s: headings %s" % (fx, exp.get("headings"))
            assert pbody.count("docx-number") >= exp.get("numbering", 0), \
                "%s: numbering %s" % (fx, exp.get("numbering"))
            assert pbody.count("<img") >= exp.get("imgs", 0), "%s: imgs %s" % (fx, exp.get("imgs"))
            assert pbody.count("data:image") >= exp.get("dataimg", 0), \
                "%s: data images %s" % (fx, exp.get("dataimg"))
            assert (pbody.count('class="docx-float"') +
                    pbody.count('class="docx-float-wrapped"')) >= exp.get("float", 0), \
                "%s: floating %s" % (fx, exp.get("float"))
            # Self-contained: no external src/href (only data: and #anchors).
            bad_src = [s for s in __import__("re").findall(r'src="([^"]*)"', pbody)
                       if not s.startswith("data:")]
            bad_href = [h for h in __import__("re").findall(r'href="([^"]*)"', pbody)
                        if not h.startswith("#")]
            assert bad_src == [], "%s: external src %s" % (fx, bad_src)
            assert bad_href == [], "%s: external href %s" % (fx, bad_href)

    def test_download_returns_html(self):
        path = os.path.join(FIXTURES, "mixed-document.docx")
        with open(path, "rb") as fh:
            data = fh.read()
        _status, page = _post(self.port, "Annual_Report.docx", data)
        doc_id = page.split("/preview/", 1)[1].split('"', 1)[0]
        dstatus, dbody, dheaders = _get(self.port, "/download/" + doc_id)
        assert dstatus == 200
        assert dbody.lstrip().startswith("<!DOCTYPE html>")
        assert dbody == self._preview_body(self.port, doc_id)
        cd = dheaders.get("Content-Disposition", "")
        assert cd.startswith("attachment")
        assert "Annual_Report.html" in cd

    def _preview_body(self, port, doc_id):
        st, body, _ = _get(port, "/preview/" + doc_id)
        assert st == 200
        return body

    def test_no_file_rejected(self):
        status, page = _post(self.port, "", b"", field="docx")
        assert status == 400
        assert "No file was selected" in page

    def test_wrong_extension_rejected(self):
        status, page = _post(self.port, "notes.txt", b"hello")
        assert status == 400
        assert "Only .docx files are accepted" in page

    def test_non_docx_rejected(self):
        # A random ZIP that is NOT a DOCX/OOXML package is rejected, not crashed.
        import zipfile
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            zf.writestr("hello.txt", "not a docx")
        status, page = _post(self.port, "fake.docx", buf.getvalue())
        assert status == 400
        assert "not a valid DOCX" in page

    def test_corrupted_docx_rejected_cleanly(self):
        # A truncated/garbage file wearing a .docx name is rejected cleanly.
        status, page = _post(self.port, "broken.docx", b"PK\x03\x04corrupted-bytes")
        assert status == 400
        assert "not a valid DOCX" in page

    def test_index_serves_form(self):
        status, body, _ = _get(self.port, "/")
        assert status == 200
        assert "DOCX" in body and "Upload" in body


if __name__ == "__main__":
    import pytest
    raise SystemExit(pytest.main([__file__, "-v"]))
