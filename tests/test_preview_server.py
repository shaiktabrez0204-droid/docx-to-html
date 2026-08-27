"""HTTP-level tests for the local DOCX -> HTML preview server (stdlib only).

Verifies the preview glue without requiring a browser:
  * a real DOCX converts through the EXISTING pipeline,
  * the generated HTML is written to a predictable location,
  * localhost serves it with HTTP 200,
  * the document is self-contained (no external src/href => no broken assets),
  * TOC / headings / images are present in the served payload.
"""

import os
import re
import sys
import threading
import unittest
import urllib.request

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from ui.preview_server import (  # noqa: E402
    convert_to_html,
    make_server,
    default_output_path,
)

FIXTURES = os.path.join(PROJECT_ROOT, "tests", "fixtures")


class PreviewServerHttpTest(unittest.TestCase):
    def _start(self, fixture, port=0):
        fx = os.path.join(FIXTURES, fixture + ".docx")
        self.assertTrue(os.path.isfile(fx), "missing fixture: %s" % fx)
        html = convert_to_html(fx)
        out = default_output_path(fx)
        with open(out, "wb") as fh:
            fh.write(html.encode("utf-8"))
        srv = make_server(out, host="127.0.0.1", port=port)
        threading.Thread(target=srv.serve_forever, daemon=True).start()
        return srv, srv.server_address[1], html

    def _get(self, port):
        with urllib.request.urlopen("http://127.0.0.1:%d/" % port, timeout=10) as r:
            return r.status, r.read().decode("utf-8")

    def test_convert_works(self):
        # Requirement 1: DOCX can be converted via the existing pipeline.
        html = convert_to_html(os.path.join(FIXTURES, "mixed-document.docx"))
        self.assertIn("<!DOCTYPE html>", html)

    def test_html_file_generated(self):
        # Requirement 2: HTML is written to a predictable location.
        srv, port, html = self._start("mixed-document")
        try:
            self.assertTrue(os.path.isfile(default_output_path(
                os.path.join(FIXTURES, "mixed-document.docx"))))
        finally:
            srv.shutdown()

    def test_localhost_serves_html(self):
        # Requirement 3: localhost serves the generated HTML (HTTP 200).
        srv, port, html = self._start("mixed-document")
        try:
            status, body = self._get(port)
            self.assertEqual(status, 200)
            self.assertEqual(body, html)  # served bytes match generated HTML
        finally:
            srv.shutdown()

    def test_assets_self_contained(self):
        # Requirement 4/6: no external resources => nothing can break.
        srv, port, html = self._start("mixed-document")
        try:
            _status, body = self._get(port)
            bad_src = [s for s in re.findall(r'src="([^"]*)"', body)
                       if not s.startswith("data:")]
            bad_href = [h for h in re.findall(r'href="([^"]*)"', body)
                        if not h.startswith("#")]
            self.assertEqual(bad_src, [], "external src references")
            self.assertEqual(bad_href, [], "external href references")
            self.assertIn('class="toc"', body)  # TOC present
            self.assertGreaterEqual(body.count("<h1"), 1)
        finally:
            srv.shutdown()

    def test_images_served_for_image_fixture(self):
        # Requirement 4: image fixtures embed data: URLs that render over HTTP.
        srv, port, html = self._start("img-multiple")
        try:
            status, body = self._get(port)
            self.assertEqual(status, 200)
            self.assertGreaterEqual(body.count("<img"), 4)
            self.assertEqual(body.count("data:image"), 4)
        finally:
            srv.shutdown()


if __name__ == "__main__":
    unittest.main(verbosity=2)
