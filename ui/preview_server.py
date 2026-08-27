"""Development-only local preview server for the DOCX -> HTML converter.

This module is NOT part of the conversion core. It is pure glue that:

  1. accepts a .docx path,
  2. runs the EXISTING converter (``semantic.pipeline.convert_docx``),
  3. writes the resulting self-contained HTML to a predictable location,
  4. serves that single file over localhost,
  5. optionally opens the browser.

Why this is safe and decoupled
------------------------------
The renderer (``output.html_renderer``) emits ONE self-contained document:
CSS lives in an inline ``<style>`` block, images are embedded as ``data:``
URLs, and TOC links are ``#anchor`` references. There are therefore NO
external CSS / image / asset files to route. Serving the single HTML file
over HTTP is sufficient and cannot produce broken asset requests. No
converter code (OOXML extraction, styles, heading classification, hierarchy,
numbering, TOC, or image anchoring) is touched by this module.
"""

import argparse
import os
import sys
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

# Make the converter package importable regardless of the current working dir.
_PKG_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PKG_ROOT not in sys.path:
    sys.path.insert(0, _PKG_ROOT)

from semantic.pipeline import convert_docx  # noqa: E402


def convert_to_html(docx_path, title=None):
    """Run the existing converter and return the generated HTML string.

    Pure pass-through to ``convert_docx``; no rendering logic lives here.
    """
    if title is None:
        title = os.path.splitext(os.path.basename(docx_path))[0]
    res = convert_docx(docx_path, title=title)
    if not res.html:
        raise RuntimeError("converter produced no HTML for %s" % docx_path)
    return res.html


def default_output_path(docx_path):
    """Predictable on-disk location for the generated preview HTML."""
    preview_dir = os.path.join(_PKG_ROOT, "ui", "preview")
    os.makedirs(preview_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(docx_path))[0]
    return os.path.join(preview_dir, stem + ".html")


class _PreviewHandler(BaseHTTPRequestHandler):
    """Serves the single self-contained HTML document for any GET path."""

    output_path = ""  # set per-server before serve_forever()

    def do_GET(self):
        try:
            with open(self.output_path, "rb") as fh:
                data = fh.read()
        except OSError as exc:
            self.send_error(500, "cannot read preview html: %s" % exc)
            return
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, *args):
        # Keep the dev server quiet; the startup banner is enough.
        pass


def make_server(html_path, host="127.0.0.1", port=8000):
    """Return a ThreadingHTTPServer serving ``html_path`` at ``host:port``."""
    _PreviewHandler.output_path = html_path
    return ThreadingHTTPServer((host, port), _PreviewHandler)


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Local DOCX->HTML preview server (development only).")
    ap.add_argument("docx", help="Path to a .docx file to preview.")
    ap.add_argument("--host", default="127.0.0.1", help="Bind host.")
    ap.add_argument("--port", type=int, default=8000, help="Bind port.")
    ap.add_argument("--title", default=None, help="Override HTML <title>.")
    ap.add_argument("--output", default=None,
                    help="Where to write the HTML (default: ui/preview/<stem>.html).")
    ap.add_argument("--no-browser", action="store_true",
                    help="Do not attempt to open a browser.")
    args = ap.parse_args(argv)

    docx_path = os.path.abspath(args.docx)
    if not os.path.isfile(docx_path):
        ap.error("DOCX not found: %s" % docx_path)

    html = convert_to_html(docx_path, title=args.title)

    out_path = os.path.abspath(args.output) if args.output else default_output_path(docx_path)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "wb") as fh:
        fh.write(html.encode("utf-8"))

    server = make_server(out_path, host=args.host, port=args.port)
    url = "http://%s:%d" % (args.host, args.port)
    print("DOCX:   %s" % docx_path)
    print("HTML:   %s (%d bytes, self-contained)" % (out_path, len(html)))
    print("Preview: %s" % url)
    if not args.no_browser:
        try:
            webbrowser.open(url)
        except Exception:
            pass
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
