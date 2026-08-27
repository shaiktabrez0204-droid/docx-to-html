"""Product web interface: DOCX upload -> existing converter -> HTML preview + download.

This module is the PRODUCT INPUT/OUTPUT LAYER only. It does NOT contain any
document-conversion logic. The conversion core (OOXML extraction, styles,
headings, hierarchy, TOC, numbering, images, floating images, rendering) lives
in ``semantic.pipeline.convert_docx``, which is the single source of truth and
is called verbatim here.

Design (stdlib only, no new dependencies):
  * The renderer emits ONE self-contained HTML document: CSS is inline, images
    are embedded as data: URLs, and TOC links are #anchor references. There are
    therefore NO external asset files to route, which means serving the single
    generated document cannot produce broken asset requests and cannot reach
    the server filesystem.
  * Uploaded bytes are validated (extension + real ZIP/DOCX check + size) before
    any conversion. The file is held in a temp location only for the duration of
    the conversion call, then removed.
  * The generated HTML is kept in an in-memory store keyed by a random id and
    served via /preview/<id> (inline) and /download/<id> (attachment). Nothing
    is written to the repository; the only temporary file is the uploaded docx,
    cleaned up immediately after conversion.

Safe preview: the converted HTML is user content, but because it is fully
self-contained (only data: URLs and #anchors) it cannot reference server paths
or arbitrary files. It is served as text/html from this same origin; no
converter internals or server paths are exposed.
"""

import html
import io
import logging
import os
import re
import sys
import tempfile
import time
import uuid
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

_PKG_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PKG_ROOT not in sys.path:
    sys.path.insert(0, _PKG_ROOT)

from semantic.pipeline import convert_docx  # noqa: E402

log = logging.getLogger("docx_web")

MAX_UPLOAD_BYTES = 25 * 1024 * 1024  # 25 MB reasonable limit

# Status codes reused by the form renderer for clean user-facing messages.
ERR_NO_FILE = "no-file"
ERR_EXT = "bad-extension"
ERR_NOT_DOCX = "not-a-docx"
ERR_TOO_LARGE = "too-large"
ERR_CONVERT = "conversion-failed"

_ERROR_MESSAGES = {
    ERR_NO_FILE: "No file was selected. Please choose a .docx file.",
    ERR_EXT: "Only .docx files are accepted.",
    ERR_NOT_DOCX: "The uploaded file is not a valid DOCX (it is not a ZIP/OOXML package).",
    ERR_TOO_LARGE: "The file is too large (maximum 25 MB).",
    ERR_CONVERT: "The document could not be converted. The file may be corrupted or use unsupported features.",
}

_FILENAME_STRIP = re.compile(r"[^A-Za-z0-9._-]+")
_LEADING_DOTS = re.compile(r"^[.\-]+")


class _DocumentStore:
    """In-memory store of generated HTML keyed by a random id.

    No generated HTML is written to disk, so nothing accumulates in the
    repository. Entries are evicted lazily when they age out.
    """

    def __init__(self, ttl_seconds=600):
        self._docs = {}
        self._ttl = ttl_seconds

    def put(self, html_text, filename):
        doc_id = uuid.uuid4().hex
        self._docs[doc_id] = {
            "html": html_text,
            "filename": filename,
            "ts": time.time(),
        }
        self._evict()
        return doc_id

    def get(self, doc_id):
        doc = self._docs.get(doc_id)
        if doc is None:
            return None
        if time.time() - doc["ts"] > self._ttl:
            self._docs.pop(doc_id, None)
            return None
        return doc

    def _evict(self):
        now = time.time()
        expired = [k for k, v in self._docs.items() if now - v["ts"] > self._ttl]
        for k in expired:
            self._docs.pop(k, None)


_STORE = _DocumentStore()


def safe_filename(docx_name, default="converted"):
    """Derive a safe .html filename from the uploaded DOCX name.

    Strips path separators and unsafe characters; never trusts the client.
    Returns something like ``Annual_Report.html``.
    """
    base = os.path.basename(docx_name or "")
    stem = os.path.splitext(base)[0]
    stem = _FILENAME_STRIP.sub("_", stem)
    stem = _LEADING_DOTS.sub("", stem)
    stem = stem.strip("._-")
    if not stem:
        stem = default
    return stem + ".html"


def is_real_docx(data):
    """Return True when ``data`` is a ZIP/OOXML package, not just any blob."""
    if len(data) < 4 or data[:2] != b"PK":
        return False
    try:
        import zipfile
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            names = zf.namelist()
        # A DOCX must contain at least the content-types part and document.xml.
        return "[Content_Types].xml" in names and "word/document.xml" in names
    except Exception:
        return False


def parse_multipart(body, boundary):
    """Minimal multipart/form-data parser (cgi was removed in 3.13).

    Returns {field_name: (filename, bytes_or_str)}. File parts carry bytes and a
    filename; text fields carry a decoded string. Only the uploaded file is
    needed here.
    """
    parts = {}
    delimiter = b"--" + boundary
    for chunk in body.split(delimiter):
        if chunk in (b"", b"--", b"--\r\n"):
            continue
        # Strip the leading CRLF that follows the boundary.
        if chunk.startswith(b"\r\n"):
            chunk = chunk[2:]
        if b"\r\n\r\n" not in chunk:
            continue
        raw_headers, raw_content = chunk.split(b"\r\n\r\n", 1)
        # Trailing CRLF before the next boundary belongs to the framing.
        if raw_content.endswith(b"\r\n"):
            raw_content = raw_content[:-2]
        headers = {}
        for line in raw_headers.split(b"\r\n"):
            if b":" in line:
                k, v = line.split(b":", 1)
                headers[k.decode("latin-1").strip().lower()] = v.decode("latin-1")
        disp = headers.get("content-disposition", "")
        name_m = re.search(r'name="([^"]*)"', disp)
        file_m = re.search(r'filename="([^"]*)"', disp)
        field = name_m.group(1) if name_m else ""
        if file_m:
            parts[field] = (file_m.group(1), raw_content)
        else:
            parts[field] = (None, raw_content.decode("utf-8", "replace"))
    return parts


def validate_and_read(handler):
    """Validate the upload and return (data_bytes, original_filename, error_code).

    Performs every check the product requires: a file is present, the extension
    is .docx, the content is a real DOCX/ZIP package, and the size is within
    limits. Filename extension alone is NEVER trusted.
    """
    ctype = handler.headers.get("Content-Type", "")
    m = re.search(r"boundary=([^;]+)", ctype)
    if not m:
        return None, None, ERR_NO_FILE
    boundary = m.group(1).strip().strip('"').encode("utf-8")
    try:
        length = int(handler.headers.get("Content-Length", "0"))
    except ValueError:
        length = 0
    if length <= 0:
        return None, None, ERR_NO_FILE
    body = handler.rfile.read(length)
    parts = parse_multipart(body, boundary)
    file_part = parts.get("docx")
    if not file_part or file_part[1] is None or not file_part[1]:
        return None, None, ERR_NO_FILE
    filename, data = file_part
    if not (filename or "").lower().endswith(".docx"):
        return None, None, ERR_EXT
    if len(data) > MAX_UPLOAD_BYTES:
        return None, None, ERR_TOO_LARGE
    if not is_real_docx(data):
        return None, None, ERR_NOT_DOCX
    return data, filename, None


def run_conversion(data, filename):
    """Save to a temp docx, call the EXISTING converter, clean up, return HTML.

    The uploaded bytes are written to a safe temp file only for the duration of
    the convert_docx() call, then removed. The converter remains the single
    source of truth; this function adds no rendering logic.
    """
    tmp_path = None
    try:
        fd, tmp_path = tempfile.mkstemp(suffix=".docx", prefix="docxup_")
        with os.fdopen(fd, "wb") as fh:
            fh.write(data)
        title = os.path.splitext(os.path.basename(filename or "document"))[0]
        res = convert_docx(tmp_path, title=title)
        if not res.html:
            raise RuntimeError("converter produced no HTML")
        return res.html
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass


def _form_page(message=None, message_kind="error", result_html=None,
               doc_id=None, download_name=None):
    """Render the upload UI (and optionally a result) as a full HTML page."""
    if message:
        msg_class = "ok" if message_kind == "ok" else "err"
        msg_block = '<div class="message %s" role="alert">%s</div>' % (
            msg_class, html.escape(message))
    else:
        msg_block = ""

    if result_html is not None and doc_id is not None:
        preview_src = "/preview/%s" % doc_id
        dl_name = html.escape(download_name or "converted.html")
        result_block = (
            '<div class="result">\n'
            '  <div class="result-bar">\n'
            '    <span class="badge">Converted</span>\n'
            '    <a class="btn" href="/download/%s" download="%s">Download HTML</a>\n'
            '    <a class="btn ghost" href="/">Convert another</a>\n'
            '  </div>\n'
            '  <iframe class="preview" src="%s" '
            'title="Converted document preview"></iframe>\n'
            '</div>' % (html.escape(doc_id), dl_name, preview_src)
        )
    else:
        result_block = ""

    return (
        "<!DOCTYPE html>\n"
        '<html lang="en">\n'
        "<head>\n"
        '  <meta charset="UTF-8">\n'
        '  <meta name="viewport" content="width=device-width, initial-scale=1">\n'
        "  <title>DOCX to HTML Converter</title>\n"
        "  <style>\n"
        "    body { font: 15px/1.5 system-ui, sans-serif; background: #f4f5f7; "
        "margin: 0; color: #1c1e21; }\n"
        "    .wrap { max-width: 980px; margin: 0 auto; padding: 32px 20px 60px; }\n"
        "    h1 { font-size: 22px; margin: 0 0 4px; }\n"
        "    .sub { color: #666; margin: 0 0 22px; }\n"
        "    .card { background: #fff; border: 1px solid #e3e5e8; border-radius: 10px; "
        "padding: 22px; }\n"
        "    .drop { border: 2px dashed #c4c9d0; border-radius: 8px; padding: 26px; "
        "text-align: center; }\n"
        "    input[type=file] { width: 100%; }\n"
        "    .btn { display: inline-block; background: #2563eb; color: #fff; "
        "border: 0; border-radius: 6px; padding: 10px 16px; font-size: 14px; "
        "cursor: pointer; text-decoration: none; }\n"
        "    .btn:hover { background: #1d4ed8; }\n"
        "    .btn.ghost { background: #eef1f5; color: #333; margin-left: 8px; }\n"
        "    .message { border-radius: 6px; padding: 10px 14px; margin: 16px 0; }\n"
        "    .message.err { background: #fde8e8; color: #9b1c1c; border: 1px solid #f5c2c2; }\n"
        "    .message.ok { background: #e7f6ec; color: #1c7c3f; border: 1px solid #b6e2c6; }\n"
        "    .result { margin-top: 22px; }\n"
        "    .result-bar { display: flex; align-items: center; gap: 10px; margin-bottom: 10px; }\n"
        "    .badge { background: #1c7c3f; color: #fff; border-radius: 4px; "
        "padding: 3px 8px; font-size: 12px; }\n"
        "    iframe.preview { width: 100%; height: 70vh; border: 1px solid #e3e5e8; "
        "border-radius: 8px; background: #fff; }\n"
        "    .hint { color: #888; font-size: 13px; margin-top: 10px; }\n"
        "  </style>\n"
        "</head>\n"
        "<body>\n"
        '<div class="wrap">\n'
        "  <h1>DOCX &rarr; HTML Converter</h1>\n"
        '  <p class="sub">Upload a .docx file to convert it to HTML and preview / '
        "download the result.</p>\n"
        '  <div class="card">\n'
        '    <form method="post" action="/upload" '
        'enctype="multipart/form-data">\n'
        '      <div class="drop">\n'
        '        <input type="file" name="docx" accept=".docx" required>\n'
        '        <p class="hint">Only .docx files (max 25 MB).</p>\n'
        "      </div>\n"
        '      <p style="margin-top:16px;">\n'
        '        <button class="btn" type="submit">Upload &amp; Convert</button>\n'
        "      </p>\n"
        "    </form>\n"
        + msg_block + "\n"
        "  </div>\n"
        + result_block + "\n"
        "</div>\n"
        "</body>\n"
        "</html>\n"
    )


class DocxWebHandler(BaseHTTPRequestHandler):
    """Handles the product endpoints: GET / , POST /upload, /preview, /download."""

    def log_message(self, fmt, *args):
        log.info("%s - %s" % (self.address_string(), fmt % args))

    def _send(self, status, body_bytes, content_type, extra=None):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body_bytes)))
        self.send_header("Cache-Control", "no-store")
        if extra:
            for k, v in extra.items():
                self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body_bytes)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        if path == "/" or path == "/index.html":
            self._send(200, _form_page().encode("utf-8"),
                       "text/html; charset=utf-8")
            return
        if path.startswith("/preview/"):
            self._serve_stored(path[len("/preview/"):], download=False)
            return
        if path.startswith("/download/"):
            self._serve_stored(path[len("/download/"):], download=True)
            return
        self._send(404, b"Not found", "text/plain; charset=utf-8")

    def _serve_stored(self, doc_id, download):
        doc = _STORE.get(doc_id)
        if doc is None:
            self._send(404, b"Preview expired or not found.",
                       "text/plain; charset=utf-8")
            return
        body = doc["html"].encode("utf-8")
        extra = {}
        if download:
            # RFC 5987 / 6266: safe ASCII filename, plus UTF-8 fallback.
            name = doc["filename"]
            ascii_name = _FILENAME_STRIP.sub("_", name)
            extra["Content-Disposition"] = (
                "attachment; filename=\"%s\"; filename*=UTF-8''%s"
                % (ascii_name, urllib.parse.quote(name))
            )
        self._send(200, body, "text/html; charset=utf-8", extra)

    def do_POST(self):
        if urllib.parse.urlparse(self.path).path != "/upload":
            self._send(404, b"Not found", "text/plain; charset=utf-8")
            return
        data, filename, err = validate_and_read(self)
        if err is not None:
            log.warning("upload rejected: %s (file=%s)", err, filename)
            self._send(400, _form_page(_ERROR_MESSAGES[err]).encode("utf-8"),
                       "text/html; charset=utf-8")
            return
        try:
            log.info("converting uploaded file: %s (%d bytes)",
                     filename, len(data))
            html_text = run_conversion(data, filename)
            log.info("RUN_CONVERSION docx-number=%d len=%d",
                     html_text.count("docx-number"), len(html_text))
        except Exception as exc:  # converter or IO failure -> clean message
            log.exception("conversion failed for %s: %s", filename, exc)
            self._send(400, _form_page(_ERROR_MESSAGES[ERR_CONVERT]).encode("utf-8"),
                       "text/html; charset=utf-8")
            return
        dl_name = safe_filename(filename)
        doc_id = _STORE.put(html_text, dl_name)
        log.info("conversion ok: %s -> doc %s (%d bytes html)", filename,
                 doc_id, len(html_text))
        self._send(200, _form_page(result_html=html_text, doc_id=doc_id,
                                   download_name=dl_name).encode("utf-8"),
                   "text/html; charset=utf-8")


def make_server(host="127.0.0.1", port=8000):
    return ThreadingHTTPServer((host, port), DocxWebHandler)


def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser(description="DOCX -> HTML web converter (product UI).")
    ap.add_argument("--host", default="127.0.0.1", help="Bind host.")
    ap.add_argument("--port", type=int, default=8000, help="Bind port.")
    args = ap.parse_args(argv)

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
    server = make_server(host=args.host, port=args.port)
    url = "http://%s:%d" % (args.host, args.port)
    print("DOCX -> HTML converter running at: %s" % url)
    print("Press Ctrl+C to stop.")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
