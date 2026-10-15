#!/usr/bin/env python3
"""Web server for the thesis format checker site and API.

Serves site/ and exposes POST /api/check, which runs the real checker on an
uploaded .docx and returns the findings as JSON. The uploaded file is written
to a temp file, checked, and deleted in the same request: nothing is stored.

Usage:
    python3 server.py [port, default $PORT or 8000]
"""
from __future__ import annotations

import json
import os
import tempfile
import threading
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import purdue_format_check as pfc

HERE = os.path.dirname(os.path.abspath(__file__))
SITE = os.path.join(HERE, "site")
MAX_BYTES = 50 * 1024 * 1024
CHECK_LOCK = threading.Lock()  # config lives in module globals, so one analysis at a time
ALLOWED_ORIGINS = {
    origin.strip()
    for origin in os.environ.get(
        "ALLOWED_ORIGINS",
        "https://mogiiee.github.io,http://127.0.0.1:8000,"
        "http://127.0.0.1:8130,http://localhost:8000",
    ).split(",")
    if origin.strip()
}

MIME = {".html": "text/html; charset=utf-8", ".css": "text/css; charset=utf-8",
        ".js": "application/javascript; charset=utf-8", ".json": "application/json"}


def _universities():
    d = os.path.join(HERE, "universities")
    return sorted(f[:-5] for f in os.listdir(d) if f.endswith(".json")) if os.path.isdir(d) else []


def _parse_multipart(body: bytes, boundary: bytes):
    """Minimal multipart/form-data parser: returns (fields, files). Files are (filename, bytes)."""
    parts = body.split(b"--" + boundary)
    fields, files = {}, {}
    for part in parts:
        if b'Content-Disposition' not in part:
            continue
        head, _, content = part.partition(b"\r\n\r\n")
        content = content[: -2] if content.endswith(b"\r\n") else content
        name = head.split(b'name="')[1].split(b'"')[0].decode()
        if b'filename="' in head:
            filename = head.split(b'filename="')[1].split(b'"')[0].decode()
            files[name] = (filename, content)
        else:
            fields[name] = content.decode()
    return fields, files


class Handler(BaseHTTPRequestHandler):
    server_version = "ThesisCheck/0.1"

    def log_message(self, *a):
        pass

    def _send(self, code, data: bytes, ctype):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        origin = self.headers.get("Origin")
        if origin in ALLOWED_ORIGINS:
            self.send_header("Access-Control-Allow-Origin", origin)
            self.send_header("Vary", "Origin")
        self.end_headers()
        self.wfile.write(data)

    def _json(self, code, obj):
        self._send(code, json.dumps(obj).encode(), "application/json")

    def do_OPTIONS(self):
        if urllib.parse.urlparse(self.path).path not in ("/api/check", "/api/universities"):
            return self._json(404, {"error": "not found"})
        self.send_response(204)
        origin = self.headers.get("Origin")
        if origin in ALLOWED_ORIGINS:
            self.send_header("Access-Control-Allow-Origin", origin)
            self.send_header("Vary", "Origin")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Max-Age", "86400")
        self.end_headers()

    def do_GET(self):
        path = urllib.parse.urlparse(self.path).path
        if path == "/api/universities":
            return self._json(200, {"universities": _universities()})
        rel = "index.html" if path in ("/", "") else path.lstrip("/").replace("..", "")
        full = os.path.join(SITE, rel)
        if not os.path.isfile(full):
            return self._json(404, {"error": "not found"})
        with open(full, "rb") as f:
            self._send(200, f.read(), MIME.get(os.path.splitext(full)[1], "application/octet-stream"))

    def do_POST(self):
        if urllib.parse.urlparse(self.path).path != "/api/check":
            return self._json(404, {"error": "not found"})
        length = int(self.headers.get("Content-Length", 0))
        if length > MAX_BYTES:
            return self._json(413, {"error": "file too large (50 MB max)"})
        ctype = self.headers.get("Content-Type", "")
        if "multipart/form-data" not in ctype or "boundary=" not in ctype:
            return self._json(400, {"error": "expected multipart/form-data"})
        boundary = ctype.split("boundary=")[1].encode()
        fields, files = _parse_multipart(self.rfile.read(length), boundary)
        if "file" not in files:
            return self._json(400, {"error": "no file part"})
        filename, data = files["file"]
        if not filename.lower().endswith(".docx") or data[:2] != b"PK":
            return self._json(400, {"error": "not a .docx file"})
        uni = fields.get("university", "purdue") or "purdue"
        try:
            cfg = pfc.load_university(uni)
        except (FileNotFoundError, KeyError) as e:
            return self._json(400, {"error": str(e)})
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            tmp.write(data)
            tmppath = tmp.name
        with CHECK_LOCK:
            pfc.apply_config(cfg)
            try:
                issues = pfc.check(tmppath)
            except Exception as e:  # corrupt file, not a thesis — say so plainly
                return self._json(422, {"error": f"could not analyze this file: {e}"})
            finally:
                try:
                    os.unlink(tmppath)
                except OSError:
                    pass
        report = pfc.serialize_findings(issues, filename, uni)
        print(f"checked {len(data)} bytes -> {report['verdict']} "
              f"({report['errors']} errors); temp file deleted, nothing stored", flush=True)
        self._json(200, report)


if __name__ == "__main__":
    import sys
    port = int(sys.argv[1]) if len(sys.argv) > 1 else int(os.environ.get("PORT", "8000"))
    host = os.environ.get("HOST", "0.0.0.0")
    srv = ThreadingHTTPServer((host, port), Handler)
    print(f"serving site/ and POST /api/check on http://{host}:{port} (Ctrl-C to stop)")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
