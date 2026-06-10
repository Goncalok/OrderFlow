from __future__ import annotations

import json
from http.server import BaseHTTPRequestHandler

from ._auth import clear_session_cookie


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        data = json.dumps({"ok": True}).encode("utf-8")
        self.send_response(200)
        clear_session_cookie(self)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)
