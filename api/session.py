from __future__ import annotations

import json
from http.server import BaseHTTPRequestHandler

from ._auth import current_user_from_headers, unauthorized_response


class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        user = current_user_from_headers(dict(self.headers))
        if user is None:
            unauthorized_response(self)
            return

        data = json.dumps({"user": user}).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)
