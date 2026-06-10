from __future__ import annotations

import json
from http.server import BaseHTTPRequestHandler

from ._auth import authenticate_user, set_session_cookie


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        content_length = int(self.headers.get("content-length", "0"))
        body = self.rfile.read(content_length)
        try:
            payload = json.loads(body.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            self._send_json({"error": "Invalid login payload."}, 400)
            return

        email = str(payload.get("email", "")).strip()
        password = str(payload.get("password", ""))
        user = authenticate_user(email, password)
        if user is None:
            self._send_json({"error": "Invalid email or password."}, 401)
            return

        data = json.dumps({"user": user}).encode("utf-8")
        self.send_response(200)
        set_session_cookie(self, user)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _send_json(self, payload, status):
        data = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)
