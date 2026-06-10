from __future__ import annotations

import json
from http.server import BaseHTTPRequestHandler

from ._auth import current_user_from_headers, unauthorized_response
from ._team_state_store import load_team_state, save_team_state


class handler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        user = current_user_from_headers(dict(self.headers))
        if user is None:
            unauthorized_response(self)
            return
        state = load_team_state()
        self._json_ok(state)

    def do_PUT(self) -> None:
        user = current_user_from_headers(dict(self.headers))
        if user is None:
            unauthorized_response(self)
            return
        content_length = int(self.headers.get("content-length", "0"))
        if content_length <= 0 or content_length > 6_000_000:
            self._json_err("Invalid payload size.", 400)
            return
        raw_body = self.rfile.read(content_length)
        try:
            payload = json.loads(raw_body.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            self._json_err("Invalid JSON body.", 400)
            return
        if not isinstance(payload, dict):
            self._json_err("Body must be a JSON object.", 400)
            return
        try:
            save_team_state(payload)
        except OSError:
            self._json_err("Could not persist team state.", 500)
            return
        self._json_ok({"ok": True})

    def _json_ok(self, payload: dict) -> None:
        data = json.dumps(payload).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _json_err(self, message: str, status: int) -> None:
        data = json.dumps({"error": message}).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)
