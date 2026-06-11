from __future__ import annotations

import json
import uuid
from datetime import date, datetime
from http.server import BaseHTTPRequestHandler

from ._auth import current_user_from_headers, unauthorized_response
from ._team_state_store import load_team_state, save_team_state


class handler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        user = current_user_from_headers(dict(self.headers))
        if user is None:
            unauthorized_response(self)
            return
        try:
            state = load_team_state()
        except OSError as exc:
            self._json_err(f"Could not load team state: {exc}", 500)
            return
        self._json_ok(state)

    def do_POST(self) -> None:
        user = current_user_from_headers(dict(self.headers))
        if user is None:
            unauthorized_response(self)
            return
        payload = self._read_json_body(max_size=200_000)
        if payload is None:
            self._json_err("Invalid JSON body.", 400)
            return
        session_date = str(payload.get("date") or "").strip()
        try:
            date.fromisoformat(session_date)
        except ValueError:
            self._json_err("Please choose a valid session date.", 400)
            return
        now = datetime.now().isoformat(timespec="seconds")
        session = {
            "id": str(uuid.uuid4()),
            "createdBy": str(user.get("email") or ""),
            "date": session_date,
            "name": str(payload.get("name") or f"Plan for {session_date}").strip(),
            "createdAt": now,
            "updatedAt": now,
            "workspaces": {},
        }
        try:
            save_team_state({"sessions": [session]})
            state = load_team_state()
        except OSError as exc:
            self._json_err(f"Could not persist work session: {exc}", 500)
            return
        if not any(entry.get("id") == session["id"] for entry in state.get("sessions", [])):
            state["sessions"] = [session, *state.get("sessions", [])]
        self._json_ok({"session": session, **state})

    def do_PUT(self) -> None:
        user = current_user_from_headers(dict(self.headers))
        if user is None:
            unauthorized_response(self)
            return
        payload = self._read_json_body(max_size=6_000_000)
        if payload is None:
            self._json_err("Invalid JSON body.", 400)
            return
        try:
            save_team_state(payload)
        except OSError as exc:
            self._json_err(f"Could not persist team state: {exc}", 500)
            return
        self._json_ok({"ok": True})

    def _read_json_body(self, max_size: int) -> dict | None:
        content_length = int(self.headers.get("content-length", "0"))
        if content_length <= 0 or content_length > max_size:
            return None
        raw_body = self.rfile.read(content_length)
        try:
            payload = json.loads(raw_body.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            return None
        return payload if isinstance(payload, dict) else None

    def _json_ok(self, payload: dict) -> None:
        data = json.dumps(payload).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _json_err(self, message: str, status: int) -> None:
        data = json.dumps({"error": message}).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)
