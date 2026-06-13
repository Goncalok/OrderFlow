from __future__ import annotations

import json
from http.server import BaseHTTPRequestHandler

from ._auth import current_user_from_headers, unauthorized_response
from ._shared import json_response
from greenops_shortage_bridge import remove_preview_sessions


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        if current_user_from_headers(dict(self.headers)) is None:
            unauthorized_response(self)
            return

        content_length = int(self.headers.get("content-length", "0"))
        if content_length <= 0 or content_length > 200_000:
            json_response(self, {"error": "Invalid payload size."}, 400)
            return

        try:
            payload = json.loads(self.rfile.read(content_length).decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            json_response(self, {"error": "Invalid JSON body."}, 400)
            return
        if not isinstance(payload, dict):
            json_response(self, {"error": "Body must be a JSON object."}, 400)
            return

        session_date = str(payload.get("date") or "").strip()
        work_session_id = str(payload.get("workSessionId") or "").strip()
        reference = str(payload.get("reference") or "").strip()
        customer = str(payload.get("customer") or "").strip()
        fatrans = str(payload.get("fatrans") or "").strip()
        delivery_point = str(payload.get("deliveryPoint") or "").strip()
        if not session_date and not work_session_id:
            json_response(self, {"error": "Missing work session context."}, 400)
            return
        if not any([reference, customer, fatrans, delivery_point]):
            json_response(self, {"error": "Missing order identity."}, 400)
            return

        try:
            removed, sessions = remove_preview_sessions(
                session_date=session_date,
                work_session_id=work_session_id,
                reference=reference,
                customer=customer,
                fatrans=fatrans,
                delivery_point=delivery_point,
            )
        except Exception as exc:
            json_response(self, {"error": f"Could not remove order from Manco's: {exc}"}, 400)
            return
        json_response(self, {"removed": removed, "sessions": sessions})
