from __future__ import annotations

import json
from http.server import BaseHTTPRequestHandler

from ._shared import json_response, parse_uploaded_email, read_form_payload
from greenops_shortage_bridge import build_shortage_previews, save_previews_as_sessions


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        content_length = int(self.headers.get("content-length", "0"))
        if content_length <= 0 or content_length > 12_000_000:
            json_response(self, {"error": "Invalid upload size."}, 400)
            return

        try:
            body = self.rfile.read(content_length)
            file_name, file_bytes, fields = read_form_payload(dict(self.headers), body)

            session_date = str(fields.get("date") or "").strip()
            session_name = str(fields.get("name") or "").strip()
            work_session_id = str(fields.get("workSessionId") or "").strip()
            selected_client = str(fields.get("selectedClient") or "").strip()
            if not session_date:
                json_response(self, {"error": "Open a daily session before uploading an email."}, 400)
                return

            raw_preview = str(fields.get("preview") or "").strip()
            if raw_preview:
                mode = str(fields.get("mode") or "standard").strip() or "standard"
                try:
                    greenops_preview = json.loads(raw_preview)
                except json.JSONDecodeError:
                    json_response(self, {"error": "Invalid OrderFlow preview payload."}, 400)
                    return
                parsed = {"mode": mode, "preview": greenops_preview}
                file_name = file_name or "OrderFlow saved orders"
            else:
                if not file_name or file_bytes is None:
                    json_response(self, {"error": "No OrderFlow email file was uploaded."}, 400)
                    return
                parsed = parse_uploaded_email(file_name, file_bytes)
                mode = str(parsed.get("mode") or "")
                greenops_preview = parsed.get("preview")
            if not isinstance(greenops_preview, dict):
                json_response(self, {"error": "OrderFlow did not return an order preview."}, 400)
                return

            shortage_previews = build_shortage_previews(file_name, mode, greenops_preview, selected_client)
            if not shortage_previews:
                json_response(self, {"error": "No order lines were found in this email."}, 400)
                return

            saved, sessions = save_previews_as_sessions(session_date, session_name, shortage_previews, work_session_id)
            json_response(
                self,
                {
                    "mode": mode,
                    "greenopsPreview": greenops_preview,
                    "savedSessions": saved,
                    "sessions": sessions,
                },
            )
        except Exception as exc:
            json_response(self, {"error": f"Could not ingest OrderFlow order: {exc}"}, 400)
