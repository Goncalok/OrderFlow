from __future__ import annotations

import json
from http.server import BaseHTTPRequestHandler

from ._auth import current_user_from_headers, unauthorized_response
from ._shared import export_leverschema_results


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        if current_user_from_headers(dict(self.headers)) is None:
            unauthorized_response(self)
            return
        content_length = int(self.headers.get("content-length", "0"))
        body = self.rfile.read(content_length)

        try:
            payload = json.loads(body.decode("utf-8") or "{}")
        except json.JSONDecodeError:
            self._send_json({"error": "Invalid Leverschema payload."}, 400)
            return

        leverschema_results = payload.get("results")
        selected_sheet = payload.get("selectedSheet")
        session_date = payload.get("sessionDate")
        if not isinstance(leverschema_results, dict):
            self._send_json({"error": "No Leverschema results were provided."}, 400)
            return
        if not isinstance(selected_sheet, str) or not selected_sheet.strip():
            self._send_json({"error": "No Leverschema sheet was selected."}, 400)
            return

        try:
            output_name, output_bytes = export_leverschema_results(
                leverschema_results,
                selected_sheet.strip(),
                session_date if isinstance(session_date, str) else None,
            )
            content_type = (
                "application/vnd.ms-excel.sheet.macroEnabled.12"
                if output_name.lower().endswith(".xlsm")
                else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Disposition", f'attachment; filename="{output_name}"')
            self.send_header("Content-Length", str(len(output_bytes)))
            self.end_headers()
            self.wfile.write(output_bytes)
        except Exception as exc:
            self._send_json({"error": str(exc)}, 400)

    def _send_json(self, payload, status):
        data = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)
