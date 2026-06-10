from __future__ import annotations

import json
from http.server import BaseHTTPRequestHandler

from ._auth import current_user_from_headers, unauthorized_response
from ._shared import export_laadschema_data


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
            self._send_json({"error": "Invalid Laadschema payload."}, 400)
            return

        laadschema_data = payload.get("data")
        custom_trucks = payload.get("customTrucks", {})
        selected_day = payload.get("selectedDay")
        selected_date = payload.get("selectedDate")
        selected_week = payload.get("selectedWeek")
        
        if not isinstance(laadschema_data, dict):
            self._send_json({"error": "No Laadschema data was provided."}, 400)
            return
        if not isinstance(selected_day, str) or not selected_day.strip():
            self._send_json({"error": "No day was selected."}, 400)
            return

        try:
            output_name, output_bytes = export_laadschema_data(
                laadschema_data,
                selected_day.strip(),
                selected_date if isinstance(selected_date, str) else None,
                selected_week if isinstance(selected_week, (str, int)) else None,
                custom_trucks if isinstance(custom_trucks, dict) else {},
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
