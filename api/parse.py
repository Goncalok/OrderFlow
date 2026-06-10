from http.server import BaseHTTPRequestHandler

from ._auth import current_user_from_headers, unauthorized_response
from ._shared import json_response, parse_uploaded_email, read_uploaded_file


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        if current_user_from_headers(dict(self.headers)) is None:
            unauthorized_response(self)
            return
        content_length = int(self.headers.get("content-length", "0"))
        body = self.rfile.read(content_length)
        file_name, file_bytes = read_uploaded_file(dict(self.headers), body)
        if not file_name or file_bytes is None:
            json_response(self, {"error": "No email file was uploaded."}, status=400)
            return

        try:
            payload = parse_uploaded_email(file_name, file_bytes)
            json_response(self, payload)
        except Exception as exc:
            json_response(self, {"error": str(exc)}, status=400)
