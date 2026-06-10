from __future__ import annotations

import json
import tempfile
import uuid
from http.server import BaseHTTPRequestHandler
from pathlib import Path

from ._auth import current_user_from_headers, unauthorized_response
from ._shared import read_form_payload


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        if current_user_from_headers(dict(self.headers)) is None:
            unauthorized_response(self)
            return
        
        content_length = int(self.headers.get("content-length", "0"))
        body = self.rfile.read(content_length)
        
        try:
            # Parse multipart form data to get uploaded files
            _, _, fields = read_form_payload(dict(self.headers), body)
            
            # Get uploaded files from the request
            # Files are sent as separate parts in multipart/form-data
            from email import policy
            from email.parser import BytesParser
            
            headers_lower = {key.lower(): value for key, value in dict(self.headers).items()}
            content_type = headers_lower.get("content-type", "")
            
            if "multipart/form-data" not in content_type:
                self._send_json({"error": "Expected multipart/form-data"}, 400)
                return
            
            # Parse the multipart data
            raw = f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8") + body
            message = BytesParser(policy=policy.default).parsebytes(raw)
            
            pdf_files = []
            temp_dir = Path(tempfile.gettempdir()) / "pakbon_uploads"
            temp_dir.mkdir(parents=True, exist_ok=True)
            
            for part in message.iter_parts():
                if part.get_content_disposition() != "form-data":
                    continue
                
                filename = part.get_filename()
                if filename and filename.lower().endswith('.pdf'):
                    payload = part.get_payload(decode=True)
                    if payload:
                        # Save to temp file
                        temp_path = temp_dir / f"{uuid.uuid4().hex}_{filename}"
                        temp_path.write_bytes(payload)
                        pdf_files.append(temp_path)
            
            if not pdf_files:
                self._send_json({"error": "No PDF files uploaded"}, 400)
                return
            
            # Parse the pakbon files
            try:
                from email_order_app.pakbon_parser import parse_multiple_pakbons, calculate_goederen_total
                
                # Parse Emballage section (for Carrefour FIF/KDC)
                merged_items = parse_multiple_pakbons(pdf_files, section="Emballage")
                
                # Calculate Goederen total (for Colruyt)
                goederen_total = calculate_goederen_total(pdf_files)
                
                # Calculate CHEP total (for Carrefour KDC - article 409)
                chep_total = 0
                for item in merged_items.values():
                    if item.article_number == "409" or "chep" in item.description.lower():
                        chep_total += item.quantity
                
                # Convert to JSON-serializable format
                result = {
                    "items": [
                        {
                            "articleNumber": item.article_number,
                            "description": item.description,
                            "quantity": item.quantity,
                            "unit": item.unit,
                        }
                        for item in merged_items.values()
                    ],
                    "totalItems": len(merged_items),
                    "goederenTotal": goederen_total,  # Total from Goederen section for Colruyt
                    "chepTotal": chep_total,  # Total CHEP quantity for Carrefour KDC
                }
                
                self._send_json(result, 200)
            
            finally:
                # Clean up temp files
                for temp_path in pdf_files:
                    try:
                        temp_path.unlink()
                    except:
                        pass
        
        except Exception as exc:
            self._send_json({"error": str(exc)}, 400)
    
    def _send_json(self, payload, status):
        data = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)
