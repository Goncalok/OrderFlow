from http.server import BaseHTTPRequestHandler
from urllib.parse import parse_qs

from ._auth import current_user_from_headers, unauthorized_response
from ._shared import export_preview_data, export_uploaded_email, read_form_payload


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        if current_user_from_headers(dict(self.headers)) is None:
            unauthorized_response(self)
            return
        content_length = int(self.headers.get("content-length", "0"))
        body = self.rfile.read(content_length)
        file_name, file_bytes, fields = read_form_payload(dict(self.headers), body)
        
        query = parse_qs(self.path.split("?", 1)[1] if "?" in self.path else "")
        export_type = query.get("type", ["selected"])[0]
        selected_client = query.get("selectedClient", [""])[0].strip()
        
        # Allow Denemark and Edeka Laatzen CMR export without email
        client_lower = (selected_client or "").strip().lower()
        allow_without_email = export_type == "print_cmr" and (
            ("denemark" in client_lower or "denmark" in client_lower) or 
            "edeka laatzen" in client_lower or
            "edeka mochmuhl" in client_lower or
            "globus" in client_lower or
            "havi nl" in client_lower or
            "havi be" in client_lower or
            "havi de" in client_lower or
            "heeren" in client_lower or
            "nettomd" in client_lower or
            "rewe" in client_lower or
            "penny" in client_lower or
            "hanos" in client_lower or
            "carrefour fif" in client_lower or
            "carrefour kdc" in client_lower or
            "colruyt" in client_lower or
            ("havi" in client_lower and "duisburg" in client_lower)
        )
        
        has_preview_export = bool(fields.get("previewExport", "").strip()) and export_type in {"selected", "merge"}
        if (not file_name or file_bytes is None) and not allow_without_email and not has_preview_export:
            self._send_json({"error": "No email file was uploaded."}, 400)
            return

        order_index = int(query.get("orderIndex", ["0"])[0])
        export_sheet = query.get("exportSheet", [""])[0].strip()
        cmr_references = []
        cmr_pallet_places = fields.get("cmrPalletPlaces", "").strip() or None
        dc_name = fields.get("dcName", "").strip() or None
        pakbon_items = None
        goederen_total = None
        quantity_overrides = None
        merge_order_indexes = []
        preview_export = None
        overrides_raw = fields.get("quantityOverrides", "").strip()
        if overrides_raw:
            try:
                import json

                quantity_overrides = json.loads(overrides_raw)
            except json.JSONDecodeError:
                self._send_json({"error": "Invalid quantity override payload."}, 400)
                return

        merge_indexes_raw = fields.get("mergeOrderIndexes", "").strip()
        if merge_indexes_raw:
            try:
                import json

                parsed_indexes = json.loads(merge_indexes_raw)
                if isinstance(parsed_indexes, list):
                    merge_order_indexes = parsed_indexes
            except json.JSONDecodeError:
                self._send_json({"error": "Invalid merge order selection payload."}, 400)
                return

        preview_export_raw = fields.get("previewExport", "").strip()
        if preview_export_raw:
            try:
                import json

                parsed_preview = json.loads(preview_export_raw)
                if isinstance(parsed_preview, dict):
                    preview_export = parsed_preview
            except json.JSONDecodeError:
                self._send_json({"error": "Invalid saved order payload."}, 400)
                return
        
        # Parse pakbon items if provided
        pakbon_items_raw = fields.get("pakbonItems", "").strip()
        if pakbon_items_raw:
            try:
                import json
                pakbon_items = json.loads(pakbon_items_raw)
            except json.JSONDecodeError:
                self._send_json({"error": "Invalid pakbon items payload."}, 400)
                return
        
        # Parse goederen total if provided (for Colruyt)
        goederen_total_raw = fields.get("goederenTotal", "").strip()
        if goederen_total_raw:
            try:
                goederen_total = float(goederen_total_raw)
            except ValueError:
                self._send_json({"error": "Invalid goederen total value."}, 400)
                return
        
        cmr_refs_raw = fields.get("cmrReferences", "").strip()
        if cmr_refs_raw:
            try:
                import json

                parsed_refs = json.loads(cmr_refs_raw)
                if isinstance(parsed_refs, list):
                    cmr_references = [str(value).strip() for value in parsed_refs if str(value).strip()]
            except json.JSONDecodeError:
                self._send_json({"error": "Invalid CMR references payload."}, 400)
                return

        try:
            if isinstance(preview_export, dict) and export_type in {"selected", "merge"}:
                output_name, output_bytes = export_preview_data(
                    preview_export,
                    export_type,
                    order_index,
                    quantity_overrides,
                    export_sheet or None,
                    selected_client or None,
                    merge_order_indexes,
                )
            else:
                output_name, output_bytes = export_uploaded_email(
                    file_name,
                    file_bytes,
                    export_type,
                    order_index,
                    quantity_overrides,
                    export_sheet or None,
                    selected_client or None,
                    cmr_references,
                    cmr_pallet_places,
                    dc_name,
                    pakbon_items,
                    goederen_total,
                    merge_order_indexes,
                )
            self.send_response(200)
            content_type = (
                "application/vnd.ms-excel.sheet.macroEnabled.12"
                if str(output_name).lower().endswith(".xlsm")
                else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Disposition", f'attachment; filename="{output_name}"')
            self.send_header("Content-Length", str(len(output_bytes)))
            self.end_headers()
            self.wfile.write(output_bytes)
        except Exception as exc:
            self._send_json({"error": str(exc)}, 400)

    def _send_json(self, payload, status):
        import json

        data = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)
