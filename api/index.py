from __future__ import annotations

import json
import os
import re
import sys
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
import uuid
from datetime import datetime
from email import policy
from email.parser import BytesParser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler
from io import BytesIO
from pathlib import Path

from email_order_app.excel_writer import fill_template, fill_cmr_template, fill_colruyt_cmr_template
from email_order_app.models import CustomerOrder, OrderItem, ParsedOrderEmail
from email_order_app.parser import parse_order_email
from email_order_app.special_havi_uien import HaviUienDestination, HaviUienOrder, parse_havi_uien_email, write_havi_uien_excel
from email_order_app.special_netto_md import NettoMdEmail, NettoMdItem, NettoMdOrder, parse_netto_md_email, write_netto_md_excel


BASE_DIR = Path(__file__).resolve().parent.parent
STATIC_DIR = BASE_DIR / "public"
OUTPUT_DIR = BASE_DIR / "output" / ".webtmp"
TEMPLATE_PATH = BASE_DIR / "ORDER REWE PENNY Print In Color.xlsx"
CMR_TEMPLATE_PATH = BASE_DIR / "CMR" / "FIF CMR DOCUMENT.xlsm"
KDC_CMR_TEMPLATE_PATH = BASE_DIR / "CMR" / "KDC CMR DOCUMENT.xlsm"
COLRUYT_CMR_TEMPLATE_PATH = BASE_DIR / "CMR" / "CMR Colruyt.xlsm"
NETTO_MD_TEMPLATE_PATH = BASE_DIR / "Order picking NETTO blanco.xlsx"
STORAGE_PATH = BASE_DIR / "data" / "web_sessions.json"
REDIS_STATE_KEY = "web_app:sessions:v1"
SESSIONS: dict[str, dict[str, object]] = {}


def _redis_rest_credentials() -> tuple[str, str]:
    url = (
        os.environ.get("UPSTASH_REDIS_REST_URL")
        or os.environ.get("KV_REST_API_URL")
        or ""
    ).strip()
    token = (
        os.environ.get("UPSTASH_REDIS_REST_TOKEN")
        or os.environ.get("KV_REST_API_TOKEN")
        or ""
    ).strip()
    return url, token


def _load_sessions() -> dict[str, dict[str, object]]:
    url, token = _redis_rest_credentials()
    if url and token:
        raw = _upstash_get(url, token, REDIS_STATE_KEY)
        if raw:
            try:
                loaded = json.loads(raw)
                if isinstance(loaded, list):
                    sessions: dict[str, dict[str, object]] = {}
                    for entry in loaded:
                        if not isinstance(entry, dict):
                            continue
                        session_id = entry.get("id")
                        mode = entry.get("mode")
                        data = entry.get("data")
                        if not isinstance(session_id, str):
                            continue
                        if mode not in {"standard", "special", "netto_md"}:
                            continue
                        if not isinstance(data, dict):
                            continue
                        sessions[session_id] = {"mode": mode, "data": data}
                    return sessions
            except json.JSONDecodeError:
                pass
        return {}

    if not STORAGE_PATH.is_file():
        return {}

    try:
        raw = json.loads(STORAGE_PATH.read_text(encoding="utf-8"))
        if not isinstance(raw, list):
            return {}

        sessions: dict[str, dict[str, object]] = {}
        for entry in raw:
            if not isinstance(entry, dict):
                continue
            session_id = entry.get("id")
            mode = entry.get("mode")
            data = entry.get("data")
            if not isinstance(session_id, str):
                continue
            if mode not in {"standard", "special", "netto_md"}:
                continue
            if not isinstance(data, dict):
                continue
            sessions[session_id] = {"mode": mode, "data": data}
        return sessions
    except (OSError, json.JSONDecodeError):
        return {}


def _save_sessions() -> None:
    payload = [
        {"id": session_id, "mode": session_data["mode"], "data": session_data["data"]}
        for session_id, session_data in SESSIONS.items()
    ]
    encoded = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))

    url, token = _redis_rest_credentials()
    if url and token:
        _upstash_set(url, token, REDIS_STATE_KEY, encoded)
        return

    STORAGE_PATH.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = STORAGE_PATH.with_suffix(".json.tmp")
    tmp_path.write_text(encoded, encoding="utf-8")
    tmp_path.replace(STORAGE_PATH)


def _upstash_get(url: str, token: str, key: str) -> str | None:
    from urllib.parse import quote

    endpoint = url.rstrip("/") + "/get/" + quote(key, safe="")
    req = urllib.request.Request(endpoint, headers={"Authorization": f"Bearer {token}"})
    try:
        with urllib.request.urlopen(req, timeout=25) as resp:
            body = json.loads(resp.read().decode("utf-8"))
    except (urllib.error.URLError, urllib.error.HTTPError, json.JSONDecodeError, TimeoutError):
        return None
    if isinstance(body, dict) and body.get("error"):
        return None
    result = body.get("result")
    if result is None:
        return None
    if isinstance(result, str):
        return result
    return json.dumps(result)


def _upstash_set(url: str, token: str, key: str, value: str) -> None:
    from urllib.parse import quote

    endpoint = url.rstrip("/") + "/set/" + quote(key, safe="")
    req = urllib.request.Request(
        endpoint,
        data=value.encode("utf-8"),
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json; charset=utf-8",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=45) as resp:
            body = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = ""
        try:
            detail = exc.read().decode("utf-8", errors="replace")[:300]
        except OSError:
            pass
        raise OSError(f"Upstash HTTP {exc.code} {detail}") from exc
    except urllib.error.URLError as exc:
        raise OSError(str(exc)) from exc
    if not isinstance(body, dict):
        raise OSError("Unexpected Upstash response")
    if body.get("error"):
        raise OSError(str(body["error"]))


def _serialize_standard_session(parsed: ParsedOrderEmail) -> dict[str, object]:
    return {
        "source_file": parsed.source_file,
        "subject": parsed.subject,
        "sender": parsed.sender,
        "received_at": parsed.received_at.isoformat() if parsed.received_at else None,
        "delivery_date_to_dc": parsed.delivery_date_to_dc,
        "leaving_date_venlo": parsed.leaving_date_venlo,
        "leaving_time_venlo": parsed.leaving_time_venlo,
        "orders": [
            {
                "customer_dc": order.customer_dc,
                "fatrans_dc": order.fatrans_dc,
                "reference": order.reference,
                "items": [
                    {
                        "article_number": item.article_number,
                        "description": item.description,
                        "quantity_boxes": item.quantity_boxes,
                        "unit": item.unit,
                    }
                    for item in order.items
                ],
            }
            for order in parsed.orders
        ],
    }


def _serialize_special_session(data: HaviUienOrder) -> dict[str, object]:
    return {
        "source_file": data.source_file,
        "subject": data.subject,
        "order_date": data.order_date,
        "delivery_date": data.delivery_date,
        "ve_per_pallet": data.ve_per_pallet,
        "product_name": data.product_name,
        "unit_label": data.unit_label,
        "case_label": data.case_label,
        "destinations": [
            {
                "destination": item.destination,
                "reference": item.reference,
                "company": item.company,
                "slices": item.slices,
                "cases": item.cases,
            }
            for item in data.destinations
        ],
        "total_slices": data.total_slices,
        "total_cases": data.total_cases,
    }


def _serialize_netto_md_session(data: NettoMdEmail) -> dict[str, object]:
    return {
        "source_file": data.source_file,
        "subject": data.subject,
        "delivery_date": data.delivery_date,
        "dc_name": data.dc_name,
        "orders": [
            {
                "customer_name": order.customer_name,
                "delivery_name": order.delivery_name,
                "sales_order": order.sales_order,
                "reference": order.reference,
                "items": [
                    {
                        "article_number": item.article_number,
                        "description": item.description,
                        "quantity": item.quantity,
                        "unit": item.unit,
                    }
                    for item in order.items
                ],
            }
            for order in data.orders
        ],
    }


def _deserialize_standard(data: dict[str, object]) -> ParsedOrderEmail:
    orders = []
    for order_data in data.get("orders", []):
        if not isinstance(order_data, dict):
            continue
        items = []
        for item_data in order_data.get("items", []):
            if not isinstance(item_data, dict):
                continue
            items.append(
                OrderItem(
                    article_number=str(item_data.get("article_number", "")),
                    description=str(item_data.get("description", "")),
                    quantity_boxes=float(item_data.get("quantity_boxes", 0)),
                    unit=str(item_data.get("unit", "")),
                )
            )
        orders.append(
            CustomerOrder(
                customer_dc=str(order_data.get("customer_dc", "")),
                fatrans_dc=str(order_data.get("fatrans_dc", "")),
                reference=str(order_data.get("reference", "")),
                items=items,
            )
        )

    received_at_raw = data.get("received_at")
    received_at = None
    if isinstance(received_at_raw, str) and received_at_raw:
        try:
            received_at = datetime.fromisoformat(received_at_raw)
        except ValueError:
            received_at = None

    return ParsedOrderEmail(
        source_file=str(data.get("source_file", "")),
        subject=str(data.get("subject", "")),
        sender=str(data.get("sender", "")),
        received_at=received_at,
        delivery_date_to_dc=str(data.get("delivery_date_to_dc", "")),
        leaving_date_venlo=str(data.get("leaving_date_venlo", "")),
        leaving_time_venlo=str(data.get("leaving_time_venlo", "")),
        orders=orders,
    )


def _deserialize_special(data: dict[str, object]) -> HaviUienOrder:
    destinations = []
    for dest_data in data.get("destinations", []):
        if not isinstance(dest_data, dict):
            continue
        destinations.append(
            HaviUienDestination(
                destination=str(dest_data.get("destination", "")),
                reference=str(dest_data.get("reference", "")),
                company=str(dest_data.get("company", "")),
                slices=int(dest_data.get("slices", 0)),
                cases=int(dest_data.get("cases", 0)),
            )
        )

    return HaviUienOrder(
        source_file=str(data.get("source_file", "")),
        subject=str(data.get("subject", "")),
        order_date=str(data.get("order_date", "")),
        delivery_date=str(data.get("delivery_date", "")),
        ve_per_pallet=str(data.get("ve_per_pallet", "")),
        product_name=str(data.get("product_name", "")),
        unit_label=str(data.get("unit_label", "")),
        case_label=str(data.get("case_label", "")),
        destinations=destinations,
        total_slices=int(data.get("total_slices", 0)),
        total_cases=int(data.get("total_cases", 0)),
    )


def _deserialize_netto_md(data: dict[str, object]) -> NettoMdEmail:
    orders = []
    for order_data in data.get("orders", []):
        if not isinstance(order_data, dict):
            continue
        items = [
            NettoMdItem(
                article_number=str(item.get("article_number", "")),
                description=str(item.get("description", "")),
                quantity=float(item.get("quantity", 0)),
                unit=str(item.get("unit", "")),
            )
            for item in order_data.get("items", [])
            if isinstance(item, dict)
        ]
        orders.append(
            NettoMdOrder(
                customer_name=str(order_data.get("customer_name", "")),
                delivery_name=str(order_data.get("delivery_name", "")),
                sales_order=str(order_data.get("sales_order", "")),
                reference=str(order_data.get("reference", "")),
                items=items,
            )
        )
    return NettoMdEmail(
        source_file=str(data.get("source_file", "")),
        subject=str(data.get("subject", "")),
        delivery_date=str(data.get("delivery_date", "")),
        dc_name=str(data.get("dc_name", "NettoMD")),
        orders=orders,
    )


def _serialize_standard(parsed: ParsedOrderEmail) -> dict[str, object]:
    return {
        "deliveryDate": parsed.delivery_date_to_dc,
        "customerCount": len(parsed.orders),
        "orders": [
            {
                "label": _build_order_label(order, parsed.orders),
                "customer": order.customer_dc,
                "reference": order.reference,
                "fatrans": order.fatrans_dc,
                "items": [
                    {
                        "primary": item.article_number,
                        "secondary": item.description,
                        "quantity": _format_number(item.quantity_boxes),
                        "unit": item.unit,
                    }
                    for item in order.items
                ],
            }
            for order in parsed.orders
        ],
        "canMerge": _can_merge_orders(parsed.orders),
    }


def _serialize_special(data: HaviUienOrder) -> dict[str, object]:
    return {
        "deliveryDate": data.delivery_date,
        "customerCount": 1,
        "customer": "Havi Logistics GmbH",
        "reference": "Multiple VCSO references",
        "items": [
            {
                "primary": row.destination,
                "secondary": row.reference,
                "slicesQuantity": row.slices,
                "quantity": row.cases,
                "unit": data.case_label,
            }
            for row in data.destinations
        ],
    }


def _serialize_netto_md(data: NettoMdEmail) -> dict[str, object]:
    return {
        "deliveryDate": data.delivery_date,
        "customerCount": len(data.orders),
        "dcName": data.dc_name,
        "orders": [
            {
                "label": order.customer_name,
                "customer": order.customer_name,
                "reference": order.reference or order.sales_order,
                "fatrans": data.dc_name,
                "items": [
                    {
                        "primary": item.article_number,
                        "secondary": item.description,
                        "quantity": _format_number(item.quantity),
                        "unit": item.unit,
                    }
                    for item in order.items
                ],
            }
            for order in data.orders
        ],
        "canMerge": False,
    }


def _format_number(value: float) -> str:
    return str(int(value)) if value.is_integer() else f"{value:.2f}"


def _build_order_label(order: CustomerOrder, orders: list[CustomerOrder]) -> str:
    names = [item.customer_dc for item in orders]
    if names.count(order.customer_dc) > 1:
        return f"{order.customer_dc} ({order.reference})"
    return order.customer_dc


def _can_merge_orders(orders: list[CustomerOrder]) -> bool:
    if len(orders) < 2:
        return False
    return len({order.customer_dc.strip().lower() for order in orders}) == 1


def _merge_orders(orders: list[CustomerOrder]) -> CustomerOrder:
    if not orders:
        raise ValueError("No orders to merge")
    first = orders[0]
    merged_items = []
    for order in orders:
        merged_items.extend(order.items)
    return CustomerOrder(
        customer_dc=first.customer_dc,
        fatrans_dc=first.fatrans_dc,
        reference=f"{first.reference} + {len(orders) - 1} more",
        items=merged_items,
    )


def _select_orders_by_indexes(orders: list[CustomerOrder], indexes: list[int], fallback_index: int = 0) -> list[CustomerOrder]:
    selected: list[CustomerOrder] = []
    seen: set[int] = set()
    for raw_index in indexes or []:
        try:
            index = int(raw_index)
        except (TypeError, ValueError):
            continue
        if index in seen or index < 0 or index >= len(orders):
            continue
        selected.append(orders[index])
        seen.add(index)
    if selected:
        return selected
    if 0 <= fallback_index < len(orders):
        return [orders[fallback_index]]
    return orders[:1]


def _build_export_name(order: CustomerOrder, selected_client: str | None = None) -> str:
    customer = order.customer_dc or ""
    customer = re.sub(r'^(netto|rewe|penny|edeka)\s+', '', customer, flags=re.IGNORECASE).strip()
    customer = re.sub(r'[^\w\-_\. ]', '_', customer)

    if selected_client:
        client_lower = selected_client.lower()
        if client_lower == "rewe":
            return f"Rewe_{customer}.xlsx"
        elif client_lower == "penny":
            return f"Penny_{customer}.xlsx"

    return f"{customer}.xlsx"


SESSIONS = _load_sessions()


# Garantir que as sessões sejam carregadas na inicialização
_load_sessions()


class handler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        parsed_url = urllib.parse.urlparse(self.path)
        if parsed_url.path in {"/", "/index.html"}:
            self._serve_file(STATIC_DIR / "index.html", "text/html; charset=utf-8")
            return
        if parsed_url.path == "/styles.css":
            self._serve_file(STATIC_DIR / "styles.css", "text/css; charset=utf-8")
            return
        if parsed_url.path == "/app.js":
            self._serve_file(STATIC_DIR / "app.js", "application/javascript; charset=utf-8")
            return
        if parsed_url.path == "/api/session":
            self._handle_session(parsed_url)
            return
        if parsed_url.path == "/api/debug":
            self._handle_debug()
            return
        self.send_error(HTTPStatus.NOT_FOUND, "Not found")

    def do_POST(self) -> None:
        if self.path == "/api/parse":
            self._handle_parse()
            return
        if self.path == "/api/export":
            self._handle_export()
            return
        self.send_error(HTTPStatus.NOT_FOUND, "Not found")

    def log_message(self, format: str, *args) -> None:
        return

    def _handle_parse(self) -> None:
        file_name, file_bytes = self._read_uploaded_file()
        if not file_name or file_bytes is None:
            self._json_response({"error": "No email file was uploaded."}, status=HTTPStatus.BAD_REQUEST)
            return

        temp_path = OUTPUT_DIR / f"{uuid.uuid4().hex}_{file_name}"
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        temp_path.write_bytes(file_bytes)

        try:
            special = parse_havi_uien_email(temp_path)
            session_id = uuid.uuid4().hex
            if special is not None:
                SESSIONS[session_id] = {"mode": "special", "data": _serialize_special_session(special)}
                _save_sessions()  # Salvar imediatamente
                self._json_response({"sessionId": session_id, "mode": "special", "preview": _serialize_special(special)})
                return

            netto_md = parse_netto_md_email(temp_path)
            if netto_md is not None:
                SESSIONS[session_id] = {"mode": "netto_md", "data": _serialize_netto_md_session(netto_md)}
                _save_sessions()
                self._json_response({"sessionId": session_id, "mode": "netto_md", "preview": _serialize_netto_md(netto_md)})
                return

            parsed = parse_order_email(temp_path)
            SESSIONS[session_id] = {"mode": "standard", "data": _serialize_standard_session(parsed)}
            _save_sessions()  # Salvar imediatamente
            self._json_response({"sessionId": session_id, "mode": "standard", "preview": _serialize_standard(parsed)})
        except Exception as exc:
            self._json_response({"error": str(exc)}, status=HTTPStatus.BAD_REQUEST)
        finally:
            if temp_path.exists():
                temp_path.unlink()

    def _handle_session(self, parsed_url: urllib.parse.ParseResult) -> None:
        query = urllib.parse.parse_qs(parsed_url.query)
        session_id = query.get("sessionId", [""])[0]
        session = SESSIONS.get(session_id)
        if not session:
            self._json_response({"error": "Session not found."}, status=HTTPStatus.BAD_REQUEST)
            return

        try:
            if session["mode"] == "special":
                data = _deserialize_special(session["data"])
                preview = _serialize_special(data)
            elif session["mode"] == "netto_md":
                data = _deserialize_netto_md(session["data"])
                preview = _serialize_netto_md(data)
            else:
                parsed = _deserialize_standard(session["data"])
                preview = _serialize_standard(parsed)

            self._json_response({"sessionId": session_id, "mode": session["mode"], "preview": preview})
        except Exception as exc:
            self._json_response({"error": str(exc)}, status=HTTPStatus.BAD_REQUEST)

    def _handle_debug(self) -> None:
        """Endpoint de debug para verificar configuração do Upstash"""
        url, token = _redis_rest_credentials()

        debug_info = {
            "upstash_configured": bool(url and token),
            "upstash_url": url[:50] + "..." if url else None,
            "upstash_token_prefix": token[:10] + "..." if token else None,
            "sessions_count": len(SESSIONS),
            "session_ids": list(SESSIONS.keys())[:5],  # primeiros 5 IDs
            "environment_vars": {
                "KV_REST_API_URL": bool(os.environ.get("KV_REST_API_URL")),
                "KV_REST_API_TOKEN": bool(os.environ.get("KV_REST_API_TOKEN")),
                "UPSTASH_REDIS_REST_URL": bool(os.environ.get("UPSTASH_REDIS_REST_URL")),
                "UPSTASH_REDIS_REST_TOKEN": bool(os.environ.get("UPSTASH_REDIS_REST_TOKEN")),
            }
        }

        # Testar conexão com Upstash se configurado
        if url and token:
            try:
                test_result = _upstash_get(url, token, "test_key")
                debug_info["upstash_connection"] = "OK" if test_result is not None else "FAIL"
            except Exception as e:
                debug_info["upstash_connection"] = f"ERROR: {str(e)}"
        else:
            debug_info["upstash_connection"] = "NOT_CONFIGURED"

        self._json_response(debug_info)

    def _handle_export(self) -> None:
        parsed_url = urllib.parse.urlparse(self.path)
        query = urllib.parse.parse_qs(parsed_url.query)
        export_type = query.get("type", ["selected"])[0]
        order_index = int(query.get("orderIndex", ["0"])[0])
        url_selected_client = query.get("selectedClient", [""])[0]

        file_name_orig, file_bytes, fields = self._read_multipart_form_data()
        form_selected_client = fields.get("selectedClient", "").strip()
        merge_order_indexes = []
        raw_merge_indexes = fields.get("mergeOrderIndexes", "").strip()
        if raw_merge_indexes:
            try:
                parsed_merge_indexes = json.loads(raw_merge_indexes)
                if isinstance(parsed_merge_indexes, list):
                    merge_order_indexes = parsed_merge_indexes
            except json.JSONDecodeError:
                self._json_response({"error": "Invalid merge order selection payload."}, status=HTTPStatus.BAD_REQUEST)
                return

        # Robust client detection: prefer form field, fallback to URL query
        selected_client = form_selected_client or url_selected_client

        if not file_name_orig or file_bytes is None:
            self._json_response({"error": "No email file was uploaded."}, status=HTTPStatus.BAD_REQUEST)
            return

        temp_path = OUTPUT_DIR / f"{uuid.uuid4().hex}_{file_name_orig}"
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        temp_path.write_bytes(file_bytes)

        try:
            special = parse_havi_uien_email(temp_path)
            if special is not None:
                file_name = "HAVI_DE_UIEN.xlsx"
                output_path = OUTPUT_DIR / f"{uuid.uuid4().hex}_{file_name}"
                write_havi_uien_excel(output_path, special)
            else:
                netto_md = parse_netto_md_email(temp_path)
                if netto_md is not None:
                    export_date = datetime.now().strftime("%d-%m-%Y")
                    file_name = f"NettoMD_Orderpicking {export_date}.xlsx"
                    output_path = OUTPUT_DIR / f"{uuid.uuid4().hex}_{file_name}"
                    write_netto_md_excel(NETTO_MD_TEMPLATE_PATH, output_path, netto_md)
                else:
                    parsed = parse_order_email(temp_path)
                    if export_type == "merge":
                        order = _merge_orders(_select_orders_by_indexes(parsed.orders, merge_order_indexes, order_index))
                    else:
                        order = parsed.orders[order_index]

                    if export_type == "print_cmr":
                        raw_references = fields.get("cmrReferences", "[]")
                        raw_pallet_places = fields.get("cmrPalletPlaces", "").strip()
                        try:
                            cmr_references = json.loads(raw_references) if raw_references else []
                        except json.JSONDecodeError:
                            cmr_references = []

                        file_name = f"CMR_{_build_export_name(order, selected_client).removesuffix('.xlsx')}.xlsm"
                        output_path = OUTPUT_DIR / f"{uuid.uuid4().hex}_{file_name}"

                        fill_cmr_template(
                            CMR_TEMPLATE_PATH,
                            output_path,
                            parsed,
                            order,
                            references=cmr_references,
                            pallet_places=raw_pallet_places or None,
                        )
                    else:
                        file_name = _build_export_name(order, selected_client)
                        output_path = OUTPUT_DIR / f"{uuid.uuid4().hex}_{file_name}"
                        fill_template(TEMPLATE_PATH, output_path, parsed, order)

            data = output_path.read_bytes()
            self.send_response(HTTPStatus.OK)
            content_type = (
                "application/vnd.ms-excel.sheet.macroEnabled.12"
                if str(file_name).lower().endswith(".xlsm")
                else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Disposition", f'attachment; filename="{file_name}"')
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        except Exception as exc:
            self._json_response({"error": str(exc)}, status=HTTPStatus.BAD_REQUEST)
        finally:
            if temp_path.exists():
                temp_path.unlink()
            if "output_path" in locals() and output_path.exists():
                output_path.unlink()

    def _serve_file(self, path: Path, content_type: str) -> None:
        if not path.exists():
            self.send_error(HTTPStatus.NOT_FOUND, "Not found")
            return
        data = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _json_response(self, payload: dict[str, object], status: HTTPStatus = HTTPStatus.OK) -> None:
        data = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _send_json(self, payload: dict[str, object], status: int) -> None:
        self._json_response(payload, HTTPStatus(status))

    def _read_uploaded_file(self) -> tuple[str | None, bytes | None]:
        content_type = self.headers.get("Content-Type", "")
        content_length = int(self.headers.get("Content-Length", "0"))
        if "multipart/form-data" not in content_type or content_length <= 0:
            return None, None

        body = self.rfile.read(content_length)
        raw = f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8") + body
        message = BytesParser(policy=policy.default).parsebytes(raw)
        for part in message.iter_parts():
            if part.get_content_disposition() != "form-data":
                continue
            filename = part.get_filename()
            if filename:
                return filename, part.get_payload(decode=True)
        return None, None

    def _read_multipart_form_data(self) -> tuple[str | None, bytes | None, dict[str, str]]:
        content_type = self.headers.get("Content-Type", "")
        content_length = int(self.headers.get("Content-Length", "0"))
        if "multipart/form-data" not in content_type or content_length <= 0:
            return None, None, {}

        body = self.rfile.read(content_length)
        raw = f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8") + body
        message = BytesParser(policy=policy.default).parsebytes(raw)

        fields: dict[str, str] = {}
        file_name: str | None = None
        file_bytes: bytes | None = None

        for part in message.iter_parts():
            if part.get_content_disposition() != "form-data":
                continue
            name = part.get_param("name", header="content-disposition")
            filename = part.get_filename()
            if filename:
                file_name = filename
                file_bytes = part.get_payload(decode=True)
            elif name:
                payload = part.get_payload(decode=True)
                if isinstance(payload, bytes):
                    payload = payload.decode(part.get_content_charset() or "utf-8", errors="replace")
                fields[name] = payload
        return file_name, file_bytes, fields
