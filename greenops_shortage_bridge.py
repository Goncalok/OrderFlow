from __future__ import annotations

import re
import uuid
from datetime import date, datetime
from pathlib import Path
from typing import Any

from shortage_app import clean_number, load_sessions, recognize_client, repair_display_texts, save_sessions


UNKNOWN_CLIENT = "Unknown client"
UNKNOWN_POINT = "Unknown delivery point"
MAIN_POINT = "Main delivery point"


def build_shortage_previews(file_name: str, mode: str, greenops_preview: dict[str, Any]) -> list[dict[str, Any]]:
    if mode == "special":
        special_orders = greenops_preview.get("specialOrders")
        if isinstance(special_orders, list):
            previews: list[dict[str, Any]] = []
            for special_preview in special_orders:
                if not isinstance(special_preview, dict):
                    continue
                special_file_name = str(special_preview.get("sourceFileName") or file_name).strip() or file_name
                preview = _build_special_preview(special_file_name, special_preview)
                if preview["items"]:
                    previews.append(preview)
            return previews
        return [_build_special_preview(str(greenops_preview.get("sourceFileName") or file_name).strip() or file_name, greenops_preview)]

    previews: list[dict[str, Any]] = []
    orders = greenops_preview.get("orders")
    if not isinstance(orders, list):
        return previews

    for index, order in enumerate(orders):
        if not isinstance(order, dict):
            continue
        order_file_name = str(order.get("sourceFileName") or file_name).strip() or file_name
        preview = _build_order_preview(order_file_name, mode, greenops_preview, order, index)
        if preview["items"]:
            previews.append(preview)
    return previews


def save_previews_as_sessions(
    session_date: str,
    session_name: str,
    previews: list[dict[str, Any]],
    work_session_id: str = "",
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    previews = [repair_display_texts(preview) for preview in previews]
    parsed_date = date.fromisoformat(session_date)
    base_name = session_name.strip() or f"{parsed_date.strftime('%A')} Manco session"
    sessions = load_sessions()
    saved: list[dict[str, Any]] = []
    new_session_ids: set[str] = set()
    existing_by_key = {
        _session_identity(entry.get("preview", {}), str(entry.get("workSessionId") or ""), str(entry.get("date") or "")): entry
        for entry in sessions
        if isinstance(entry, dict)
    }
    timestamp = datetime.now().isoformat(timespec="seconds")

    for preview in previews:
        identity = _session_identity(preview, work_session_id, session_date)
        existing = existing_by_key.get(identity)
        if existing:
            _refresh_existing_session_metadata(existing, preview)
            saved.append(existing)
            continue
        session = {
            "id": str(uuid.uuid4()),
            "date": session_date,
            "weekday": parsed_date.strftime("%A"),
            "name": base_name,
            "workSessionId": work_session_id,
            "createdAt": timestamp,
            "preview": preview,
        }
        saved.append(session)
        new_session_ids.add(session["id"])
        existing_by_key[identity] = session

    new_sessions = [entry for entry in saved if str(entry.get("id") or "") in new_session_ids]
    sessions = new_sessions + sessions
    save_sessions(sessions[:1000])
    return saved, sessions[:1000]


def _refresh_existing_session_metadata(session: dict[str, Any], preview: dict[str, Any]) -> None:
    existing_preview = session.get("preview")
    if not isinstance(existing_preview, dict):
        session["preview"] = preview
        return

    for key in [
        "sourceFile",
        "client",
        "deliveryPoint",
        "confidence",
        "greenopsMode",
        "greenopsOrderIndex",
        "greenopsReference",
        "greenopsCustomer",
        "greenopsFatrans",
        "greenopsDeliveryDate",
    ]:
        if key in preview:
            existing_preview[key] = preview[key]

    if not _has_recorded_shortages(existing_preview):
        for key in ["orderedTotal", "deliveredTotal", "shortageTotal", "shortagePercentage", "lineCount", "items"]:
            if key in preview:
                existing_preview[key] = preview[key]


def _has_recorded_shortages(preview: dict[str, Any]) -> bool:
    items = preview.get("items")
    if not isinstance(items, list):
        return False
    return any(_number(item.get("shortageQuantity")) > 0 for item in items if isinstance(item, dict))


def _build_order_preview(
    file_name: str,
    mode: str,
    greenops_preview: dict[str, Any],
    order: dict[str, Any],
    order_index: int,
) -> dict[str, Any]:
    customer = str(order.get("customer") or "").strip()
    fatrans = str(order.get("fatrans") or "").strip()
    reference = str(order.get("reference") or "").strip()
    client, point, confidence = _infer_client_delivery(file_name, customer, fatrans, reference)
    explicit_point = str(order.get("deliveryPoint") or "").strip()
    if explicit_point:
        point = explicit_point
        confidence = max(confidence, 90)
    if client.casefold() == "nettomd":
        netto_point = _clean_nettomd_delivery_point(customer) or _clean_nettomd_delivery_point(point)
        if netto_point:
            point = netto_point
            confidence = max(confidence, 95)
    items = [_to_shortage_item(item) for item in order.get("items", []) if isinstance(item, dict)]
    items = [item for item in items if item["orderedQuantity"] > 0]
    return _preview_payload(
        file_name=file_name,
        mode=mode,
        greenops_preview=greenops_preview,
        client=client,
        delivery_point=point,
        confidence=confidence,
        items=items,
        order_index=order_index,
        reference=reference,
        customer=customer,
        fatrans=fatrans,
    )


def _build_special_preview(file_name: str, greenops_preview: dict[str, Any]) -> dict[str, Any]:
    items = []
    for item in greenops_preview.get("items", []):
        if not isinstance(item, dict):
            continue
        ordered = _number(item.get("quantity"))
        if ordered <= 0:
            continue
        items.append(
            {
                "article": str(item.get("primary") or "").strip(),
                "description": str(item.get("secondary") or "").strip(),
                "orderedQuantity": clean_number(ordered),
                "deliveredQuantity": clean_number(ordered),
                "shortageQuantity": 0,
                "shortagePercentage": 0,
            }
        )
    return _preview_payload(
        file_name=file_name,
        mode="special",
        greenops_preview=greenops_preview,
        client="HAVI",
        delivery_point=str(greenops_preview.get("deliveryPoint") or _infer_havi_delivery_point(file_name, greenops_preview)),
        confidence=95,
        items=items,
        order_index=0,
        reference=_special_reference(greenops_preview),
        customer=str(greenops_preview.get("customer") or "Havi Logistics GmbH"),
        fatrans="HAVI DE UIEN",
    )


def _preview_payload(
    *,
    file_name: str,
    mode: str,
    greenops_preview: dict[str, Any],
    client: str,
    delivery_point: str,
    confidence: int,
    items: list[dict[str, Any]],
    order_index: int,
    reference: str,
    customer: str,
    fatrans: str,
) -> dict[str, Any]:
    ordered_total = sum(float(item["orderedQuantity"]) for item in items)
    shortage_total = sum(float(item["shortageQuantity"]) for item in items)
    shortage_percentage = (shortage_total / ordered_total * 100) if ordered_total else 0
    return {
        "sourceFile": Path(file_name).name,
        "client": client or UNKNOWN_CLIENT,
        "deliveryPoint": delivery_point or UNKNOWN_POINT,
        "confidence": confidence,
        "orderedTotal": clean_number(ordered_total),
        "shortageTotal": clean_number(shortage_total),
        "shortagePercentage": round(shortage_percentage, 2),
        "lineCount": len(items),
        "items": items,
        "greenopsMode": mode,
        "greenopsOrderIndex": order_index,
        "greenopsReference": reference,
        "greenopsCustomer": customer,
        "greenopsFatrans": fatrans,
        "greenopsDeliveryDate": greenops_preview.get("deliveryDate") or "",
    }


def _to_shortage_item(item: dict[str, Any]) -> dict[str, Any]:
    ordered = _number(item.get("quantity"))
    return {
        "article": str(item.get("primary") or "").strip(),
        "description": str(item.get("secondary") or "").strip(),
        "orderedQuantity": clean_number(ordered),
        "deliveredQuantity": clean_number(ordered),
        "shortageQuantity": 0,
        "shortagePercentage": 0,
    }


def _infer_client_delivery(file_name: str, *parts: str) -> tuple[str, str, int]:
    context = " ".join(part for part in parts if part)
    recognized = recognize_client(file_name, context)
    client = str(recognized.get("client") or UNKNOWN_CLIENT)
    point = str(recognized.get("deliveryPoint") or UNKNOWN_POINT)
    confidence = int(recognized.get("confidence") or 0)

    fallback_point = _fallback_delivery_point(context)
    explicit_point = _explicit_delivery_point(*parts)
    first_part = str(parts[0] or "").strip() if parts else ""
    if _looks_like_denemark_order(*parts):
        client = "Denemark"
        denemark_point = _clean_denemark_delivery_point(first_part) or _clean_denemark_delivery_point(context)
        if denemark_point:
            point = denemark_point
        confidence = max(confidence, 90)
    if client.casefold() == "carrefour" and explicit_point in {"FIF", "KDC"}:
        point = explicit_point
        confidence = max(confidence, 90)
    if client.casefold() == "havi":
        havi_point = _infer_havi_delivery_point(file_name, {"customer": context, "items": []})
        if havi_point != "HAVI":
            point = havi_point
            confidence = max(confidence, 90)
    if client.casefold() == "denemark" and first_part:
        denemark_point = _clean_denemark_delivery_point(first_part)
        if denemark_point:
            point = denemark_point
            confidence = max(confidence, 80)
    if client.casefold() == "nettomd":
        netto_point = _clean_nettomd_delivery_point(context)
        if netto_point:
            point = netto_point
            confidence = max(confidence, 95)
    if client.casefold() in {"rewe", "penny"} and first_part and first_part.casefold() not in {"rewe", "penny", "rewe/penny"}:
        point = first_part
        confidence = max(confidence, 75)
    if point in {"", UNKNOWN_POINT, MAIN_POINT} and fallback_point:
        point = fallback_point
        confidence = max(confidence, 50)
    if client == UNKNOWN_CLIENT:
        guessed = _guess_client(context)
        if guessed:
            client = guessed
            confidence = max(confidence, 50)
    return client, point, min(confidence, 95)


def _fallback_delivery_point(context: str) -> str:
    cleaned = re.sub(r"\s+", " ", context).strip()
    return cleaned[:80] if cleaned else UNKNOWN_POINT


def _clean_denemark_delivery_point(value: str) -> str:
    cleaned = re.sub(r"\s+", " ", str(value or "").strip(" -:\t"))
    cleaned = re.sub(r"^netto\s+", "", cleaned, flags=re.IGNORECASE).strip()
    cleaned = re.sub(r"\bdenem?ark\b", "", cleaned, flags=re.IGNORECASE).strip()
    cleaned = re.sub(r"\b(VCSO|SSCO)\d+\b", "", cleaned, flags=re.IGNORECASE).strip()
    if not cleaned:
        return ""
    if re.search(r"\b(VCSO|SSCO)\d+\b", cleaned, re.IGNORECASE):
        return ""
    if cleaned.casefold() in {"denemark", "denmark", "netto"}:
        return ""
    return cleaned[:80]


def _looks_like_denemark_order(*parts: str) -> bool:
    normalized = " ".join(str(part or "") for part in parts).casefold()
    return (
        "denemark" in normalized
        or "denmark" in normalized
        or bool(re.search(r"\b(brabrand|køge|koge|koege)\b", normalized))
    )


def _clean_nettomd_delivery_point(value: str) -> str:
    normalized = re.sub(r"[^a-z0-9]+", "", str(value or "").casefold())
    for needle, label in {
        "kerpen": "Kerpen",
        "hodenhagen": "Hodenhagen",
        "henstedt": "Henstedt-Ulzburg",
        "henstedtulzburg": "Henstedt-Ulzburg",
        "hamm": "Hamm",
        "ganderkesee": "Ganderkesee",
        "bottrop": "Bottrop",
        "krefeld": "Krefeld",
    }.items():
        if needle in normalized:
            return label
    return ""


def _explicit_delivery_point(*parts: str) -> str:
    context = " ".join(str(part or "") for part in parts)
    normalized = context.casefold()
    if re.search(r"\bkdc\b", normalized):
        return "KDC"
    if re.search(r"\bfif\b", normalized):
        return "FIF"
    return ""


def _infer_havi_delivery_point(file_name: str, greenops_preview: dict[str, Any]) -> str:
    item_text = " ".join(
        f"{item.get('primary') or ''} {item.get('secondary') or ''}"
        for item in greenops_preview.get("items", [])
        if isinstance(item, dict)
    )
    context = f"{file_name} {greenops_preview.get('customer') or ''} {greenops_preview.get('reference') or ''} {item_text}".casefold()
    if "uien" in context or "onion" in context:
        return "HAVI DE UIEN"
    if "neu-wulmstorf" in context or "neu wulmstorf" in context:
        return "Neu Wulmstorf"
    if "duisburg" in context:
        return "Duisburg"
    if "wunstorf" in context:
        return "Wunstorf"
    if "havi nl" in context or "netherlands" in context or "nederland" in context:
        return "HAVI NL"
    if "havi be" in context or "belg" in context:
        return "HAVI BE"
    if "havi de" in context or "gmbh" in context:
        return "HAVI DE"
    return "HAVI"


def _special_reference(greenops_preview: dict[str, Any]) -> str:
    reference = str(greenops_preview.get("reference") or "").strip()
    if reference and reference.casefold() != "multiple vcso references":
        return reference
    for item in greenops_preview.get("items", []):
        if not isinstance(item, dict):
            continue
        value = str(item.get("secondary") or "").strip()
        if value.upper().startswith("VCSO"):
            return value
    return reference


def _guess_client(context: str) -> str:
    normalized = context.casefold()
    for label, needles in {
        "HAVI": ["havi"],
        "Carrefour": ["carrefour"],
        "Colruyt": ["colruyt"],
        "Edeka": ["edeka", "mochmuhl", "mockmuhl", "laatzen"],
        "NettoMD": ["netto md", "nettomd", "netto"],
        "Rewe": ["rewe"],
        "Penny": ["penny"],
        "Globus": ["globus"],
        "Hanos": ["hanos"],
        "Heeren": ["heeren", "herren"],
        "HelloFresh": ["hellofresh", "hello fresh"],
    }.items():
        if any(needle in normalized for needle in needles):
            return label
    return ""


def _number(value: Any) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(" ", "")
    if not text:
        return 0.0
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0


def _session_identity(preview: Any, work_session_id: str, session_date: str) -> str:
    if not isinstance(preview, dict):
        preview = {}
    scope = work_session_id or session_date
    reference = str(preview.get("greenopsReference") or "").strip().casefold()
    customer = str(preview.get("greenopsCustomer") or preview.get("client") or "").strip().casefold()
    fatrans = str(preview.get("greenopsFatrans") or preview.get("deliveryPoint") or "").strip().casefold()
    source = str(preview.get("sourceFile") or "").strip().casefold()
    order_index = str(preview.get("greenopsOrderIndex") or "")
    if reference and reference != "multiple vcso references":
        return "|".join([scope, reference, customer, fatrans])
    return "|".join([scope, source, order_index, customer, fatrans])
