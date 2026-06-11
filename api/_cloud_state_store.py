from __future__ import annotations

import json
import os
import urllib.error
import urllib.parse
import urllib.request
from typing import Any


SUPABASE_STATE_TABLE = os.environ.get("SUPABASE_STATE_TABLE", "orderflow_app_state")


def supabase_enabled() -> bool:
    return bool(_supabase_credentials()[0] and _supabase_credentials()[1])


def load_json_state(key: str) -> Any | None:
    url, service_key = _supabase_credentials()
    if not url or not service_key:
        return None

    encoded_key = urllib.parse.quote(str(key), safe="")
    endpoint = (
        f"{url.rstrip('/')}/rest/v1/{SUPABASE_STATE_TABLE}"
        f"?key=eq.{encoded_key}&select=value"
    )
    req = urllib.request.Request(endpoint, headers=_supabase_headers(service_key))
    try:
        with urllib.request.urlopen(req, timeout=25) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = ""
        try:
            detail = exc.read().decode("utf-8", errors="replace")[:400]
        except OSError:
            pass
        raise OSError(f"Supabase HTTP {exc.code} {detail}") from exc
    except (urllib.error.URLError, json.JSONDecodeError, TimeoutError) as exc:
        raise OSError(f"Supabase read failed: {exc}") from exc
    if not isinstance(payload, list) or not payload:
        return None
    value = payload[0].get("value") if isinstance(payload[0], dict) else None
    return value


def save_json_state(key: str, value: Any) -> None:
    url, service_key = _supabase_credentials()
    if not url or not service_key:
        raise RuntimeError("Supabase credentials are not configured.")

    endpoint = f"{url.rstrip('/')}/rest/v1/{SUPABASE_STATE_TABLE}?on_conflict=key"
    data = json.dumps(
        {
            "key": str(key),
            "value": value,
        },
        ensure_ascii=False,
        separators=(",", ":"),
    ).encode("utf-8")
    headers = _supabase_headers(service_key)
    headers.update(
        {
            "Content-Type": "application/json; charset=utf-8",
            "Prefer": "resolution=merge-duplicates,return=minimal",
        }
    )
    req = urllib.request.Request(endpoint, data=data, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=45):
            return
    except urllib.error.HTTPError as exc:
        detail = ""
        try:
            detail = exc.read().decode("utf-8", errors="replace")[:400]
        except OSError:
            pass
        raise OSError(f"Supabase HTTP {exc.code} {detail}") from exc
    except urllib.error.URLError as exc:
        raise OSError(str(exc)) from exc


def _supabase_credentials() -> tuple[str, str]:
    url = _normalize_supabase_url(os.environ.get("SUPABASE_URL", ""))
    service_key = (
        os.environ.get("SUPABASE_SERVICE_ROLE_KEY")
        or os.environ.get("SUPABASE_SERVICE_KEY")
        or ""
    ).strip()
    return url, service_key


def _normalize_supabase_url(value: str) -> str:
    url = str(value or "").strip().rstrip("/")
    if url.endswith("/rest/v1"):
        url = url[: -len("/rest/v1")]
    return url


def _supabase_headers(service_key: str) -> dict[str, str]:
    return {
        "apikey": service_key,
        "Authorization": f"Bearer {service_key}",
        "Accept": "application/json",
    }
