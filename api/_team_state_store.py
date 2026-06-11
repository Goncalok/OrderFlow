from __future__ import annotations

# Shared work sessions + Leverschema for all logged-in users (any browser).
# Production: Upstash REST — either
#   UPSTASH_REDIS_REST_URL + UPSTASH_REDIS_REST_TOKEN (console), or
#   KV_REST_API_URL + KV_REST_API_TOKEN (Vercel Upstash integration).
# Local dev: data/team_state.json (gitignored).

import json
import os
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any

from ._cloud_state_store import load_json_state, save_json_state, supabase_enabled

BASE_DIR = Path(__file__).resolve().parent.parent
LOCAL_TEAM_STATE_PATH = BASE_DIR / "data" / "team_state.json"
REDIS_STATE_KEY = "greenops:team_state:v1"
SUPABASE_STATE_KEY = "team_state:v1"


def _redis_rest_credentials() -> tuple[str, str]:
    """Upstash console uses UPSTASH_*; Vercel integration often injects KV_REST_*."""
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


def load_team_state() -> dict[str, Any]:
    if supabase_enabled():
        data = load_json_state(SUPABASE_STATE_KEY)
        if isinstance(data, dict):
            return _coerce_team_state(data)
        return {
            "sessions": [],
            "leverschemaResults": {},
            "laadschemaData": {},
            "laadschemaCustomTrucks": {},
        }

    url, token = _redis_rest_credentials()
    if url and token:
        raw = _upstash_get(url, token, REDIS_STATE_KEY)
        if raw:
            try:
                data = json.loads(raw)
                if isinstance(data, dict):
                    return _coerce_team_state(data)
            except json.JSONDecodeError:
                pass
        return {
            "sessions": [],
            "leverschemaResults": {},
            "laadschemaData": {},
            "laadschemaCustomTrucks": {},
        }

    if LOCAL_TEAM_STATE_PATH.is_file():
        try:
            data = json.loads(LOCAL_TEAM_STATE_PATH.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return _coerce_team_state(data)
        except (OSError, json.JSONDecodeError):
            pass
    return {
        "sessions": [],
        "leverschemaResults": {},
        "laadschemaData": {},
        "laadschemaCustomTrucks": {},
    }


def save_team_state(data: dict[str, Any]) -> None:
    normalized = _merge_team_state(_load_current_team_state_for_merge(), data)
    if supabase_enabled():
        save_json_state(SUPABASE_STATE_KEY, normalized)
        return

    url, token = _redis_rest_credentials()
    payload = json.dumps(normalized, ensure_ascii=False, separators=(",", ":"))
    if url and token:
        _upstash_set(url, token, REDIS_STATE_KEY, payload)
        return

    LOCAL_TEAM_STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    tmp = LOCAL_TEAM_STATE_PATH.with_suffix(".json.tmp")
    tmp.write_text(payload, encoding="utf-8")
    tmp.replace(LOCAL_TEAM_STATE_PATH)


def _load_current_team_state_for_merge() -> dict[str, Any]:
    if supabase_enabled():
        data = load_json_state(SUPABASE_STATE_KEY)
        if isinstance(data, dict):
            return _coerce_team_state(data)
        return _coerce_team_state({})

    url, token = _redis_rest_credentials()
    if url and token:
        raw = _upstash_get(url, token, REDIS_STATE_KEY)
        if raw:
            try:
                data = json.loads(raw)
                if isinstance(data, dict):
                    return _coerce_team_state(data)
            except json.JSONDecodeError:
                pass
        return _coerce_team_state({})

    if LOCAL_TEAM_STATE_PATH.is_file():
        try:
            data = json.loads(LOCAL_TEAM_STATE_PATH.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return _coerce_team_state(data)
        except (OSError, json.JSONDecodeError):
            pass
    return _coerce_team_state({})


def _merge_team_state(existing: dict[str, Any], incoming: dict[str, Any]) -> dict[str, Any]:
    current = _coerce_team_state(existing)
    new = _coerce_team_state(incoming)
    deleted_ids = {
        str(value)
        for value in incoming.get("deletedSessionIds", [])
        if value is not None and str(value).strip()
    }

    sessions_by_id: dict[str, dict[str, Any]] = {}
    for session in current["sessions"]:
        session_id = str(session.get("id") or "").strip()
        if session_id and session_id not in deleted_ids:
            sessions_by_id[session_id] = session

    for session in new["sessions"]:
        session_id = str(session.get("id") or "").strip()
        if not session_id or session_id in deleted_ids:
            continue
        previous = sessions_by_id.get(session_id)
        if previous is None or _session_timestamp(session) >= _session_timestamp(previous):
            sessions_by_id[session_id] = session

    sessions = sorted(
        sessions_by_id.values(),
        key=lambda entry: str(entry.get("createdAt") or entry.get("updatedAt") or ""),
        reverse=True,
    )[:500]

    return {
        "sessions": sessions,
        "leverschemaResults": {**current["leverschemaResults"], **new["leverschemaResults"]},
        "laadschemaData": {**current["laadschemaData"], **new["laadschemaData"]},
        "laadschemaCustomTrucks": {**current["laadschemaCustomTrucks"], **new["laadschemaCustomTrucks"]},
    }


def _session_timestamp(session: dict[str, Any]) -> str:
    return str(session.get("updatedAt") or session.get("createdAt") or "")


def _coerce_team_state(data: dict[str, Any]) -> dict[str, Any]:
    sessions = data.get("sessions")
    lev = data.get("leverschemaResults")
    laadschema = data.get("laadschemaData")
    custom_trucks = data.get("laadschemaCustomTrucks")
    
    if not isinstance(sessions, list):
        sessions = []
    if not isinstance(lev, dict):
        lev = {}
    if not isinstance(laadschema, dict):
        laadschema = {}
    if not isinstance(custom_trucks, dict):
        custom_trucks = {}
    
    if len(sessions) > 500:
        sessions = sessions[:500]
    if len(lev) > 8000:
        lev = dict(list(lev.items())[:8000])
    if len(laadschema) > 8000:
        laadschema = dict(list(laadschema.items())[:8000])
    if len(custom_trucks) > 8000:
        custom_trucks = dict(list(custom_trucks.items())[:8000])
    
    return {
        "sessions": sessions,
        "leverschemaResults": lev,
        "laadschemaData": laadschema,
        "laadschemaCustomTrucks": custom_trucks,
    }


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
    """POST body to /set/{key} — large JSON breaks when passed inside /pipeline array."""
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
