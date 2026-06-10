from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import time
from http import cookies
from typing import Any


SESSION_COOKIE = "greenops_session"
SESSION_TTL_SECONDS = 60 * 60 * 12
PBKDF2_ITERATIONS = 120000

USERS = [
    {
        "email": "beatriz@greenops.app",
        "name": "Beatriz",
        "role": "admin",
        "salt": "1583966f2e5991be1cd62a828b740fd9",
        "password_hash": "a3fbd7558f1c2a5f4cbceedea4d6d325dfbc1795c4a382a4cf97c57028a3caaf",
    },
    {
        "email": "planner@greenops.app",
        "name": "Planner",
        "role": "planner",
        "salt": "f3dfc5116561c5a7b653540e443f4377",
        "password_hash": "e735a4062a7a5da6ee01c0673b3bfddeda53e36c4f2fb885bdbcc4852550f7b8",
    },
]


def authenticate_user(email: str, password: str) -> dict[str, str] | None:
    normalized_email = email.strip().lower()
    for user in _configured_users():
        if user["email"] != normalized_email:
            continue
        if _verify_user_password(user, password):
            return {"email": user["email"], "name": user["name"], "role": user["role"]}
    return None


def current_user_from_headers(headers: dict[str, str]) -> dict[str, str] | None:
    morsel = _read_cookie(headers).get(SESSION_COOKIE)
    if morsel is None:
        return None
    return _decode_session_token(morsel.value)


def set_session_cookie(handler, user: dict[str, str]) -> None:
    payload = {
        "email": user["email"],
        "name": user["name"],
        "role": user["role"],
        "exp": int(time.time()) + SESSION_TTL_SECONDS,
    }
    token = _encode_session_token(payload)
    cookie = cookies.SimpleCookie()
    cookie[SESSION_COOKIE] = token
    cookie[SESSION_COOKIE]["path"] = "/"
    cookie[SESSION_COOKIE]["httponly"] = True
    cookie[SESSION_COOKIE]["samesite"] = "Lax"
    cookie[SESSION_COOKIE]["max-age"] = str(SESSION_TTL_SECONDS)
    if _should_use_secure_cookie(handler.headers):
        cookie[SESSION_COOKIE]["secure"] = True
    handler.send_header("Set-Cookie", cookie.output(header="").strip())


def clear_session_cookie(handler) -> None:
    cookie = cookies.SimpleCookie()
    cookie[SESSION_COOKIE] = ""
    cookie[SESSION_COOKIE]["path"] = "/"
    cookie[SESSION_COOKIE]["httponly"] = True
    cookie[SESSION_COOKIE]["samesite"] = "Lax"
    cookie[SESSION_COOKIE]["max-age"] = "0"
    cookie[SESSION_COOKIE]["expires"] = "Thu, 01 Jan 1970 00:00:00 GMT"
    if _should_use_secure_cookie(handler.headers):
        cookie[SESSION_COOKIE]["secure"] = True
    handler.send_header("Set-Cookie", cookie.output(header="").strip())


def unauthorized_response(handler, message: str = "Authentication required.") -> None:
    data = json.dumps({"error": message}).encode("utf-8")
    handler.send_response(401)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(data)))
    handler.end_headers()
    handler.wfile.write(data)


def _verify_password(password: str, salt_hex: str, expected_hash_hex: str) -> bool:
    salt = bytes.fromhex(salt_hex)
    derived = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ITERATIONS)
    return hmac.compare_digest(derived.hex(), expected_hash_hex)


def _verify_user_password(user: dict[str, str], password: str) -> bool:
    if user.get("salt") and user.get("password_hash"):
        return _verify_password(password, user["salt"], user["password_hash"])
    plain_password = user.get("password")
    if plain_password is not None:
        return hmac.compare_digest(password, plain_password)
    return False


def _configured_users() -> list[dict[str, str]]:
    raw = os.environ.get("ORDERFLOW_USERS_JSON", "").strip()
    if not raw:
        return USERS
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        return USERS
    if not isinstance(parsed, list):
        return USERS

    users: list[dict[str, str]] = []
    for entry in parsed:
        if not isinstance(entry, dict):
            continue
        email = str(entry.get("email") or "").strip().lower()
        name = str(entry.get("name") or email.split("@")[0] or "User").strip()
        role = str(entry.get("role") or "planner").strip()
        password = str(entry.get("password") or "")
        salt = str(entry.get("salt") or "")
        password_hash = str(entry.get("password_hash") or "")
        if not email or (not password and not (salt and password_hash)):
            continue
        user = {"email": email, "name": name, "role": role}
        if salt and password_hash:
            user.update({"salt": salt, "password_hash": password_hash})
        else:
            user["password"] = password
        users.append(user)
    return users or USERS


def _encode_session_token(payload: dict[str, Any]) -> str:
    payload_bytes = json.dumps(payload, separators=(",", ":"), sort_keys=True).encode("utf-8")
    payload_token = base64.urlsafe_b64encode(payload_bytes).rstrip(b"=").decode("ascii")
    signature = hmac.new(_secret_key(), payload_token.encode("ascii"), hashlib.sha256).digest()
    signature_token = base64.urlsafe_b64encode(signature).rstrip(b"=").decode("ascii")
    return f"{payload_token}.{signature_token}"


def _decode_session_token(token: str) -> dict[str, str] | None:
    try:
        payload_token, signature_token = token.split(".", 1)
    except ValueError:
        return None

    expected_signature = hmac.new(_secret_key(), payload_token.encode("ascii"), hashlib.sha256).digest()
    actual_signature = _urlsafe_b64decode(signature_token)
    if actual_signature is None or not hmac.compare_digest(expected_signature, actual_signature):
        return None

    payload_bytes = _urlsafe_b64decode(payload_token)
    if payload_bytes is None:
        return None

    try:
        payload = json.loads(payload_bytes.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError):
        return None

    if int(payload.get("exp", 0)) < int(time.time()):
        return None

    return {
        "email": str(payload.get("email", "")),
        "name": str(payload.get("name", "")),
        "role": str(payload.get("role", "")),
    }


def _urlsafe_b64decode(value: str) -> bytes | None:
    padding = "=" * (-len(value) % 4)
    try:
        return base64.urlsafe_b64decode(value + padding)
    except (ValueError, TypeError):
        return None


def _read_cookie(headers: dict[str, str]) -> cookies.SimpleCookie:
    jar = cookies.SimpleCookie()
    cookie_header = headers.get("Cookie") or headers.get("cookie") or ""
    if cookie_header:
        jar.load(cookie_header)
    return jar


def _secret_key() -> bytes:
    secret = os.environ.get("GREENOPS_AUTH_SECRET", "greenops-local-secret-change-me")
    return secret.encode("utf-8")


def _should_use_secure_cookie(headers: dict[str, str]) -> bool:
    host = (headers.get("Host") or headers.get("host") or "").lower()
    forwarded_proto = (headers.get("X-Forwarded-Proto") or headers.get("x-forwarded-proto") or "").lower()
    if forwarded_proto == "https":
        return True
    return host not in {
        "127.0.0.1:8000",
        "localhost:8000",
        "127.0.0.1:8030",
        "localhost:8030",
        "127.0.0.1",
        "localhost",
    }
