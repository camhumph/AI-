"""Local email credential storage for the CMS quoting webapp.

Credentials live in a JSON file under CMS_DATA_DIR (never committed to git).
The webapp Settings page is the only place you enter your Gmail app password —
not a separate gmail_app_password.txt file.

Module6121.bas and cms_gmail_search.py read the same JSON file so SMTP/IMAP
stay in sync without duplicating secrets.
"""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

from . import config

CREDENTIALS_PATH = Path(
    os.environ.get(
        "CMS_EMAIL_CREDENTIALS",
        config.DATA_DIR / "email_credentials.json",
    )
)

# Fields stored in the JSON file (passwords are written but never returned by API).
_FIELDS = (
    "imap_host",
    "imap_port",
    "imap_user",
    "imap_password",
    "imap_folder",
    "imap_ssl",
    "smtp_host",
    "smtp_port",
    "smtp_user",
    "smtp_password",
    "smtp_from",
    "gmail_address",
)


def _defaults() -> dict[str, Any]:
    return {
        "imap_host": "imap.gmail.com",
        "imap_port": 993,
        "imap_user": "",
        "imap_password": "",
        "imap_folder": "INBOX",
        "imap_ssl": True,
        "smtp_host": "smtp.gmail.com",
        "smtp_port": 587,
        "smtp_user": "",
        "smtp_password": "",
        "smtp_from": "",
        "gmail_address": "cms1engineering@gmail.com",
    }


def load() -> dict[str, Any]:
    data = _defaults()
    if CREDENTIALS_PATH.exists():
        try:
            stored = json.loads(CREDENTIALS_PATH.read_text(encoding="utf-8"))
            if isinstance(stored, dict):
                data.update({k: stored[k] for k in _FIELDS if k in stored})
        except Exception:
            pass
    # Environment variables override file (useful for cloud agents / CI).
    env_map = {
        "imap_host": "CMS_IMAP_HOST",
        "imap_port": "CMS_IMAP_PORT",
        "imap_user": "CMS_IMAP_USER",
        "imap_password": "CMS_IMAP_PASSWORD",
        "imap_folder": "CMS_IMAP_FOLDER",
        "imap_ssl": "CMS_IMAP_SSL",
        "smtp_host": "CMS_SMTP_HOST",
        "smtp_port": "CMS_SMTP_PORT",
        "smtp_user": "CMS_SMTP_USER",
        "smtp_password": "CMS_SMTP_PASSWORD",
        "smtp_from": "CMS_SMTP_FROM",
        "gmail_address": "CMS_GMAIL_ADDRESS",
    }
    for key, env_name in env_map.items():
        val = os.environ.get(env_name)
        if val is None or val == "":
            continue
        if key in ("imap_port", "smtp_port"):
            try:
                data[key] = int(val)
            except ValueError:
                pass
        elif key == "imap_ssl":
            data[key] = val.lower() not in ("false", "0", "no")
        else:
            data[key] = val
    if not data.get("smtp_from"):
        data["smtp_from"] = data.get("smtp_user") or data.get("gmail_address", "")
    # Treat gmail_address as imap_user when user only filled the address field.
    if data.get("gmail_address") and not data.get("imap_user"):
        data["imap_user"] = data["gmail_address"]
    return data


def save(updates: dict[str, Any]) -> dict[str, Any]:
    current = load()
    for key in _FIELDS:
        if key not in updates:
            continue
        val = updates[key]
        if key.endswith("_password") and (val is None or val == ""):
            continue  # blank password = keep existing
        current[key] = val

    # Auto-sync Gmail fields so one app password configures both IMAP and SMTP.
    if current.get("gmail_address") and not current.get("imap_user"):
        current["imap_user"] = current["gmail_address"]
    if current.get("imap_user") and not current.get("smtp_user"):
        current["smtp_user"] = current["imap_user"]
    if current.get("imap_password") and not current.get("smtp_password"):
        current["smtp_password"] = current["imap_password"]
    if not current.get("smtp_from"):
        current["smtp_from"] = current.get("smtp_user") or current.get("gmail_address", "")

    CREDENTIALS_PATH.parent.mkdir(parents=True, exist_ok=True)
    CREDENTIALS_PATH.write_text(json.dumps(current, indent=2), encoding="utf-8")
    return public_view(current)


def public_view(data: dict[str, Any] | None = None) -> dict[str, Any]:
    """Safe subset for API responses — never includes passwords."""
    src = data or load()
    return {
        "imap_host": src.get("imap_host") or "",
        "imap_port": int(src.get("imap_port") or 993),
        "imap_user": src.get("imap_user") or "",
        "imap_password_set": bool(src.get("imap_password")),
        "imap_folder": src.get("imap_folder") or "INBOX",
        "imap_ssl": bool(src.get("imap_ssl", True)),
        "smtp_host": src.get("smtp_host") or "",
        "smtp_port": int(src.get("smtp_port") or 587),
        "smtp_user": src.get("smtp_user") or "",
        "smtp_password_set": bool(src.get("smtp_password")),
        "smtp_from": src.get("smtp_from") or "",
        "gmail_address": src.get("gmail_address") or "",
        "configured": bool(src.get("imap_host") and src.get("imap_user") and src.get("imap_password")),
        "smtp_configured": bool(
            src.get("smtp_host") and src.get("smtp_user") and src.get("smtp_password")
        ),
        "credentials_path": str(CREDENTIALS_PATH),
    }
