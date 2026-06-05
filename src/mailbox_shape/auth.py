"""MSAL device-code auth with on-disk token cache."""
from __future__ import annotations

import os
from pathlib import Path

import msal

SCOPES = ["Mail.Read"]
CACHE_PATH = Path.home() / ".mailbox-shape" / "token_cache.bin"


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if CACHE_PATH.exists():
        cache.deserialize(CACHE_PATH.read_text())
    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
        CACHE_PATH.write_text(cache.serialize())


def get_access_token() -> str:
    client_id = os.environ.get("MAILBOX_SHAPE_CLIENT_ID")
    if not client_id:
        raise RuntimeError("MAILBOX_SHAPE_CLIENT_ID is not set. See README for app registration.")
    tenant = os.environ.get("MAILBOX_SHAPE_TENANT_ID", "common")
    authority = f"https://login.microsoftonline.com/{tenant}"

    cache = _load_cache()
    app = msal.PublicClientApplication(client_id, authority=authority, token_cache=cache)

    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Failed to start device flow: {flow}")
        print(flow["message"], flush=True)
        result = app.acquire_token_by_device_flow(flow)

    _save_cache(cache)

    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description') or result}")
    return result["access_token"]
