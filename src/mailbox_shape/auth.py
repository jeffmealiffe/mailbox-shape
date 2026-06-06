"""MSAL device-code auth with on-disk token cache."""
from __future__ import annotations

import os
from pathlib import Path

import msal

BASE_SCOPES = ["Mail.Read"]
SHARED_SCOPES = ["Mail.Read", "Mail.Read.Shared"]
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


def get_access_token(shared: bool = False) -> str:
    """Acquire an access token via MSAL.

    When `shared` is True, requests Mail.Read.Shared in addition to Mail.Read,
    which is required to read mailboxes the authenticated user has delegated
    access to (shared mailboxes, "open another user's folder", etc.). The
    delegated Mail.Read.Shared permission must be granted in the Azure AD app
    registration for this to succeed.
    """
    client_id = os.environ.get("MAILBOX_SHAPE_CLIENT_ID")
    if not client_id:
        raise RuntimeError("MAILBOX_SHAPE_CLIENT_ID is not set. See README for app registration.")
    tenant = os.environ.get("MAILBOX_SHAPE_TENANT_ID", "common")
    authority = f"https://login.microsoftonline.com/{tenant}"
    scopes = SHARED_SCOPES if shared else BASE_SCOPES

    cache = _load_cache()
    app = msal.PublicClientApplication(client_id, authority=authority, token_cache=cache)

    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            raise RuntimeError(f"Failed to start device flow: {flow}")
        print(flow["message"], flush=True)
        result = app.acquire_token_by_device_flow(flow)

    _save_cache(cache)

    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description') or result}")
    return result["access_token"]
