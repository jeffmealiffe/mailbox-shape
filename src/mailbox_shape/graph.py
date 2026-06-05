"""Thin wrapper over Microsoft Graph for paged GETs."""
from __future__ import annotations

import time
from collections.abc import Iterator
from typing import Any

import httpx

GRAPH = "https://graph.microsoft.com/v1.0"

# Status codes worth retrying on — gateway and throttling errors. Graph returns
# 429 with a Retry-After header for throttling; 5xx for transient backend hiccups.
_RETRYABLE_STATUSES = {429, 502, 503, 504}


class GraphError(RuntimeError):
    def __init__(self, response: httpx.Response) -> None:
        try:
            body = response.json()
            err = body.get("error", {})
            detail = f"{err.get('code', '?')}: {err.get('message', response.text)}"
        except Exception:
            detail = response.text
        super().__init__(
            f"Graph {response.status_code} on {response.request.method} {response.request.url}\n{detail}"
        )
        self.response = response


def _raise(r: httpx.Response) -> None:
    if r.is_error:
        raise GraphError(r)


class GraphClient:
    def __init__(self, token: str, timeout: float = 60.0, retries: int = 3) -> None:
        self._client = httpx.Client(
            base_url=GRAPH,
            headers={"Authorization": f"Bearer {token}"},
            timeout=timeout,
        )
        self._retries = retries

    def close(self) -> None:
        self._client.close()

    def __enter__(self) -> "GraphClient":
        return self

    def __exit__(self, *_: object) -> None:
        self.close()

    def _request(self, url: str, params: dict[str, Any] | None) -> httpx.Response:
        delay = 1.0
        for attempt in range(self._retries):
            r = self._client.get(url, params=params)
            if r.status_code not in _RETRYABLE_STATUSES or attempt == self._retries - 1:
                return r
            # Honor Retry-After if present (throttling), else exponential backoff.
            retry_after = r.headers.get("Retry-After")
            sleep_for = float(retry_after) if retry_after and retry_after.isdigit() else delay
            time.sleep(sleep_for)
            delay *= 2
        return r  # unreachable, satisfies type checker

    def get(self, path: str, **params: Any) -> dict[str, Any]:
        r = self._request(path, params)
        _raise(r)
        return r.json()

    def paged(self, path: str, **params: Any) -> Iterator[dict[str, Any]]:
        """Yield items across all @odata.nextLink pages."""
        url: str | None = path
        first = True
        while url:
            r = self._request(url, params if first else None)
            _raise(r)
            body = r.json()
            yield from body.get("value", [])
            url = body.get("@odata.nextLink")
            first = False
