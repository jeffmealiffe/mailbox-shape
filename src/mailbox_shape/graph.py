"""Thin wrapper over Microsoft Graph for paged GETs."""
from __future__ import annotations

from collections.abc import Iterator
from typing import Any

import httpx

GRAPH = "https://graph.microsoft.com/v1.0"


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
    def __init__(self, token: str, timeout: float = 60.0) -> None:
        self._client = httpx.Client(
            base_url=GRAPH,
            headers={"Authorization": f"Bearer {token}"},
            timeout=timeout,
        )

    def close(self) -> None:
        self._client.close()

    def __enter__(self) -> "GraphClient":
        return self

    def __exit__(self, *_: object) -> None:
        self.close()

    def get(self, path: str, **params: Any) -> dict[str, Any]:
        r = self._client.get(path, params=params)
        _raise(r)
        return r.json()

    def paged(self, path: str, **params: Any) -> Iterator[dict[str, Any]]:
        """Yield items across all @odata.nextLink pages."""
        url: str | None = path
        first = True
        while url:
            r = self._client.get(url, params=params if first else None)
            _raise(r)
            body = r.json()
            yield from body.get("value", [])
            url = body.get("@odata.nextLink")
            first = False
