"""Message size percentiles, split by sent vs. received."""
from __future__ import annotations

from collections.abc import Iterable
from statistics import quantiles

from ..graph import GraphClient

PERCENTILES = (50, 75, 90, 95, 99)


def _percentiles(values: list[int]) -> dict[int, int]:
    if not values:
        return {p: 0 for p in PERCENTILES}
    # quantiles with n=100 gives 99 cut points covering p1..p99.
    cuts = quantiles(values, n=100, method="inclusive")
    out: dict[int, int] = {}
    for p in PERCENTILES:
        # cuts[p-1] is the boundary that ~p% of data falls below.
        out[p] = int(cuts[p - 1]) if p - 1 < len(cuts) else int(values[-1])
    return out


def _collect_sizes(client: GraphClient, folder_id: str) -> list[int]:
    sizes: list[int] = []
    for msg in client.paged(
        f"/me/mailFolders/{folder_id}/messages",
        **{"$select": "id,size", "$top": 999},
    ):
        size = msg.get("size")
        if isinstance(size, int):
            sizes.append(size)
    return sizes


def size_percentiles(client: GraphClient) -> dict[str, dict[int, int]]:
    """Return {'sent': {p: bytes}, 'received': {p: bytes}}.

    "sent" pulls from the Sent Items well-known folder; "received" pulls from
    Inbox plus its descendants would be more accurate — for now this samples
    Inbox only as a starting point.
    """
    return {
        "sent": _percentiles(_collect_sizes(client, "sentitems")),
        "received": _percentiles(_collect_sizes(client, "inbox")),
    }


def summarize(values: Iterable[int]) -> dict[int, int]:
    return _percentiles(list(values))
