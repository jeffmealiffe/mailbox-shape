"""Message size percentiles, split by sent vs. received."""
from __future__ import annotations

from collections.abc import Iterable
from statistics import quantiles

from . import folders as folders_mod
from ..graph import GraphClient

PERCENTILES = (50, 75, 90, 95, 99)

# Graph 504s on large $top values when combined with the MAPI expand. Keep pages
# small so each request stays under the gateway timeout.
PAGE_SIZE = 100
# Sample target — percentile estimates converge well below 10k; default of 5k
# is a reasonable compromise between accuracy and runtime.
DEFAULT_SAMPLE = 5000


def _percentiles(values: list[int]) -> dict[int, int]:
    if not values:
        return {p: 0 for p in PERCENTILES}
    cuts = quantiles(values, n=100, method="inclusive")
    out: dict[int, int] = {}
    for p in PERCENTILES:
        out[p] = int(cuts[p - 1]) if p - 1 < len(cuts) else int(values[-1])
    return out


def _collect_sizes(
    client: GraphClient,
    folder_id: str,
    ts_field: str,
    limit: int | None,
) -> list[int]:
    """Pull message sizes from a folder, newest first, capped at `limit`.

    Uses an indexed $orderby on the timestamp field so Graph can stream pages
    without scanning the full folder. Pass limit=None to fetch everything.
    """
    sizes: list[int] = []
    params = {
        "$select": "id",
        "$expand": f"singleValueExtendedProperties($filter=id eq '{folders_mod.PR_MESSAGE_SIZE}')",
        "$orderby": f"{ts_field} desc",
        "$top": PAGE_SIZE,
    }
    for msg in client.paged(f"/me/mailFolders/{folder_id}/messages", **params):
        s = folders_mod._msg_size(msg)
        if s is not None:
            sizes.append(s)
        if limit is not None and len(sizes) >= limit:
            break
    return sizes


def size_percentiles(
    client: GraphClient,
    limit: int | None = DEFAULT_SAMPLE,
) -> dict[str, dict[int, int]]:
    """Return {'sent': {p: bytes}, 'received': {p: bytes}}.

    Samples the `limit` most recent messages from Sent Items and Inbox.
    """
    return {
        "sent": _percentiles(_collect_sizes(client, "sentitems", "sentDateTime", limit)),
        "received": _percentiles(_collect_sizes(client, "inbox", "receivedDateTime", limit)),
    }


def summarize(values: Iterable[int]) -> dict[int, int]:
    return _percentiles(list(values))
