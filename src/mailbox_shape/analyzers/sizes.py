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
    ts_field: str,  # kept for API compat; no longer used for $orderby
    limit: int | None,
) -> list[int]:
    """Pull message sizes from a folder, capped at `limit`.

    No $orderby — on some folders the sort interacts badly with the MAPI
    expand and causes gateway timeouts. Order doesn't matter for percentile
    estimation as long as the sample is large enough to converge.
    """
    del ts_field  # accepted for symmetry with other analyzers; unused here
    sizes: list[int] = []
    params = {
        "$select": "id",
        "$expand": f"singleValueExtendedProperties($filter=id eq '{folders_mod.PR_MESSAGE_SIZE}')",
        "$top": PAGE_SIZE,
    }
    for msg in client.paged(f"/me/mailFolders/{folder_id}/messages", **params):
        s = folders_mod._msg_size(msg)
        if s is not None:
            sizes.append(s)
        if limit is not None and len(sizes) >= limit:
            break
    return sizes


def _allocate_per_folder(
    folders: list[folders_mod.FolderNode],
    limit: int | None,
) -> list[tuple[folders_mod.FolderNode, int | None]]:
    """Split a sample budget across folders proportionally to item count.

    When `limit` is None, each folder gets None (no cap). Folders with zero
    items are dropped. Allocations are clamped to each folder's own item count
    so we never ask for more than it holds.
    """
    active = [f for f in folders if f.total_item_count > 0]
    if limit is None:
        return [(f, None) for f in active]
    total = sum(f.total_item_count for f in active)
    if total == 0:
        return []
    out: list[tuple[folders_mod.FolderNode, int | None]] = []
    for f in active:
        share = f.total_item_count / total
        cap = max(1, min(f.total_item_count, int(round(share * limit))))
        out.append((f, cap))
    return out


def _collect_received_sizes(client: GraphClient, limit: int | None) -> list[int]:
    """Sample message sizes from the entire Inbox subtree.

    Walks Inbox + descendants, allocates the sample budget proportionally to
    each folder's item count, and pulls newest-first from each.
    """
    inbox = folders_mod.walk_subtree(client, "inbox")
    sizes: list[int] = []
    for folder, cap in _allocate_per_folder(folders_mod.flatten(inbox), limit):
        sizes.extend(_collect_sizes(client, folder.id, "receivedDateTime", cap))
    return sizes


def size_percentiles(
    client: GraphClient,
    limit: int | None = DEFAULT_SAMPLE,
) -> dict[str, dict[int, int]]:
    """Return {'sent': {p: bytes}, 'received': {p: bytes}}.

    `sent` samples Sent Items. `received` samples Inbox + all subfolders,
    proportionally to each folder's item count, so heavy newsletter/ads
    subfolders don't get overweighted just because Inbox root is fetched first.
    """
    return {
        "sent": _percentiles(_collect_sizes(client, "sentitems", "sentDateTime", limit)),
        "received": _percentiles(_collect_received_sizes(client, limit)),
    }


def summarize(values: Iterable[int]) -> dict[int, int]:
    return _percentiles(list(values))
