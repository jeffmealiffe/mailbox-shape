"""Volume over time: messages sent, received, and filed into subfolders."""
from __future__ import annotations

from collections import Counter
from datetime import datetime
from typing import Literal

from dateutil import parser as dateparser

from . import folders as folders_mod
from ..graph import GraphClient

Bucket = Literal["day", "week", "month"]


def _bucket_key(ts: datetime, bucket: Bucket) -> str:
    if bucket == "day":
        return ts.date().isoformat()
    if bucket == "week":
        iso = ts.isocalendar()
        return f"{iso.year}-W{iso.week:02d}"
    return f"{ts.year}-{ts.month:02d}"


def _count(client: GraphClient, folder_id: str, ts_field: str, bucket: Bucket) -> Counter[str]:
    counts: Counter[str] = Counter()
    params = {"$select": f"id,{ts_field}", "$top": 500}
    for msg in client.paged(f"/me/mailFolders/{folder_id}/messages", **params):
        raw = msg.get(ts_field)
        if not raw:
            continue
        counts[_bucket_key(dateparser.isoparse(raw), bucket)] += 1
    return counts


def volume_breakdown(client: GraphClient, bucket: Bucket = "month") -> dict[str, Counter[str]]:
    """Return {'sent', 'received', 'filed'} counters keyed by date bucket.

    - **sent**: messages in Sent Items, bucketed by sentDateTime.
    - **received**: every message in the Inbox tree (root + descendants),
      bucketed by receivedDateTime.
    - **filed**: subset of `received` that ended up in an Inbox subfolder
      rather than staying in Inbox root — proxy for messages caught by a
      rule or hand-moved.
    """
    inbox = folders_mod.walk_subtree(client, "inbox")
    flat = folders_mod.flatten(inbox)

    inbox_root: Counter[str] = Counter()
    filed: Counter[str] = Counter()
    for i, f in enumerate(flat):
        if f.total_item_count == 0:
            continue
        counts = _count(client, f.id, "receivedDateTime", bucket)
        target = inbox_root if i == 0 else filed
        target.update(counts)

    sent = _count(client, "sentitems", "sentDateTime", bucket)
    return {
        "sent": sent,
        "received": inbox_root + filed,
        "filed": filed,
    }
