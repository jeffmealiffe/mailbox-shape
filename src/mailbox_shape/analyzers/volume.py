"""Volume over time: messages sent, received, and moved into folders."""
from __future__ import annotations

from collections import Counter
from datetime import date, datetime
from typing import Literal

from dateutil import parser as dateparser

from ..graph import GraphClient

Bucket = Literal["day", "week", "month"]


def _bucket_key(ts: datetime, bucket: Bucket) -> str:
    if bucket == "day":
        return ts.date().isoformat()
    if bucket == "week":
        iso = ts.isocalendar()
        return f"{iso.year}-W{iso.week:02d}"
    return f"{ts.year}-{ts.month:02d}"


def _count(client: GraphClient, folder: str, ts_field: str, bucket: Bucket) -> Counter[str]:
    counts: Counter[str] = Counter()
    for msg in client.paged(
        f"/me/mailFolders/{folder}/messages",
        **{"$select": f"id,{ts_field}", "$top": 999, "$orderby": f"{ts_field} desc"},
    ):
        raw = msg.get(ts_field)
        if not raw:
            continue
        ts = dateparser.isoparse(raw)
        counts[_bucket_key(ts, bucket)] += 1
    return counts


def sent_volume(client: GraphClient, bucket: Bucket = "month") -> Counter[str]:
    return _count(client, "sentitems", "sentDateTime", bucket)


def received_volume(client: GraphClient, bucket: Bucket = "month") -> Counter[str]:
    return _count(client, "inbox", "receivedDateTime", bucket)


def today() -> date:
    return datetime.utcnow().date()
