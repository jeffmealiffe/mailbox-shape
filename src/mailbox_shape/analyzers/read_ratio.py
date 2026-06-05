"""Read vs. ignored: share of received messages ever marked read."""
from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass

from ..graph import GraphClient


@dataclass
class ReadStats:
    total: int
    read: int

    @property
    def read_ratio(self) -> float:
        return self.read / self.total if self.total else 0.0


def read_ratio_by_sender_domain(client: GraphClient, folder_id: str = "inbox") -> dict[str, ReadStats]:
    totals: Counter[str] = Counter()
    reads: Counter[str] = Counter()
    for msg in client.paged(
        f"/me/mailFolders/{folder_id}/messages",
        **{"$select": "id,isRead,from", "$top": 999},
    ):
        sender = (msg.get("from") or {}).get("emailAddress", {}).get("address", "")
        domain = sender.split("@", 1)[1].lower() if "@" in sender else "(unknown)"
        totals[domain] += 1
        if msg.get("isRead"):
            reads[domain] += 1
    out: dict[str, ReadStats] = defaultdict(lambda: ReadStats(0, 0))
    for domain, n in totals.items():
        out[domain] = ReadStats(total=n, read=reads[domain])
    return dict(out)
