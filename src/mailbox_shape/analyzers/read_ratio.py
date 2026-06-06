"""Read vs. ignored: share of received messages ever marked read."""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass

from . import folders as folders_mod
from ..graph import GraphClient


@dataclass
class ReadStats:
    total: int
    read: int

    @property
    def read_ratio(self) -> float:
        return self.read / self.total if self.total else 0.0


def read_ratio_by_sender_domain(client: GraphClient) -> dict[str, ReadStats]:
    """Walk Inbox + descendants, group by sender domain.

    Receiving a 10,000-message newsletter you never opened is a strong
    signal — the read ratio for that domain ends up near 0%. Filtering on
    that signal is the typical use case.
    """
    inbox = folders_mod.walk_subtree(client, "inbox")
    totals: Counter[str] = Counter()
    reads: Counter[str] = Counter()
    params = {"$select": "id,isRead,from", "$top": 500}

    for f in folders_mod.flatten(inbox):
        if f.total_item_count == 0:
            continue
        for msg in client.paged(f"{client.mailbox}/mailFolders/{f.id}/messages", **params):
            sender = (msg.get("from") or {}).get("emailAddress", {}).get("address", "")
            domain = sender.split("@", 1)[1].lower() if "@" in sender else "(unknown)"
            totals[domain] += 1
            if msg.get("isRead"):
                reads[domain] += 1

    return {d: ReadStats(total=n, read=reads[d]) for d, n in totals.items()}
