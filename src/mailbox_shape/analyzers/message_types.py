"""Per-folder distribution of message subtypes (mail vs meeting requests vs ...)."""
from __future__ import annotations

from collections import Counter

from . import folders as folders_mod
from ..graph import GraphClient


def message_types_by_folder(client: GraphClient, root: str = "inbox") -> dict[str, Counter[str]]:
    """Walk `root` + descendants, count items by @odata.type per folder.

    Graph's minimal-metadata convention omits `@odata.type` on plain
    `microsoft.graph.message` items and emits it on subtypes like
    `eventMessage`, `eventMessageRequest`, and `eventMessageResponse`.
    Absent annotation here therefore means a plain mail message.

    `root` is a well-known folder name or id. 'inbox' covers Inbox + its
    descendants; 'msgfolderroot' covers the whole visible mailbox.
    """
    subtree = folders_mod.walk_subtree(client, root)
    out: dict[str, Counter[str]] = {}
    params = {"$select": "id", "$top": 500}

    for f in folders_mod.flatten(subtree):
        if f.total_item_count == 0:
            continue
        counts: Counter[str] = Counter()
        for msg in client.paged(f"/me/mailFolders/{f.id}/messages", **params):
            raw_type = msg.get("@odata.type", "#microsoft.graph.message")
            counts[raw_type.removeprefix("#microsoft.graph.")] += 1
        out[f.display_name] = counts
    return out
