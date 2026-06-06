"""Top senders (Inbox subtree) and top recipients (Sent Items)."""
from __future__ import annotations

from collections import Counter

from . import folders as folders_mod
from ..graph import GraphClient


def top_senders(client: GraphClient, limit: int = 20) -> list[tuple[str, int]]:
    """Most frequent sender addresses across the Inbox subtree."""
    inbox = folders_mod.walk_subtree(client, "inbox")
    counts: Counter[str] = Counter()
    params = {"$select": "id,from", "$top": 500}
    for f in folders_mod.flatten(inbox):
        if f.total_item_count == 0:
            continue
        for msg in client.paged(f"{client.mailbox}/mailFolders/{f.id}/messages", **params):
            addr = (msg.get("from") or {}).get("emailAddress", {}).get("address", "")
            if addr:
                counts[addr.lower()] += 1
    return counts.most_common(limit)


def top_recipients(client: GraphClient, limit: int = 20) -> list[tuple[str, int]]:
    """Most frequent recipient addresses across Sent Items (To + Cc).

    Each recipient on a multi-recipient send counts once. Bcc isn't returned
    by Graph on the message resource after send, so it's not included.
    """
    counts: Counter[str] = Counter()
    params = {"$select": "id,toRecipients,ccRecipients", "$top": 500}
    for msg in client.paged(f"{client.mailbox}/mailFolders/sentitems/messages", **params):
        for r in (msg.get("toRecipients") or []) + (msg.get("ccRecipients") or []):
            addr = (r or {}).get("emailAddress", {}).get("address", "")
            if addr:
                counts[addr.lower()] += 1
    return counts.most_common(limit)
