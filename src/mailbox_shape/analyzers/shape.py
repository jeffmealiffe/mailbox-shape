"""Attachment share and item-type distribution, per folder."""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field

from . import folders as folders_mod
from ..graph import GraphClient


@dataclass
class FolderShape:
    total: int = 0
    with_attachments: int = 0
    type_counts: Counter[str] = field(default_factory=Counter)

    @property
    def attachment_pct(self) -> float:
        return (self.with_attachments / self.total * 100) if self.total else 0.0


def shape_by_folder(client: GraphClient, root: str = "inbox") -> dict[str, FolderShape]:
    """Per-folder attachment share and item-type breakdown.

    `root` is a well-known folder name or id. 'inbox' covers Inbox + its
    descendants; 'msgfolderroot' covers the entire visible mailbox.

    Item type comes from Graph's `@odata.type` annotation, which is omitted
    on plain `microsoft.graph.message` items and present on subtypes like
    `eventMessage`, `eventMessageRequest`, `eventMessageResponse`. So absent
    annotation = plain message; this is the standard minimal-metadata
    convention, no extra Accept header needed.
    """
    subtree = folders_mod.walk_subtree(client, root)
    out: dict[str, FolderShape] = {}
    params = {"$select": "id,hasAttachments", "$top": 500}

    for f in folders_mod.flatten(subtree):
        if f.total_item_count == 0:
            continue
        shape = FolderShape()
        for msg in client.paged(f"/me/mailFolders/{f.id}/messages", **params):
            shape.total += 1
            if msg.get("hasAttachments"):
                shape.with_attachments += 1
            raw_type = msg.get("@odata.type", "#microsoft.graph.message")
            short = raw_type.removeprefix("#microsoft.graph.")
            shape.type_counts[short] += 1
        out[f.display_name] = shape
    return out
