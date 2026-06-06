"""Per-folder share of messages that have attachments."""
from __future__ import annotations

from dataclasses import dataclass

from . import folders as folders_mod
from ..graph import GraphClient


@dataclass
class AttachmentStats:
    total: int
    with_attachments: int

    @property
    def ratio(self) -> float:
        return self.with_attachments / self.total if self.total else 0.0


def attachments_by_folder(client: GraphClient, root: str = "inbox") -> dict[str, AttachmentStats]:
    """Walk `root` + descendants, count messages with hasAttachments=true per folder.

    `root` is a well-known folder name or id. 'inbox' covers Inbox + its
    descendants; 'msgfolderroot' covers the whole visible mailbox.
    """
    subtree = folders_mod.walk_subtree(client, root)
    out: dict[str, AttachmentStats] = {}
    params = {"$select": "id,hasAttachments", "$top": 500}

    for f in folders_mod.flatten(subtree):
        if f.total_item_count == 0:
            continue
        total = 0
        with_att = 0
        for msg in client.paged(f"{client.mailbox}/mailFolders/{f.id}/messages", **params):
            total += 1
            if msg.get("hasAttachments"):
                with_att += 1
        out[f.display_name] = AttachmentStats(total=total, with_attachments=with_att)
    return out
