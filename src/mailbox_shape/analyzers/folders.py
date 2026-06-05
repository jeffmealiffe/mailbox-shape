"""Folder shape: tree, item counts, sizes, attachment share, item-type breakdown."""
from __future__ import annotations

from dataclasses import dataclass, field

from ..graph import GraphClient

# microsoft.graph.message / mailFolder have no schema-declared size property.
# Both surface via the MAPI extended property at ID 0x0E08:
#   - On a message: PR_MESSAGE_SIZE        (type PT_LONG  → "Integer 0x0E08")
#   - On a folder:  PR_MESSAGE_SIZE_EXTENDED (type PT_I8 → "Long 0x0E08")
# The folder variant is 64-bit so totals beyond 2 GB don't overflow.
PR_MESSAGE_SIZE = "Integer 0x0E08"
PR_MESSAGE_SIZE_EXTENDED = "Long 0x0E08"


@dataclass
class FolderNode:
    id: str
    display_name: str
    parent_id: str | None
    total_item_count: int
    unread_item_count: int
    size_in_bytes: int | None
    children: list["FolderNode"] = field(default_factory=list)


def _ext_value(item: dict, prop_id: str) -> int | None:
    for prop in item.get("singleValueExtendedProperties", []) or []:
        if prop.get("id") == prop_id:
            try:
                return int(prop.get("value", 0))
            except (TypeError, ValueError):
                return None
    return None


def _msg_size(msg: dict) -> int | None:
    return _ext_value(msg, PR_MESSAGE_SIZE)


_SELECT = "id,displayName,parentFolderId,totalItemCount,unreadItemCount"


def walk_folders(client: GraphClient) -> list[FolderNode]:
    """Return top-level folders with children populated recursively.

    Folder size comes from PR_MESSAGE_SIZE_EXTENDED, expanded inline — one
    extended property per folder, one round-trip per page (not per message).
    """
    params = {
        "$select": _SELECT,
        "$expand": f"singleValueExtendedProperties($filter=id eq '{PR_MESSAGE_SIZE_EXTENDED}')",
        "$top": 100,
        "includeHiddenFolders": "true",
    }

    def fetch(parent: str | None) -> list[FolderNode]:
        path = "/me/mailFolders" if parent is None else f"/me/mailFolders/{parent}/childFolders"
        nodes: list[FolderNode] = []
        for item in client.paged(path, **params):
            nodes.append(
                FolderNode(
                    id=item["id"],
                    display_name=item.get("displayName", ""),
                    parent_id=item.get("parentFolderId"),
                    total_item_count=item.get("totalItemCount", 0),
                    unread_item_count=item.get("unreadItemCount", 0),
                    size_in_bytes=_ext_value(item, PR_MESSAGE_SIZE_EXTENDED),
                    children=fetch(item["id"]),
                )
            )
        return nodes

    return fetch(None)
