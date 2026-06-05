"""Folder shape: tree, item counts, sizes, attachment share, item-type breakdown."""
from __future__ import annotations

from dataclasses import dataclass, field

from ..graph import GraphClient


@dataclass
class FolderNode:
    id: str
    display_name: str
    parent_id: str | None
    total_item_count: int
    unread_item_count: int
    # Graph does not expose folder size as a property. Populated only when the
    # caller explicitly asks for it via compute_sizes — that walks all messages
    # in the folder and sums their `size` field.
    size_in_bytes: int | None
    children: list["FolderNode"] = field(default_factory=list)


_SELECT = "id,displayName,parentFolderId,totalItemCount,unreadItemCount"


def walk_folders(client: GraphClient) -> list[FolderNode]:
    """Return top-level folders with children populated recursively."""

    def fetch(parent: str | None) -> list[FolderNode]:
        path = "/me/mailFolders" if parent is None else f"/me/mailFolders/{parent}/childFolders"
        nodes: list[FolderNode] = []
        for item in client.paged(path, **{"$select": _SELECT, "$top": 100, "includeHiddenFolders": "true"}):
            nodes.append(
                FolderNode(
                    id=item["id"],
                    display_name=item.get("displayName", ""),
                    parent_id=item.get("parentFolderId"),
                    total_item_count=item.get("totalItemCount", 0),
                    unread_item_count=item.get("unreadItemCount", 0),
                    size_in_bytes=None,
                    children=fetch(item["id"]),
                )
            )
        return nodes

    return fetch(None)


# microsoft.graph.message has no schema-declared `size` — $select=size fails.
# Read the MAPI property PR_MESSAGE_SIZE (tag 0x0E08, type Integer) instead.
PR_MESSAGE_SIZE = "Integer 0x0E08"


def _msg_size(msg: dict) -> int | None:
    """Pull the PR_MESSAGE_SIZE value out of an expanded singleValueExtendedProperties."""
    for prop in msg.get("singleValueExtendedProperties", []) or []:
        if prop.get("id") == PR_MESSAGE_SIZE:
            try:
                return int(prop.get("value", 0))
            except (TypeError, ValueError):
                return None
    return None


def compute_size(client: GraphClient, folder_id: str) -> int:
    """Sum the size of every message directly in this folder.

    Does NOT include child folders — callers must recurse.
    """
    total = 0
    params = {
        "$select": "id",
        "$expand": f"singleValueExtendedProperties($filter=id eq '{PR_MESSAGE_SIZE}')",
        "$top": 999,
    }
    for msg in client.paged(f"/me/mailFolders/{folder_id}/messages", **params):
        s = _msg_size(msg)
        if s is not None:
            total += s
    return total


def populate_sizes(client: GraphClient, nodes: list[FolderNode]) -> None:
    """Walk the tree and fill in size_in_bytes for every node, in place."""
    for n in nodes:
        if n.total_item_count > 0:
            n.size_in_bytes = compute_size(client, n.id)
        else:
            n.size_in_bytes = 0
        populate_sizes(client, n.children)
