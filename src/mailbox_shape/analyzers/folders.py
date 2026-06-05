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


def compute_size(client: GraphClient, folder_id: str) -> int:
    """Sum the `size` field of every message directly in this folder.

    Does NOT include child folders — callers must recurse.
    """
    total = 0
    for msg in client.paged(
        f"/me/mailFolders/{folder_id}/messages",
        **{"$select": "id,size", "$top": 999},
    ):
        size = msg.get("size")
        if isinstance(size, int):
            total += size
    return total


def populate_sizes(client: GraphClient, nodes: list[FolderNode]) -> None:
    """Walk the tree and fill in size_in_bytes for every node, in place."""
    for n in nodes:
        if n.total_item_count > 0:
            n.size_in_bytes = compute_size(client, n.id)
        else:
            n.size_in_bytes = 0
        populate_sizes(client, n.children)
