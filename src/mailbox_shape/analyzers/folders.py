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
    size_in_bytes: int
    children: list["FolderNode"] = field(default_factory=list)


def walk_folders(client: GraphClient) -> list[FolderNode]:
    """Return top-level folders with children populated recursively.

    Uses /me/mailFolders with includeHiddenFolders=true and reads sizeInBytes
    from the Graph mailFolder resource (available in v1.0 since 2022).
    """
    select = "id,displayName,parentFolderId,totalItemCount,unreadItemCount,sizeInBytes"

    def fetch(parent: str | None) -> list[FolderNode]:
        path = "/me/mailFolders" if parent is None else f"/me/mailFolders/{parent}/childFolders"
        nodes: list[FolderNode] = []
        for item in client.paged(path, **{"$select": select, "$top": 100, "includeHiddenFolders": "true"}):
            node = FolderNode(
                id=item["id"],
                display_name=item.get("displayName", ""),
                parent_id=item.get("parentFolderId"),
                total_item_count=item.get("totalItemCount", 0),
                unread_item_count=item.get("unreadItemCount", 0),
                size_in_bytes=item.get("sizeInBytes", 0),
            )
            node.children = fetch(node.id)
            nodes.append(node)
        return nodes

    return fetch(None)
