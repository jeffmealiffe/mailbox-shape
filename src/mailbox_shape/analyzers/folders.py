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
    # sizeInBytes is only present on work/school mailboxes — None on consumer.
    size_in_bytes: int | None
    children: list["FolderNode"] = field(default_factory=list)


# Base fields, available on both consumer and work/school mailboxes.
_BASE_SELECT = "id,displayName,parentFolderId,totalItemCount,unreadItemCount"


def walk_folders(client: GraphClient) -> list[FolderNode]:
    """Return top-level folders with children populated recursively.

    Tries to include sizeInBytes (work/school accounts only). Falls back to the
    base selection if Graph rejects it — consumer (personal) mailboxes do not
    have sizeInBytes on the mailFolder resource.
    """
    select = _select_for(client)

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
                size_in_bytes=item.get("sizeInBytes"),
            )
            node.children = fetch(node.id)
            nodes.append(node)
        return nodes

    return fetch(None)


def _select_for(client: GraphClient) -> str:
    """Probe whether sizeInBytes is selectable; cache the answer on the client."""
    cached = getattr(client, "_folders_select", None)
    if cached:
        return cached
    try:
        client.get("/me/mailFolders", **{"$select": _BASE_SELECT + ",sizeInBytes", "$top": 1})
        select = _BASE_SELECT + ",sizeInBytes"
    except Exception:
        select = _BASE_SELECT
    client._folders_select = select  # type: ignore[attr-defined]
    return select
