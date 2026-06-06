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
    # `total_item_count` and `size_in_bytes` are own-folder values — items
    # directly in this folder, not in descendants. Use the `tree_*` properties
    # below for the recursive rollup.
    total_item_count: int
    unread_item_count: int
    size_in_bytes: int | None
    children: list["FolderNode"] = field(default_factory=list)

    @property
    def tree_size_in_bytes(self) -> int:
        return (self.size_in_bytes or 0) + sum(c.tree_size_in_bytes for c in self.children)

    @property
    def tree_item_count(self) -> int:
        return self.total_item_count + sum(c.tree_item_count for c in self.children)


def _normalize_prop_id(s: str) -> str:
    """Normalize a Graph extended-property id for comparison.

    Graph echoes ids back with lowercase type+hex and leading zeros stripped
    (e.g. our 'Long 0x0E08' filter comes back as 'Long 0xe08' in the response).
    """
    parts = s.split(maxsplit=1)
    if len(parts) != 2:
        return s.lower()
    typ, val = parts
    if val.lower().startswith("0x"):
        digits = val[2:].lstrip("0").lower() or "0"
        val = "0x" + digits
    return f"{typ.lower()} {val}"


def _ext_value(item: dict, prop_id: str) -> int | None:
    want = _normalize_prop_id(prop_id)
    for prop in item.get("singleValueExtendedProperties", []) or []:
        if _normalize_prop_id(prop.get("id", "")) == want:
            try:
                return int(prop.get("value", 0))
            except (TypeError, ValueError):
                return None
    return None


def _msg_size(msg: dict) -> int | None:
    return _ext_value(msg, PR_MESSAGE_SIZE)


_SELECT = "id,displayName,parentFolderId,totalItemCount,unreadItemCount"
_COLLECTION_PARAMS = {
    "$select": _SELECT,
    "$expand": f"singleValueExtendedProperties($filter=id eq '{PR_MESSAGE_SIZE_EXTENDED}')",
    "$top": 100,
    "includeHiddenFolders": "true",
}
_ITEM_PARAMS = {
    "$select": _SELECT,
    "$expand": f"singleValueExtendedProperties($filter=id eq '{PR_MESSAGE_SIZE_EXTENDED}')",
}


def _build_node(client: GraphClient, item: dict) -> FolderNode:
    return FolderNode(
        id=item["id"],
        display_name=item.get("displayName", ""),
        parent_id=item.get("parentFolderId"),
        total_item_count=item.get("totalItemCount", 0),
        unread_item_count=item.get("unreadItemCount", 0),
        size_in_bytes=_ext_value(item, PR_MESSAGE_SIZE_EXTENDED),
        children=_fetch_children(client, item["id"]),
    )


def _fetch_children(client: GraphClient, parent_id: str) -> list[FolderNode]:
    path = f"{client.mailbox}/mailFolders/{parent_id}/childFolders"
    return [_build_node(client, item) for item in client.paged(path, **_COLLECTION_PARAMS)]


def walk_folders(client: GraphClient) -> list[FolderNode]:
    """Return top-level folders with descendants populated recursively.

    Folder size comes from PR_MESSAGE_SIZE_EXTENDED, expanded inline — one
    extended property per folder, one round-trip per page (not per message).
    """
    return [_build_node(client, item) for item in client.paged(f"{client.mailbox}/mailFolders", **_COLLECTION_PARAMS)]


def walk_subtree(client: GraphClient, root: str) -> FolderNode:
    """Return one folder (by id or well-known name like 'inbox') with descendants.

    Used to scope analyses to a specific subtree without walking the whole mailbox.
    """
    item = client.get(f"{client.mailbox}/mailFolders/{root}", **_ITEM_PARAMS)
    return _build_node(client, item)


def flatten(node: FolderNode) -> list[FolderNode]:
    """Pre-order traversal: root, then descendants."""
    out = [node]
    for c in node.children:
        out.extend(flatten(c))
    return out
