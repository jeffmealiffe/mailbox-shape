"""mailbox-shape command-line interface."""
from __future__ import annotations

import base64
import json

import click
from dotenv import load_dotenv
from rich.console import Console
from rich.table import Table

from .analyzers import folders as folders_mod
from .analyzers import read_ratio as read_mod
from .analyzers import sizes as sizes_mod
from .analyzers import volume as volume_mod
from .auth import get_access_token
from .graph import GraphClient

load_dotenv()
console = Console()

# Microsoft's well-known "consumer / MSA" tenant id. If a token's `tid` claim
# matches this, the account is a personal Microsoft account regardless of which
# email it shows.
MSA_TENANT_ID = "9188040d-6c67-4c5b-b112-36a304b66dad"


def _client() -> GraphClient:
    return GraphClient(get_access_token())


def _decode_token_claims(token: str) -> dict:
    """Decode the JWT payload without verifying — diagnostic only."""
    parts = token.split(".")
    if len(parts) < 2:
        return {}
    pad = "=" * (-len(parts[1]) % 4)
    try:
        return json.loads(base64.urlsafe_b64decode(parts[1] + pad))
    except Exception:
        return {}


@click.group()
def main() -> None:
    """Analyze the shape of a Microsoft 365 mailbox."""


@main.command("raw-folder")
@click.argument("folder", default="inbox")
@click.option("--prop", default="Long 0x0E08", help="Extended property id to expand (default: PR_MESSAGE_SIZE_EXTENDED).")
def raw_folder(folder: str, prop: str) -> None:
    """Print the raw Graph JSON for a single folder with the size extended property expanded."""
    params = {
        "$expand": f"singleValueExtendedProperties($filter=id eq '{prop}')",
    }
    with _client() as c:
        body = c.get(f"/me/mailFolders/{folder}", **params)
    console.print_json(data=body)


@main.command("raw-message")
@click.argument("folder", default="inbox")
@click.option("--no-expand", is_flag=True, help="Omit the singleValueExtendedProperties expand.")
def raw_message(folder: str, no_expand: bool) -> None:
    """Print the raw Graph JSON for the newest message in a folder.

    Use to check whether 'size' is in the default response without $expand —
    if so, $expand is unnecessary and slow folders can skip it.
    """
    params: dict = {"$top": 1, "$orderby": "receivedDateTime desc"}
    if not no_expand:
        params["$expand"] = "singleValueExtendedProperties($filter=id eq 'Integer 0x0E08')"
    with _client() as c:
        body = c.get(f"/me/mailFolders/{folder}/messages", **params)
    console.print_json(data=body)


@main.command()
def whoami() -> None:
    """Print the account / tenant the cached token is bound to."""
    token = get_access_token()
    claims = _decode_token_claims(token)
    tid = claims.get("tid", "?")
    account_kind = "personal Microsoft account (consumer)" if tid == MSA_TENANT_ID else "work/school (Entra ID)"

    table = Table(title="Cached token identity")
    table.add_column("field")
    table.add_column("value")
    table.add_row("account kind", account_kind)
    table.add_row("tenant id (tid claim)", tid)
    table.add_row("upn claim", claims.get("upn", claims.get("preferred_username", "(none)")))
    table.add_row("name claim", claims.get("name", "(none)"))
    table.add_row("scopes (scp claim)", claims.get("scp", "(none)"))

    with GraphClient(token) as c:
        try:
            me = c.get("/me")
            table.add_row("/me userPrincipalName", me.get("userPrincipalName", "(none)"))
            table.add_row("/me mail", me.get("mail", "(none)"))
        except Exception as e:
            table.add_row("/me", f"[red]{e}[/]")
    console.print(table)


@main.command()
@click.option("--own", "show_own", is_flag=True, help="Show per-folder own values instead of recursive subtree rollups.")
def folders(show_own: bool) -> None:
    """Print folder tree with item counts and sizes.

    Counts and sizes are recursive by default — each row reflects that folder
    plus everything beneath it. Pass --own to show only the messages stored
    directly in each folder.
    """
    with _client() as c:
        tree = folders_mod.walk_folders(c)

    scope = "own folder only" if show_own else "subtree rollup"
    console.print(f"[bold]Folder tree[/]  [dim]({scope}; sizes from PR_MESSAGE_SIZE_EXTENDED)[/]")

    def render(nodes: list[folders_mod.FolderNode], depth: int = 0) -> None:
        for n in nodes:
            indent = "  " * depth
            if show_own:
                count = n.total_item_count
                size = n.size_in_bytes
            else:
                count = n.tree_item_count
                size = n.tree_size_in_bytes
            mb = f"{size / (1024 * 1024):,.1f} MB" if size is not None else "—"
            console.print(f"{indent}{n.display_name}  [dim]{count:,} items, {mb}[/]")
            render(n.children, depth + 1)

    render(tree)


@main.command()
@click.option("--limit", type=int, default=sizes_mod.DEFAULT_SAMPLE, show_default=True,
              help="Sample size per direction. Use 0 for no cap (may take a long time).")
def sizes(limit: int) -> None:
    """Print message size percentiles for sent and received."""
    cap = None if limit == 0 else limit
    with _client() as c, console.status(f"Sampling up to {limit or 'all'} messages per direction..."):
        result = sizes_mod.size_percentiles(c, cap)
    table = Table(title=f"Message size (bytes), newest {cap or 'all'} per direction")
    table.add_column("direction")
    for p in sizes_mod.PERCENTILES:
        table.add_column(f"p{p}", justify="right")
    for direction, pcts in result.items():
        table.add_row(direction, *(f"{pcts[p]:,}" for p in sizes_mod.PERCENTILES))
    console.print(table)


@main.command("read-ratio")
@click.option("--folder", default="inbox", help="Well-known folder name or id (default: inbox).")
def read_ratio(folder: str) -> None:
    """Print read-vs-ignored share by sender domain."""
    with _client() as c:
        result = read_mod.read_ratio_by_sender_domain(c, folder)
    rows = sorted(result.items(), key=lambda kv: kv[1].total, reverse=True)
    table = Table(title=f"Read ratio by sender domain — {folder}")
    table.add_column("domain")
    table.add_column("total", justify="right")
    table.add_column("read", justify="right")
    table.add_column("% read", justify="right")
    for domain, stats in rows[:50]:
        table.add_row(domain, str(stats.total), str(stats.read), f"{stats.read_ratio * 100:.0f}%")
    console.print(table)


@main.command()
@click.option("--by", "bucket", type=click.Choice(["day", "week", "month"]), default="month")
def volume(bucket: str) -> None:
    """Print sent / received / filed volume bucketed by day/week/month.

    'filed' is the subset of received messages that ended up in an Inbox
    subfolder rather than staying in Inbox root — a proxy for how much of
    your incoming mail got auto-sorted by rules.
    """
    with _client() as c, console.status("Counting messages across folders (this can take a while)..."):
        result = volume_mod.volume_breakdown(c, bucket)  # type: ignore[arg-type]
    keys = sorted(set(result["sent"]) | set(result["received"]))
    table = Table(title=f"Volume by {bucket}")
    table.add_column(bucket)
    table.add_column("sent", justify="right")
    table.add_column("received", justify="right")
    table.add_column("filed", justify="right")
    for k in keys:
        table.add_row(
            k,
            str(result["sent"].get(k, 0)),
            str(result["received"].get(k, 0)),
            str(result["filed"].get(k, 0)),
        )
    console.print(table)


if __name__ == "__main__":
    main()
