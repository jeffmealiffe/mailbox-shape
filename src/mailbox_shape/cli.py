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
@click.option("--with-sizes", is_flag=True, help="Sum per-message sizes to compute folder size (slow on large mailboxes).")
def folders(with_sizes: bool) -> None:
    """Print folder tree with item counts (and optionally sizes)."""
    with _client() as c:
        tree = folders_mod.walk_folders(c)
        if with_sizes:
            with console.status("Computing folder sizes..."):
                folders_mod.populate_sizes(c, tree)

    def render(nodes: list[folders_mod.FolderNode], depth: int = 0) -> None:
        for n in nodes:
            indent = "  " * depth
            if n.size_in_bytes is None:
                tail = f"{n.total_item_count} items"
            else:
                tail = f"{n.total_item_count} items, {n.size_in_bytes / (1024 * 1024):,.1f} MB"
            console.print(f"{indent}{n.display_name}  [dim]{tail}[/]")
            render(n.children, depth + 1)

    render(tree)


@main.command()
def sizes() -> None:
    """Print message size percentiles for sent and received."""
    with _client() as c:
        result = sizes_mod.size_percentiles(c)
    table = Table(title="Message size (bytes)")
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
    """Print sent/received volume bucketed by day/week/month."""
    with _client() as c:
        sent = volume_mod.sent_volume(c, bucket)  # type: ignore[arg-type]
        recv = volume_mod.received_volume(c, bucket)  # type: ignore[arg-type]
    keys = sorted(set(sent) | set(recv))
    table = Table(title=f"Volume by {bucket}")
    table.add_column(bucket)
    table.add_column("sent", justify="right")
    table.add_column("received", justify="right")
    for k in keys:
        table.add_row(k, str(sent.get(k, 0)), str(recv.get(k, 0)))
    console.print(table)


if __name__ == "__main__":
    main()
