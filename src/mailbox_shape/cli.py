"""mailbox-shape command-line interface."""
from __future__ import annotations

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


def _client() -> GraphClient:
    return GraphClient(get_access_token())


@click.group()
def main() -> None:
    """Analyze the shape of a Microsoft 365 mailbox."""


@main.command()
def folders() -> None:
    """Print folder tree with item counts and sizes."""
    with _client() as c:
        tree = folders_mod.walk_folders(c)

    def render(nodes: list[folders_mod.FolderNode], depth: int = 0) -> None:
        for n in nodes:
            indent = "  " * depth
            mb = n.size_in_bytes / (1024 * 1024)
            console.print(f"{indent}{n.display_name}  [dim]{n.total_item_count} items, {mb:,.1f} MB[/]")
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
