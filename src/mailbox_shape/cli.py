"""mailbox-shape command-line interface."""
from __future__ import annotations

import base64
import json
from pathlib import Path

import click
from dotenv import load_dotenv
from rich.console import Console
from rich.table import Table

from . import report as report_mod
from .analyzers import attachments as attachments_mod
from .analyzers import folders as folders_mod
from .analyzers import message_types as types_mod
from .analyzers import people as people_mod
from .analyzers import rates as rates_mod
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
    """Build a GraphClient from current Click context — respects --mailbox."""
    ctx = click.get_current_context(silent=True)
    mailbox: str | None = ctx.obj.get("mailbox") if ctx and ctx.obj else None
    return GraphClient(get_access_token(shared=bool(mailbox)), mailbox=mailbox)


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
@click.option(
    "--mailbox",
    "mailbox",
    default=None,
    metavar="UPN_OR_ID",
    help=(
        "Target another user's mailbox (UPN or user id). Defaults to the "
        "authenticated user's own mailbox. Requires Mail.Read.Shared (or "
        "Mail.Read.All) in the app registration and that the authenticated "
        "user has delegated access to the target mailbox."
    ),
)
@click.pass_context
def main(ctx: click.Context, mailbox: str | None) -> None:
    """Analyze the shape of a Microsoft 365 mailbox."""
    ctx.ensure_object(dict)
    ctx.obj["mailbox"] = mailbox


@main.command("raw-folder")
@click.argument("folder", default="inbox")
@click.option("--prop", default="Long 0x0E08", help="Extended property id to expand (default: PR_MESSAGE_SIZE_EXTENDED).")
def raw_folder(folder: str, prop: str) -> None:
    """Print the raw Graph JSON for a single folder with the size extended property expanded."""
    params = {
        "$expand": f"singleValueExtendedProperties($filter=id eq '{prop}')",
    }
    with _client() as c:
        body = c.get(f"{c.mailbox}/mailFolders/{folder}", **params)
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
        body = c.get(f"{c.mailbox}/mailFolders/{folder}/messages", **params)
    console.print_json(data=body)


@main.command()
def whoami() -> None:
    """Print the account / tenant the cached token is bound to."""
    token = get_access_token()
    claims = _decode_token_claims(token)
    tid = claims.get("tid", "?")
    account_kind = "personal Microsoft account (consumer)" if tid == MSA_TENANT_ID else "work/school (Entra ID)"

    ctx = click.get_current_context(silent=True)
    target_mailbox = ctx.obj.get("mailbox") if ctx and ctx.obj else None

    table = Table(title="Cached token identity")
    table.add_column("field")
    table.add_column("value")
    table.add_row("account kind", account_kind)
    table.add_row("tenant id (tid claim)", tid)
    table.add_row("upn claim", claims.get("upn", claims.get("preferred_username", "(none)")))
    table.add_row("name claim", claims.get("name", "(none)"))
    table.add_row("scopes (scp claim)", claims.get("scp", "(none)"))
    table.add_row("target mailbox", target_mailbox or "(authenticated user)")

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
@click.option("--top", "top_n", type=int, default=50, show_default=True, help="Number of domains to show.")
def read_ratio(top_n: int) -> None:
    """Print read-vs-ignored share by sender domain across the Inbox subtree."""
    with _client() as c, console.status("Walking Inbox subtree..."):
        result = read_mod.read_ratio_by_sender_domain(c)
    rows = sorted(result.items(), key=lambda kv: kv[1].total, reverse=True)
    table = Table(title="Read ratio by sender domain (Inbox + subfolders)")
    table.add_column("domain")
    table.add_column("total", justify="right")
    table.add_column("read", justify="right")
    table.add_column("% read", justify="right")
    for domain, stats in rows[:top_n]:
        table.add_row(domain, f"{stats.total:,}", f"{stats.read:,}", f"{stats.read_ratio * 100:.0f}%")
    console.print(table)


@main.command()
@click.option("--top", "top_n", type=int, default=20, show_default=True)
def senders(top_n: int) -> None:
    """Top sender addresses across the Inbox subtree."""
    with _client() as c, console.status("Counting senders..."):
        rows = people_mod.top_senders(c, top_n)
    table = Table(title=f"Top {top_n} sender addresses (Inbox + subfolders)")
    table.add_column("address")
    table.add_column("count", justify="right")
    for addr, n in rows:
        table.add_row(addr, f"{n:,}")
    console.print(table)


@main.command("attachment-ratio")
@click.option("--root", default="inbox", show_default=True,
              help="Well-known folder to scope to. Use 'msgfolderroot' for the whole mailbox.")
def attachment_ratio(root: str) -> None:
    """Per-folder share of messages with attachments."""
    with _client() as c, console.status(f"Walking '{root}' subtree..."):
        result = attachments_mod.attachments_by_folder(c, root)
    rows = sorted(result.items(), key=lambda kv: kv[1].total, reverse=True)
    table = Table(title=f"Attachment ratio by folder ({root} + subfolders)")
    table.add_column("folder")
    table.add_column("total", justify="right")
    table.add_column("w/ attach", justify="right")
    table.add_column("% w/ attach", justify="right")
    for name, s in rows:
        table.add_row(name, f"{s.total:,}", f"{s.with_attachments:,}", f"{s.ratio * 100:.0f}%")
    console.print(table)


@main.command("message-types")
@click.option("--root", default="inbox", show_default=True,
              help="Well-known folder to scope to. Use 'msgfolderroot' for the whole mailbox.")
def message_types(root: str) -> None:
    """Per-folder item-type distribution.

    Folders that are 100% plain mail show '—' in the non-message-types
    column — Graph only annotates @odata.type on subtypes like
    eventMessage, so absent annotation means a plain mail message.
    """
    with _client() as c, console.status(f"Walking '{root}' subtree..."):
        result = types_mod.message_types_by_folder(c, root)
    rows = sorted(result.items(), key=lambda kv: sum(kv[1].values()), reverse=True)
    table = Table(title=f"Message types by folder ({root} + subfolders)")
    table.add_column("folder")
    table.add_column("total", justify="right")
    table.add_column("message", justify="right")
    table.add_column("non-message types")
    for name, counts in rows:
        total = sum(counts.values())
        msg_count = counts.get("message", 0)
        non_msg = [(t, n) for t, n in counts.most_common() if t != "message"]
        non_msg_str = ", ".join(f"{t}: {n:,}" for t, n in non_msg) or "—"
        table.add_row(name, f"{total:,}", f"{msg_count:,}", non_msg_str)
    console.print(table)


@main.command()
@click.option("--top", "top_n", type=int, default=20, show_default=True)
def recipients(top_n: int) -> None:
    """Top recipient addresses across Sent Items (To + Cc)."""
    with _client() as c, console.status("Counting recipients..."):
        rows = people_mod.top_recipients(c, top_n)
    table = Table(title=f"Top {top_n} recipient addresses (Sent Items, To + Cc)")
    table.add_column("address")
    table.add_column("count", justify="right")
    for addr, n in rows:
        table.add_row(addr, f"{n:,}")
    console.print(table)


@main.command()
@click.option("--output", "-o", type=click.Path(dir_okay=False), default="mailbox-report.html", show_default=True)
@click.option("--days", type=int, default=30, show_default=True, help="Window for the rates section.")
@click.option("--work-start", type=int, default=9, show_default=True)
@click.option("--work-end", type=int, default=17, show_default=True)
@click.option("--tz", default="America/Los_Angeles", show_default=True)
@click.option("--quick", is_flag=True, help="Skip the slow Inbox-subtree analyzers (read-ratio, senders, attachments, types, monthly volume).")
def report(output: str, days: int, work_start: int, work_end: int, tz: str, quick: bool) -> None:
    """Generate a self-contained HTML report covering every analyzer.

    Without --quick this runs a single fused walk over the Inbox subtree to
    populate read-ratio + senders + attachments + types + monthly received
    volume in one pass, so it's roughly as slow as a single 'volume' run
    (5-15 minutes on a large mailbox), not the sum of all five.
    """
    out_path = Path(output).resolve()

    def progress(label: str) -> None:
        console.print(f"[dim]→[/] {label}")

    with _client() as c:
        data = report_mod.build_report(
            c, days=days, work_start=work_start, work_end=work_end, tz=tz, quick=quick, progress=progress
        )
    progress("Rendering HTML...")
    report_mod.write_report(data, out_path)
    console.print(f"[bold green]✓[/] Wrote {out_path}")


@main.command()
@click.option("--days", type=int, default=30, show_default=True, help="Window size in days, ending today.")
@click.option("--work-start", type=int, default=9, show_default=True, help="First hour of the working day (0-23).")
@click.option("--work-end", type=int, default=17, show_default=True, help="Hour after the working day ends (1-24).")
@click.option("--tz", default="America/Los_Angeles", show_default=True, help="IANA timezone for weekday/hour classification.")
def rates(days: int, work_start: int, work_end: int, tz: str) -> None:
    """Per-day timeline of sent / received / filed, with working-hour-normalized rates.

    Each day shows absolute counts plus '/wh' columns — messages received or
    sent during configured working hours, divided by hours-per-workday. Weekend
    rows show '—' under the /wh columns since no working hours apply.
    """
    with _client() as c, console.status(f"Counting messages in the last {days} days..."):
        window = rates_mod.compute_daily_rates(
            c, days=days, work_start=work_start, work_end=work_end, tz=tz
        )

    console.print(
        f"[bold]Rates: last {days} days[/]  "
        f"[dim]({window.weekdays} weekdays, working hours {work_start:02d}:00–{work_end:02d}:00 {tz})[/]"
    )

    table = Table()
    table.add_column("date")
    table.add_column("dow")
    table.add_column("sent", justify="right")
    table.add_column("recv", justify="right")
    table.add_column("filed", justify="right")
    table.add_column("sent/wh", justify="right")
    table.add_column("recv/wh", justify="right")
    table.add_column("filed/wh", justify="right")

    hours_per_workday = window.hours_per_workday
    for d in sorted(window.daily.keys(), reverse=True):
        c = window.daily[d]
        is_weekday = d.weekday() < 5
        if is_weekday and hours_per_workday:
            sent_rate = f"{c.sent_working / hours_per_workday:.2f}"
            recv_rate = f"{c.received_working / hours_per_workday:.2f}"
            filed_rate = f"{c.filed_working / hours_per_workday:.2f}"
            dow_style = ""
        else:
            sent_rate = recv_rate = filed_rate = "—"
            dow_style = "[dim]"
        dow_name = d.strftime("%a")
        table.add_row(
            f"{dow_style}{d.isoformat()}",
            f"{dow_style}{dow_name}",
            f"{dow_style}{c.sent:,}",
            f"{dow_style}{c.received:,}",
            f"{dow_style}{c.filed:,}",
            f"{dow_style}{sent_rate}",
            f"{dow_style}{recv_rate}",
            f"{dow_style}{filed_rate}",
        )

    # Summary row across the window.
    agg = window.aggregate()
    total_wh = window.total_working_hours
    if total_wh:
        table.add_section()
        table.add_row(
            "[bold]window total[/]",
            "",
            f"[bold]{agg.sent:,}[/]",
            f"[bold]{agg.received:,}[/]",
            f"[bold]{agg.filed:,}[/]",
            f"[bold]{agg.sent_working / total_wh:.2f}[/]",
            f"[bold]{agg.received_working / total_wh:.2f}[/]",
            f"[bold]{agg.filed_working / total_wh:.2f}[/]",
        )
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
