"""HTML report orchestrator — runs analyzers and renders a self-contained file."""
from __future__ import annotations

import io
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from dateutil import parser as dateparser  # noqa: E402
from jinja2 import Environment, PackageLoader, select_autoescape  # noqa: E402

from .analyzers import attachments as attachments_mod  # noqa: E402
from .analyzers import folders as folders_mod  # noqa: E402
from .analyzers import message_types as types_mod  # noqa: E402
from .analyzers import people as people_mod  # noqa: E402
from .analyzers import rates as rates_mod  # noqa: E402
from .analyzers import read_ratio as read_mod  # noqa: E402
from .analyzers import sizes as sizes_mod  # noqa: E402
from .analyzers import volume as volume_mod  # noqa: E402
from .graph import GraphClient  # noqa: E402

CHART_FG = "#1f3a5f"
CHART_ACCENT = "#e07a37"
CHART_MUTED = "#7da46f"
CHART_BG = "#fafbfc"


# -----------------------------------------------------------------------------
# Fused walk — runs read_ratio + senders + attachments + types + received-volume
# in a single pass over the Inbox subtree.
# -----------------------------------------------------------------------------

@dataclass
class _FusedAccum:
    read_totals: Counter[str] = field(default_factory=Counter)
    read_reads: Counter[str] = field(default_factory=Counter)
    senders: Counter[str] = field(default_factory=Counter)
    attachments: dict[str, attachments_mod.AttachmentStats] = field(default_factory=dict)
    types: dict[str, Counter[str]] = field(default_factory=dict)
    received: Counter[str] = field(default_factory=Counter)  # month bucket -> count
    filed: Counter[str] = field(default_factory=Counter)


def _fused_inbox_walk(client: GraphClient) -> _FusedAccum:
    accum = _FusedAccum()
    inbox = folders_mod.walk_subtree(client, "inbox")
    params = {"$select": "id,isRead,from,hasAttachments,receivedDateTime", "$top": 500}

    for i, f in enumerate(folders_mod.flatten(inbox)):
        if f.total_item_count == 0:
            continue
        att = attachments_mod.AttachmentStats(total=0, with_attachments=0)
        type_counts: Counter[str] = Counter()
        for msg in client.paged(f"/me/mailFolders/{f.id}/messages", **params):
            # read ratio + senders
            sender = (msg.get("from") or {}).get("emailAddress", {}).get("address", "")
            if sender:
                accum.senders[sender.lower()] += 1
            domain = sender.split("@", 1)[1].lower() if "@" in sender else "(unknown)"
            accum.read_totals[domain] += 1
            if msg.get("isRead"):
                accum.read_reads[domain] += 1
            # attachments
            att.total += 1
            if msg.get("hasAttachments"):
                att.with_attachments += 1
            # message types
            raw_type = msg.get("@odata.type", "#microsoft.graph.message")
            type_counts[raw_type.removeprefix("#microsoft.graph.")] += 1
            # volume by month
            raw = msg.get("receivedDateTime")
            if raw:
                key = volume_mod._bucket_key(dateparser.isoparse(raw), "month")
                accum.received[key] += 1
                if i != 0:
                    accum.filed[key] += 1
        accum.attachments[f.display_name] = att
        accum.types[f.display_name] = type_counts

    return accum


def _sent_volume_by_month(client: GraphClient) -> Counter[str]:
    counts: Counter[str] = Counter()
    params = {"$select": "id,sentDateTime", "$top": 500}
    for msg in client.paged("/me/mailFolders/sentitems/messages", **params):
        raw = msg.get("sentDateTime")
        if raw:
            counts[volume_mod._bucket_key(dateparser.isoparse(raw), "month")] += 1
    return counts


# -----------------------------------------------------------------------------
# Build the report data
# -----------------------------------------------------------------------------

def build_report(
    client: GraphClient,
    days: int = 30,
    work_start: int = 9,
    work_end: int = 17,
    tz: str = "America/Los_Angeles",
    quick: bool = False,
    progress=None,
) -> dict:
    def step(label: str) -> None:
        if progress is not None:
            progress(label)

    data: dict = {"generated_at": datetime.now()}

    step("Walking folder tree...")
    data["folders"] = folders_mod.walk_folders(client)

    step("Sampling message sizes...")
    data["sizes"] = sizes_mod.size_percentiles(client, limit=5000)

    step("Counting recipients...")
    data["recipients"] = people_mod.top_recipients(client, limit=20)

    step("Computing rates...")
    data["rates"] = rates_mod.compute_daily_rates(client, days, work_start, work_end, tz)
    data["rates_params"] = {"days": days, "work_start": work_start, "work_end": work_end, "tz": tz}

    if not quick:
        step("Fused Inbox-subtree walk (read-ratio + senders + attachments + types + received-volume)...")
        accum = _fused_inbox_walk(client)
        step("Counting sent volume...")
        sent = _sent_volume_by_month(client)

        data["read_ratio"] = {
            d: read_mod.ReadStats(total=n, read=accum.read_reads[d])
            for d, n in accum.read_totals.items()
        }
        data["senders"] = accum.senders.most_common(20)
        data["attachments"] = accum.attachments
        data["message_types"] = accum.types
        data["volume"] = {"sent": sent, "received": accum.received, "filed": accum.filed}

    data["quick"] = quick
    return data


# -----------------------------------------------------------------------------
# Charts — matplotlib → inline SVG
# -----------------------------------------------------------------------------

def _fig_to_inline_svg(fig) -> str:
    buf = io.StringIO()
    fig.savefig(buf, format="svg", bbox_inches="tight")
    plt.close(fig)
    svg = buf.getvalue()
    # Drop XML declaration / DOCTYPE so the SVG drops cleanly into HTML body.
    lines = svg.splitlines()
    while lines and not lines[0].lstrip().startswith("<svg"):
        lines.pop(0)
    return "\n".join(lines)


def chart_top_folders(folders_tree, top_n: int = 15) -> str:
    flat: list[tuple[str, int]] = []

    def walk(n: folders_mod.FolderNode) -> None:
        flat.append((n.display_name, n.tree_size_in_bytes))
        for c in n.children:
            walk(c)

    for top in folders_tree:
        walk(top)
    flat.sort(key=lambda kv: kv[1], reverse=True)
    top = flat[:top_n][::-1]
    names = [n for n, _ in top]
    sizes_mb = [s / (1024 * 1024) for _, s in top]

    fig, ax = plt.subplots(figsize=(10, 0.4 * len(top) + 1))
    ax.barh(names, sizes_mb, color=CHART_FG)
    ax.set_xlabel("Size (MB, includes descendants)")
    ax.set_title(f"Top {top_n} folders by total size")
    ax.set_facecolor(CHART_BG)
    fig.patch.set_facecolor(CHART_BG)
    ax.grid(axis="x", color="#dddddd", linewidth=0.5)
    ax.set_axisbelow(True)
    return _fig_to_inline_svg(fig)


def chart_volume_timeline(volume: dict, last_n: int = 36) -> str:
    months = sorted(set(volume["sent"]) | set(volume["received"]))[-last_n:]
    sent = [volume["sent"].get(m, 0) for m in months]
    received = [volume["received"].get(m, 0) for m in months]
    filed = [volume["filed"].get(m, 0) for m in months]

    fig, ax = plt.subplots(figsize=(14, 4.5))
    ax.plot(months, received, marker="o", markersize=4, label="received", color=CHART_FG)
    ax.plot(months, filed, marker="o", markersize=4, label="filed (subset)", color=CHART_MUTED, linestyle="--")
    ax.plot(months, sent, marker="o", markersize=4, label="sent", color=CHART_ACCENT)
    ax.set_ylabel("Messages")
    ax.set_title(f"Monthly volume — last {last_n} months")
    ax.legend(loc="upper left")
    ax.set_facecolor(CHART_BG)
    fig.patch.set_facecolor(CHART_BG)
    ax.grid(axis="y", color="#dddddd", linewidth=0.5)
    ax.set_axisbelow(True)
    plt.xticks(rotation=60, fontsize=8)
    fig.tight_layout()
    return _fig_to_inline_svg(fig)


def chart_rates_timeline(window: rates_mod.TimeWindow) -> str:
    days_sorted = sorted(window.daily.keys())
    received = [window.daily[d].received for d in days_sorted]
    sent = [window.daily[d].sent for d in days_sorted]
    colors = [CHART_FG if d.weekday() < 5 else "#a8b5c4" for d in days_sorted]

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 6), sharex=True)
    ax1.bar(days_sorted, received, color=colors)
    ax1.set_title(f"Received per day — last {window.days} days (weekend days dimmed)")
    ax1.set_ylabel("messages")
    ax2.bar(days_sorted, sent, color=[CHART_ACCENT if d.weekday() < 5 else "#e8c4a8" for d in days_sorted])
    ax2.set_title("Sent per day")
    ax2.set_ylabel("messages")
    for ax in (ax1, ax2):
        ax.set_facecolor(CHART_BG)
        ax.grid(axis="y", color="#dddddd", linewidth=0.5)
        ax.set_axisbelow(True)
    fig.patch.set_facecolor(CHART_BG)
    plt.xticks(rotation=45, fontsize=8)
    fig.tight_layout()
    return _fig_to_inline_svg(fig)


def chart_top_addresses(rows: list[tuple[str, int]], title: str, color: str) -> str:
    rows = rows[::-1]
    fig, ax = plt.subplots(figsize=(10, 0.35 * len(rows) + 1))
    ax.barh([a for a, _ in rows], [n for _, n in rows], color=color)
    ax.set_xlabel("messages")
    ax.set_title(title)
    ax.set_facecolor(CHART_BG)
    fig.patch.set_facecolor(CHART_BG)
    ax.grid(axis="x", color="#dddddd", linewidth=0.5)
    ax.set_axisbelow(True)
    return _fig_to_inline_svg(fig)


def chart_sizes(sizes: dict) -> str:
    pcts = sizes_mod.PERCENTILES
    x = list(range(len(pcts)))
    width = 0.4
    sent = [sizes["sent"].get(p, 0) / 1024 for p in pcts]
    recv = [sizes["received"].get(p, 0) / 1024 for p in pcts]
    fig, ax = plt.subplots(figsize=(8, 4))
    ax.bar([i - width / 2 for i in x], sent, width=width, label="sent", color=CHART_ACCENT)
    ax.bar([i + width / 2 for i in x], recv, width=width, label="received", color=CHART_FG)
    ax.set_xticks(x)
    ax.set_xticklabels([f"p{p}" for p in pcts])
    ax.set_ylabel("KB")
    ax.set_yscale("log")
    ax.set_title("Message size percentiles (KB, log scale)")
    ax.legend()
    ax.set_facecolor(CHART_BG)
    fig.patch.set_facecolor(CHART_BG)
    ax.grid(axis="y", color="#dddddd", linewidth=0.5, which="both")
    ax.set_axisbelow(True)
    return _fig_to_inline_svg(fig)


# -----------------------------------------------------------------------------
# Render the template
# -----------------------------------------------------------------------------

def render_report(data: dict) -> str:
    charts: dict[str, str] = {}
    charts["folders"] = chart_top_folders(data["folders"])
    charts["sizes"] = chart_sizes(data["sizes"])
    charts["rates"] = chart_rates_timeline(data["rates"])
    if data["recipients"]:
        charts["recipients"] = chart_top_addresses(data["recipients"], "Top recipients (Sent Items)", CHART_ACCENT)
    if not data["quick"]:
        charts["volume"] = chart_volume_timeline(data["volume"])
        if data["senders"]:
            charts["senders"] = chart_top_addresses(data["senders"], "Top senders (Inbox + subfolders)", CHART_FG)

    env = Environment(
        loader=PackageLoader("mailbox_shape", "templates"),
        autoescape=select_autoescape(),
    )
    env.filters["humanbytes"] = _humanbytes
    env.filters["pct"] = lambda v: f"{v * 100:.0f}%"
    env.filters["intcomma"] = lambda v: f"{int(v):,}"
    template = env.get_template("report.html")

    # Pre-compute totals for header.
    total_items = sum(f.tree_item_count for f in data["folders"])
    total_size = sum(f.tree_size_in_bytes for f in data["folders"])

    # Sort read-ratio rows by total descending; cap at 60.
    if "read_ratio" in data:
        read_rows = sorted(data["read_ratio"].items(), key=lambda kv: kv[1].total, reverse=True)[:60]
    else:
        read_rows = []

    # Pre-sort attachments and message_types rows for the template.
    attach_rows: list = []
    types_rows: list = []
    if not data["quick"]:
        attach_rows = sorted(data["attachments"].items(), key=lambda kv: kv[1].total, reverse=True)
        types_rows = sorted(data["message_types"].items(), key=lambda kv: sum(kv[1].values()), reverse=True)

    return template.render(
        data=data,
        charts=charts,
        total_items=total_items,
        total_size=total_size,
        read_rows=read_rows,
        attach_rows=attach_rows,
        types_rows=types_rows,
    )


def write_report(data: dict, output: Path) -> None:
    html = render_report(data)
    output.write_text(html, encoding="utf-8")


def _humanbytes(n: int | None) -> str:
    if n is None:
        return "—"
    units = ["B", "KB", "MB", "GB", "TB"]
    f = float(n)
    for u in units:
        if f < 1024 or u == units[-1]:
            return f"{f:,.1f} {u}"
        f /= 1024
    return f"{n} B"
