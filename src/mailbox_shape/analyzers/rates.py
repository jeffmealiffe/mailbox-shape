"""Per-day timeline of message rates, normalized to working hours."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone
from zoneinfo import ZoneInfo

from dateutil import parser as dateparser

from . import folders as folders_mod
from ..graph import GraphClient


@dataclass
class DayCounts:
    # `*_working` are the subset of the same-named count received/sent during
    # configured working hours of a weekday.
    sent: int = 0
    sent_working: int = 0
    received: int = 0
    received_working: int = 0
    filed: int = 0
    filed_working: int = 0

    def add(self, other: "DayCounts") -> None:
        self.sent += other.sent
        self.sent_working += other.sent_working
        self.received += other.received
        self.received_working += other.received_working
        self.filed += other.filed
        self.filed_working += other.filed_working


@dataclass
class TimeWindow:
    days: int
    work_start: int
    work_end: int
    tz: str
    daily: dict[date, DayCounts] = field(default_factory=dict)

    @property
    def hours_per_workday(self) -> int:
        return self.work_end - self.work_start

    @property
    def weekdays(self) -> int:
        return sum(1 for d in self.daily if d.weekday() < 5)

    @property
    def total_working_hours(self) -> int:
        return self.weekdays * self.hours_per_workday

    def aggregate(self) -> DayCounts:
        agg = DayCounts()
        for c in self.daily.values():
            agg.add(c)
        return agg


def _iso_z(ts: datetime) -> str:
    return ts.astimezone(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def _scan(
    client: GraphClient,
    folder_id: str,
    ts_field: str,
    since_iso_z: str,
    tzinfo: ZoneInfo,
    work_start: int,
    work_end: int,
    window: TimeWindow,
    series: str,  # 'sent', 'received', or 'filed'
) -> None:
    params = {
        "$select": f"id,{ts_field}",
        "$top": 500,
        "$filter": f"{ts_field} ge {since_iso_z}",
    }
    for msg in client.paged(f"{client.mailbox}/mailFolders/{folder_id}/messages", **params):
        raw = msg.get(ts_field)
        if not raw:
            continue
        local = dateparser.isoparse(raw).astimezone(tzinfo)
        d = local.date()
        if d not in window.daily:
            continue  # boundary slop — Graph may include a stray msg just outside the window
        bucket = window.daily[d]
        in_work_hours = local.weekday() < 5 and work_start <= local.hour < work_end
        if series == "sent":
            bucket.sent += 1
            if in_work_hours:
                bucket.sent_working += 1
        elif series == "received":
            bucket.received += 1
            if in_work_hours:
                bucket.received_working += 1
        else:  # filed
            bucket.filed += 1
            if in_work_hours:
                bucket.filed_working += 1


def compute_daily_rates(
    client: GraphClient,
    days: int = 30,
    work_start: int = 9,
    work_end: int = 17,
    tz: str = "America/Los_Angeles",
) -> TimeWindow:
    """Per-day counts for the last `days` days, in `tz`.

    For each day we track: total sent / received / filed, and the subset of
    each that landed inside [work_start, work_end) on a weekday. 'filed' is
    the subset of received that came in to an Inbox subfolder rather than
    Inbox root.

    Implementation uses Graph $filter on the timestamp field to scope each
    folder fetch to the window — fast even on a huge mailbox.
    """
    if not 0 <= work_start < work_end <= 24:
        raise ValueError(f"Invalid work hours: {work_start}–{work_end}")

    tzinfo = ZoneInfo(tz)
    now_local = datetime.now(tzinfo)
    since = (now_local - timedelta(days=days)).astimezone(timezone.utc)
    since_iso_z = _iso_z(since)

    window = TimeWindow(days=days, work_start=work_start, work_end=work_end, tz=tz)
    today = now_local.date()
    for offset in range(days):
        window.daily[today - timedelta(days=offset)] = DayCounts()

    _scan(client, "sentitems", "sentDateTime", since_iso_z, tzinfo, work_start, work_end, window, "sent")

    inbox = folders_mod.walk_subtree(client, "inbox")
    for i, f in enumerate(folders_mod.flatten(inbox)):
        if f.total_item_count == 0:
            continue
        _scan(client, f.id, "receivedDateTime", since_iso_z, tzinfo, work_start, work_end, window, "received")
        if i != 0:
            _scan(client, f.id, "receivedDateTime", since_iso_z, tzinfo, work_start, work_end, window, "filed")

    return window
