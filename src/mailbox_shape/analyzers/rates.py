"""Per-day / per-weekday / per-working-hour rates over a recent time window."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone
from zoneinfo import ZoneInfo

from dateutil import parser as dateparser

from . import folders as folders_mod
from ..graph import GraphClient


@dataclass
class Counts:
    total: int = 0
    weekday: int = 0  # received on Mon-Fri (any hour)
    weekend: int = 0  # received on Sat-Sun (any hour)
    working: int = 0  # received Mon-Fri inside [work_start, work_end)


@dataclass
class WindowRates:
    days: int
    weekdays: int
    weekend_days: int
    working_hours: int
    work_start: int
    work_end: int
    tz: str
    sent: Counts = field(default_factory=Counts)
    received: Counts = field(default_factory=Counts)
    filed: Counts = field(default_factory=Counts)  # subset of received: messages outside Inbox root


def _iso_z(ts: datetime) -> str:
    """ISO 8601 with trailing Z — Graph's preferred date-time literal."""
    return ts.astimezone(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def _denominators(now_local: datetime, days: int, work_start: int, work_end: int) -> tuple[int, int, int]:
    """Count weekdays, weekend days, and working hours in the last `days` window."""
    weekdays = 0
    weekend_days = 0
    today: date = now_local.date()
    for offset in range(days):
        d = today - timedelta(days=offset)
        if d.weekday() < 5:
            weekdays += 1
        else:
            weekend_days += 1
    return weekdays, weekend_days, weekdays * (work_end - work_start)


def _accumulate(
    client: GraphClient,
    folder_id: str,
    ts_field: str,
    since_iso_z: str,
    tzinfo: ZoneInfo,
    work_start: int,
    work_end: int,
    bucket: Counts,
) -> None:
    params = {
        "$select": f"id,{ts_field}",
        "$top": 500,
        "$filter": f"{ts_field} ge {since_iso_z}",
    }
    for msg in client.paged(f"/me/mailFolders/{folder_id}/messages", **params):
        raw = msg.get(ts_field)
        if not raw:
            continue
        local = dateparser.isoparse(raw).astimezone(tzinfo)
        bucket.total += 1
        if local.weekday() < 5:
            bucket.weekday += 1
            if work_start <= local.hour < work_end:
                bucket.working += 1
        else:
            bucket.weekend += 1


def compute_rates(
    client: GraphClient,
    days: int = 30,
    work_start: int = 9,
    work_end: int = 17,
    tz: str = "America/Los_Angeles",
) -> WindowRates:
    if not 0 <= work_start < work_end <= 24:
        raise ValueError(f"Invalid work hours: {work_start}–{work_end}")
    tzinfo = ZoneInfo(tz)
    now_local = datetime.now(tzinfo)
    since = (now_local - timedelta(days=days)).astimezone(timezone.utc)
    since_iso_z = _iso_z(since)

    weekdays, weekend_days, working_hours = _denominators(now_local, days, work_start, work_end)
    result = WindowRates(
        days=days,
        weekdays=weekdays,
        weekend_days=weekend_days,
        working_hours=working_hours,
        work_start=work_start,
        work_end=work_end,
        tz=tz,
    )

    # Sent.
    _accumulate(client, "sentitems", "sentDateTime", since_iso_z, tzinfo, work_start, work_end, result.sent)

    # Received: Inbox root + descendants. Track 'filed' = anything that's not in root.
    inbox = folders_mod.walk_subtree(client, "inbox")
    for i, f in enumerate(folders_mod.flatten(inbox)):
        if f.total_item_count == 0:
            continue
        target = Counts()
        _accumulate(client, f.id, "receivedDateTime", since_iso_z, tzinfo, work_start, work_end, target)
        result.received.total += target.total
        result.received.weekday += target.weekday
        result.received.weekend += target.weekend
        result.received.working += target.working
        if i != 0:
            result.filed.total += target.total
            result.filed.weekday += target.weekday
            result.filed.weekend += target.weekend
            result.filed.working += target.working

    return result
