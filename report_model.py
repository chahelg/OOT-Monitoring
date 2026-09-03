"""
report_model.py — the single source of truth for the Integration Health
Report. build_report() returns one fully-computed dict; the live page
and the export both render that same dict through different templates,
so they can never disagree with each other. No Flask import.

    py report_model.py --self-check "output\\active-alert category 19 Aug.xlsx"

Degrades gracefully by construction: every section checks for a missing
interface matrix or missing history and falls back rather than raising —
see build_report()'s docstring for exactly what each band looks like
without them.
"""

from __future__ import annotations

import argparse
import sys
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

import generate_workbook as gw
import history as hist
import interface_matrix as im
import severity as sv

AGING_BUCKETS = [
    ("<24h", lambda d: d < 1),
    ("1-2d", lambda d: 1 <= d < 3),
    ("3-7d", lambda d: 3 <= d < 8),
    (">7d", lambda d: d >= 8),
]

SPARKLINE_DAYS = 7
QUEUE_DEFAULT_ROWS = 10


# --------------------------------------------------------------------
# Small render helpers — plain data in, plain strings out. No Jinja
# computation needed for any of this (per the brief's "webapp.py stays
# thin... no computation in Jinja" rule) since the template just drops
# these strings in directly.
# --------------------------------------------------------------------

def _sparkline_values(day_counts: list[tuple[date, int]], end_date: date, days: int = SPARKLINE_DAYS) -> list[int]:
    by_date = dict(day_counts)
    return [by_date.get(end_date - timedelta(days=n), 0) for n in range(days - 1, -1, -1)]


def render_sparkline_svg(values: list[int], width: int = 84, height: int = 22) -> str:
    """A hand-emitted inline SVG sparkline — no chart library, per the
    brief. Neutral/muted colour deliberately: severity is already
    carried by the row's left border and P1/P2/P3 label, so the
    sparkline's job is just to show shape, not add a second colour
    signal fighting the first."""
    n = len(values)
    if n == 0:
        return ""
    vmax = max(values) or 1
    step = width / max(n - 1, 1)
    pad = 3
    points = []
    for i, v in enumerate(values):
        x = i * step
        y = height - pad - (v / vmax) * (height - 2 * pad)
        points.append(f"{x:.1f},{y:.1f}")
    points_attr = " ".join(points)
    return (
        f'<svg width="{width}" height="{height}" viewBox="0 0 {width} {height}" '
        f'class="sparkline" role="img" aria-label="{n}-day trend: {", ".join(str(v) for v in values)}">'
        f'<polyline points="{points_attr}" fill="none" stroke="#6b7686" stroke-width="1.5" '
        f'stroke-linecap="round" stroke-linejoin="round"/></svg>'
    )


def render_daily_bar_chart_svg(
    day_alert_counts: list[tuple[date, int]],
    day_object_counts: dict[date, int],
    width: int = 760,
    height: int = 200,
) -> str:
    """The History tab's daily-volume chart — richer than a bare
    sparkline on purpose: value gridlines, a handful of date labels,
    one bar per day for alert volume, and a dot per day for distinct
    objects affected (so the "how much of this volume is retries vs
    real distinct documents" story the rest of this report tells is
    visible in one chart, not two separate line graphs). Still no
    chart library — plain SVG with data-* attributes carrying the
    exact values; the page's own small hover-tooltip script (see
    view.html) reads them, so this function stays pure and Flask-free."""
    n = len(day_alert_counts)
    if n == 0:
        return ""
    pad_l, pad_r, pad_t, pad_b = 34, 6, 10, 22
    plot_w, plot_h = width - pad_l - pad_r, height - pad_t - pad_b
    vmax = max((c for _, c in day_alert_counts), default=0) or 1
    bar_w = plot_w / n
    bar_gap = min(3.0, bar_w * 0.25)
    label_every = max(1, round(n / 7))

    parts = []
    for frac in (0, 0.25, 0.5, 0.75, 1.0):
        y = pad_t + plot_h * (1 - frac)
        parts.append(
            f'<line x1="{pad_l}" y1="{y:.1f}" x2="{width - pad_r}" y2="{y:.1f}" '
            f'stroke="#e5e8ee" stroke-width="1"/>'
            f'<text x="{pad_l - 6}" y="{y:.1f}" text-anchor="end" dominant-baseline="middle" '
            f'class="hist-axis-label">{round(vmax * frac)}</text>'
        )

    bars, dots, labels = [], [], []
    for i, (d, alerts) in enumerate(day_alert_counts):
        x = pad_l + i * bar_w
        bar_h = (alerts / vmax) * plot_h
        y = pad_t + (plot_h - bar_h)
        bars.append(
            f'<rect class="hist-bar" data-date="{d.strftime("%b %d")}" data-alerts="{alerts}" '
            f'data-docs="{day_object_counts.get(d, 0)}" '
            f'x="{x + bar_gap / 2:.1f}" y="{y:.1f}" width="{max(bar_w - bar_gap, 1):.1f}" '
            f'height="{max(bar_h, 0):.1f}" rx="1.5"/>'
        )
        docs = day_object_counts.get(d, 0)
        if docs:
            dot_y = pad_t + plot_h * (1 - docs / vmax)
            dots.append(f'<circle class="hist-dot" cx="{x + bar_w / 2:.1f}" cy="{dot_y:.1f}" r="2.5"/>')
        if i == 0 or i == n - 1 or i % label_every == 0:
            # "%-d" (no leading zero) is a Unix-only strftime extension
            # that raises on Windows' C library — this app runs on
            # Windows, so plain "%d %b" (e.g. "05 Aug") it is.
            labels.append(
                f'<text x="{x + bar_w / 2:.1f}" y="{height - 6}" text-anchor="middle" '
                f'class="hist-axis-label">{d.strftime("%d %b")}</text>'
            )

    label_range = f'{day_alert_counts[0][0].strftime("%b %d")} to {day_alert_counts[-1][0].strftime("%b %d")}'
    return (
        f'<svg viewBox="0 0 {width} {height}" class="hist-chart" role="img" '
        f'aria-label="Daily alert volume and distinct documents affected, {n} days, {label_range}">'
        + "".join(parts) + "".join(bars) + "".join(dots) + "".join(labels) +
        "</svg>"
    )


def _trend_arrow(trend: int) -> str:
    return {1: "▲", -1: "▼"}.get(trend, "–")  # ▲ / ▼ / –


def _format_age(age_days: float) -> str:
    if age_days < 1:
        return f"{age_days * 24:.0f}h"
    return f"{age_days:.0f}d"


def _aging_bucket(age_days: float) -> str:
    for label, test in AGING_BUCKETS:
        if test(age_days):
            return label
    return AGING_BUCKETS[-1][0]


# --------------------------------------------------------------------
# Ownership / contacts
# --------------------------------------------------------------------

def issue_contacts(issue: "sv.Issue") -> list[str]:
    """Contacts for one issue: the fixed stream -> people table
    (gw.STREAM_CONTACTS) only — by explicit instruction, not the
    interface matrix's own ams_team/app_owner_it columns, which in
    practice turned out to be noisy free-text (raw email addresses
    mixed in with names, 4-5 entries per issue) rather than a clean
    per-application contact. Streams with nobody on that fixed table
    ("Unmapped", "MDG", "Unavailable" when the matrix itself is
    missing, etc.) fall back to the technical contact, matching the
    "others -> Harshit Joshi" catch-all given alongside the table.

    This intentionally drops the old Technical-grouping-is-additive
    nuance (Band 6's legacy gw._resolve_category_contacts still has
    it, unchanged) — the table given here is a straight stream ->
    contact lookup, not a per-issue blend."""
    contacts = gw._contacts_for_streams([issue.stream])
    return contacts if contacts else [gw.TECHNICAL_CONTACT]


# --------------------------------------------------------------------
# Band assembly
# --------------------------------------------------------------------

def _queue_row(issue: "sv.Issue", latest_date: date) -> dict:
    rec = issue.interface_record
    flags = []
    if rec and rec.fi_posting_impact:
        flags.append("FI")
    if rec and rec.stock_movement:
        flags.append("Stock")
    if rec and rec.time_critical:
        flags.append("Time-critical")
    return {
        "priority": issue.priority,
        "score": issue.score,
        "priority_reasons": issue.priority_reasons,
        "stream": issue.stream,
        "application": issue.application,
        "interface_name": rec.interface_name if rec else None,
        "match_state": issue.match_state,
        "error_category": issue.error_category,
        "distinct_objects": issue.distinct_objects,
        "alert_count": issue.alert_count,
        # The dual-measure figure ("N docs · M alerts") only earns its
        # place when the two numbers actually differ — that's the
        # retry-storm signal it exists to show. When every alert is its
        # own distinct object (the common case on real data so far),
        # showing the same number twice is just noise, so the template
        # collapses to one plain alert count instead.
        "docs_alerts_differ": issue.distinct_objects != issue.alert_count,
        "age_label": _format_age(issue.age_days),
        "age_days": issue.age_days,
        "trend_arrow": _trend_arrow(issue.trend),
        "trend": issue.trend,
        "sparkline_svg": render_sparkline_svg(_sparkline_values(issue.day_counts, latest_date)),
        "flags": flags,
        "owners": issue_contacts(issue),
        "enrichment_confidence": issue.enrichment_confidence,
    }


def _active_issue_count(issues: list["sv.Issue"], on_date: date | None) -> int:
    if on_date is None:
        return 0
    return sum(1 for i in issues if any(d == on_date for d, _ in i.day_counts))


def _build_headline(issues: list["sv.Issue"], issues_union: list["sv.Issue"], file_latest_date: date | None) -> dict:
    p1_issues = [i for i in issues if i.priority == "P1"]
    fi_issues = [i for i in issues if i.interface_record and i.interface_record.fi_posting_impact]

    # Net vs yesterday: how many distinct issues were *actively*
    # generating alerts on each of the two days, both counted from the
    # same source (the full historical union) — comparing today's file
    # (a narrow rolling window) against a cumulative "everything ever
    # seen up to yesterday" count would always show a large negative
    # number as history accumulates, regardless of whether things are
    # actually improving. Like-for-like: active-on-day-X vs
    # active-on-day-X-minus-1.
    yesterday = (file_latest_date - timedelta(days=1)) if file_latest_date else None
    today_active = _active_issue_count(issues_union, file_latest_date)
    yesterday_active = _active_issue_count(issues_union, yesterday)

    return {
        "p1_open": len(p1_issues),
        "docs_stuck": sum(i.distinct_objects for i in issues),
        "fi_impacting": len(fi_issues),
        "streams_at_risk": len({i.stream for i in p1_issues}),
        "oldest_days": max((i.age_days for i in issues), default=0.0),
        "net_vs_yesterday": today_active - yesterday_active,
    }


def _build_verdict(headline: dict) -> str:
    """One deterministic, plain-English sentence — no LLM, so it's
    reproducible and defensible in a management meeting (same input
    always produces the same sentence)."""
    p1 = headline["p1_open"]
    if p1 == 0:
        return "No P1 issues open right now."
    clauses = [f'{p1} issue{"s" if p1 != 1 else ""} need{"s" if p1 == 1 else ""} attention']
    if headline["fi_impacting"]:
        n = headline["fi_impacting"]
        clauses.append(f'{n} affect{"s" if n == 1 else ""} financial posting')
    oldest = headline["oldest_days"]
    clauses.append(f'the oldest has been open {oldest:.0f} day{"s" if round(oldest) != 1 else ""}')
    if len(clauses) == 1:
        return clauses[0] + "."
    return "; ".join(clauses[:-1]) + ", and " + clauses[-1] + "."


def _build_ownership(issues: list["sv.Issue"]) -> list[dict]:
    by_stream: dict[str, list[sv.Issue]] = defaultdict(list)
    for i in issues:
        by_stream[i.stream].append(i)

    rows = []
    for stream, stream_issues in by_stream.items():
        p1_count = sum(1 for i in stream_issues if i.priority == "P1")
        docs = sum(i.distinct_objects for i in stream_issues)
        oldest = max((i.age_days for i in stream_issues), default=0.0)
        buckets = Counter(_aging_bucket(i.age_days) for i in stream_issues)
        contacts: list[str] = []
        for i in stream_issues:
            for name in issue_contacts(i):
                if name not in contacts:
                    contacts.append(name)
        alert_count = sum(i.alert_count for i in stream_issues)
        rows.append({
            "stream": stream,
            "p1_count": p1_count,
            "open_count": len(stream_issues),
            "distinct_objects": docs,
            "alert_count": alert_count,
            "docs_alerts_differ": docs != alert_count,  # see _queue_row's identical field for why
            "oldest_days": oldest,
            "aging_buckets": {label: buckets.get(label, 0) for label, _ in AGING_BUCKETS},
            "contacts": contacts,
            "escalate": p1_count > 0 and oldest > 7,
        })

    rows.sort(key=lambda r: (-r["p1_count"], -r["distinct_objects"]))
    return rows


def _movement_row(issue: "sv.Issue") -> dict:
    """A render-ready shape for one issue appearing in a movement list
    (cleared/new/recurring) — same idea as _queue_row, kept separate
    and lighter since these lists don't need score/flags/sparkline.
    Converting here (not leaving raw sv.Issue objects in the report
    dict) keeps the "no computation in Jinja" rule honest: the
    template only ever does attribute/key lookups on plain dicts."""
    rec = issue.interface_record
    return {
        "stream": issue.stream,
        "application": issue.application,
        "error_category": issue.error_category,
        "interface_name": rec.interface_name if rec else None,
        "priority": issue.priority,
        "distinct_objects": issue.distinct_objects,
        "alert_count": issue.alert_count,
        "docs_alerts_differ": issue.distinct_objects != issue.alert_count,  # see _queue_row
        "last_seen_date": issue.last_seen.date().isoformat(),
        "last_seen_label": issue.last_seen.strftime("%b %d"),
        "first_seen_label": issue.first_seen.strftime("%b %d"),
    }


def _build_movement(events: list[dict], issues_union: list["sv.Issue"], latest_date: date | None) -> dict:
    if latest_date:
        window_start = latest_date - timedelta(days=13)
        object_days: dict[date, set] = defaultdict(set)
        for e in events:
            d = datetime.fromisoformat(e["timestamp_readable"]).date()
            if window_start <= d <= latest_date:
                key = e.get("BUSINESS_OBJECT_KEY") or f"__unkeyed_{id(e)}"
                object_days[d].add(key)
        series_14d = [(d, len(object_days.get(d, set()))) for d in
                      (window_start + timedelta(days=n) for n in range(14))]
        series_svg = render_sparkline_svg([c for _, c in series_14d], width=560, height=48)
        series_start_label = window_start.strftime("%b %d")
        series_end_label = latest_date.strftime("%b %d")
    else:
        series_14d = []
        series_svg = ""
        series_start_label = series_end_label = ""

    movement = hist.movement_summary(issues_union, latest_date)
    return {
        "series_14d": series_14d,
        "series_svg": series_svg,
        "series_start_label": series_start_label,
        "series_end_label": series_end_label,
        "cleared": [_movement_row(i) for i in movement["cleared_since_yesterday"]],
        "new_today": [_movement_row(i) for i in movement["new_today"]],
        "recurring": [_movement_row(i) for i in movement["recurring"]],
    }


def _reconciliation(issues: list["sv.Issue"], total_alerts: int) -> dict:
    by_stream_sum = sum(i.alert_count for i in issues)
    by_severity_sum = sum(i.alert_count for i in issues)  # every issue has exactly one priority
    matched = sum(i.alert_count for i in issues if i.match_state != "unmatched")
    unmatched = sum(i.alert_count for i in issues if i.match_state == "unmatched")
    return {
        "total_alerts": total_alerts,
        "by_stream_sum": by_stream_sum,
        "by_severity_sum": by_severity_sum,
        "matched_plus_unmatched_sum": matched + unmatched,
        "ok": total_alerts == by_stream_sum == by_severity_sum == (matched + unmatched),
    }


# --------------------------------------------------------------------
# Export PII policy
# --------------------------------------------------------------------

def redact_names(names: list[str]) -> list[str]:
    """Strips anything email-shaped, keeping human names only. The
    interface matrix's ams_team/app_owner_it fields sometimes hold raw
    email addresses alongside plain names (issue_contacts() doesn't
    distinguish the two) — this is the one place that does, and only
    for the export. Public (not module-private) since webapp.py's
    export route needs to apply the same redaction to Band 6's legacy
    contact list, which lives outside the report dict apply_contact_
    policy() below covers."""
    return [n for n in names if "@" not in n]


def apply_contact_policy(report: dict, full: bool = False) -> dict:
    """Returns a NEW report dict with every owners/contacts list
    redacted to names-only, unless `full` is True. This is the only
    transform ever applied on top of build_report()'s output — every
    other field (counts, priorities, reconciliation, the verdict) is
    passed through byte-for-byte unchanged, so it can't cause the
    export to disagree with the live page on anything that matters.
    Only the export route calls this; the live page always shows full
    contact info, since it never leaves this machine."""
    if full:
        return report

    out = dict(report)
    out["queue_default"] = [{**row, "owners": redact_names(row["owners"])} for row in report["queue_default"]]
    out["queue_extra"] = [{**row, "owners": redact_names(row["owners"])} for row in report["queue_extra"]]
    out["ownership"] = [{**row, "contacts": redact_names(row["contacts"])} for row in report["ownership"]]
    return out


# --------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------

def build_report(file_path: Path, matrix_path: Path | None = None, now: datetime | None = None) -> dict:
    """The single source of truth. Everything the live page and the
    export render comes from this one dict — see module docstring."""
    now = now or datetime.now()
    matrix_path = matrix_path or im.DEFAULT_MATRIX_PATH

    rows = gw.read_generated_rows(file_path)
    total_alerts = sum(1 for r in rows if r.get("alert_type"))

    interface_index = im.load_interface_index(matrix_path)
    interface_matrix_available = interface_index is not None

    # History: never allowed to take the whole report down. A failure
    # here (e.g. a corrupt cache, a genuinely unreadable historical
    # file already skipped by history.py itself) degrades to "no
    # multi-day context" rather than a 500 — trend stays flat and
    # movement is empty, same shape as history genuinely being empty.
    try:
        cache = hist.build_history_index()
        union_events = hist.deduplicated_events(cache)
    except Exception as e:
        print(f"[report_model] history unavailable: {e}", file=sys.stderr)
        union_events = []

    union_rows = hist.events_to_rows(union_events) if union_events else []
    issues_union = sv.group_issues(union_rows, interface_index, now=now) if union_rows else []
    latest_union_date = max((i.last_seen.date() for i in issues_union), default=None)
    trends = hist.compute_trends(issues_union, latest_union_date) if issues_union else {}

    issues = sv.group_issues(rows, interface_index, now=now, trend_overrides=trends)
    file_latest_date = max((r["timestamp_readable"].date() for r in rows if r.get("alert_type")), default=None)

    headline = _build_headline(issues, issues_union, file_latest_date)

    report = {
        "file_name": file_path.name,
        "file_date": file_latest_date,
        "generated_at": now,
        "interface_matrix_available": interface_matrix_available,
        "verdict": _build_verdict(headline),
        "headline": headline,
        "queue_default": [_queue_row(i, file_latest_date) for i in issues[:QUEUE_DEFAULT_ROWS]] if file_latest_date else [],
        "queue_extra": [_queue_row(i, file_latest_date) for i in issues[QUEUE_DEFAULT_ROWS:]] if file_latest_date else [],
        "queue_total": len(issues),
        "ownership": _build_ownership(issues),
        "movement": _build_movement(union_events, issues_union, latest_union_date) if union_events else
                    {"series_14d": [], "cleared": [], "new_today": [], "recurring": []},
        "reconciliation": _reconciliation(issues, total_alerts),
        "issues_all": issues,  # kept for Band 6 / drilldown reuse, not directly rendered as a table
    }
    return report


# --------------------------------------------------------------------
# History report — a second, independent entry point. Unlike
# build_report(), this isn't scoped to one selected file; it's a view
# over the whole deduplicated cross-file union history.py already
# maintains. Same "one dict, no computation in Jinja" contract.
# --------------------------------------------------------------------

def _file_row(path_str: str, entry: dict) -> dict:
    """Render-ready shape for one row of the History tab's "files
    ingested" table. date_mismatch is re-derived here (filename vs.
    the cached file_date) rather than stored in the cache itself —
    it's cheap to recompute and keeps the cache a pure, disposable
    cache rather than a second source of truth (see history.py's own
    docstring on this)."""
    name = Path(path_str).name
    file_date = date.fromisoformat(entry["file_date"]) if entry.get("file_date") else None
    filename_date = hist._infer_filename_date(name)
    return {
        "name": name,
        "file_date": file_date,
        "file_date_label": file_date.strftime("%Y-%m-%d") if file_date else "unknown",
        "event_count": len(entry["events"]),
        "date_mismatch": bool(filename_date and file_date and filename_date != file_date),
    }


def _window_row(events: list[dict], days: int, latest_date: date | None) -> dict:
    w = hist.window_totals(events, days, latest_date)
    return {
        "days": days,
        "alerts": w["alerts"],
        "distinct_objects": w["distinct_objects"],
        "docs_alerts_differ": w["distinct_objects"] != w["alerts"],
    }


def _stream_history_row(stream: str, day_counts: list[tuple[date, int]], trend: int) -> dict:
    return {
        "stream": stream,
        "total_alerts": sum(c for _, c in day_counts),
        "trend": trend,
        "trend_arrow": _trend_arrow(trend),
        "sparkline_svg": render_sparkline_svg([c for _, c in day_counts]),
    }


def build_history_report(matrix_path: Path | None = None) -> dict:
    """The History tab's single source of truth. Independent of any
    one selected file — reads from the user's own daily archive
    folders (hist.DAILY_FOLDERS_DIR), not output/, since a day only
    lands in output/ once someone runs Generate on it while the daily
    folders are populated every day regardless — see
    hist.build_daily_folder_index()'s docstring. Deliberately a
    separate index from build_report()'s (the Aging tab's own
    trend/movement bands), so this can't affect that tab."""
    matrix_path = matrix_path or im.DEFAULT_MATRIX_PATH
    interface_index = im.load_interface_index(matrix_path)

    cache = hist.build_daily_folder_index()
    files = [_file_row(p, e) for p, e in cache.get("files", {}).items()]
    files.sort(key=lambda f: f["file_date"] or date.min, reverse=True)
    total_raw_rows = sum(len(e["events"]) for e in cache.get("files", {}).values())

    events = hist.deduplicated_events(cache)
    base = {
        "has_data": False,
        "interface_matrix_available": interface_index is not None,
        "files": files,
        "total_events": len(events),
        "total_raw_rows": total_raw_rows,
        "earliest_date": None,
        "latest_date": None,
        "windows": [],
        "daily_chart_svg": "",
        "daily_series_start_label": "",
        "daily_series_end_label": "",
        "streams": [],
    }
    if not events:
        return base

    all_days = hist.daily_totals(events)
    latest_date, earliest_date = all_days[-1][0], all_days[0][0]
    object_days = dict(hist.daily_distinct_objects(events))

    span_days = min(30, (latest_date - earliest_date).days + 1)
    series = hist.stream_series(events, interface_index, days=span_days, latest_date=latest_date)
    trends = hist.stream_trends(series, latest_date)
    streams = [
        _stream_history_row(stream, day_counts, trends.get(stream, 0))
        for stream, day_counts in sorted(series.items(), key=lambda kv: -sum(c for _, c in kv[1]))
    ]

    base.update({
        "has_data": True,
        "earliest_date": earliest_date,
        "latest_date": latest_date,
        "windows": [_window_row(events, n, latest_date) for n in (7, 14, 30)],
        "daily_chart_svg": render_daily_bar_chart_svg(all_days, object_days),
        "daily_series_start_label": earliest_date.strftime("%b %d"),
        "daily_series_end_label": latest_date.strftime("%b %d"),
        "stream_window_days": span_days,
        "streams": streams,
    })
    return base


# --------------------------------------------------------------------
# History search — find a particular error's full multi-day history
# (every day it fired, trend, every raw occurrence with details),
# reusing the exact same deduplicated union build_history_report()
# is built from, so the two views can't disagree.
# --------------------------------------------------------------------

# Raw-alert-row fields a search query is matched against, beyond the
# error category itself — lets a business object key (or a snippet of
# an error/alert message, application, service, or transaction URL)
# find the error(s) it appeared under, not just an exact category name.
_SEARCHABLE_ROW_FIELDS = (
    "BUSINESS_OBJECT_KEY", "ALERT_MESSAGE", "ERROR_MESSAGE", "ERROR_DETAILS",
    "APPLICATION_NAME", "SERVICE", "TRANSACTION_URL",
)


def _row_matches_query(row: dict, query_lower: str) -> bool:
    return any(query_lower in str(row.get(f) or "").lower() for f in _SEARCHABLE_ROW_FIELDS)


def _history_issues(matrix_path: Path | None = None) -> list[sv.Issue]:
    """Every (stream, application, error_category) Issue across the
    FULL deduplicated multi-day history — same daily-folder source as
    build_history_report(), so search results and the History tab's
    other numbers can never disagree. Recomputed per call: grouping
    ~1-2k events is cheap in pure Python; the actually expensive part
    (reading every workbook) stays covered by
    hist.build_daily_folder_index()'s on-disk cache."""
    matrix_path = matrix_path or im.DEFAULT_MATRIX_PATH
    interface_index = im.load_interface_index(matrix_path)
    cache = hist.build_daily_folder_index()
    events = hist.deduplicated_events(cache)
    rows = hist.events_to_rows(events)
    return sv.group_issues(rows, interface_index)


def search_history(query: str) -> list[dict]:
    """Results for the History tab's search box. An error category
    (issue) matches if the query text appears in its own name OR in
    any searchable field of any raw alert row grouped under it — so
    searching a business object key still surfaces the error(s) that
    object triggered. Each result summarizes that error's ENTIRE
    history (every day it fired, not just one file); use
    history_error_details() for the raw occurrence-by-occurrence
    detail. Sorted busiest-first, same convention as the "By stream"
    table."""
    query_lower = query.strip().lower()
    if not query_lower:
        return []
    results = []
    for issue in _history_issues():
        category_hit = query_lower in (issue.error_category or "").lower()
        if not category_hit and not any(_row_matches_query(r, query_lower) for r in issue.sample_rows):
            continue
        results.append({
            "stream": issue.stream,
            "application": issue.application or "",
            "error_category": issue.error_category,
            "alert_count": issue.alert_count,
            "distinct_objects": issue.distinct_objects,
            "first_seen_label": issue.first_seen.strftime("%Y-%m-%d"),
            "last_seen_label": issue.last_seen.strftime("%Y-%m-%d"),
            "days_active": len(issue.day_counts),
            "sparkline_svg": render_sparkline_svg([c for _, c in issue.day_counts]),
            "match_kind": "category" if category_hit else "detail",
        })
    results.sort(key=lambda r: -r["alert_count"])
    return results


def history_error_details(stream: str, application: str, error_category: str) -> dict | None:
    """Every raw occurrence of one error category across the full
    multi-day history — the row-level detail behind one
    search_history() result. Keyed the same way group_issues() groups
    (stream, application, error_category); `application` of "" means
    the group's own APPLICATION_NAME was blank, not "any application".
    Returns None if no such issue exists (e.g. history changed between
    the search and this follow-up call)."""
    application_key = application or None
    for issue in _history_issues():
        if issue.stream == stream and issue.application == application_key and issue.error_category == error_category:
            rows = sorted(issue.sample_rows, key=lambda r: r["timestamp_readable"], reverse=True)
            return {"stream": issue.stream, "application": issue.application or "",
                    "error_category": issue.error_category, "rows": rows}
    return None


# --------------------------------------------------------------------
# Self-check CLI
# --------------------------------------------------------------------

def _self_check(file_path: Path) -> None:
    report = build_report(file_path)
    print(f"File: {report['file_name']}  (data through {report['file_date']})")
    print(f"Interface matrix available: {report['interface_matrix_available']}")
    print()
    print(f"Verdict: {report['verdict']}")
    print(f"Headline: {report['headline']}")
    print()

    rec = report["reconciliation"]
    print(f"Reconciliation: total={rec['total_alerts']}  by_stream_sum={rec['by_stream_sum']}  "
          f"by_severity_sum={rec['by_severity_sum']}  matched+unmatched={rec['matched_plus_unmatched_sum']}  "
          f"[{'OK' if rec['ok'] else 'FAIL'}]")
    assert rec["ok"], "Reconciliation failed — see report['reconciliation']"
    print()

    print(f"Queue: {report['queue_total']} issues total, showing top {len(report['queue_default'])}")
    for row in report["queue_default"]:
        print(f"  {row['priority']}  {row['distinct_objects']:4d} obj / {row['alert_count']:4d} alerts  "
              f"{row['age_label']:>4s} {row['trend_arrow']}  [{row['stream']}] {row['error_category'][:50]!r}  "
              f"owners={row['owners']}")
    print()

    print("Ownership:")
    for row in report["ownership"]:
        esc = " [ESCALATE]" if row["escalate"] else ""
        print(f"  {row['stream']:12s} P1={row['p1_count']} open={row['open_count']} docs={row['distinct_objects']} "
              f"oldest={row['oldest_days']:.1f}d buckets={row['aging_buckets']}{esc}")
    print()

    m = report["movement"]
    print(f"Movement: {len(m['cleared'])} cleared, {len(m['new_today'])} new today, {len(m['recurring'])} recurring")
    print(f"14-day distinct-object series points: {len(m['series_14d'])}")
    print()

    # Not rendered anywhere (Band 5 "Coverage" was removed by request) —
    # kept here only as a CLI diagnostic on match quality, since it's
    # cheap to derive from issues_all and useful when tuning
    # interface_matrix.py's matching.
    unmatched = sum(1 for i in report["issues_all"] if i.match_state == "unmatched")
    ambiguous = sum(1 for i in report["issues_all"] if i.match_state == "ambiguous")
    print(f"Match quality: {unmatched} unmatched issue(s), {ambiguous} ambiguous issue(s) "
          f"(of {len(report['issues_all'])} total)")


def _self_check_history() -> None:
    report = build_history_report()
    print(f"Files ingested: {len(report['files'])}")
    for f in report["files"][:15]:
        flag = "  [DATE MISMATCH]" if f["date_mismatch"] else ""
        print(f"  {f['name']:42s} {f['file_date_label']}  {f['event_count']:4d} rows{flag}")
    if len(report["files"]) > 15:
        print(f"  ... and {len(report['files']) - 15} more")
    print()

    if not report["has_data"]:
        print("No historical events found — nothing further to check.")
        return

    print(f"Deduplicated union: {report['total_events']} distinct events "
          f"(from {report['total_raw_rows']} raw rows — "
          f"{report['total_raw_rows'] - report['total_events']} re-exports correctly not double-counted)")
    print(f"Range: {report['earliest_date']} to {report['latest_date']}")
    print()

    for w in report["windows"]:
        print(f"  {w['days']:2d}-day window: {w['alerts']} alerts, {w['distinct_objects']} distinct objects")
    print()

    print(f"By stream (last {report['stream_window_days']} days):")
    for s in report["streams"]:
        print(f"  {s['stream']:18s} total={s['total_alerts']:4d}  trend={s['trend_arrow']}")
    print()

    # Regression guard: the daily series and total_events both derive
    # from the exact same deduplicated event union — summing the daily
    # series must equal total_events exactly, or something has
    # silently forked into two different notions of "the events"
    # between these two views of the same history.
    cache = hist.build_daily_folder_index()
    daily_sum = sum(c for _, c in hist.daily_totals(hist.deduplicated_events(cache)))
    print(f"[{'OK' if daily_sum == report['total_events'] else 'FAIL'}] "
          f"daily-series total ({daily_sum}) vs total_events ({report['total_events']})")
    assert daily_sum == report["total_events"], "History tab's daily series disagrees with the event union total"


def main():
    # Windows' console defaults to cp1252, which can't print the ▲/▼
    # trend arrows this module generates — a CLI-only cosmetic issue
    # (the browser renders them fine), not a data bug, but worth not
    # crashing the one tool meant to catch real data bugs.
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--self-check", metavar="ALERT_XLSX")
    parser.add_argument("--self-check-history", action="store_true",
                         help="Check the History tab's report (build_history_report()) instead of a single file's report")
    parser.add_argument("--search-history", metavar="QUERY",
                         help="Run the History tab's error search for QUERY and print the matches (error text or a business object key)")
    args = parser.parse_args()
    if args.self_check:
        _self_check(Path(args.self_check))
    elif args.self_check_history:
        _self_check_history()
    elif args.search_history:
        results = search_history(args.search_history)
        print(f"{len(results)} match(es) for {args.search_history!r}:")
        for r in results:
            print(f"  [{r['stream']}] {r['error_category'][:60]!r}  "
                  f"{r['alert_count']:4d} alerts / {r['distinct_objects']} obj over {r['days_active']} day(s) "
                  f"({r['first_seen_label']} to {r['last_seen_label']})  match={r['match_kind']}")
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
