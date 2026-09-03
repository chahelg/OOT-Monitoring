"""
history.py — builds a true multi-day event history from the daily
workbooks in output/, for trend, cleared/new/recurring detection, and
14/30-day totals. Pure, Flask-free, independently runnable:

    py history.py --self-check

Each daily export is a ROLLING multi-day window, so the files overlap
heavily — reading them as a stack of independent snapshots and summing
would massively over-count. Instead this builds one deduplicated event
set, keyed by (business_object_key, error_category, timestamp) as
specified, across every file, then derives everything else from that
single union.

File discovery is filtered to the app's own canonical output name
(`generate_workbook.default_output_name()`'s pattern,
"active-alert category <day> <Mon>.xlsx") — output/ was found on
inspection to also contain test/trial files and at least one file
("new active-alert category 17 Aug.xlsx") already confirmed corrupted
earlier this session (data silently written to the wrong sheet).
Reading those into a management trend report would either crash or
quietly poison it with garbage events, so anything not matching the
canonical name is skipped rather than guessed about.

Caching exception, deliberate: reading every daily workbook on every
page load is too slow for a live page, so this is the one place in the
app that keeps state between requests (everywhere else recomputes from
disk every time, on purpose). output/.history_index.json is keyed per
file by (mtime, size); only changed or new files are re-read. A
missing or corrupt index just triggers a full rebuild — it's a cache,
never a second source of truth, so "restart the server" (or "delete
the index") stays a complete fix for anything going wrong with it.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
import tempfile
from collections import defaultdict
from dataclasses import asdict
from datetime import date, datetime, timedelta
from pathlib import Path

import generate_workbook as gw
import interface_matrix as im
import severity as sv

OUTPUT_DIR = gw.OUTPUT_DIR
HISTORY_INDEX_PATH = OUTPUT_DIR / ".history_index.json"

# The History tab reads from a SECOND, separate source: the user's own
# daily archive folders next to this app's own folder — one per day,
# named "<day><ordinal> <Mon>" (e.g. "19th Aug"), each holding one
# workbook named identically ("19th Aug/19th Aug.xlsx"). Confirmed by
# direct inspection to be the exact same generated-workbook shape as
# output/'s own files (same sheet names, read_generated_rows() returns
# identical rows for a day present in both) — but this folder set is
# typically MORE complete, since a day only lands in output/ once
# someone actually clicks Generate on it, while the daily folder gets
# populated every day regardless. Deliberately a separate index/cache
# from HISTORY_INDEX_PATH (see build_daily_folder_index()) so this
# only affects the History tab — the Aging tab's trend/movement bands
# keep using the output/-only index unchanged.
DAILY_FOLDERS_DIR = gw.APP_DIR.parent
DAILY_FOLDERS_INDEX_PATH = OUTPUT_DIR / ".history_daily_folders_index.json"

# Matches generate_workbook.default_output_name()'s own pattern exactly
# — see module docstring for why this filter exists at all.
_CANONICAL_NAME_RE = re.compile(r"^active-alert category (\d{1,2}) (\w{3})\.xlsx$", re.IGNORECASE)

# "19th Aug", "5th Aug", "21st Aug" — the user's daily-folder naming.
# Any ordinal suffix is accepted regardless of the number (no attempt
# to validate "21st" vs "21th") — being lenient here is safe; a
# genuinely misnamed folder just won't match and is silently skipped,
# same as any other non-canonical name in this module.
_DAILY_FOLDER_RE = re.compile(r"^(\d{1,2})(?:st|nd|rd|th) (\w{3})$", re.IGNORECASE)

_EVENT_FIELDS = (
    "TIMESTAMP", "SERVICE", "APPLICATION_NAME", "BUSINESS_OBJECT_KEY", "SEQUENCE",
    "TRANSACTION_URL", "ALERT_MESSAGE", "ERROR_MESSAGE", "ERROR_DETAILS",
    "alert_type", "grouping", "check",
)


def discover_daily_files(output_dir: Path = OUTPUT_DIR) -> list[Path]:
    """Every file in output/ matching the canonical daily-export name,
    sorted. Anything else (test files, "new"/"trial" duplicates,
    already-known-corrupted attempts) is silently excluded here, not
    hidden later — see module docstring."""
    if not output_dir.exists():
        return []
    return sorted(p for p in output_dir.glob("*.xlsx") if _CANONICAL_NAME_RE.match(p.name))


def discover_daily_folder_files(root_dir: Path = DAILY_FOLDERS_DIR) -> list[Path]:
    """Every "<day><ordinal> <Mon>/<day><ordinal> <Mon>.xlsx" folder
    directly under root_dir — see DAILY_FOLDERS_DIR above. A folder
    matching the naming pattern but without its own same-named .xlsx
    inside (still mid-export, or genuinely empty) is silently skipped,
    same "don't guess about it" policy as discover_daily_files()."""
    if not root_dir.exists():
        return []
    found = []
    for entry in root_dir.iterdir():
        if not entry.is_dir() or not _DAILY_FOLDER_RE.match(entry.name):
            continue
        candidate = entry / f"{entry.name}.xlsx"
        if candidate.exists():
            found.append(candidate)
    return sorted(found)


def _resolve_day_month(day: int, mon: str, reference: date | None) -> date | None:
    """Shared by date_from_filename() and _date_from_daily_folder_name()
    — both filenames carry a day+month but no year, resolved the same
    way: relative to `reference` (defaults to today), preferring last
    year if the naive current-year reading would land more than 60
    days in the future (handles a December file still being read in
    January without every caller having to think about it)."""
    reference = reference or date.today()
    for year in (reference.year, reference.year - 1):
        try:
            candidate = datetime.strptime(f"{day} {mon} {year}", "%d %b %Y").date()
        except ValueError:
            continue
        if (candidate - reference).days <= 60:
            return candidate
    return None


def date_from_filename(name: str, reference: date | None = None) -> date | None:
    """"active-alert category 19 Aug.xlsx" -> 2026-08-19. Public (not
    module-private) since report_model.py's History tab reuses it to
    flag filename/data date disagreement per file without re-reading
    the file itself."""
    m = _CANONICAL_NAME_RE.match(name)
    return _resolve_day_month(int(m.group(1)), m.group(2), reference) if m else None


def _date_from_daily_folder_name(name: str, reference: date | None = None) -> date | None:
    """"19th Aug.xlsx" -> 2026-08-19 (also matches the bare folder name
    "19th Aug", so it works on either)."""
    m = _DAILY_FOLDER_RE.match(Path(name).stem)
    return _resolve_day_month(int(m.group(1)), m.group(2), reference) if m else None


def _infer_filename_date(name: str, reference: date | None = None) -> date | None:
    """Tries every daily-file naming convention this app deals with —
    the canonical output/ name and the user's own daily-folder name —
    returning whichever one actually matches."""
    return date_from_filename(name, reference) or _date_from_daily_folder_name(name, reference)


def _file_signature(path: Path) -> list:
    st = path.stat()
    return [st.st_mtime, st.st_size]


def _read_rows_tolerating_lock(path: Path) -> list[dict]:
    """gw.read_generated_rows(path), but falls back to a throwaway copy
    if Excel currently has the file open — the exact scenario for
    *today's* daily file, which the user routinely has open while
    working, and a very plausible reason a file this module expects to
    find would otherwise silently vanish from the History tab every
    single day. Same "copy first, never touch the real file" workaround
    generate_workbook.py's Excel-COM path already uses for this same
    Windows file-locking quirk (see its own module for the precedent)."""
    try:
        return gw.read_generated_rows(path)
    except PermissionError:
        tmp_dir = Path(tempfile.mkdtemp(prefix="history_read_"))
        try:
            tmp_copy = tmp_dir / path.name
            shutil.copy2(path, tmp_copy)
            return gw.read_generated_rows(tmp_copy)
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)


def _extract_file_events(path: Path) -> tuple[date | None, list[dict]]:
    """Reads one workbook and returns (file_date, events). file_date
    prefers the max real TIMESTAMP found inside the file over the
    filename's date, logging when they disagree — the filename is
    just what someone typed when saving, the data inside is the fact."""
    rows = _read_rows_tolerating_lock(path)
    events = []
    max_ts_date = None
    for r in rows:
        if not r.get("alert_type"):
            continue
        ts_date = r["timestamp_readable"].date()
        if max_ts_date is None or ts_date > max_ts_date:
            max_ts_date = ts_date
        event = {k: r.get(k) for k in _EVENT_FIELDS}
        event["timestamp_readable"] = r["timestamp_readable"].isoformat()
        events.append(event)

    filename_date = _infer_filename_date(path.name)
    file_date = max_ts_date or filename_date
    if filename_date and max_ts_date and filename_date != max_ts_date:
        print(f"[history] {path.name}: filename implies {filename_date}, data's latest "
              f"TIMESTAMP is {max_ts_date} — trusting the data.", file=sys.stderr)
    return file_date, events


def _load_cache(index_path: Path) -> dict:
    if not index_path.exists():
        return {}
    try:
        with open(index_path, encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict) or "files" not in data:
            raise ValueError("unexpected shape")
        return data
    except Exception as e:
        print(f"[history] index at {index_path} missing or corrupt ({e}) — rebuilding from scratch.",
              file=sys.stderr)
        return {}


def _save_cache(index_path: Path, cache: dict) -> None:
    tmp = index_path.with_suffix(".json.tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cache, f)
    tmp.replace(index_path)  # atomic swap, same pattern generate_workbook uses for output files


def build_history_index(
    files: list[Path] | None = None,
    index_path: Path = HISTORY_INDEX_PATH,
) -> dict:
    """The cached, per-file event index — rebuilds only files that are
    new or have changed (by mtime+size) since it was last written, and
    drops entries for files that no longer exist. Returns the raw
    {"files": {path: {...}}} structure; use deduplicated_events() to
    get the actual cross-file union most callers want.

    `files` defaults to discover_daily_files(OUTPUT_DIR) (the Aging
    tab's source) but accepts an explicit list so a caller with a
    different discovery source — see build_daily_folder_index() below,
    the History tab's own index — can reuse this same caching logic
    without duplicating it."""
    if files is None:
        files = discover_daily_files(OUTPUT_DIR)
    cache = _load_cache(index_path)
    files_cache = cache.setdefault("files", {})

    current_files = files
    current_paths = {str(p) for p in current_files}

    for stale in set(files_cache) - current_paths:
        del files_cache[stale]

    changed = 0
    for path in current_files:
        key = str(path)
        sig = _file_signature(path)
        entry = files_cache.get(key)
        if entry is not None and entry.get("signature") == sig:
            continue
        try:
            file_date, events = _extract_file_events(path)
        except Exception as e:
            print(f"[history] skipping {path.name}: failed to read ({e})", file=sys.stderr)
            files_cache.pop(key, None)
            continue
        files_cache[key] = {
            "signature": sig,
            "file_date": file_date.isoformat() if file_date else None,
            "events": events,
        }
        changed += 1

    if changed:
        _save_cache(index_path, cache)
    return cache


def build_daily_folder_index(
    root_dir: Path = DAILY_FOLDERS_DIR,
    index_path: Path = DAILY_FOLDERS_INDEX_PATH,
) -> dict:
    """The History tab's own file index, sourced from the user's daily
    archive folders (DAILY_FOLDERS_DIR) rather than output/. Kept as a
    genuinely separate cache from build_history_index()'s default —
    not just a different discovery list reusing the same cache file —
    so the Aging tab's trend/movement bands (which call
    build_history_index() with no arguments) are completely unaffected
    by this; only the History tab calls this function."""
    return build_history_index(files=discover_daily_folder_files(root_dir), index_path=index_path)


def _event_key(event: dict) -> tuple:
    return (event.get("BUSINESS_OBJECT_KEY"), event.get("alert_type"), event.get("TIMESTAMP"))


def deduplicated_events(cache: dict) -> list[dict]:
    """The true cross-file union: every distinct real alert event, once,
    regardless of how many overlapping daily exports it appeared in.
    Keyed by (business_object_key, error_category, timestamp) as
    specified — two events with the same key are the same real alert
    re-exported, not two different alerts."""
    merged: dict[tuple, dict] = {}
    for entry in cache.get("files", {}).values():
        for event in entry["events"]:
            merged[_event_key(event)] = event
    return list(merged.values())


def events_to_rows(events: list[dict]) -> list[dict]:
    """Converts cached event dicts back into the row shape
    severity.group_issues()/interface_matrix expect (real datetime
    object instead of an ISO string, same field names)."""
    rows = []
    for e in events:
        row = dict(e)
        row["timestamp_readable"] = datetime.fromisoformat(e["timestamp_readable"])
        rows.append(row)
    return rows


# --------------------------------------------------------------------
# Derived views over the deduplicated union
# --------------------------------------------------------------------

def daily_totals(events: list[dict]) -> list[tuple[date, int]]:
    counts: dict[date, int] = defaultdict(int)
    for e in events:
        counts[datetime.fromisoformat(e["timestamp_readable"]).date()] += 1
    return sorted(counts.items())


def daily_distinct_objects(events: list[dict]) -> list[tuple[date, int]]:
    """Distinct BUSINESS_OBJECT_KEY count per day — the "documents
    affected" measure used everywhere else in this app (see
    severity._distinct_object_count), computed once here so the
    History tab doesn't reimplement the "a blank key is still its own
    distinct unknown object, never merged with other blanks" rule
    separately and risk it drifting out of sync."""
    by_day: dict[date, set] = defaultdict(set)
    unkeyed_by_day: dict[date, int] = defaultdict(int)
    for e in events:
        d = datetime.fromisoformat(e["timestamp_readable"]).date()
        key = e.get("BUSINESS_OBJECT_KEY")
        if key:
            by_day[d].add(key)
        else:
            unkeyed_by_day[d] += 1
    days = set(by_day) | set(unkeyed_by_day)
    return sorted((d, len(by_day.get(d, set())) + unkeyed_by_day.get(d, 0)) for d in days)


def window_totals(events: list[dict], days: int, latest_date: date | None = None) -> dict:
    """Distinct objects + alert count in the trailing `days`-day window,
    ending at `latest_date` (defaults to the union's own true latest
    day — never derived per-caller, same anti-drift principle as the
    single-file "today" fix elsewhere in this app)."""
    if latest_date is None:
        latest_date = max((datetime.fromisoformat(e["timestamp_readable"]).date() for e in events), default=None)
    if latest_date is None:
        return {"alerts": 0, "distinct_objects": 0, "window_start": None, "window_end": None}
    window_start = latest_date - timedelta(days=days - 1)
    in_window = [e for e in events if window_start <= datetime.fromisoformat(e["timestamp_readable"]).date() <= latest_date]
    objects = {e.get("BUSINESS_OBJECT_KEY") for e in in_window if e.get("BUSINESS_OBJECT_KEY")}
    unkeyed = sum(1 for e in in_window if not e.get("BUSINESS_OBJECT_KEY"))
    return {
        "alerts": len(in_window),
        "distinct_objects": len(objects) + unkeyed,
        "window_start": window_start,
        "window_end": latest_date,
    }


def stream_series(events: list[dict], interface_index: dict | None, days: int = 14, latest_date: date | None = None) -> dict[str, list[tuple[date, int]]]:
    """Per-stream daily alert counts over the trailing `days` days —
    resolves each event's stream the same way severity.group_issues()
    does (single deterministic stream per event, never fanned out)."""
    if latest_date is None:
        latest_date = max((datetime.fromisoformat(e["timestamp_readable"]).date() for e in events), default=None)
    if latest_date is None:
        return {}
    window_start = latest_date - timedelta(days=days - 1)

    series: dict[str, dict[date, int]] = defaultdict(lambda: defaultdict(int))
    for e in events:
        d = datetime.fromisoformat(e["timestamp_readable"]).date()
        if not (window_start <= d <= latest_date):
            continue
        if interface_index is None:
            stream = "Unavailable"
        else:
            result = im.resolve_alert_interface(e.get("APPLICATION_NAME"), interface_index, e)
            # Same "label by actual application name, not a generic
            # bucket" rule as severity.group_issues() — kept identical
            # here so this module's own --self-check output can't
            # disagree with what the live report shows.
            stream = result.stream or e.get("APPLICATION_NAME") or im.STREAM_UNMAPPED
        series[stream][d] += 1

    return {stream: sorted(day_counts.items()) for stream, day_counts in series.items()}


def _issue_key(issue: sv.Issue) -> tuple:
    return (issue.stream, issue.application, issue.error_category)


def _trend_from_day_counts(day_counts: list[tuple[date, int]], latest_date: date) -> int:
    """+1 rising / 0 flat / -1 falling, comparing the most recent 3
    days of activity against the preceding 3 days. Deliberately a
    plain count comparison, not magnitude-thresholded — matches the
    brief's plain "rising/flat/falling" framing; retune here if it
    turns out too noisy on low-volume issues once more real days
    exist. Shared by compute_trends() (per issue) and stream_trends()
    (per stream) so the two can't disagree on what "trending" means."""
    recent_start = latest_date - timedelta(days=2)
    prior_start = latest_date - timedelta(days=5)
    prior_end = latest_date - timedelta(days=3)
    recent = sum(c for d, c in day_counts if recent_start <= d <= latest_date)
    prior = sum(c for d, c in day_counts if prior_start <= d <= prior_end)
    if recent > prior:
        return 1
    if recent < prior:
        return -1
    return 0


def compute_trends(issues: list["sv.Issue"], latest_date: date | None = None) -> dict[tuple, int]:
    """+1 rising / 0 flat / -1 falling per issue — this is what
    replaces severity.group_issues()'s Phase-2 placeholder (always 0)
    once real multi-day history is available."""
    if latest_date is None:
        latest_date = max((i.last_seen.date() for i in issues), default=None)
    if latest_date is None:
        return {}
    return {_issue_key(issue): _trend_from_day_counts(issue.day_counts, latest_date) for issue in issues}


def stream_trends(series: dict[str, list[tuple[date, int]]], latest_date: date | None) -> dict[str, int]:
    """Same +1/0/-1 rising/flat/falling idea as compute_trends(), one
    level up — per stream instead of per issue, from stream_series()'s
    output. Powers the History tab's by-stream trend column."""
    if latest_date is None:
        return {}
    return {stream: _trend_from_day_counts(day_counts, latest_date) for stream, day_counts in series.items()}


def _has_gap(day_counts: list[tuple[date, int]]) -> bool:
    """True if there's at least one wholly-inactive day strictly
    between this issue's first and last occurrence — i.e. it went
    quiet and then came back, not just a steady run of consecutive
    days."""
    if len(day_counts) < 2:
        return False
    active_dates = {d for d, _ in day_counts}
    span_start, span_end = day_counts[0][0], day_counts[-1][0]
    d = span_start + timedelta(days=1)
    while d < span_end:
        if d not in active_dates:
            return True
        d += timedelta(days=1)
    return False


def movement_summary(issues: list["sv.Issue"], latest_date: date | None = None, recent_window_days: int = 2) -> dict:
    """The three lists Band 4 needs: cleared since yesterday, new
    today, and recurring (cleared at some point, now back). All three
    are computed off the SAME issues (built from the deduplicated
    union, not a single file), so an issue can only ever land in one
    "new" bucket and one "cleared" bucket per day — no double-counting
    across these lists is possible by construction."""
    if latest_date is None:
        latest_date = max((i.last_seen.date() for i in issues), default=None)
    result = {"cleared_since_yesterday": [], "new_today": [], "recurring": []}
    if latest_date is None:
        return result

    yesterday = latest_date - timedelta(days=1)
    for issue in issues:
        active_dates = {d for d, _ in issue.day_counts}
        if yesterday in active_dates and latest_date not in active_dates:
            result["cleared_since_yesterday"].append(issue)
        if issue.first_seen.date() == latest_date:
            result["new_today"].append(issue)
        if _has_gap(issue.day_counts) and (latest_date - issue.last_seen.date()).days <= recent_window_days:
            result["recurring"].append(issue)

    return result


# --------------------------------------------------------------------
# Self-check CLI
# --------------------------------------------------------------------

def _self_check() -> None:
    print(f"Discovering daily files in {OUTPUT_DIR} ...")
    files = discover_daily_files()
    for p in files:
        print(f"  {p.name}  -> filename date: {date_from_filename(p.name)}")
    print(f"{len(files)} canonical daily file(s) found.")
    print()

    cache = build_history_index()
    for path_str, entry in cache["files"].items():
        print(f"  {Path(path_str).name}: file_date={entry['file_date']}  {len(entry['events'])} events")
    print()

    events = deduplicated_events(cache)
    total_raw = sum(len(e["events"]) for e in cache["files"].values())
    print(f"Deduplicated union: {len(events)} distinct events (from {total_raw} raw rows across all files "
          f"— {total_raw - len(events)} were re-exports of the same real event, correctly not double-counted)")
    print()

    interface_index = im.load_interface_index()
    rows = events_to_rows(events)
    issues = sv.group_issues(rows, interface_index)
    latest_date = max((i.last_seen.date() for i in issues), default=None)
    print(f"Union's true latest day: {latest_date}")

    trends = compute_trends(issues, latest_date)
    trend_counts = {1: 0, 0: 0, -1: 0}
    for v in trends.values():
        trend_counts[v] += 1
    print(f"Trend distribution across {len(issues)} issues: rising={trend_counts[1]} "
          f"flat={trend_counts[0]} falling={trend_counts[-1]}")
    print()

    movement = movement_summary(issues, latest_date)
    print(f"Cleared since yesterday: {len(movement['cleared_since_yesterday'])}")
    for i in movement["cleared_since_yesterday"][:10]:
        print(f"    [{i.stream}] {i.error_category[:60]!r}  (last seen {i.last_seen.date()})")
    print(f"New today: {len(movement['new_today'])}")
    for i in movement["new_today"][:10]:
        print(f"    [{i.stream}] {i.error_category[:60]!r}")
    print(f"Recurring (gap then back): {len(movement['recurring'])}")
    for i in movement["recurring"][:10]:
        print(f"    [{i.stream}] {i.error_category[:60]!r}  days active: {[d.isoformat() for d, _ in i.day_counts]}")
    print()

    for days in (14, 30):
        w = window_totals(events, days, latest_date)
        print(f"{days}-day window ({w['window_start']} to {w['window_end']}): "
              f"{w['alerts']} alerts, {w['distinct_objects']} distinct objects")
    print()

    series = stream_series(events, interface_index, days=14, latest_date=latest_date)
    print("14-day per-stream series:")
    for stream, day_counts in sorted(series.items(), key=lambda kv: -sum(c for _, c in kv[1])):
        total = sum(c for _, c in day_counts)
        print(f"  {stream:12s} total={total:4d}  {day_counts}")

    # --- Regression check: every view that has an idea of "today" must
    # agree, whether it derives that date itself or has it passed in —
    # this is the exact class of bug fixed earlier this session (the
    # old aging page's spike detection anchoring to a per-channel
    # "latest day" instead of the file's real one). Deriving it four
    # independent ways here and asserting they match is what actually
    # catches a regression, not just asserting the value looks plausible.
    independently_derived = {
        latest_date,
        max((d for d, _ in daily_totals(events)), default=None),
        window_totals(events, 1)["window_end"],
        max((i.last_seen.date() for i in sv.group_issues(rows, interface_index)), default=None),
    }
    assert len(independently_derived) == 1, (
        f"latest-day anchor disagrees across views: {independently_derived}"
    )
    print(f"[OK] latest-day anchor agrees everywhere it's derived independently: {latest_date}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--self-check", action="store_true", help="Build the history index and print discovery, dedup, trend, and movement diagnostics")
    args = parser.parse_args()

    if args.self_check:
        _self_check()
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
