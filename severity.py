"""
severity.py — groups alerts into Issues (the real unit of analysis for
management, not raw alert rows) and scores each one into a P1/P2/P3
priority. Pure, Flask-free, independently runnable:

    py severity.py --self-check "output\\active-alert category 19 Aug.xlsx"

Why "Issue" instead of "alert row": a category with 168 rows against 3
stuck documents is a retry storm on 3 documents, not 168 separate
problems — that distinction (documents vs. alerts) is the single
biggest readability fix available on the old alert-count-ranked page,
and it's why every issue below carries both numbers.

Trend is a placeholder here (see Issue.trend / group_issues): real
multi-day trend needs history.py (Phase 3), which doesn't exist yet.
score_issue() takes trend as a plain parameter — it doesn't know or
care where it came from — so Phase 3 only has to change the value
group_issues() passes in, not this module's scoring logic.
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass, field
from datetime import date, datetime
from math import log1p
from pathlib import Path

import generate_workbook as gw
import interface_matrix as im

# --------------------------------------------------------------------
# Tunable weights — retune here, not in the scoring logic below.
# --------------------------------------------------------------------
WEIGHTS = {
    "CRIT": 2,    # criticality_rank() is 1-4, so this term contributes 2-8
    "FI": 5,      # FI posting impact is binary — weighted heavily on purpose
    "STOCK": 2,
    "TIME": 3,
    "AGE": 4,     # scaled 0-1 (age_days / 7, capped), so this term contributes 0-4
    "BLAST": 2,   # 2 * log1p(distinct_objects) — blast radius
    "VOL": 1,     # 1 * log1p(alert_count) — deliberately small, log not linear:
                  # raw volume must not dominate (that's the old page's problem)
    "TREND": 2,   # -2 falling / 0 flat / +2 rising
}

# Score buckets — tunable alongside the weights above. Calibrated by
# running --self-check against a real file and eyeballing the top-20
# ordering (per the brief's own verification method for this phase;
# there's no theoretically "correct" cutoff, only a plausible one to
# refine against real days as they come in).
#
# CRIT started at 3 (max 12) with thresholds at 20/12 — on real data
# that put 11 of 21 issues in P1, more than half, because one dominant
# application (ASTRO-WMS-BENE) is rated "Critical" and that single
# app-level property alone contributed over half the P1 threshold to
# nearly every issue tied to it, regardless of that specific issue's
# actual size. Criticality is a property of the *application*, shared
# across many differently-sized issues from it — it shouldn't alone be
# enough to nearly clear the bar. Dropping CRIT's weight and raising
# both thresholds produced a real triage shape on the same file (3
# P1 / 10 P2 / 8 P3) where P1 tracks genuinely large or override-
# qualified issues, not just "which app is this."
P1_THRESHOLD = 24.0
P2_THRESHOLD = 14.0

_CRITICALITY_RANK = {"Critical": 4, "High": 3, "Medium": 2, "Low": 1}
_UNKNOWN_CRITICALITY_RANK = 2  # neutral, NOT 0 — see criticality_rank()'s docstring


def criticality_rank(*values: str | None) -> float:
    """Combines interface/application/business criticality into one
    rank, 1-4. Takes the MAX of whichever of the (up to 3) values are
    known — if any one dimension says "Critical", the issue is treated
    as critical, rather than averaging it away. If all are unknown,
    returns a neutral middle value (2), never 0 — the brief is explicit
    that unknown enrichment must not silently score as zero-risk; see
    Issue.enrichment_confidence for the separate flag that tells a
    reader *when* a score rests on unknowns like this."""
    known = [_CRITICALITY_RANK[v] for v in values if v in _CRITICALITY_RANK]
    return float(max(known)) if known else float(_UNKNOWN_CRITICALITY_RANK)


@dataclass
class ScoreResult:
    score: float
    priority: str  # "P1" | "P2" | "P3"
    reasons: list[str] = field(default_factory=list)  # override rules that fired, if any


def score_issue(
    interface_criticality: str | None,
    application_criticality: str | None,
    business_criticality: str | None,
    fi_posting_impact: bool | None,
    stock_movement: bool | None,
    time_critical: bool | None,
    age_days: float,
    distinct_objects: int,
    alert_count: int,
    trend: int,  # -1 falling / 0 flat / +1 rising
) -> ScoreResult:
    crit = criticality_rank(interface_criticality, application_criticality, business_criticality)
    score = (
        WEIGHTS["CRIT"] * crit
        + WEIGHTS["FI"] * (1 if fi_posting_impact else 0)
        + WEIGHTS["STOCK"] * (1 if stock_movement else 0)
        + WEIGHTS["TIME"] * (1 if time_critical else 0)
        + WEIGHTS["AGE"] * min(age_days / 7, 1)
        + WEIGHTS["BLAST"] * log1p(max(distinct_objects, 0))
        + WEIGHTS["VOL"] * log1p(max(alert_count, 0))
        + WEIGHTS["TREND"] * trend
    )

    if score >= P1_THRESHOLD:
        priority = "P1"
    elif score >= P2_THRESHOLD:
        priority = "P2"
    else:
        priority = "P3"

    reasons = []
    # Override rules — these can only ever escalate, never downgrade
    # what the score already produced.
    if fi_posting_impact and age_days > 2 and priority == "P3":
        priority = "P2"
        reasons.append("FI posting impact + age > 2 days (min P2)")
    if interface_criticality == "High" and age_days > 7 and priority != "P1":
        priority = "P1"
        reasons.append("Interface Criticality = High + age > 7 days (forced P1)")

    return ScoreResult(score=round(score, 1), priority=priority, reasons=reasons)


# --------------------------------------------------------------------
# Issue grouping
# --------------------------------------------------------------------

@dataclass
class Issue:
    stream: str
    application: str | None
    error_category: str
    match_state: str  # "exact" | "inferred" | "ambiguous" | "unmatched"
    interface_record: "im.InterfaceRecord | None"
    also_streams: list[str]

    alert_count: int
    distinct_objects: int
    first_seen: datetime
    last_seen: datetime
    age_days: float
    alerts_today: int
    alerts_last_7d: int
    day_counts: list[tuple[date, int]]  # ascending, for the sparkline

    trend: int  # -1/0/+1 — Phase 2 placeholder, always 0 (see module docstring)
    enrichment_confidence: str  # "unknown" | "partial" | "full"

    score: float
    priority: str
    priority_reasons: list[str]

    sample_rows: list[dict] = field(default_factory=list)  # raw alert rows for drilldown


_KEY_ENRICHMENT_FIELDS = (
    "interface_criticality", "application_criticality", "business_criticality",
    "fi_posting_impact", "time_critical",
)


def _enrichment_confidence(record: "im.InterfaceRecord | None") -> str:
    if record is None:
        return "unknown"
    missing = sum(1 for f in _KEY_ENRICHMENT_FIELDS if getattr(record, f) is None)
    if missing >= 2:
        return "partial"
    return "full"


def _distinct_object_count(keys: list) -> int:
    """Distinct BUSINESS_OBJECT_KEY values, treating each blank/missing
    key as its OWN distinct unknown object rather than collapsing them
    into one shared bucket — merging blanks together would understate
    how many real documents are actually affected. Not hit in practice
    on real data checked so far (0 blanks), but wrong-by-default here
    would be a silent undercount, so it's handled explicitly rather
    than assumed away."""
    named = {k for k in keys if k}
    blank_count = sum(1 for k in keys if not k)
    return len(named) + blank_count


def group_issues(
    rows: list[dict],
    interface_index: dict | None,
    now: datetime | None = None,
    trend_overrides: dict[tuple, int] | None = None,
) -> list[Issue]:
    """Groups alert rows into Issues — (stream, application,
    error_category) — the unit a manager actually assigns to someone,
    instead of the raw alert-row list the old page ranked on.

    Every row is resolved to exactly one stream via
    interface_matrix.resolve_alert_interface (never fanned out — see
    that module for why), so summing alert_count across every returned
    Issue always equals len(rows) exactly. --self-check asserts this.

    `trend_overrides` maps (stream, application, error_category) -> -1/
    0/+1, i.e. history.compute_trends()'s output — when given, replaces
    the plain 0 (flat) placeholder this function used alone in Phase 2.
    Kept optional and additive rather than required: callers with only
    a single file's rows (no multi-day history available) still get a
    correct — just less informed — score with trend simply flat.
    """
    if now is None:
        now = datetime.now()
    trend_overrides = trend_overrides or {}

    # "Today" and the 7-day window are anchored to the file's own true
    # latest day, computed once and shared across every issue — the
    # same fix already applied to the old aging page's spike detection
    # (a per-issue "latest day" derived only from that issue's own rows
    # would misjudge activity that just happens to predate the file's
    # actual latest day as if it were current).
    dated_rows = [r for r in rows if r.get("alert_type") and r.get("timestamp_readable")]
    latest_overall_date = max((r["timestamp_readable"].date() for r in dated_rows), default=None)

    groups: dict[tuple, dict] = {}
    for r in dated_rows:
        if interface_index is None:
            # Matrix not loaded at all — every alert is grouped, just
            # with no interface/criticality/contact data attached.
            # Distinct from "unmatched" (matrix loaded, this particular
            # application just isn't in it) — report_model.py needs to
            # tell these apart to show the right banner.
            match_state, record, also_streams, stream = "unavailable", None, [], "Unavailable"
        else:
            result = im.resolve_alert_interface(r.get("APPLICATION_NAME"), interface_index, r)
            match_state, record, also_streams = result.state, result.record, result.also_streams
            # An unmatched alert (application not in the interface
            # matrix at all — in practice MULESOFT or "Intransit from
            # SAP") is labeled by that actual application name, not a
            # generic "Unmapped" bucket that would lump unrelated
            # applications together — by request, and matching the
            # same idea generate_workbook.build_stream_channels already
            # used for the old channel-cards view. Only falls back to
            # the generic label in the (currently never-seen) case
            # where even the application name itself is blank.
            stream = result.stream or r.get("APPLICATION_NAME") or im.STREAM_UNMAPPED

        key = (stream, r.get("APPLICATION_NAME"), r["alert_type"])
        g = groups.setdefault(key, {
            "rows": [],
            "match_state": match_state,
            "record": record,
            "also_streams": also_streams,
        })
        g["rows"].append(r)

    issues = []
    for (stream, application, category), g in groups.items():
        grows = g["rows"]
        timestamps = [r["timestamp_readable"] for r in grows]
        first_seen, last_seen = min(timestamps), max(timestamps)
        age_days = (now - first_seen).total_seconds() / 86400

        day_counter: dict[date, int] = {}
        for r in grows:
            d = r["timestamp_readable"].date()
            day_counter[d] = day_counter.get(d, 0) + 1
        day_counts = sorted(day_counter.items())

        alerts_today = day_counter.get(latest_overall_date, 0) if latest_overall_date else 0
        if latest_overall_date:
            window_start = date.fromordinal(latest_overall_date.toordinal() - 6)
            alerts_last_7d = sum(c for d, c in day_counts if d >= window_start)
        else:
            alerts_last_7d = len(grows)

        distinct_objects = _distinct_object_count([r.get("BUSINESS_OBJECT_KEY") for r in grows])
        confidence = _enrichment_confidence(g["record"])
        trend = trend_overrides.get((stream, application, category), 0)

        rec = g["record"]
        result = score_issue(
            interface_criticality=rec.interface_criticality if rec else None,
            application_criticality=rec.application_criticality if rec else None,
            business_criticality=rec.business_criticality if rec else None,
            fi_posting_impact=rec.fi_posting_impact if rec else None,
            stock_movement=rec.stock_movement if rec else None,
            time_critical=rec.time_critical if rec else None,
            age_days=age_days,
            distinct_objects=distinct_objects,
            alert_count=len(grows),
            trend=trend,
        )

        issues.append(Issue(
            stream=stream,
            application=application,
            error_category=category,
            match_state=g["match_state"],
            interface_record=rec,
            also_streams=g["also_streams"],
            alert_count=len(grows),
            distinct_objects=distinct_objects,
            first_seen=first_seen,
            last_seen=last_seen,
            age_days=age_days,
            alerts_today=alerts_today,
            alerts_last_7d=alerts_last_7d,
            day_counts=day_counts,
            trend=trend,
            enrichment_confidence=confidence,
            score=result.score,
            priority=result.priority,
            priority_reasons=result.reasons,
            sample_rows=grows,
        ))

    issues.sort(key=lambda i: -i.score)
    return issues


# --------------------------------------------------------------------
# Self-check CLI
# --------------------------------------------------------------------

def _self_check(alert_file: Path, matrix_path: Path) -> None:
    rows = gw.read_generated_rows(alert_file)
    interface_index = im.load_interface_index(matrix_path)
    print(f"Alert file: {alert_file}  ({len(rows)} rows)")
    print(f"Interface matrix: {'loaded' if interface_index else 'UNAVAILABLE'}")
    print()

    issues = group_issues(rows, interface_index)

    # --- Reconciliation: sum(per-issue alert_count) must equal file total ---
    total_from_issues = sum(i.alert_count for i in issues)
    real_total = sum(1 for r in rows if r.get("alert_type"))
    ok = total_from_issues == real_total
    print(f"Reconciliation: sum(issue alert_count)={total_from_issues} vs real total={real_total}  "
          f"[{'OK' if ok else 'FAIL'}]")
    assert ok, "Issue grouping does not reconcile to the file total — see group_issues()"

    priority_counts = {"P1": 0, "P2": 0, "P3": 0}
    for i in issues:
        priority_counts[i.priority] += 1
    print(f"Priorities: {priority_counts}")
    print()

    print(f"Top 20 issues by score ({len(issues)} total):")
    for i in issues[:20]:
        override = f"  [{'; '.join(i.priority_reasons)}]" if i.priority_reasons else ""
        print(f"  {i.priority}  score={i.score:5.1f}  {i.distinct_objects:3d} obj / {i.alert_count:4d} alerts  "
              f"conf={i.enrichment_confidence:7s}  [{i.stream}] {i.error_category[:55]!r}{override}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--self-check", metavar="ALERT_XLSX", help="Group issues and print reconciliation + a scored top-20")
    parser.add_argument("--matrix", default=str(im.DEFAULT_MATRIX_PATH))
    args = parser.parse_args()

    if args.self_check:
        _self_check(Path(args.self_check), Path(args.matrix))
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
