"""
checks/render_aging_tab.py — permanent regression check for the live
Aging and History tab pages (webapp.py's /view route + templates/
view.html) plus /export/health-report. Unlike report_model.py/
severity.py/history.py/interface_matrix.py, webapp.py is deliberately
kept thin (routes only, no --self-check CLI of its own — see
report_model.py's module docstring: "webapp.py stays thin"). This
script is the equivalent check for the one layer those modules' CLIs
can't cover: does the page actually render, band by band, both with
the interface matrix present and with it missing (graceful
degradation), without a live server or a browser.

Deliberately kept as a real, rerunnable file (not written-and-deleted)
per the project's own rule against throwaway verification scripts —
run it after any change to webapp.py or templates/view.html:

    py checks/render_aging_tab.py [ALERT_XLSX]

Uses Flask's in-process test client — never binds a port, so it's safe
to run even while the real app is up on 8765.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(APP_DIR))

import webapp  # noqa: E402


REQUIRED_BAND_MARKERS = [
    "verdict-sentence",   # Band 1: verdict + headline
    "queue-table",        # Band 2: priority queue
    "ownership-table",    # Band 3: ownership
    "Movement",           # Band 4: movement
    # Band 5 (coverage) was removed by request — no marker for it.
    "Full detail table",  # Band 6 (renumbered from the brief's Band 6): legacy appendix, unchanged
    "error-modal-overlay",  # drilldown modal, still wired
]


def _default_file() -> str:
    templates = sorted(webapp.gw.OUTPUT_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not templates:
        raise SystemExit(f"No .xlsx files found in {webapp.gw.OUTPUT_DIR} to test against.")
    return templates[0].name


def _body_content_only(html: str) -> str:
    """Strips the inline <style> block before scanning for a leaked
    '@' — the export page's own CSS legitimately contains @media/@page
    rules, which aren't PII and would otherwise false-positive the
    names-only check below."""
    return re.sub(r"<style\b.*?</style>", "", html, flags=re.DOTALL | re.IGNORECASE)


def check(file_name: str) -> None:
    client = webapp.app.test_client()

    print(f"File: {file_name}")
    print()

    print("=== tab=aging, interface matrix PRESENT ===")
    resp = client.get(f"/view?file={file_name}&tab=aging")
    assert resp.status_code == 200, f"expected 200, got {resp.status_code}"
    body = resp.get_data(as_text=True)
    for marker in REQUIRED_BAND_MARKERS:
        assert marker in body, f"MISSING band marker: {marker!r}"
    print(f"  [OK] status 200, all {len(REQUIRED_BAND_MARKERS)} band markers present ({len(body)} bytes)")
    print()

    print("=== tab=aging, interface matrix ABSENT (simulated) ===")
    # Same degrade-gracefully path the real app takes if
    # reference/Oxygen_Interface_Matrix_w1.xlsx is ever missing — forced
    # here via a nonexistent matrix_path rather than by touching the
    # real reference file.
    orig_build_report = webapp.report_model.build_report

    def _fake_build_report(path, matrix_path=None, now=None):
        return orig_build_report(path, matrix_path=Path("__nonexistent_matrix_for_check__.xlsx"), now=now)

    webapp.report_model.build_report = _fake_build_report
    try:
        resp2 = client.get(f"/view?file={file_name}&tab=aging")
        assert resp2.status_code == 200, f"expected 200, got {resp2.status_code}"
        body2 = resp2.get_data(as_text=True)
        assert "Interface matrix not loaded" in body2, "missing degradation banner"
        for marker in REQUIRED_BAND_MARKERS[:4]:  # Bands 1-4 must still render without the matrix
            assert marker in body2, f"MISSING band marker in degraded mode: {marker!r}"
        print(f"  [OK] status 200, degradation banner shown, Bands 1-4 still render ({len(body2)} bytes)")
    finally:
        webapp.report_model.build_report = orig_build_report
    print()

    print("=== regression: other tabs / routes unaffected ===")
    for tab in ("data", "email_draft"):  # excel_pivot needs Excel installed via COM; skipped here
        r = client.get(f"/view?file={file_name}&tab={tab}")
        assert r.status_code == 200, f"tab={tab} expected 200, got {r.status_code}"
        print(f"  [OK] tab={tab}: 200")
    for route in ("/", "/match"):
        r = client.get(route)
        assert r.status_code == 200, f"{route} expected 200, got {r.status_code}"
        print(f"  [OK] {route}: 200")

    print()

    print("=== /export/health-report ===")
    # full mode, names-only (default) — must not contain a raw '@'
    # anywhere in the rendered contact lists.
    r = client.get(f"/export/health-report?file={file_name}&mode=full")
    assert r.status_code == 200, f"export full mode expected 200, got {r.status_code}"
    body = r.get_data(as_text=True)
    for marker in ["export-content", "Priority queue", "Ownership", "Movement", "Full detail table"]:
        assert marker in body, f"MISSING in export (full): {marker!r}"
    assert "@" not in _body_content_only(body), "names-only (default) export must not leak an email address anywhere on the page"
    print(f"  [OK] mode=full: 200, all bands + Band 6 present, no '@' leaked ({len(body)} bytes)")

    # brief mode must NOT include Band 6's legacy table
    r = client.get(f"/export/health-report?file={file_name}&mode=brief")
    assert r.status_code == 200, f"export brief mode expected 200, got {r.status_code}"
    body_brief = r.get_data(as_text=True)
    assert "Full detail table" not in body_brief, "brief mode should not include Band 6"
    for marker in ["export-content", "Priority queue", "Ownership", "Movement"]:
        assert marker in body_brief, f"MISSING in export (brief): {marker!r}"
    print(f"  [OK] mode=brief: 200, Bands 1-4 present, Band 6 correctly absent ({len(body_brief)} bytes)")

    # contacts=full opts back into raw contact strings (may include
    # emails); default (names-only) must never leak an '@' into the
    # rendered owners/contacts columns.
    r = client.get(f"/export/health-report?file={file_name}&mode=full&contacts=full")
    assert r.status_code == 200
    body_full_contacts = r.get_data(as_text=True)
    note = "" if "@" in _body_content_only(body_full_contacts) else " (no email-shaped contact in this file's matrix rows to prove it with — not a failure)"
    print(f"  [OK] contacts=full: 200 ({len(body_full_contacts)} bytes){note}")

    # unresolvable file -> a clean error render, not a 500
    r = client.get("/export/health-report?file=__does_not_exist__.xlsx&mode=full")
    assert r.status_code == 200, f"export with missing file expected a clean 200 error page, got {r.status_code}"
    assert "Can't build the report" in r.get_data(as_text=True)
    print("  [OK] missing file: renders the error card instead of crashing")

    print()

    print("=== tab=history ===")
    r = client.get(f"/view?file={file_name}&tab=history")
    assert r.status_code == 200, f"tab=history expected 200, got {r.status_code}"
    body = r.get_data(as_text=True)
    for marker in ["Daily volume", "By stream", "Files ingested", "daily-volume-chart", "hist-bar"]:
        assert marker in body, f"MISSING in History tab: {marker!r}"
    print(f"  [OK] tab=history: 200, all sections present ({len(body)} bytes)")

    # interface matrix absent -> degrade, don't crash (By stream still
    # renders, just labeled "Unavailable" per severity.group_issues()'s
    # own handling of a missing index).
    orig_build_history_report = webapp.report_model.build_history_report

    def _fake_build_history_report(matrix_path=None):
        return orig_build_history_report(matrix_path=Path("__nonexistent_matrix_for_check__.xlsx"))

    webapp.report_model.build_history_report = _fake_build_history_report
    try:
        r2 = client.get(f"/view?file={file_name}&tab=history")
        assert r2.status_code == 200, f"tab=history (no matrix) expected 200, got {r2.status_code}"
        body2 = r2.get_data(as_text=True)
        assert "Interface matrix not loaded" in body2, "missing degradation banner on History tab"
        print(f"  [OK] tab=history, matrix absent: 200, degradation banner shown ({len(body2)} bytes)")
    finally:
        webapp.report_model.build_history_report = orig_build_history_report

    print()
    print("ALL CHECKS PASSED")


def main() -> None:
    file_name = sys.argv[1] if len(sys.argv) > 1 else _default_file()
    check(file_name)


if __name__ == "__main__":
    main()
