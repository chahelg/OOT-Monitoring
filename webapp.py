"""
Alert Monitoring Workbook Generator — local web app.

Run by double-clicking "Launch Workbook Generator.bat", or:
    py webapp.py
then open http://127.0.0.1:8765/ in a browser (the launcher does this for
you automatically).

This is a *local* Flask server, not a hosted service — it has to run on
your own PC because generation needs your local Datadog export files and
(optionally) Excel via COM automation, neither of which can move to a
remote machine. The browser tab is just the UI; all the work happens in
this same process. Binds to 127.0.0.1 only (not reachable from the
network) since it's a single-user local tool with no login.
"""

import os
import re
import shutil
import tempfile
import threading
import time
from pathlib import Path

from flask import Flask, redirect, render_template, request, send_from_directory, url_for
from werkzeug.utils import secure_filename

import generate_workbook as gw
import report_model

PORT = 8765

app = Flask(__name__)

# Filenames throughout this project routinely have spaces and parentheses
# ("active-alert category 6 Aug.xlsx", "technical-active-alert (4).csv") —
# werkzeug's secure_filename() strips those, which both breaks matching
# against real files on disk and mangles the established naming
# convention. This only strips characters Windows actually forbids, plus
# path traversal.
_UNSAFE_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')


def _safe_output_name(name: str) -> str:
    name = name.strip().replace("..", "")
    name = _UNSAFE_FILENAME_CHARS.sub("", name).strip(" .")
    return name or gw.default_output_name()


def _is_within_output_dir(path: Path) -> bool:
    try:
        path.resolve().relative_to(gw.OUTPUT_DIR.resolve())
        return True
    except (ValueError, OSError):
        return False


def _template_choices():
    return sorted(gw.OUTPUT_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)


@app.route("/", methods=["GET"])
def index():
    return render_template(
        "index.html",
        templates=_template_choices(),
        default_output_name=gw.default_output_name(),
        result=None,
    )


@app.route("/generate", methods=["POST"])
def generate():
    result = {"ok": False, "error": None, "log": [], "unmatched": []}

    tmp_dir = Path(tempfile.mkdtemp(prefix="workbook_upload_"))

    try:
        technical_file = request.files.get("technical")
        functional_file = request.files.get("functional")
        template_upload = request.files.get("template_upload")
        template_choice = request.form.get("template_choice", "")
        output_name = request.form.get("output_name", "").strip() or gw.default_output_name()
        do_validate = request.form.get("validate") == "on"

        if not technical_file or not technical_file.filename:
            raise ValueError("Please choose a Technical export file.")
        if not functional_file or not functional_file.filename:
            raise ValueError("Please choose a Functional export file.")

        technical_path = tmp_dir / secure_filename(technical_file.filename)
        functional_path = tmp_dir / secure_filename(functional_file.filename)
        technical_file.save(technical_path)
        functional_file.save(functional_path)

        if template_upload and template_upload.filename:
            template_path = tmp_dir / secure_filename(template_upload.filename)
            template_upload.save(template_path)
        elif template_choice:
            # Allowlist against the actual files we just offered in the
            # dropdown, rather than sanitizing — these are real on-disk
            # names (often with spaces) that must match exactly.
            known = {t.name: t for t in _template_choices()}
            if template_choice not in known:
                raise ValueError(f"Selected template file not found: {template_choice}")
            template_path = known[template_choice]
        else:
            raise ValueError("Please choose a template workbook (or upload one).")

        output_name = _safe_output_name(output_name)
        if not output_name.lower().endswith(".xlsx"):
            output_name += ".xlsx"
        output_path = gw.OUTPUT_DIR / output_name

        gen_result = gw.generate_workbook(technical_path, functional_path, template_path, output_path)
        result["ok"] = True
        result["row_count"] = gen_result["row_count"]
        result["technical_count"] = gen_result["technical_count"]
        result["functional_count"] = gen_result["functional_count"]
        result["unmatched"] = gen_result["unmatched"]
        result["output_path"] = str(output_path)
        result["output_name"] = output_path.name

        if do_validate:
            gw.validate_in_excel(output_path, result["log"].append)

    except Exception as e:
        result["ok"] = False
        result["error"] = str(e)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    return render_template(
        "index.html",
        templates=_template_choices(),
        default_output_name=gw.default_output_name(),
        result=result,
    )


def _resolve_file(filename: str):
    """Allowlist lookup against real files in OUTPUT_DIR (same pattern as
    /generate's template_choice) — falls back to the most recent output
    when no filename is given."""
    if filename:
        known = {t.name: t for t in _template_choices()}
        return known.get(filename)
    return gw.find_latest_template()


def _dedupe_output_name(name: str) -> Path:
    """Pick a destination for a newly-opened file under OUTPUT_DIR,
    never silently overwriting something already there — if the picked
    file's name collides with an existing one, append " (2)", " (3)",
    etc., the same way Windows Explorer avoids clobbering on copy."""
    dest = gw.OUTPUT_DIR / name
    if not dest.exists():
        return dest
    stem, suffix = dest.stem, dest.suffix
    n = 2
    while True:
        candidate = gw.OUTPUT_DIR / f"{stem} ({n}){suffix}"
        if not candidate.exists():
            return candidate
        n += 1


def _build_legacy_aging(path, rows=None):
    """Band 6's enrichment shape (compute_aging + the interface-matrix
    join) — shared by /view's aging tab and /export/health-report's
    full mode so they can't drift apart from computing it two
    different ways."""
    if rows is None:
        rows = gw.read_generated_rows(path)
    aging = gw.compute_aging(rows)
    interface_index = gw.load_interface_index()
    aging = gw.enrich_aging_with_interfaces(aging, rows, interface_index)
    return aging, interface_index is None


@app.route("/view/open", methods=["POST"])
def view_open():
    """Backs the View Table page's "Choose File…" button — lets the user
    pick any .xlsx from anywhere on disk via the native OS file dialog,
    rather than being limited to the dropdown of files already in
    OUTPUT_DIR. Browsers never expose a picked file's real path to the
    page (by design), so the only way to load it is to upload its bytes;
    this drops a copy into OUTPUT_DIR — same as manually copying the file
    there and picking it — which is also what makes every other tab
    (download, history, match rules) work on it unchanged."""
    tab = request.form.get("tab", "data")
    f = request.files.get("file")
    if not f or not f.filename:
        return redirect(url_for("view", tab=tab))

    name = _safe_output_name(f.filename)
    if not name.lower().endswith(".xlsx"):
        return redirect(url_for("view", tab=tab, open_error="Please choose an .xlsx workbook."))

    dest = _dedupe_output_name(name)
    f.save(dest)
    return redirect(url_for("view", file=dest.name, tab=tab))


@app.route("/view", methods=["GET"])
def view():
    filename = request.args.get("file", "")
    tab = request.args.get("tab", "data")
    path = _resolve_file(filename)

    error = request.args.get("open_error") or None
    rows = []
    excel_pivot = None
    aging = []
    email_draft = None
    interface_matrix_missing = False
    report = None
    history_report = None
    if error:
        pass
    elif path is None:
        error = "No generated workbook found yet — generate one first."
    else:
        try:
            rows = gw.read_generated_rows(path)
            # Cheap pure-Python — computed unconditionally so the Email
            # Draft tab (which needs it to build the text) can use it
            # without recomputation.
            aging = gw.compute_aging(rows)
            if tab == "excel_pivot":
                # Only actually opens Excel when this tab is requested —
                # every other tab stays instant and Excel-free.
                excel_pivot = gw.read_excel_pivot(path)
            elif tab == "email_draft":
                email_draft = gw.build_email_draft(rows, aging)
            elif tab == "aging":
                # report_model.build_report() is the single source of
                # truth for Bands 1-5 (per its own module docstring) —
                # this route just calls it and hands the dict to the
                # template, no computation here. Band 6 (the collapsed
                # legacy detail table) is unchanged and still needs the
                # old enrichment shape, so that's kept alongside it.
                report = report_model.build_report(path)
                aging, interface_matrix_missing = _build_legacy_aging(path, rows)
            elif tab == "history":
                # Independent of the selected file — build_history_report()
                # covers every canonical daily export in output/, not just
                # the one currently chosen in the dropdown above.
                history_report = report_model.build_history_report()
        except Exception as e:
            error = str(e)

    return render_template(
        "view.html",
        selected_file=path.name if path else "",
        tab=tab,
        rows=rows,
        report=report,
        history_report=history_report,
        excel_pivot=excel_pivot,
        aging=aging,
        email_draft=email_draft,
        interface_matrix_missing=interface_matrix_missing,
        error=error,
    )


@app.route("/view/copy-excel-pivot", methods=["POST"])
def copy_excel_pivot():
    filename = request.form.get("file", "")
    path = _resolve_file(filename)
    if path is None:
        return {"ok": False, "error": f"File not found: {filename}"}, 404
    return gw.copy_excel_pivot_to_clipboard(path)


@app.route("/view/generate-ai-email", methods=["POST"])
def generate_ai_email():
    filename = request.form.get("file", "")
    path = _resolve_file(filename)
    if path is None:
        return {"ok": False, "error": f"File not found: {filename}"}, 404
    try:
        rows = gw.read_generated_rows(path)
        aging = gw.compute_aging(rows)
        text = gw.build_email_draft_ai(rows, aging)
        return {"ok": True, "text": text}
    except Exception as e:
        return {"ok": False, "error": str(e)}, 500


def _serialize_alert_row(r: dict) -> dict:
    """One raw alert row -> the flat JSON shape the drill-down modal
    (view.html's ERROR_MODAL_COLUMNS) renders — shared by the Aging
    tab's single-file category drill-down and the History tab's
    cross-file error-search drill-down so the two can't render the
    same kind of row two different ways."""
    return {
        "timestamp": r["timestamp_readable"].strftime("%Y-%m-%d %H:%M:%S"),
        "grouping": r["grouping"],
        "service": r["SERVICE"] or "",
        "application_name": r["APPLICATION_NAME"] or "",
        "business_object_key": r["BUSINESS_OBJECT_KEY"] or "",
        "sequence": r["SEQUENCE"] or "",
        "transaction_url": r["TRANSACTION_URL"] or "",
        "alert_message": r["ALERT_MESSAGE"] or "",
        "error_message": r["ERROR_MESSAGE"] or "",
        "error_details": r["ERROR_DETAILS"] or "",
    }


@app.route("/view/category-errors", methods=["POST"])
def category_errors():
    """Powers the Aging tab's "view errors" drill-down: every raw alert
    row for one category, in full (no truncation — that's the point of
    this view, unlike the Data tab's compact cells). Returns JSON since
    it's fetched into a modal, not a page navigation."""
    filename = request.form.get("file", "")
    category = request.form.get("category", "")
    path = _resolve_file(filename)
    if path is None:
        return {"ok": False, "error": f"File not found: {filename}"}, 404
    try:
        rows = gw.read_generated_rows(path)
        cat_rows = [r for r in rows if r["alert_type"] == category]
        cat_rows.sort(key=lambda r: r["timestamp_readable"], reverse=True)
        serialized = [_serialize_alert_row(r) for r in cat_rows]
        return {"ok": True, "alert_type": category, "count": len(serialized), "rows": serialized}
    except Exception as e:
        return {"ok": False, "error": str(e)}, 500


@app.route("/view/history-search", methods=["POST"])
def history_search():
    """Powers the History tab's error search box: every error category
    (across ALL daily folders, not just one file) whose name or whose
    underlying alert rows (business object key, messages, application,
    service, transaction URL) match the query text. See
    report_model.search_history()."""
    query = request.form.get("query", "")
    try:
        results = report_model.search_history(query)
        return {"ok": True, "query": query, "count": len(results), "results": results}
    except Exception as e:
        return {"ok": False, "error": str(e)}, 500


@app.route("/view/history-error-details", methods=["POST"])
def history_error_details():
    """Row-level drill-down for one History-search result: every raw
    occurrence of that (stream, application, error_category) across the
    full multi-day history. Same response shape as /view/category-errors
    so the existing modal JS renders either without a fork."""
    stream = request.form.get("stream", "")
    application = request.form.get("application", "")
    category = request.form.get("error_category", "")
    try:
        detail = report_model.history_error_details(stream, application, category)
        if detail is None:
            return {"ok": False, "error": "That error's history has changed since this search — try searching again."}, 404
        serialized = [_serialize_alert_row(r) for r in detail["rows"]]
        return {"ok": True, "alert_type": category, "count": len(serialized), "rows": serialized}
    except Exception as e:
        return {"ok": False, "error": str(e)}, 500


@app.route("/export/health-report", methods=["GET"])
def export_health_report():
    """A single self-contained, printable/copyable rendering of the
    exact same report_model.build_report() dict the live Aging tab
    uses — see that module's docstring: "the live page and the export
    both render that same dict through different templates, so they
    can never disagree." The only transform applied here and nowhere
    else is report_model.apply_contact_policy(), which redacts emails
    out of the contact lists by default — this page is meant to leave
    the machine (printed, PDF'd, pasted into an email), so it defaults
    to the more conservative PII posture; the live page never applies
    it, since it never leaves this machine.

    mode=full (default) includes Band 6, the legacy raw detail table;
    mode=brief stops after Band 5 — a shorter management summary.
    contacts=full opts back into showing raw contact strings (which
    may include email addresses) instead of names only.
    """
    filename = request.args.get("file", "")
    mode = request.args.get("mode", "full")
    if mode not in ("full", "brief"):
        mode = "full"
    contacts_full = request.args.get("contacts", "") == "full"
    path = _resolve_file(filename)

    error = None
    report = None
    aging = None
    if path is None:
        error = "No generated workbook found yet — generate one first."
    else:
        try:
            report = report_model.build_report(path)
            report = report_model.apply_contact_policy(report, full=contacts_full)
            if mode == "full":
                aging, _ = _build_legacy_aging(path)
                if not contacts_full:
                    # Band 6's legacy contact list lives outside the
                    # report dict apply_contact_policy() covers, so the
                    # same names-only default is applied here too —
                    # otherwise "contacts=full not requested" would be
                    # true everywhere on the page except this one table.
                    aging = [
                        {**cat, "interface": {**cat["interface"], "contacts": report_model.redact_names(cat["interface"].get("contacts", []))}}
                        for cat in aging
                    ]
        except Exception as e:
            error = str(e)

    return render_template(
        "export_health_report.html",
        selected_file=path.name if path else filename,
        mode=mode,
        contacts_full=contacts_full,
        report=report,
        aging=aging,
        error=error,
    )


@app.route("/match", methods=["GET"])
def match():
    filename = request.args.get("file", "")
    path = _resolve_file(filename)

    error = None
    rules = []
    if path is None:
        error = "No generated workbook found yet — generate one first."
    else:
        try:
            rules = gw.load_match_rules_full(path)
        except Exception as e:
            error = str(e)

    return render_template(
        "match.html",
        templates=_template_choices(),
        selected_file=path.name if path else "",
        rules=rules,
        draft=request.args.get("draft", ""),
        error=error,
        saved=request.args.get("saved") == "1",
    )


@app.route("/match/save", methods=["POST"])
def match_save():
    filename = request.form.get("file", "")
    path = _resolve_file(filename)
    if path is None:
        return redirect(url_for("match", file=filename))

    # Rows are submitted with indexed field names (string_a_<id> etc.)
    # rather than array-style, because plain HTML checkboxes only submit
    # when checked — array positions would drift out of sync between
    # rows the moment one row's box is unchecked. `row_ids` (kept in sync
    # by the page's JS as rows are added/removed) is the authoritative
    # list of which indices are actually present.
    row_ids = [i for i in request.form.get("row_ids", "").split(",") if i]
    rules = []
    for i in row_ids:
        string_a = request.form.get(f"string_a_{i}", "").strip()
        string_b = request.form.get(f"string_b_{i}", "").strip()
        result = request.form.get(f"result_{i}", "").strip()
        if not string_a and not string_b and not result:
            continue
        rules.append({
            "string_a": string_a,
            "string_b": string_b,
            "result": result,
            "check": 1 if request.form.get(f"check_{i}") == "on" else None,
        })

    try:
        gw.set_match_rules(path, rules)
        return redirect(url_for("match", file=path.name, saved="1"))
    except Exception as e:
        return render_template(
            "match.html",
            templates=_template_choices(),
            selected_file=path.name,
            rules=gw.load_match_rules_full(path),
            draft="",
            error=str(e),
            saved=False,
        )


@app.route("/download/<path:filename>")
def download(filename):
    return send_from_directory(gw.OUTPUT_DIR, filename, as_attachment=True)


@app.route("/open-file", methods=["POST"])
def open_file():
    path = Path(request.form.get("path", ""))
    if _is_within_output_dir(path) and path.exists():
        os.startfile(str(path))
        return {"ok": True}
    return {"ok": False}, 400


@app.route("/open-folder", methods=["POST"])
def open_folder():
    path = Path(request.form.get("path", ""))
    if _is_within_output_dir(path):
        os.startfile(str(path.parent))
        return {"ok": True}
    return {"ok": False}, 400


@app.route("/shutdown", methods=["POST"])
def shutdown():
    def _stop():
        time.sleep(0.3)
        os._exit(0)

    threading.Thread(target=_stop, daemon=True).start()
    return render_template("shutdown.html")


if __name__ == "__main__":
    print(f"Open http://127.0.0.1:{PORT}/ in your browser.")
    # threaded=True: the AI email draft (/view/generate-ai-email) can take
    # a minute or more running a local LLM — without this, that single
    # request would block every other tab/request in the app until it
    # finishes. Safe here: the Excel-COM work (the one thing that ever
    # needed single-process isolation) already runs in its own subprocess
    # per call (see generate_workbook._run_excel_worker), so it's
    # unaffected by Flask's threading model either way.
    app.run(host="127.0.0.1", port=PORT, debug=False, threaded=True)
