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


@app.route("/view", methods=["GET"])
def view():
    filename = request.args.get("file", "")
    tab = request.args.get("tab", "data")
    path = _resolve_file(filename)

    error = None
    rows = []
    excel_pivot = None
    aging = []
    aging_stats = None
    email_draft = None
    if path is None:
        error = "No generated workbook found yet — generate one first."
    else:
        try:
            rows = gw.read_generated_rows(path)
            # Cheap pure-Python — computed unconditionally so both the
            # Aging tab and the Email Draft tab (which needs it to build
            # the text) can use it without recomputation.
            aging = gw.compute_aging(rows)
            if tab == "excel_pivot":
                # Only actually opens Excel when this tab is requested —
                # every other tab stays instant and Excel-free.
                excel_pivot = gw.read_excel_pivot(path)
            elif tab == "email_draft":
                email_draft = gw.build_email_draft(rows, aging)
            elif tab == "aging":
                aging_stats = gw.compute_aging_stats(aging)
        except Exception as e:
            error = str(e)

    return render_template(
        "view.html",
        templates=_template_choices(),
        selected_file=path.name if path else "",
        tab=tab,
        rows=rows,
        aging_stats=aging_stats,
        excel_pivot=excel_pivot,
        aging=aging,
        email_draft=email_draft,
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
