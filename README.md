# OOT Monitoring — Alert Workbook Generator

A local web app that turns two daily Datadog alert exports (Technical +
Functional) into a single live, formula-driven Excel workbook — merging
the raw rows, classifying each error against a maintained "Match" rule
sheet, and carrying forward a real, refreshable pivot table — plus a
dashboard for aging/unresolved issues and a daily observations email
draft.

## Why

The manual process was: export two CSVs from a Datadog dashboard, copy
each into an existing Excel workbook by hand, drag formulas down,
manually classify any new error types, and eyeball a pivot table for a
daily status email. This automates all of that while keeping the output
a completely normal, editable `.xlsx` file — nothing here locks you into
a proprietary format or a hosted service.

## What it does

- **Generate**: upload today's Technical + Functional exports (`.csv` or
  `.xlsx`), pick the most recent day's workbook as a template (its Match
  sheet and pivot table carry forward), get back a new workbook with the
  same live formulas and a refreshable pivot — no static, pre-computed
  values.
- **View Table**: browse the generated Data sheet and the real pivot
  table (refreshed via Excel automation, not a reconstruction) right in
  the browser, with one-click copy that preserves Excel's actual
  formatting on paste.
- **Aging**: a dashboard (stat tiles, a ranked bar chart, a legend) of
  which alert categories are still open past a 48-hour window, based on
  how long each has been appearing in the data, plus the full detail
  table below it.
- **Email Draft**: auto-writes a draft of the daily observations email
  from the aging data, using one disclosed rule for "recurring vs. new
  spike" framing — always a draft to review, never auto-sent. A second,
  optional **AI-drafted email** sits alongside it: a local language
  model (via [Ollama](https://ollama.com), nothing sent over the
  network) writes the analytical commentary — which categories are
  worth calling out, spike/recurring/declining framing — while every
  number, date, and count stays computed by plain Python, never by the
  model. See "Optional: AI-drafted email" below.
- **Match Rules**: view, add, edit, and delete the classification rules
  directly in the browser instead of hand-editing the sheet.

## Why local, not hosted

Generation needs your local Datadog export files, and several features
use real Excel via COM automation (for a byte-for-byte accurate pivot
table and rich-formatted copy/paste) — both only make sense running on
your own machine. This is a local Flask server (binds to
`127.0.0.1` only); the browser tab is just the UI.

## Requirements

- Windows, with Microsoft Excel installed (for pivot refresh / rich copy
  — the rest of the app works without it, just with those features
  disabled)
- Python 3.11+
- `pip install flask openpyxl pywin32`

## Running it

Double-click `Launch Workbook Generator.bat`, or:

```
py webapp.py
```

then open `http://127.0.0.1:8765/`.

## Optional: AI-drafted email

The rule-based Email Draft always works with no setup. The AI-drafted
version next to it needs [Ollama](https://ollama.com) installed and
running locally, with a model pulled:

```
ollama pull qwen2.5:7b
```

Nothing about this feature calls out to an external API — it's a
plain HTTP request to `localhost:11434`, so today's alert data (error
text, business object keys, service names, internal URLs) never leaves
the machine. Without Ollama running, that one button just shows an
error; every other feature is unaffected. On a CPU-only machine expect
it to take anywhere from under a minute to several minutes, since it
makes several small, isolated model calls per email rather than one —
see `build_email_draft_ai`'s docstring in `generate_workbook.py` for
why.

## Files

- `generate_workbook.py` — all core logic: reading exports, merging,
  classification, building the live-formula Data sheet, the Excel-COM
  pivot/validation helpers, and the Match-sheet read/write primitives.
  Also runnable standalone from the CLI (`py generate_workbook.py --help`).
- `webapp.py` — the Flask routes; a thin layer over the above.
- `templates/`, `static/` — the UI.
