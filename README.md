# Re-Auth Report Builder

Automated Excel/VBA report builder for **weekly pending and upcoming patient re-authorizations**.

This project converts a raw re-auth export (CSV or Excel) into a formatted workbook with a KPI **Dashboard**, one **manager-specific tab** per Assigned Sales Manager (including unassigned records), and a full **Raw Data** tab. It also includes unattended run scripts and realtime logging/error-capture tooling to make report generation and troubleshooting reliable.

- **Repository:** https://github.com/smd3235/Re-Auth-Project
- **Primary template:** `Re-Auth Template.xlsm`
- **Main macro:** `CreateReAuthReport` (button: **CREATE RE-AUTH REPORT** on `START HERE`)
- **Sample input:** `reauth raw.csv`
- **Report output folder:** `Reports\`

---

## Table of contents

- [Purpose](#purpose)
- [Quick start](#quick-start)
  - [End users (Excel)](#end-users-excel)
  - [Maintainers / unattended run (Python)](#maintainers--unattended-run-python)
- [Input data](#input-data)
  - [Supported file types](#supported-file-types)
  - [Expected headers / columns](#expected-headers--columns)
  - [Data handling notes](#data-handling-notes)
- [Output report specification](#output-report-specification)
  - [Dashboard](#dashboard)
  - [Manager tabs](#manager-tabs)
  - [Raw Data](#raw-data)
  - [Styling and visual cues](#styling-and-visual-cues)
- [Repository layout](#repository-layout)
- [VBA architecture](#vba-architecture)
- [Scripts and tooling](#scripts-and-tooling)
  - [Sync VBA into template](#sync-vba-into-template)
  - [Run report (robust harness)](#run-report-robust-harness)
  - [Answer dialogs](#answer-dialogs)
  - [Compile checks](#compile-checks)
  - [View/reorder logs](#viewreorder-logs)
- [Logging, diagnostics, and troubleshooting](#logging-diagnostics-and-troubleshooting)
  - [Log files](#log-files)
  - [Common issues](#common-issues)
  - [How to investigate a failed run](#how-to-investigate-a-failed-run)
- [Configuration notes](#configuration-notes)
- [Development workflow](#development-workflow)
- [GitHub / version control notes](#github--version-control-notes)
- [Known limitations](#known-limitations)

---

## Purpose

The weekly re-auth workflow needs a fast, consistent way to review **pending and upcoming patient re-authorizations** (and related authorization statuses) by manager and by sales rep (Wound Care Consultant).

This builder:

1. Reads the raw re-auth dataset.
2. Normalizes key date fields.
3. Groups records by **Assigned Sales Manager** (blank manager -> `Unassigned`).
4. Produces a clean, formatted report workbook with summary KPIs, manager-level breakdowns, and per-manager working tabs.
5. Captures step-by-step runtime logs and completion/error markers for auditability and troubleshooting.

---

## Quick start

### End users (Excel)

1. Open **`Re-Auth Template.xlsm`** in Microsoft Excel (Windows).
2. Enable macros/content if prompted (the macro must run).
3. Click **`CREATE RE-AUTH REPORT`** on the **`START HERE`** sheet.
4. When prompted, select the raw re-auth file (`.csv`, `.xlsx`, `.xlsm`, or `.xls`).
5. Wait for processing to finish. The generated report is saved to:
   - `Reports\ReAuth_Report_YYYY-mm-dd_HHMM.xlsx`

> Tip: If Excel blocks macros or shows security dialogs, allow macros for this file/location or place the project in a trusted location.

### Maintainers / unattended run (Python)

Prerequisites (Windows):
- Microsoft Excel installed (COM automation required)
- Python 3.x with `pywin32` (`win32com`) available

Typical unattended run:

```powershell
cd "C:\Users\smd32\Downloads\Re-Auth Project"
python .\scripts\run-report.py --source ".\reauth raw.csv"
```

What the harness does:
- Copies the template to a temp workbook, sets auto-source for VBA, and launches Excel via COM
- Runs `CreateReAuthReport` unattended
- Auto-answers common Excel/VBA/file dialogs and enforces a watchdog timeout
- Writes detailed Python + VBA logs and checks for the completion marker/output file
- Prints `NEW_REPORT=...` and `OUTCOME=SUCCESS|...`

To reorder logs newest-first for review:

```powershell
python .\scripts\view-logs.py
```

---

## Input data

### Supported file types

The report can read:
- **CSV:** `.csv`
- **Excel:** `.xlsx`, `.xlsm`, `.xls`

Blank cells are normal and are preserved as blanks where applicable.

### Expected headers / columns

The macro detects the header row by looking for key columns. At minimum it requires:
- `Assigned Sales Manager`
- `Auth Status`
- `Patient Name`
- `Re-auth Due Date`

The sample/raw schema used by this project includes (original order in `reauth raw.csv`):

| Column | Purpose / notes |
|---|---|
| `Wound Care Consultant` | Sales rep (primary rep grouping on manager tabs) |
| `Associate WCC` | Associate rep |
| `Re-auth Due Date` | Key due date (normalized/formatted) |
| `HC Company` | Home care/company |
| `Patient Name` | Patient identifier/name in report context |
| `Patient/Packet Status` | Patient/packet state |
| `Task Required` | Required action/task |
| `Auth Start Date` | Authorization start (normalized/formatted) |
| `Auth Status` | Drives KPI/status grouping & highlighting (e.g. Past Due, Upcoming, Denied) |
| `Authorization Status` | Authorization workflow status |
| `Authorization Sub Status` | Sub-status/detail |
| `Pending SO Amount` | Amount field (kept in Raw Data) |
| `Initial Authorization Denial Date` | Denial date (normalized in Raw Data/source handling) |
| `Authorization Denial Reason` | Denial reason |
| `Denial Reason Sub List` | Denial sub-reason |
| `Sales Order RX Qty Requested` | Qty requested |
| `Sales Order RX Qty Approved` | Qty approved |
| `Treatment Physician` | Physician |
| `Treatment Physician NPI Number` | NPI |
| `Assigned Sales Manager` | Manager grouping for tabs/dashboard |

> Manager tabs intentionally show a curated subset/reordered columns (see [Manager tabs](#manager-tabs)). **Raw Data** retains all original columns in original order.

### Data handling notes

- **Header detection:** Scans near the top of the sheet to locate required headers.
- **Manager grouping:** Records are grouped by `Assigned Sales Manager`; empty manager -> **`Unassigned`** tab/group.
- **Date normalization:** Key date columns (e.g. `Re-auth Due Date`, `Auth Start Date`, and denial date where present) are normalized to Excel dates and formatted `m/d/yyyy` when parseable.
- **Status bucketing (dashboard/summary):** `Auth Status` is evaluated case-insensitively:
  - contains `past due` -> **Past Due**
  - contains `upcoming` -> **Upcoming**
  - contains `denied` -> **Denied**
  - otherwise -> **Other**
- **CSV caveat:** Standard comma-delimited exports work. Malformed CSVs (wrong delimiter/encoding) may import poorly; if that happens, verify the file opens correctly in Excel or re-export as `.xlsx`.

---

## Output report specification

Each run creates: `Reports\ReAuth_Report_YYYY-mm-dd_HHMM.xlsx`

Sheets (in order created/rendered):
1. `Dashboard`
2. One tab per manager (sorted; includes `Unassigned` when present)
3. `Raw Data`

### Dashboard

A high-level weekly view of pending/upcoming re-auth workload:

- **Title:** “RE-AUTHORIZATION DASHBOARD” with generation timestamp and source filename
- **KPI cards:**
  - `TOTAL RE-AUTHS`
  - `PAST DUE`
  - `UPCOMING`
  - `DENIED`
  - `OTHER`
- **Manager Breakdown table:**
  - `Assigned Sales Manager`
  - `Total`
  - `Past Due`
  - `% Past Due` (conditional formatting/highlighting for elevated rates)
  - `Tab` (hyperlink labeled “Open” to the manager’s sheet)
- Frozen header area for easy scanning; styled for readability

### Manager tabs

One worksheet per `Assigned Sales Manager` (safe/truncated sheet names as needed). Includes **`Unassigned`** when applicable.

**Layout:**
- Title block: `RE-AUTH | <MANAGER NAME>`, total re-auth count, generated timestamp
- **`RE-AUTHS BY SALES REP` summary** (merged block, rows 5–9): counts per `Wound Care Consultant` within that manager’s tab (e.g. `Rep Name: N | ...`)
- **Data table header at row 15**, data starts at row 16, **frozen below header** (`A16`) with autofilter

**Manager tab columns (curated, in order):**
1. `Wound Care Consultant` (Sales Rep)
2. `Associate WCC`
3. `Patient Name`
4. `HC Company`
5. `Re-auth Due Date`
6. `Auth Start Date`
7. `Auth Status`
8. `Authorization Status`
9. `Task Required`
10. `Patient/Packet Status`
11. `Treatment Physician`

**Intentionally excluded from manager tabs** (still available in Raw Data): `Assigned Sales Manager` (implied by tab), `Treatment Physician NPI Number`, `Pending SO Amount`, `Sales Order RX Qty Requested`, `Sales Order RX Qty Approved`, `Authorization Sub Status`, `Initial Authorization Denial Date`, `Authorization Denial Reason`, `Denial Reason Sub List`.

### Raw Data

- Contains **all records** from the source (all rows) with **all original columns in original order**
- Header formatting, borders, filters, frozen header row, and normalized date formatting on key date columns
- Use this tab for full-fidelity review/export and to access fields not shown on manager tabs

### Styling and visual cues

- Consistent **Segoe UI** typography, bordered tables, alternating row shading, wrapped headers, and AutoFit column widths (clamped for readability)
- **Status highlighting** (manager tabs): `Auth Status` values such as **Past Due / Upcoming / Denied** are color-highlighted for quick triage
- Dashboard KPI cards and manager breakdown use contrasting fills/conditional emphasis for at-a-glance review

---

## Repository layout

```text
Re-Auth Project/
├─ Re-Auth Template.xlsm        # Master macro-enabled template (START HERE + VBA)
├─ reauth raw.csv               # Sample/raw input data (example)
├─ modCore.bas                  # VBA: main build workflow (CreateReAuthReport)
├─ modDeclarations.bas          # VBA: constants, headers, paths, shared state
├─ modFormatting.bas            # VBA: dashboard/manager/raw formatting & highlights
├─ modLogging.bas               # VBA: realtime logging, step/error logging, marker
├─ modUtils.bas                 # VBA: helpers (headers, dates, arrays, output path, etc.)
├─ scripts/
│  ├─ sync_template.py          # Import mod*.bas into template, set button/macro
│  ├─ run-report.py             # Robust unattended run (watchdog, dialogs, logs)
│  ├─ answer-dialogs.py         # Standalone Excel/VBA dialog auto-answerer
│  ├─ watch-compile-error.py    # Capture/dismiss VBA compile/runtime dialogs
│  ├─ test-compile.py           # VBA compile probe harness
│  └─ view-logs.py              # Reorder log/*.log|*.txt newest-first
├─ log/                         # Runtime logs, markers, diagnostics (generated)
└─ Reports/                     # Generated report workbooks (.xlsx)
```

> You may also see local debug/temp artifacts (e.g. `_run_temp.xlsm`, `_dbg*.xlsm`, `_check*.py`, `_inspect_report.py`) created during development/troubleshooting. These are not required for normal use and can be excluded from commits.

---

## VBA architecture

| Module | Responsibility |
|---|---|
| `modDeclarations.bas` | Public constants (sheet/header names, manager header row, project/log/marker paths), shared runtime state |
| `modCore.bas` | Entry point `CreateReAuthReport`: source selection/open, header/data parsing, grouping/counts, workbook/sheet creation, save/cleanup, full step logging and error handling |
| `modFormatting.bas` | `FormatDashboard`, `FormatManagerSheet`, `BuildRepTotalsSection`, `FormatRawData`, KPI cards, alignment/widths, status highlighting |
| `modLogging.bas` | `LogMsg`, `LogStep`, `LogErr`, per-run log path, completion marker writer, log directory handling |
| `modUtils.bas` | Header/column discovery, date normalization/parsing, safe sheet names, sorting, array builders, output path, unused-row/col hiding, file picker helpers |

**Entry point:** `CreateReAuthReport()` (wired to the `CREATE RE-AUTH REPORT` button on `START HERE`).

---

## Scripts and tooling

All scripts live in `scripts\` and assume the project root is the parent of `scripts\` unless overridden by arguments.

### Sync VBA into template

Re-imports `mod*.bas` modules into `Re-Auth Template.xlsm`, refreshes the report button (`btnCreateReAuthReport` -> `CreateReAuthReport`), and removes stale module/button artifacts.

```powershell
python .\scripts\sync_template.py
# optional:
python .\scripts\sync_template.py --template "C:\path\to\Re-Auth Template.xlsm" --bas-dir "C:\path\to\project"
```

Run this after any `.bas` change.

### Run report (robust harness)

Unattended end-to-end execution with realtime logging and recovery aids.

```powershell
python .\scripts\run-report.py
python .\scripts\run-report.py --source ".\reauth raw.csv" --max-run-seconds 90
```

Common options:
- `--template` : path to template `.xlsm`
- `--source` : raw input `.csv/.xlsx/.xlsm/.xls`
- `--temp-copy` : temp workbook path (default `_run_temp.xlsm` in project root)
- `--max-run-seconds` : overall timeout/watchdog cap

Behavior highlights:
- Writes auto-source for VBA (`log\auto-source.txt`) to avoid interactive picker
- Starts dialog auto-answer/watchdog threads
- Captures VB components, macro return/result, completion marker, new report detection
- Appends this run’s VBA `macro-debug.log` tail and completion marker to the run log

### Answer dialogs

Standalone helper that watches for Excel/VBA dialogs and auto-answers/dismisses them (useful during manual or custom automated runs).

```powershell
python .\scripts\answer-dialogs.py --source ".\reauth raw.csv" --seconds 300
```

Logs to `log\answer-dialogs.log`.

### Compile checks

**Compile probe (recommended before/after VBA edits):**

```powershell
python .\scripts\test-compile.py
```

- Copies template, injects a probe macro, runs it to force full-project compile
- Uses `watch-compile-error.py` to capture/dismiss compile/runtime dialogs
- Writes `log\compile-test.log` and `COMPILE_OK=1|0`

**Compile/runtime dialog watcher:**

```powershell
python .\scripts\watch-compile-error.py
```

- Monitors Excel/VBA error dialogs, logs text to `log\watcher.log`, attempts dismissal

### View/reorder logs

Rewrites runtime log/text files newest-first with a dated header for faster review:

```powershell
python .\scripts\view-logs.py
# optional: --log-dir "C:\path\to\log"
```

---

## Logging, diagnostics, and troubleshooting

### Log files

| File | Description |
|---|---|
| `log\macro-debug.log` | Always-on VBA log (all runs appended) with timestamps and step/error entries |
| `log\reauth_YYYY-mm-dd_HHMMSS.log` | Per-run VBA log and/or Python run log (depending on execution path); harness appends VBA tail/marker context |
| `log\reauth-complete.txt` | Completion marker (`completed`, `saved`, `run_log`) written on success/failure path |
| `log\answer-dialogs.log` | Dialog detection/auto-answer actions |
| `log\watcher.log` | Captured VBA/Excel dialog text (compile/runtime) |
| `log\compile-test.log` | Compile harness results and captured dialog info |

> Note: Some runs may emit `ERROR | ... | Err=0 | Erl=0` trace entries on non-fatal handled paths; use context and final `OUTCOME`/marker to determine actual success.

### Common issues

- **Macros disabled / security prompts:** Enable macros or add the project folder to Excel Trusted Locations. Unattended runs attempt to configure trusted location/macro access in HKCU for Office 16.0.
- **“Could not find required headers”:** Ensure source has `Assigned Sales Manager`, `Auth Status`, `Patient Name`, `Re-auth Due Date` in a detectable header row.
- **No data rows:** Source contains headers but no records below them.
- **COM/RPC hang or Excel dialog blocks:** Use `scripts\run-report.py` (watchdog + dialog auto-answer). Check `log\answer-dialogs.log`, `log\watcher.log`, and `macro-debug.log` tail.
- **VBA compile error:** Run `python scripts\test-compile.py` and inspect `log\compile-test.log` / `log\watcher.log` for the reported error line/module.
- **CSV import issues:** Re-save/export as standard CSV or `.xlsx` if delimiter/encoding causes columns to collapse.
- **Wrong dates/blanks:** Non-parseable date values remain as-is; blank date cells are expected in some cases.

### How to investigate a failed run

1. Check Python console output for `OUTCOME=` and `NEW_REPORT=`.
2. Open the latest `log\reauth_*.log` and/or run `python .\scripts\view-logs.py`.
3. Review `log\macro-debug.log` (end of file) for the last `STEP`, `ERROR`, and `CreateReAuthReport COMPLETE/failed` entries.
4. Check `log\reauth-complete.txt` for saved path or failure message.
5. If dialogs/compile suspected, inspect `log\answer-dialogs.log` and `log\watcher.log`, then run `python .\scripts\test-compile.py`.

---

## Configuration notes

- **VBA constants:** `modDeclarations.bas` defines project-local paths (e.g. `PROJECT_ROOT`, `LOG_PATH`, completion marker, auto-source file) and header/sheet names. If you move the project, update these constants and re-sync the template (`python .\scripts\sync_template.py`).
- **Output naming:** Reports are saved as `Reports\ReAuth_Report_YYYY-mm-dd_HHMM.xlsx`.
- **Auto/unattended mode:** Presence of `log\auto-source.txt` allows VBA to read the source path without prompting; `scripts\run-report.py` manages this automatically.
- **Manager tab layout:** Header row is fixed at row 15 with summary block above (rows 5–9) and freeze below header; changing layout requires coordinated updates in `modCore.bas`, `modFormatting.bas`, and `modDeclarations.bas`.

---

## Development workflow

1. Edit VBA source modules (`mod*.bas`) in the repo.
2. Sync into template:
   ```powershell
   python .\scripts\sync_template.py
   ```
3. Compile-check:
   ```powershell
   python .\scripts\test-compile.py
   ```
4. Run and validate:
   ```powershell
   python .\scripts\run-report.py --source ".\reauth raw.csv"
   ```
5. Inspect output in `Reports\` and logs in `log\`.
6. Commit `.bas`, scripts, template, and docs (avoid committing generated reports/logs/temp files).

---

## GitHub / version control notes

Suggested `.gitignore` entries for a clean repo:

```gitignore
# Generated outputs and runtime artifacts
Reports/*.xlsx
log/*
!log/.gitkeep

# Excel/temp/lock/debug artifacts
~$*
*.tmp
_run_temp.xlsm
_compile_test.xlsm
_dbg*.xlsm
_check*.py
_debug*.py
_inspect_report.py
_verify_tpl.py
_force_sync.py
_compile.py
probe_ran.txt

# Python
__pycache__/
*.pyc
```

- Keep **`mod*.bas`** as the VBA source of truth; sync them into the template for distribution/use.
- Commit **`Re-Auth Template.xlsm`** if you want a ready-to-run artifact in the repo (note: binary diffs are not human-readable).
- Do not commit PHI/real patient exports. `reauth raw.csv` in this repo should remain fake/sample data only.

---

## Known limitations

- **Windows + Excel required:** VBA execution and COM automation depend on desktop Excel on Windows.
- **Hardcoded project paths in VBA:** Default log/output/marker paths are defined in `modDeclarations.bas` and must be updated if the project directory changes.
- **Untrusted/interactive environments:** Strict macro policies, protected view, or modal prompts can interrupt unattended runs (the harness mitigates common cases but cannot bypass all organizational controls).
- **Schema dependence:** Column detection/formatting relies on expected header names; significant source-schema changes may require VBA updates.
- **Large exports:** Very large datasets may increase runtime/memory usage in Excel; monitor `macro-debug.log` step timing if scaling up.
