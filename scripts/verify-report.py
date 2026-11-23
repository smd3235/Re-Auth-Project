# verify-report.py - validate a generated Re-Auth report workbook.
# Checks structure (sheets), Raw Data volume, manager tab headers, formula errors,
# and optional row-count parity against the source CSV. Logs to log/verify-report.log.
import argparse
import csv
from datetime import datetime
import glob
import os
import sys

import win32com.client

parser = argparse.ArgumentParser(description="Verify a generated Re-Auth report")
parser.add_argument("--report", default=None, help="Report path (default: newest in Reports/)")
parser.add_argument(
    "--source", default=None, help="Optional source CSV/Excel to compare row counts"
)
parser.add_argument("--log-dir", default=None)
parser.add_argument("--json", action="store_true")
args = parser.parse_args()

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
log_dir = os.path.abspath(args.log_dir or os.path.join(project_root, "log"))
os.makedirs(log_dir, exist_ok=True)
log_path = os.path.join(log_dir, "verify-report.log")

EXPECTED_MANAGER_HEADERS = [
    "Wound Care Consultant",
    "Associate WCC",
    "Auth Status",
    "Patient Name",
    "HC Company",
    "Re-auth Due Date",
    "Auth Start Date",
    "Authorization Status",
    "Task Required",
    "Patient/Packet Status",
    "Treatment Physician",
]

issues = []


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def add(sev, kind, detail):
    issues.append({"severity": sev, "kind": kind, "detail": detail})
    log(f"{sev.upper()} [{kind}] {detail}")


def find_report():
    if args.report:
        return os.path.abspath(args.report)
    files = [
        p
        for p in glob.glob(os.path.join(project_root, "Reports", "*.xlsx"))
        if not os.path.basename(p).startswith("~$")
    ]
    if not files:
        return None
    return max(files, key=os.path.getmtime)


def source_row_count():
    src = args.source or os.path.join(project_root, "reauth raw.csv")
    if not os.path.exists(src) or not src.lower().endswith((".csv", ".txt")):
        return None
    for enc in ("utf-8-sig", "cp1252"):
        try:
            with open(src, encoding=enc, newline="") as f:
                return max(0, sum(1 for _ in csv.reader(f)) - 1)
        except Exception:
            continue
    return None


report = find_report()
if not report or not os.path.exists(report):
    add("error", "report_missing", "No report workbook found to verify")
    if args.json:
        import json

        print(json.dumps({"ok": False, "issues": issues}, indent=2))
    sys.exit(1)

log(f"=== Verify started: {report} bytes={os.path.getsize(report)} ===")
try:
    import zipfile

    if not zipfile.is_zipfile(report):
        add("error", "not_xlsx", "Report is not a valid xlsx (zip) container")
except Exception as e:
    add("warning", "zip_check_failed", str(e))

excel = None
wb = None
try:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(report, ReadOnly=True, UpdateLinks=False)
    names = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]
    log(f"Sheets ({len(names)}): {', '.join(names)}")

    for required in ("Dashboard", "Raw Data"):
        if required not in names:
            add("error", "missing_sheet", f"Required sheet missing: {required}")

    manager_sheets = [n for n in names if n not in ("Dashboard", "Raw Data")]
    if not manager_sheets:
        add("error", "no_manager_tabs", "No manager tabs found")
    else:
        log(f"Manager tabs ({len(manager_sheets)}): {', '.join(manager_sheets)}")
    for n in manager_sheets:
        if len(n) > 31:
            add("error", "sheet_name_too_long", f"Sheet name exceeds 31 chars: {n}")

    # Raw Data volume
    if "Raw Data" in names:
        ws = wb.Worksheets("Raw Data")
        used = ws.UsedRange
        rows, cols = used.Rows.Count, used.Columns.Count
        log(f"Raw Data dims: rows={rows} cols={cols}")
        if rows < 2:
            add("error", "raw_data_empty", f"Raw Data has {rows} row(s)")
        expected = source_row_count()
        if expected is not None:
            log(f"Source data rows={expected} raw_data_rows={rows - 1}")
            if rows - 1 != expected:
                add(
                    "warning",
                    "rowcount_mismatch",
                    f"Source rows {expected} != Raw Data rows {rows - 1}",
                )

    # Manager tab headers + KPI text
    for n in manager_sheets:
        ws = wb.Worksheets(n)
        headers = [str(ws.Cells(7, c).Value or "").strip() for c in range(1, 12)]
        if headers != EXPECTED_MANAGER_HEADERS:
            add(
                "warning",
                "manager_header_mismatch",
                f"{n}: row 7 headers differ: {headers}",
            )
        title = str(ws.Range("A1").Value or "")
        if not title.upper().startswith("RE-AUTH"):
            add(
                "warning",
                "manager_title_missing",
                f"{n}: A1 title unexpected: {title!r}",
            )

    # Error cells anywhere in used ranges (formulas and constants)
    xlCellTypeConstants, xlCellTypeFormulas, xlErrors = 2, -4123, 16
    for n in names:
        ws = wb.Worksheets(n)
        for cell_type, label in (
            (xlCellTypeConstants, "constants"),
            (xlCellTypeFormulas, "formulas"),
        ):
            try:
                rng = ws.UsedRange.SpecialCells(cell_type, xlErrors)
                add(
                    "error",
                    "error_cells",
                    f"{n}: {rng.Count} error cell(s) in {label} at {rng.Address}",
                )
            except Exception:
                pass

    # Dashboard KPI sanity: look for the Total KPI value cell ranges actually containing numbers
    if "Dashboard" in names:
        ws = wb.Worksheets("Dashboard")
        kpi_hits = 0
        for row in ws.Range("A1:J30").Value:
            for v in row:
                if isinstance(v, (int, float)):
                    kpi_hits += 1
        log(f"Dashboard numeric cells in A1:J30={kpi_hits}")
        if kpi_hits == 0:
            add(
                "warning",
                "dashboard_no_numbers",
                "No numeric KPI cells found on Dashboard A1:J30",
            )
finally:
    try:
        if wb is not None:
            wb.Close(False)
    except Exception:
        pass
    try:
        if excel is not None:
            excel.Quit()
    except Exception:
        pass

errors = [i for i in issues if i["severity"] == "error"]
warnings = [i for i in issues if i["severity"] == "warning"]
log(f"=== Verify finished: {len(errors)} error(s), {len(warnings)} warning(s) ===")

if args.json:
    import json

    print(json.dumps({"ok": not errors, "report": report, "issues": issues}, indent=2))
sys.exit(1 if errors else 0)
