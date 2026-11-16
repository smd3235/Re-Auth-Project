# snapshot-report.py - export Re-Auth report sheets as PNG screenshots
import argparse, os, re, sys, time, subprocess
import win32com.client

parser = argparse.ArgumentParser(description="Export Re-Auth report sheets as PNG screenshots")
parser.add_argument("--path", default=None, help="Workbook to export (defaults to newest in Reports/)")
parser.add_argument("--sheets", nargs="*", default=None, help="Specific sheet names to export")
parser.add_argument("--out", default=None, help="Output folder (defaults to documentation/screenshots/)")
parser.add_argument("--prefix", default="report", help="Filename prefix")
parser.add_argument("--list", action="store_true", help="List sheets and exit")
parser.add_argument("--no-kill", action="store_true", help="Do not kill Excel before starting")
args = parser.parse_args()

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
path = args.path
out_folder = args.out or os.path.join(project_root, "documentation", "screenshots")
prefix = args.prefix

if not path:
    reports_dir = os.path.join(project_root, "Reports")
    candidates = []
    if os.path.isdir(reports_dir):
        for f in os.listdir(reports_dir):
            if f.lower().endswith(".xlsx"):
                candidates.append(os.path.join(reports_dir, f))
    if candidates:
        candidates.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        path = candidates[0]
        print(f"Using newest report: {path}")
    else:
        print("No workbook given and Reports\\ holds no .xlsx. Pass --path.", file=sys.stderr)
        sys.exit(2)

if not os.path.exists(path):
    print(f"Workbook not found: {path}", file=sys.stderr)
    sys.exit(2)
path = os.path.abspath(path)
os.makedirs(out_folder, exist_ok=True)

for f in os.listdir(out_folder):
    if f.lower().endswith(".png"):
        try:
            os.remove(os.path.join(out_folder, f))
        except OSError:
            pass


def save_sheet_image(workbook, worksheet, png_path):
    used = worksheet.UsedRange
    first_row = used.Row
    first_col = used.Column
    last_row = first_row + used.Rows.Count - 1
    last_col = first_col + used.Columns.Count - 1
    rng = worksheet.Range(worksheet.Cells(first_row, first_col), worksheet.Cells(last_row, last_col))

    worksheet.Activate()
    try:
        workbook.Application.ActiveWindow.Zoom = 100
    except Exception:
        pass
    time.sleep(0.5)

    copied = False
    for attempt in range(3):
        try:
            rng.CopyPicture(Appearance=1, Format=2)  # xlScreen, xlBitmap
            copied = True
            break
        except Exception:
            time.sleep(1.0)
    if not copied:
        # fallback picture format
        for attempt in range(2):
            try:
                rng.CopyPicture(Appearance=1, Format=1)
                copied = True
                break
            except Exception:
                time.sleep(1.0)
    if not copied:
        raise RuntimeError(f"CopyPicture failed for {worksheet.Name}")

    rows = int(used.Rows.Count)
    cols = int(used.Columns.Count)
    w = max(1400, cols * 95)
    h = max(700, rows * 26)

    chart = workbook.Charts.Add()
    try:
        try:
            chart.ChartArea.Clear()
        except Exception:
            pass
        chart.Paste()
        time.sleep(0.5)
        try:
            chart.ChartArea.Width = w
            chart.ChartArea.Height = h
        except Exception:
            pass
        if os.path.exists(png_path):
            os.remove(png_path)
        chart.Export(png_path)
    finally:
        try:
            chart.Delete()
        except Exception:
            pass


if not args.no_kill:
    try:
        subprocess.run(["taskkill", "/F", "/IM", "excel.exe"], capture_output=True)
    except Exception:
        pass
    time.sleep(2.0)

xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
try:
    xl.AutomationSecurity = 3
except Exception:
    pass
try:
    xl.ScreenUpdating = True
except Exception:
    pass
wb = None
exported = 0
try:
    wb = xl.Workbooks.Open(path, False, True)
    time.sleep(1.5)

    if args.list or not args.sheets:
        print(f"Sheets in {os.path.basename(path)}:")
        for s in wb.Worksheets:
            print(f"  {s.Name:<28} used {s.UsedRange.Address}")

    targets = []
    if args.sheets:
        for name in args.sheets:
            found = None
            for s in wb.Worksheets:
                if s.Name == name:
                    found = s
                    break
            if found:
                targets.append(found)
            else:
                print(f"Sheet not found, skipped: {name}", file=sys.stderr)
    else:
        targets = list(wb.Worksheets)

    for ws in targets:
        safe = re.sub(r"[^\w\-]+", "-", ws.Name).strip("-").lower()
        png = os.path.join(out_folder, f"{prefix}-{safe}.png")
        save_sheet_image(wb, ws, png)
        kb = round(os.path.getsize(png) / 1024, 1)
        print(f"  exported {png}  ({kb} KB)")
        exported += 1
finally:
    if wb:
        try:
            wb.Close(False)
        except Exception:
            pass
    try:
        xl.Quit()
    except Exception:
        pass
    try:
        subprocess.run(["taskkill", "/F", "/IM", "excel.exe"], capture_output=True)
    except Exception:
        pass

if exported == 0 and not args.list and args.sheets:
    print("Nothing exported.", file=sys.stderr)
    sys.exit(3)
sys.exit(0)
