# sync_template.py - import VBA modules into Re-Auth Template.xlsm (reliable replace)
# Can also build a local macro-enabled template from the macro-free base (.xlsx) to avoid MOTW issues.
import argparse
from datetime import datetime
import glob
import hashlib
import os
import sys

import win32com.client

parser = argparse.ArgumentParser(
    description="Sync VBA modules into the Re-Auth template / build local .xlsm"
)
parser.add_argument(
    "--template", default=None, help="Template to open (.xlsm or macro-free .xlsx base)"
)
parser.add_argument(
    "--output",
    default=None,
    help="Output path (default: overwrite template; or .xlsm when building from .xlsx)",
)
parser.add_argument("--bas-dir", default=None)
args = parser.parse_args()

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
template = os.path.abspath(args.template or os.path.join(project_root, "Re-Auth Template.xlsm"))
bas_dir = args.bas_dir or os.path.join(project_root, "vba")
button_name = "btnCreateReAuthReport"
log_path = os.path.join(project_root, "log", "sync-template.log")
os.makedirs(os.path.dirname(log_path), exist_ok=True)


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def _log_uncaught(exc_type, exc, tb):
    log(f"ERROR: uncaught {exc_type.__name__}: {exc}")
    sys.__excepthook__(exc_type, exc, tb)


sys.excepthook = _log_uncaught


bas_files = sorted(glob.glob(os.path.join(bas_dir, "mod*.bas")))
if not bas_files:
    log("ERROR: no mod*.bas files found")
    sys.exit(1)

for b in bas_files:
    if not os.path.exists(b):
        log(f"ERROR: bas file not found: {b}")
        sys.exit(1)

log(f"=== Sync started: template={template} modules={len(bas_files)} bas_dir={bas_dir} ===")
for b in bas_files:
    try:
        with open(b, "rb") as f:
            data = f.read()
        log(
            f"Source: {os.path.basename(b)} lines={data.count(bytes([10])) + 1} sha256={hashlib.sha256(data).hexdigest()[:12]}"
        )
    except Exception as e:
        log(f"Source read failed: {os.path.basename(b)}: {e}")

excel = win32com.client.DispatchEx("Excel.Application")
excel.DisplayAlerts = False
try:
    excel.AutomationSecurity = 3
except Exception:
    pass

try:
    wb = excel.Workbooks.Open(template)
    vbp = wb.VBProject

    imported_names = {os.path.splitext(os.path.basename(b))[0] for b in bas_files}
    # remove standard modules we manage
    for c in list(vbp.VBComponents):
        try:
            if c.Type == 1 and (
                c.Name in imported_names or c.Name.startswith("PendingPickupReport")
            ):
                vbp.VBComponents.Remove(c)
        except Exception as e:
            log(f"WARN: could not remove component {c.Name}: {e}")

    for bas in bas_files:
        comp = vbp.VBComponents.Import(bas)
        log(f"Imported: {comp.Name} lines={comp.CodeModule.CountOfLines}")

    ws = wb.Worksheets("START HERE")
    btn = None
    for shp in ws.Shapes:
        if shp.Name in (button_name, "btnCreatePendingPickupReport"):
            btn = shp
            btn.Name = button_name
            break
    if btn is None:
        cta = ws.Range("B14")
        btn = ws.Shapes.AddShape(5, cta.Left + 20, cta.Top + 8, 360, 48)
        btn.Name = button_name
    try:
        btn.Fill.ForeColor.RGB = 7949855
        btn.Line.ForeColor.RGB = 4116214
        try:
            btn.TextFrame2.TextRange.Text = "CREATE RE-AUTH REPORT"
            btn.TextFrame2.TextRange.Font.Bold = True
            btn.TextFrame2.TextRange.Font.Size = 14
            btn.TextFrame2.TextRange.Font.Name = "Segoe UI"
            btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
        except Exception:
            btn.TextFrame.Characters().Text = "CREATE RE-AUTH REPORT"
    except Exception:
        pass
    btn.OnAction = "CreateReAuthReport"

    for name in ["btnSetupEmailCC"]:
        for shp in list(ws.Shapes):
            try:
                if shp.Name == name:
                    shp.Delete()
            except Exception:
                pass

    out_path = os.path.abspath(args.output) if args.output else template
    # If building from macro-free base (.xlsx/.xltx), save as macro-enabled workbook
    if out_path.lower().endswith((".xlsx", ".xltx")):
        out_path = os.path.splitext(out_path)[0] + ".xlsm"
    if out_path.lower().endswith(".xlsm"):
        wb.SaveAs(out_path, FileFormat=52)  # xlOpenXMLWorkbookMacroEnabled
    else:
        wb.Save()
    log(
        f"Saved: {out_path} bytes={os.path.getsize(out_path) if os.path.exists(out_path) else 'MISSING'}"
    )
    if not os.path.exists(out_path):
        log("ERROR: save did not produce the expected file")
    wb.Close(True)
    log(f"Synced {len(bas_files)} module(s) into {out_path}")
    for b in bas_files:
        log(f"  - {os.path.basename(b)}")
    log("=== Sync finished OK ===")
finally:
    try:
        excel.Quit()
    except Exception:
        pass
