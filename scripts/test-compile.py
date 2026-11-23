# test-compile.py - compile-test Re-Auth VBA modules
import argparse
from datetime import datetime
import hashlib
import os
import shutil
import subprocess
import sys
import threading
import time

import win32com.client

parser = argparse.ArgumentParser(description="Compile-test Re-Auth VBA")
parser.add_argument("--template", default=None)
args = parser.parse_args()

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
template = args.template or os.path.join(project_root, "Re-Auth Template.xlsm")
test_copy = os.path.join(project_root, "_compile_test.xlsm")
marker = os.path.join(project_root, "log", "probe_ran.txt")
watcher = os.path.join(script_dir, "watch-compile-error.py")
log_dir = os.path.join(project_root, "log")
os.makedirs(log_dir, exist_ok=True)
watch_log = os.path.join(log_dir, "watcher.log")
test_log = os.path.join(log_dir, "compile-test.log")


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(test_log, "a", encoding="utf-8") as f:
        f.write(line + "\n")


if os.path.exists(test_copy):
    try:
        os.remove(test_copy)
    except Exception:
        pass
shutil.copy2(template, test_copy)
if os.path.exists(marker):
    try:
        os.remove(marker)
    except Exception:
        pass
log(f"=== Compile test started: {template} ===")
compile_start = time.time()
bas_files = sorted(
    f
    for f in os.listdir(os.path.join(project_root, "vba"))
    if f.lower().startswith("mod") and f.lower().endswith(".bas")
)
for name in bas_files:
    p = os.path.join(project_root, "vba", name)
    try:
        with open(p, "rb") as f:
            data = f.read()
        digest = hashlib.sha256(data).hexdigest()[:12]
        log(f"Source: {name} lines={data.count(bytes([10])) + 1} sha256={digest}")
    except Exception as e:
        log(f"Source read failed: {name}: {e}")

watch_proc = subprocess.Popen(
    [sys.executable, watcher], creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
)
log(f"Watcher pid={watch_proc.pid}")
time.sleep(2)

excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
try:
    log(f"Excel version={excel.Version} build={excel.Build}")
except Exception as e:
    log(f"Excel version query failed: {e}")
try:
    excel.AutomationSecurity = 1
except Exception:
    pass


def kill_excel_bg():
    time.sleep(20)
    try:
        subprocess.run(["taskkill", "/F", "/IM", "excel.exe"], capture_output=True)
        log("Watchdog killed Excel")
    except Exception:
        pass


threading.Thread(target=kill_excel_bg, daemon=True).start()

log_mark = 0
if os.path.exists(watch_log):
    with open(watch_log, encoding="utf-8", errors="replace") as f:
        log_mark = sum(1 for _ in f)

run_error = None
wb = None
try:
    wb = excel.Workbooks.Open(test_copy)
    vbp = wb.VBProject
    mod = vbp.VBComponents.Add(1)
    mod.Name = "zzCompileProbe"
    mod.CodeModule.AddFromString(
        "Option Explicit\r\n"
        "Public Sub Probe()\r\n"
        f'    Open "{marker.replace(chr(92), chr(92) + chr(92))}" For Output As #1\r\n'
        '    Print #1, "ran"\r\n'
        "    Close #1\r\n"
        "End Sub\r\n"
    )
    log("Running Probe...")
    excel.Run("Probe")
    log("Probe returned")
except Exception as e:
    run_error = str(e)
    log(f"Run error: {run_error}")

time.sleep(1)
new_lines = []
if os.path.exists(watch_log):
    with open(watch_log, encoding="utf-8", errors="replace") as f:
        all_lines = f.readlines()
    new_lines = all_lines[log_mark:]

ok = os.path.exists(marker)
marker_text = ""
if ok:
    try:
        with open(marker, encoding="utf-8", errors="replace") as f:
            marker_text = f.read().strip()
    except Exception as e:
        log(f"Marker read failed: {e}")
    log(f"Probe marker: {marker_text!r}")
if ok:
    log("RESULT: COMPILE OK")
else:
    log("RESULT: COMPILE FAILED")
    if new_lines:
        log("Watcher output:")
        for line in new_lines[-60:]:
            log("  " + line.rstrip())
    else:
        log("No watcher diagnostics captured (compile dialog not observed)")
    if run_error:
        log(f"Error: {run_error}")
log(f"Compile test duration(s)={time.time() - compile_start:.2f}")

try:
    if wb:
        wb.Close(False)
except Exception:
    pass
try:
    excel.Quit()
except Exception:
    pass
try:
    watch_proc.kill()
except Exception:
    pass
try:
    os.remove(test_copy)
except Exception:
    pass
log("=== Compile test finished ===")
print("COMPILE_OK=" + ("1" if ok else "0"))
sys.exit(0 if ok else 1)
