# run-report.py - robust unattended run for Re-Auth report
# Includes watchdog, dialog auto-answerer, realtime logging, and diagnostics.
import argparse
import ctypes
import ctypes.wintypes as wintypes
import os
import shutil
import subprocess
import sys
import threading
import time
from datetime import datetime

import win32com.client

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

EnumWindows = user32.EnumWindows
EnumChildWindows = user32.EnumChildWindows
IsWindowVisible = user32.IsWindowVisible
GetWindowTextW = user32.GetWindowTextW
GetClassNameW = user32.GetClassNameW
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
SetForegroundWindow = user32.SetForegroundWindow
SendMessageW = user32.SendMessageW
PostMessageW = user32.PostMessageW

EnumProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

WM_SETTEXT = 0x000C
WM_COMMAND = 0x0111
IDOK = 1
VK_RETURN = 0x0D
VK_ESCAPE = 0x1B

parser = argparse.ArgumentParser(description="Run the Re-Auth report macro (robust)")
parser.add_argument("--template", default=None)
parser.add_argument("--source", default=None)
parser.add_argument("--temp-copy", default=None)
parser.add_argument("--max-run-seconds", type=int, default=90)
parser.add_argument("--stall-seconds", type=int, default=15)
args = parser.parse_args()

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
template = args.template or os.path.join(project_root, "Re-Auth Template.xlsm")
source_path = os.path.abspath(args.source or os.path.join(project_root, "reauth raw.csv"))
temp_copy = args.temp_copy or os.path.join(project_root, "_run_temp.xlsm")
max_run_seconds = args.max_run_seconds

reports_dir = os.path.join(project_root, "Reports")
log_dir = os.path.join(project_root, "log")
os.makedirs(reports_dir, exist_ok=True)
os.makedirs(log_dir, exist_ok=True)

run_stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
run_log = os.path.join(log_dir, f"reauth_{run_stamp}.log")
ans_log = os.path.join(log_dir, "answer-dialogs.log")
completion_marker = os.path.join(log_dir, "reauth-complete.txt")
debug_log = os.path.join(log_dir, "macro-debug.log")

before = set()
if os.path.isdir(reports_dir):
    before = {f.name for f in os.scandir(reports_dir) if f.name.lower().endswith(".xlsx")}

for p in [completion_marker]:
    try:
        if os.path.exists(p):
            os.remove(p)
    except Exception:
        pass

debug_line_offset = 0
if os.path.exists(debug_log):
    try:
        with open(debug_log, "r", encoding="utf-8", errors="replace") as f:
            debug_line_offset = sum(1 for _ in f)
    except Exception:
        debug_line_offset = 0


def write_log(path, msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    try:
        with open(path, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def append_run_log(msg):
    write_log(run_log, msg)


append_run_log("=== Re-Auth report run started ===")
append_run_log(f"Template: {template}")
append_run_log(f"Source: {source_path}")
append_run_log(f"Temp copy: {temp_copy}")
append_run_log(f"Completion marker: {completion_marker}")
append_run_log(f"Python PID: {os.getpid()}")
append_run_log(f"Max run seconds: {max_run_seconds}")

# Auto-mode files for VBA
auto_source_file = os.path.join(log_dir, "auto-source.txt")
with open(auto_source_file, "w", encoding="utf-8") as f:
    f.write(source_path)
append_run_log(f"Wrote auto source path: {source_path}")

shutil.copy2(template, temp_copy)
append_run_log("Copied template to temp workbook")

outcome = "UNKNOWN"
run_completed = False
run_start_time = time.time()
cancel_event = threading.Event()


def excel_pids():
    try:
        import psutil
        pids = set()
        for p in psutil.process_iter(["name"]):
            try:
                if p.info["name"] and p.info["name"].lower() == "excel.exe":
                    pids.add(p.pid)
            except Exception:
                pass
        return pids
    except Exception:
        return set()


def win_text(hwnd):
    buf = ctypes.create_unicode_buffer(1024)
    GetWindowTextW(hwnd, buf, 1024)
    return buf.value


def win_class(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    GetClassNameW(hwnd, buf, 256)
    return buf.value


def find_child(hwnd, cls_name):
    found = [None]

    def cb(child, _):
        if win_class(child) == cls_name:
            found[0] = child
            return False
        return True

    EnumChildWindows(hwnd, EnumProc(cb), 0)
    return found[0]


def excel_dialogs():
    pids = excel_pids()
    if not pids:
        return []
    result = []

    def cb(hwnd, _):
        if not IsWindowVisible(hwnd):
            return True
        pid = wintypes.DWORD()
        GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if pid.value not in pids:
            return True
        result.append(hwnd)
        return True

    EnumWindows(EnumProc(cb), 0)
    return result


def send_keys(text):
    class KEYBDINPUT(ctypes.Structure):
        _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD), ("dwFlags", wintypes.DWORD),
                    ("time", wintypes.DWORD), ("dwExtraInfo", ctypes.c_void_p)]

    class INPUT(ctypes.Structure):
        _fields_ = [("type", wintypes.DWORD), ("ki", KEYBDINPUT)]

    VkKeyScanW = user32.VkKeyScanW
    inputs = []
    for ch in text:
        vk = VkKeyScanW(ord(ch) if isinstance(ch, str) else ch)
        shift = (vk >> 8) & 0xFF
        if shift & 1:
            inputs.append(INPUT(1, KEYBDINPUT(0x10, 0, 0, 0, 0)))
        inputs.append(INPUT(1, KEYBDINPUT(vk & 0xFF, 0, 0, 0, 0)))
        inputs.append(INPUT(1, KEYBDINPUT(vk & 0xFF, 0, 2, 0, 0)))
        if shift & 1:
            inputs.append(INPUT(1, KEYBDINPUT(0x10, 0, 2, 0, 0)))
    if inputs:
        n = len(inputs)
        arr = (INPUT * n)(*inputs)
        user32.SendInput(n, arr, ctypes.sizeof(INPUT))


def key_down_once(vk):
    class KEYBDINPUT(ctypes.Structure):
        _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD), ("dwFlags", wintypes.DWORD),
                    ("time", wintypes.DWORD), ("dwExtraInfo", ctypes.c_void_p)]

    class INPUT(ctypes.Structure):
        _fields_ = [("type", wintypes.DWORD), ("ki", KEYBDINPUT)]

    arr = (INPUT * 2)(INPUT(1, KEYBDINPUT(vk, 0, 0, 0, 0)), INPUT(1, KEYBDINPUT(vk, 0, 2, 0, 0)))
    user32.SendInput(2, arr, ctypes.sizeof(INPUT))


def kill_excel():
    try:
        subprocess.run(["taskkill", "/F", "/IM", "excel.exe"], capture_output=True)
        append_run_log("Issued taskkill for EXCEL.EXE")
    except Exception as e:
        append_run_log(f"kill_excel error: {e}")


def watchdog_thread():
    start = time.time()
    while not cancel_event.is_set() and (time.time() - start) < max_run_seconds:
        time.sleep(0.4)
        if os.path.exists(completion_marker):
            append_run_log("WATCHDOG: completion marker detected")
            return
        try:
            new_files = [f.name for f in os.scandir(reports_dir)
                         if f.name.lower().endswith(".xlsx") and f.name not in before]
            if new_files:
                append_run_log(f"WATCHDOG: new report file(s) detected: {', '.join(new_files)}")
        except Exception:
            pass
    if not cancel_event.is_set():
        append_run_log(f"WATCHDOG: exceeded {max_run_seconds}s, killing Excel")
        kill_excel()


def answerer_thread():
    answered = {}
    start = time.time()
    while not cancel_event.is_set() and (time.time() - start) < (max_run_seconds + 30):
        time.sleep(0.25)
        try:
            for h in excel_dialogs():
                k = str(h)
                if k in answered:
                    continue
                cls = win_class(h)
                title = win_text(h)
                if not (cls in ("#32770", "bosa_sdm_MsoFiles", "wndclass_desked_gsk")
                        or "Select" in title or "Choose" in title or "Microsoft Excel" in title
                        or "Visual Basic" in title):
                    continue
                write_log(ans_log, f"dialog: class={cls} title={title}")
                if cls == "#32770" and title == "Microsoft Excel":
                    # try to capture child static text
                    try:
                        child_texts = []
                        def ccb(child, _):
                            t = win_text(child)
                            c = win_class(child)
                            if t and c == "Static":
                                child_texts.append(t)
                            return True
                        EnumChildWindows(h, EnumProc(ccb), 0)
                        if child_texts:
                            write_log(ans_log, f"  msg: {' | '.join(child_texts)}")
                    except Exception:
                        pass
                    key_down_once(VK_RETURN)
                    write_log(ans_log, "dismissed Excel error dialog")
                    answered[k] = 1
                    continue
                if "Visual Basic" in title:
                    key_down_once(VK_RETURN)
                    write_log(ans_log, f"dismissed VBA dialog: {title}")
                    answered[k] = 1
                    continue
                SetForegroundWindow(h)
                time.sleep(0.2)
                if "PRIOR" in title:
                    key_down_once(VK_ESCAPE)
                    write_log(ans_log, "cancelled PRIOR picker")
                elif "CURRENT" in title or "Select the Re-Auth" in title or "raw data" in title.lower():
                    edit = find_child(h, "Edit")
                    if edit is not None:
                        SendMessageW(edit, WM_SETTEXT, 0, source_path)
                    else:
                        send_keys(source_path)
                    time.sleep(0.1)
                    key_down_once(VK_RETURN)
                    write_log(ans_log, f"answered source picker -> {source_path}")
                elif "Choose where" in title or "Save" in title:
                    key_down_once(VK_RETURN)
                    write_log(ans_log, "accepted save dialog")
                else:
                    key_down_once(VK_RETURN)
                    write_log(ans_log, f"dismissed dialog: {title}")
                answered[k] = 1
        except Exception as e:
            write_log(ans_log, f"answerer error: {e}")


t_watch = threading.Thread(target=watchdog_thread, daemon=True)
t_ans = threading.Thread(target=answerer_thread, daemon=True)
t_watch.start()
t_ans.start()

excel = None
wb = None
try:
    # Trusted location + macro settings
    try:
        import winreg
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\16.0\Excel\Security")
        winreg.SetValueEx(key, "VBAWarnings", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(key)
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\ReAuthProject")
        winreg.SetValueEx(key, "Path", 0, winreg.REG_SZ, project_root)
        winreg.SetValueEx(key, "Description", 0, winreg.REG_SZ, "Re-Auth Project")
        winreg.SetValueEx(key, "AllowSubfolders", 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(key)
        append_run_log("Registry: macro access/trusted location configured")
    except Exception as e:
        append_run_log(f"Registry config warning: {e}")

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    try:
        excel.AutomationSecurity = 1
    except Exception:
        pass
    append_run_log("Excel started (DispatchEx)")

    wb = excel.Workbooks.Open(temp_copy)
    append_run_log(f"Opened temp workbook: {temp_copy}")

    # Diagnostics
    try:
        vbp = wb.VBProject
        comps = [c.Name for c in vbp.VBComponents]
        append_run_log(f"VBComponents: {', '.join(comps)}")
    except Exception as e:
        append_run_log(f"Could not enumerate VBComponents: {e}")

    macro_result = None
    try:
        append_run_log("Running macro: CreateReAuthReport")
        macro_result = excel.Run("CreateReAuthReport")
        append_run_log(f"Excel.Run returned: {macro_result}")
        if macro_result is True:
            run_completed = True
            outcome = "SUCCESS"
        elif macro_result is False:
            outcome = "FAILED"
        else:
            outcome = f"RETURNED:{macro_result}"
    except Exception as e:
        outcome = f"ERROR: {e}"
        append_run_log(f"Excel.Run exception: {e}")

    # Wait for completion marker or output
    wait_deadline = time.time() + min(30, max_run_seconds)
    while time.time() < wait_deadline and not cancel_event.is_set():
        time.sleep(0.5)
        if os.path.exists(completion_marker):
            append_run_log("Completion marker found")
            run_completed = True
            if outcome.startswith("ERROR"):
                outcome = "COMPLETED_WITH_ERROR"
            elif outcome == "UNKNOWN":
                outcome = "SUCCESS"
            break
        try:
            new_files = [f.name for f in os.scandir(reports_dir)
                         if f.name.lower().endswith(".xlsx") and f.name not in before]
            if new_files:
                append_run_log(f"Detected output file(s): {', '.join(new_files)}")
                run_completed = True
                if outcome == "UNKNOWN":
                    outcome = "SUCCESS"
                break
        except Exception:
            pass

finally:
    cancel_event.set()
    append_run_log("Closing workbook/Excel")
    try:
        if wb is not None:
            wb.Close(False)
            append_run_log("Temp workbook closed")
    except Exception as e:
        append_run_log(f"Workbook close warning: {e}")
    try:
        if excel is not None:
            excel.Quit()
            append_run_log("Excel quit")
    except Exception as e:
        append_run_log(f"Excel quit warning: {e}")
    try:
        if os.path.exists(temp_copy):
            os.remove(temp_copy)
    except Exception:
        pass
    try:
        lock = os.path.join(project_root, "~$_run_temp.xlsm")
        if os.path.exists(lock):
            os.remove(lock)
    except Exception:
        pass

time.sleep(0.5)

# Collect outputs
try:
    after = {f.name for f in os.scandir(reports_dir) if f.name.lower().endswith(".xlsx")}
    new_files = sorted(after - before)
except Exception:
    new_files = []

if new_files and outcome in ("UNKNOWN", "COMPLETED_WITH_ERROR"):
    outcome = "SUCCESS"

run_duration = time.time() - run_start_time
append_run_log(f"Final outcome: {outcome}")
append_run_log(f"Total duration: {run_duration:.2f}s")
append_run_log(f"New report files: {new_files}")

# Append macro-debug tail for this run
try:
    if os.path.exists(debug_log):
        with open(debug_log, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
        tail = lines[debug_line_offset:]
        if len(tail) > 200:
            tail = tail[-200:]
        append_run_log("--- macro-debug.log (this run tail) ---")
        for ln in tail:
            append_run_log("VBA " + ln.rstrip("\n\r"))
except Exception as e:
    append_run_log(f"Could not read macro-debug.log: {e}")

# Completion marker contents
try:
    if os.path.exists(completion_marker):
        append_run_log("--- completion marker ---")
        with open(completion_marker, "r", encoding="utf-8", errors="replace") as f:
            for ln in f:
                append_run_log("MARK " + ln.rstrip("\n\r"))
    else:
        append_run_log("Completion marker not found")
except Exception as e:
    append_run_log(f"Could not read completion marker: {e}")

if new_files:
    print("NEW_REPORT=" + os.path.join(reports_dir, new_files[-1]))
else:
    print("NEW_REPORT=")
print("OUTCOME=" + outcome)

sys.exit(0 if outcome == "SUCCESS" else 1)
