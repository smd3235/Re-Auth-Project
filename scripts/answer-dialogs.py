# answer-dialogs.py - standalone dialog auto-answerer for Re-Auth unattended runs
import argparse
import ctypes
from ctypes import wintypes
from datetime import datetime
import os
import time

user32 = ctypes.windll.user32
EnumWindows = user32.EnumWindows
EnumChildWindows = user32.EnumChildWindows
IsWindowVisible = user32.IsWindowVisible
GetWindowTextW = user32.GetWindowTextW
GetClassNameW = user32.GetClassNameW
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
SetForegroundWindow = user32.SetForegroundWindow
SendMessageW = user32.SendMessageW
EnumProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
WM_SETTEXT = 0x000C
VK_RETURN = 0x0D
VK_ESCAPE = 0x1B


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


def key_down_once(vk):
    class KEYBDINPUT(ctypes.Structure):
        _fields_ = [
            ("wVk", wintypes.WORD),
            ("wScan", wintypes.WORD),
            ("dwFlags", wintypes.DWORD),
            ("time", wintypes.DWORD),
            ("dwExtraInfo", ctypes.c_void_p),
        ]

    class INPUT(ctypes.Structure):
        _fields_ = [("type", wintypes.DWORD), ("ki", KEYBDINPUT)]

    arr = (INPUT * 2)(INPUT(1, KEYBDINPUT(vk, 0, 0, 0, 0)), INPUT(1, KEYBDINPUT(vk, 0, 2, 0, 0)))
    user32.SendInput(2, arr, ctypes.sizeof(INPUT))


project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
log_dir = os.path.join(project_root, "log")
os.makedirs(log_dir, exist_ok=True)

parser = argparse.ArgumentParser(description="Answer Excel dialogs for Re-Auth unattended runs")
parser.add_argument("--source", default=os.path.join(project_root, "reauth raw.csv"))
parser.add_argument("--log", default=os.path.join(log_dir, "answer-dialogs.log"))
parser.add_argument("--seconds", type=int, default=300)
args = parser.parse_args()


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    try:
        with open(args.log, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception as e:
        print(f"(log write failed: {e}) {line}")


log(f"answer-dialogs started pid={os.getpid()} source={args.source}")
answered = {}
deadline = time.time() + args.seconds
try:
    while time.time() < deadline:
        time.sleep(0.3)
        try:
            live = {str(h): h for h in excel_dialogs()}
            for k in list(answered.keys()):
                if k not in live:
                    del answered[k]
            for k, h in live.items():
                if k in answered:
                    continue
                cls = win_class(h)
                title = win_text(h)
                if (
                    cls in ("#32770", "bosa_sdm_MsoFiles", "wndclass_desked_gsk")
                    or "Select" in title
                    or "Visual Basic" in title
                    or title == "Microsoft Excel"
                ):
                    log(f"dialog class={cls} title={title}")
                    SetForegroundWindow(h)
                    time.sleep(0.2)
                    if "PRIOR" in title:
                        key_down_once(VK_ESCAPE)
                        log("cancelled PRIOR")
                    elif (
                        "Select" in title
                        or "raw data" in title.lower()
                        or cls == "bosa_sdm_MsoFiles"
                    ):
                        edit = find_child(h, "Edit")
                        if edit is not None:
                            SendMessageW(edit, WM_SETTEXT, 0, args.source)
                        key_down_once(VK_RETURN)
                        log(f"answered source -> {args.source}")
                    else:
                        key_down_once(VK_RETURN)
                        log(f"dismissed: {title}")
                    answered[k] = 1
        except Exception as e:
            log(f"answerer error: {e}")
            time.sleep(0.5)
finally:
    log("answer-dialogs stopped")
