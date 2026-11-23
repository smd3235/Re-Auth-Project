# watch-compile-error.py - capture VBA compile/runtime error dialogs for Re-Auth
import ctypes
from ctypes import wintypes
from datetime import datetime
import os
import time

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
EnumWindows = user32.EnumWindows
EnumChildWindows = user32.EnumChildWindows
IsWindowVisible = user32.IsWindowVisible
GetWindowTextW = user32.GetWindowTextW
GetClassNameW = user32.GetClassNameW
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
PostMessageW = user32.PostMessageW
EnumProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
WM_COMMAND = 0x0111
IDOK = 1

project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
log_dir = os.path.join(project_root, "log")
os.makedirs(log_dir, exist_ok=True)
log_path = os.path.join(log_dir, "watcher.log")


def write_log(text):
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for ln in [line for line in str(text).splitlines() if line.strip()]:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"{stamp} | {ln}\n")


def win_text(hwnd):
    buf = ctypes.create_unicode_buffer(1024)
    GetWindowTextW(hwnd, buf, 1024)
    return buf.value


def win_class(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    GetClassNameW(hwnd, buf, 256)
    return buf.value


def dialog_text(hwnd):
    parts = [win_text(hwnd)]
    sb = []

    def cb(child, _):
        t = win_text(child)
        if t:
            sb.append(t)
        return True

    EnumChildWindows(hwnd, EnumProc(cb), 0)
    if sb:
        parts.append(" | " + " | ".join(sb))
    return "".join(parts)


def excel_pids():
    try:
        import psutil

        return {
            p.pid
            for p in psutil.process_iter(["name"])
            if p.info.get("name", "").lower() == "excel.exe"
        }
    except Exception:
        return set()


class Watcher:
    def __init__(self):
        self.seen = set()
        self.stuck = {}

    def poll(self):
        hits, killed = [], []
        pids = excel_pids()
        if not pids:
            return hits, killed

        def cb(hwnd, _):
            if not IsWindowVisible(hwnd):
                return True
            pid = wintypes.DWORD()
            GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value not in pids or win_class(hwnd) != "#32770":
                return True
            text = dialog_text(hwnd)
            if "error" not in text.lower() and "visual basic" not in text.lower():
                return True
            if text not in self.seen:
                self.seen.add(text)
                hits.append(text)
            c = self.stuck.get(hwnd, 0) + 1
            self.stuck[hwnd] = c
            if c >= 5:
                try:
                    proc = kernel32.OpenProcess(1, False, pid.value)
                    if proc:
                        kernel32.TerminateProcess(proc, 1)
                        kernel32.CloseHandle(proc)
                        killed.append(pid.value)
                except Exception:
                    pass
            else:
                PostMessageW(hwnd, WM_COMMAND, IDOK, 0)
            return True

        EnumWindows(EnumProc(cb), 0)
        return hits, killed


def main():
    w = Watcher()
    write_log(f"watcher started pid={os.getpid()}")
    print(f"Watching Excel/VBA dialogs -> {log_path}")
    deadline = time.time() + 2 * 60 * 60
    try:
        while time.time() < deadline:
            time.sleep(0.5)
            try:
                hits, killed = w.poll()
            except Exception as e:
                write_log(f"watcher poll error: {e}")
                continue
            for h in hits:
                write_log(f"DIALOG: {h}")
                print(f"Captured: {h}")
            for kp in killed:
                write_log(f"KILLED EXCEL pid={kp}")
    finally:
        write_log("watcher stopped")


if __name__ == "__main__":
    main()
