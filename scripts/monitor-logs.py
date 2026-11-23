# monitor-logs.py
# Realtime monitor for Re-Auth run logs.
# Tails the newest log\reauth_*.log and flags:
#   * error lines (VBA ERROR with Err<>0, macro failure, Excel.Run exception, watchdog)
#   * stalls (gap between consecutive log timestamps above the threshold)
# Also shows recent watcher.log dialog captures.
# Usage:
#   python scripts/monitor-logs.py                 # one-shot summary of newest run
#   python scripts/monitor-logs.py --follow        # tail live during a run (Ctrl+C to stop)
#   python scripts/monitor-logs.py --fragment 20261116
import argparse
import glob
import os
import re
import sys
import time

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
log_dir = os.path.join(project_root, "log")

TS_RE = re.compile(r"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})")
ERR_RE = re.compile(
    r"Excel\.Run exception|CreateReAuthReport failed|WATCHDOG: exceeded|Traceback \(most recent call last\)"
)


def is_error(line):
    if ERR_RE.search(line):
        return True
    if "| ERROR |" in line:
        m = re.search(r"Err=(-?\d+)", line)
        return not (m and m.group(1) == "0")
    return False


def parse_ts(line):
    m = TS_RE.match(line)
    if not m:
        return None
    stamp = m.group(1)
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M:%S,%f"):
        try:
            return time.mktime(time.strptime(stamp, fmt))
        except ValueError:
            pass
    return None


def newest_logs(fragment=""):
    files = [
        f
        for f in glob.glob(os.path.join(log_dir, "reauth_*.log"))
        if fragment in os.path.basename(f)
    ]
    files.sort(key=os.path.getmtime)
    return files


def watcher_tail(limit=5):
    path = os.path.join(log_dir, "watcher.log")
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8", errors="replace") as f:
        lines = [ln.rstrip("\n") for ln in f if ln.strip()]
    return lines[-limit:]


def follow(path, stall_ms=10000):
    print(f"Following {path} (stall threshold {stall_ms}ms). Ctrl+C to stop.")
    prev_ts = None
    with open(path, encoding="utf-8", errors="replace") as f:
        f.seek(0, os.SEEK_END)
        while True:
            line = f.readline()
            if not line:
                time.sleep(0.5)
                continue
            line = line.rstrip("\n")
            if not line.strip():
                continue
            ts = parse_ts(line)
            tag = ""
            if is_error(line):
                tag = " [ERROR]"
            if prev_ts is not None and ts is not None:
                gap_ms = int((ts - prev_ts) * 1000)
                if gap_ms > stall_ms:
                    tag += f" [STALL +{gap_ms}ms]"
            print(f"{line}{tag}" if tag else line)
            if ts is not None:
                prev_ts = ts


def summarize(fragment=""):
    files = newest_logs(fragment)
    if not files:
        print(
            f"No reauth_*.log found in {log_dir}" + (f" matching {fragment!r}" if fragment else "")
        )
        return 0
    path = files[-1]
    print(f"Newest run log: {os.path.basename(path)}  ({os.path.getsize(path)} bytes)")
    with open(path, encoding="utf-8", errors="replace") as f:
        lines = [ln.rstrip("\n") for ln in f if ln.strip()]
    if not lines:
        print("  (empty)")
        return 0
    errors = 0
    stalls = 0
    prev_ts = None
    prev_msg = None
    for ln in lines:
        ts = parse_ts(ln)
        if is_error(ln):
            errors += 1
            print(f"  ERROR: {ln}")
        if prev_ts is not None and ts is not None:
            gap_ms = int((ts - prev_ts) * 1000)
            if gap_ms > 10000:
                stalls += 1
                print(f"  STALL +{gap_ms}ms after: {prev_msg}")
        if ts is not None:
            prev_ts = ts
        prev_msg = ln
    outcome = ""
    for ln in lines:
        if "Final outcome:" in ln or "| Outcome:" in ln:
            outcome = ln
    print(f"  lines={len(lines)} errors={errors} stalls(>10000ms)={stalls}")
    if outcome:
        print(f"  outcome: {outcome}")
    else:
        print("  outcome: (none recorded)")
    wl = watcher_tail()
    if wl:
        print(f"  watcher.log recent ({len(wl)}):")
        for ln in wl:
            print(f"    {ln}")
    return 1 if errors else 0


def main():
    parser = argparse.ArgumentParser(description="Monitor Re-Auth run logs in realtime")
    parser.add_argument("--follow", "-f", action="store_true", help="tail the newest run log live")
    parser.add_argument(
        "--fragment", default="", help="only consider logs whose name contains this"
    )
    parser.add_argument("--stall-ms", type=int, default=10000, help="stall threshold in ms")
    args = parser.parse_args()

    if not os.path.isdir(log_dir):
        print(f"Log dir not found: {log_dir}")
        return 1
    if args.follow:
        files = newest_logs(args.fragment)
        if not files:
            print(
                f"No reauth_*.log found in {log_dir}"
                + (f" matching {args.fragment!r}" if args.fragment else "")
            )
            return 1
        try:
            follow(files[-1], stall_ms=args.stall_ms)
        except KeyboardInterrupt:
            print("\nStopped.")
            return 0
    return summarize(args.fragment)


if __name__ == "__main__":
    sys.exit(main())
