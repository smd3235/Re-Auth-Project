# view-logs.py - reorder runtime logs newest-first for easy review
import argparse
import os
import sys
from datetime import datetime

parser = argparse.ArgumentParser(description="Reorder runtime logs newest first with dated headers")
parser.add_argument("--log-dir", default=None)
args = parser.parse_args()

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
log_dir = args.log_dir or os.path.join(project_root, "log")

if not os.path.isdir(log_dir):
    print(f"Log directory not found: {log_dir}", file=sys.stderr)
    sys.exit(1)

stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
log_files = sorted([f for f in os.listdir(log_dir)
                    if os.path.isfile(os.path.join(log_dir, f)) and f.lower().endswith((".log", ".txt"))])
count = 0
for log_file in log_files:
    fpath = os.path.join(log_dir, log_file)
    try:
        with open(fpath, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
    except OSError:
        continue
    lines.reverse()
    header = [f"{stamp} | Log file: {log_file}", f"{stamp} | Display order: newest first", ""]
    with open(fpath, "w", encoding="utf-8") as f:
        f.write("\n".join(header + [ln.rstrip("\n\r") for ln in lines]) + "\n")
    print(f"Reordered: {fpath}")
    count += 1

print(f"Updated {count} runtime log files.")
