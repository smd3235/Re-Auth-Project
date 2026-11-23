# view-logs.py - present runtime logs newest-first without corrupting originals.
# By default writes reordered copies to log/views/. Use --in-place to rewrite
# the originals (a one-time .orig backup is created before the first rewrite).
import argparse
from datetime import datetime
import os
import re
import sys

parser = argparse.ArgumentParser(description="Reorder runtime logs newest first for review")
parser.add_argument("--log-dir", default=None)
parser.add_argument(
    "--out-dir",
    default=None,
    help="destination for reordered copies (default: <log-dir>/views)",
)
parser.add_argument(
    "--in-place",
    action="store_true",
    help="rewrite originals instead of writing copies",
)
parser.add_argument("--no-header", action="store_true", help="omit the synthetic header block")
args = parser.parse_args()

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
log_dir = args.log_dir or os.path.join(project_root, "log")
out_dir = args.out_dir or os.path.join(log_dir, "views")

TS_RE = re.compile(r"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})")
SYNTH_RE = re.compile(
    r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} \| (Log file: .*|Display order: newest first)$"
)


def strip_synthetic(lines):
    """Remove headers injected by older versions of this script."""
    kept = []
    for ln in lines:
        s = ln.rstrip("\n\r")
        if SYNTH_RE.match(s):
            continue
        kept.append(s)
    while kept and not kept[0].strip():
        kept.pop(0)
    return kept


def order_newest_first(lines):
    """Sort by leading timestamp when available, else reverse file order."""
    stamped = [ln for ln in lines if TS_RE.match(ln)]
    if stamped and len(stamped) == len([ln for ln in lines if ln.strip()]):

        def ts_key(ln: str) -> str:
            m = TS_RE.match(ln)
            return m.group(1) if m else ""

        return sorted(lines, key=ts_key, reverse=True)
    return list(reversed(lines))


if not os.path.isdir(log_dir):
    print(f"Log directory not found: {log_dir}", file=sys.stderr)
    sys.exit(1)

log_files = sorted(
    [
        f
        for f in os.listdir(log_dir)
        if os.path.isfile(os.path.join(log_dir, f)) and f.lower().endswith((".log", ".txt"))
    ]
)
if not log_files:
    print(f"No .log/.txt files found in {log_dir}")
    sys.exit(0)

if not args.in_place:
    os.makedirs(out_dir, exist_ok=True)

stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
count = 0
for name in log_files:
    src = os.path.join(log_dir, name)
    dst = src if args.in_place else os.path.join(out_dir, name)
    try:
        with open(src, encoding="utf-8", errors="replace") as f:
            lines = strip_synthetic(f.readlines())
    except OSError as e:
        print(f"Skipped {src}: {e}", file=sys.stderr)
        continue

    ordered = order_newest_first(lines)
    header = (
        []
        if args.no_header
        else [
            f"{stamp} | Log file: {name}",
            f"{stamp} | Display order: newest first",
            "",
        ]
    )

    if args.in_place:
        backup = src + ".orig"
        if not os.path.exists(backup):
            try:
                with open(backup, "w", encoding="utf-8") as f:
                    f.write("\n".join(lines) + "\n")
            except OSError as e:
                print(f"Backup failed for {src}: {e}", file=sys.stderr)
                continue

    with open(dst, "w", encoding="utf-8") as f:
        f.write("\n".join(header + ordered) + "\n")
    print(f"{'Reordered' if args.in_place else 'Wrote view'}: {dst}")
    count += 1

print(f"Updated {count} runtime log file(s)." + ("" if args.in_place else f" Views in {out_dir}"))
