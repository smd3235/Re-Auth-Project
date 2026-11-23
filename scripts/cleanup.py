# cleanup.py - tidy runtime artifacts (logs, reports, temp files) with retention.
# Dry-run by default; pass --apply to make changes.
import argparse
from datetime import datetime
import glob
import os
import re
import shutil
import sys

parser = argparse.ArgumentParser(description="Clean up Re-Auth runtime artifacts")
parser.add_argument(
    "--apply", action="store_true", help="actually delete/move files (default: dry run)"
)
parser.add_argument(
    "--keep-logs",
    type=int,
    default=20,
    help="number of newest run logs to keep (default 20)",
)
parser.add_argument(
    "--keep-reports",
    type=int,
    default=5,
    help="number of newest report workbooks to keep (default 5)",
)
parser.add_argument(
    "--archive",
    action="store_true",
    help="zip removed logs/reports into log/archive and Reports/archive instead of deleting",
)
parser.add_argument(
    "--keep-backups",
    action="store_true",
    help="keep *.orig log backups and log/views copies",
)
parser.add_argument("--log-dir", default=None)
parser.add_argument("--reports-dir", default=None)
args = parser.parse_args()

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
log_dir = os.path.abspath(args.log_dir or os.path.join(project_root, "log"))
reports_dir = os.path.abspath(args.reports_dir or os.path.join(project_root, "Reports"))
log_archive = os.path.join(log_dir, "archive")
reports_archive = os.path.join(reports_dir, "archive")

RUN_LOG_RE = re.compile(r"^reauth_(\d{4}-\d{2}-\d{2})_(\d{6})\.log$")
KEEP_NAMED_REPORTS = {"ReAuth_Report_2026-11-16_1322.xlsx"}  # referenced by README screenshots

actions: list[tuple] = []


def plan_delete(path, reason):
    actions.append(("delete", path, reason))


def plan_zip(folder, paths, archive_name, reason):
    actions.append(("zip", (folder, paths, archive_name), reason))


def runlog_key(path):
    m = RUN_LOG_RE.match(os.path.basename(path))
    if not m:
        return None
    try:
        return datetime.strptime(m.group(1) + " " + m.group(2), "%Y-%m-%d %H%M%S")
    except ValueError:
        return None


def collect_run_logs():
    logs = []
    if os.path.isdir(log_dir):
        for p in glob.glob(os.path.join(log_dir, "reauth_*.log")):
            key = runlog_key(p)
            if key:
                logs.append((key, p))
    logs.sort(reverse=True)
    return logs


def collect_reports():
    files = []
    if os.path.isdir(reports_dir):
        for p in glob.glob(os.path.join(reports_dir, "*.xlsx")):
            if os.path.basename(p).startswith("~$"):
                continue
            files.append((os.path.getmtime(p), p))
    files.sort(reverse=True)
    return files


logs = collect_run_logs()
old_logs = [p for _, p in logs[args.keep_logs :]]
if old_logs:
    if args.archive:
        plan_zip(
            log_dir,
            old_logs,
            "reauth_logs_older.zip",
            f"{len(old_logs)} run log(s) beyond newest {args.keep_logs}",
        )
    else:
        for p in old_logs:
            plan_delete(p, f"run log beyond newest {args.keep_logs}")

reports = collect_reports()
kept_reports = {p for _, p in reports[: args.keep_reports]}
for _, p in reports:
    if os.path.basename(p) in KEEP_NAMED_REPORTS:
        kept_reports.add(p)
old_reports = [p for _, p in reports if p not in kept_reports]
if old_reports:
    if args.archive:
        plan_zip(
            reports_dir,
            old_reports,
            "reports_older.zip",
            f"{len(old_reports)} report(s) beyond newest {args.keep_reports}",
        )
    else:
        for p in old_reports:
            plan_delete(p, f"report beyond newest {args.keep_reports}")

if not args.keep_backups:
    for p in glob.glob(os.path.join(log_dir, "*.orig")):
        plan_delete(p, "log repair backup")
    views_dir = os.path.join(log_dir, "views")
    if os.path.isdir(views_dir):
        plan_delete(views_dir, "derived log views")

for folder in (project_root, reports_dir):
    for pat in ("_run_temp.xlsm", "_compile_test.xlsm", "~$*"):
        for p in glob.glob(os.path.join(folder, pat)):
            plan_delete(p, "temp/lock file")

if not actions:
    print("Nothing to clean up.")
    sys.exit(0)

print(("APPLYING" if args.apply else "DRY RUN (use --apply)") + f": {len(actions)} action(s)")
total_bytes = 0
for kind, target, reason in actions:
    if kind == "delete":
        if os.path.isdir(target):
            size = sum(
                os.path.getsize(os.path.join(dp, f)) for dp, _, fs in os.walk(target) for f in fs
            )
            label = target
        else:
            try:
                size = os.path.getsize(target)
            except OSError:
                size = 0
            label = target
        total_bytes += size
        print(f"  delete {os.path.relpath(label, project_root)} ({size} bytes) - {reason}")
        if args.apply:
            if os.path.isdir(target):
                shutil.rmtree(target)
            else:
                os.remove(target)
    else:
        folder, paths, archive_name = target
        archived_bytes = sum(os.path.getsize(p) for p in paths if os.path.exists(p))
        total_bytes += archived_bytes
        print(
            f"  archive {len(paths)} file(s) ({archived_bytes} bytes) -> {os.path.join(folder, 'archive', archive_name)} - {reason}"
        )
        if args.apply:
            os.makedirs(os.path.join(folder, "archive"), exist_ok=True)
            zip_base = os.path.join(folder, "archive", archive_name[:-4])
            staging = os.path.join(folder, "archive", "_staging")
            os.makedirs(staging, exist_ok=True)
            for p in paths:
                shutil.copy2(p, os.path.join(staging, os.path.basename(p)))
            shutil.make_archive(zip_base, "zip", staging)
            shutil.rmtree(staging)
            for p in paths:
                os.remove(p)

print(f"Reclaimable: {total_bytes / 1024:.1f} KiB")
if args.apply:
    with open(os.path.join(log_dir, "cleanup.log"), "a", encoding="utf-8") as f:
        f.write(
            f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | cleanup applied: {len(actions)} action(s), {total_bytes} bytes\n"
        )
