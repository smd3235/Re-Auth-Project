# scan-logs.py - scan Re-Auth runtime logs for errors, warnings, and anomalies.
# Non-destructive (read-only). Checks log/*.log and log/*.txt plus Reports/ outputs.
import argparse
from collections import defaultdict
from datetime import datetime
import glob
import json
import os
import re
import sys

parser = argparse.ArgumentParser(description="Scan Re-Auth logs for issues and anomalies")
parser.add_argument("--log-dir", default=None)
parser.add_argument("--reports-dir", default=None)
parser.add_argument("--json", action="store_true", help="emit machine-readable JSON summary")
parser.add_argument(
    "--history",
    action="store_true",
    help="also list issues from logs older than the last successful run",
)
parser.add_argument("--skip-syntax", action="store_true", help="skip the Python syntax check")
parser.add_argument("--skip-vba-lint", action="store_true", help="skip the pyvbaanalysis VBA lint")
parser.add_argument(
    "--repair",
    action="store_true",
    help="repair corrupted logs: strip synthetic headers and restore chronological order (.orig backup)",
)
args = parser.parse_args()

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
log_dir = os.path.abspath(args.log_dir or os.path.join(project_root, "log"))
reports_dir = os.path.abspath(args.reports_dir or os.path.join(project_root, "Reports"))

issues: list[dict] = []


def add(sev, kind, path, detail, line_no=None):
    try:
        mtime = os.path.getmtime(path)
    except OSError:
        mtime = 0.0
    issues.append(
        {
            "severity": sev,
            "kind": kind,
            "file": os.path.relpath(path, project_root) if path else "",
            "line": line_no,
            "detail": detail,
            "mtime": mtime,
            "historical": False,
        }
    )


def read_lines(path):
    try:
        with open(path, encoding="utf-8", errors="replace") as f:
            return f.read().splitlines()
    except Exception as e:
        add("error", "read_fail", path, str(e))
        return None


log_files = []
if os.path.isdir(log_dir):
    for p in glob.glob(os.path.join(log_dir, "*")):
        if os.path.isfile(p) and p.lower().endswith((".log", ".txt")):
            log_files.append(p)
log_files.sort()

SYNTH_LINE_RE = re.compile(
    r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} \| (Log file: .*|Display order: newest first)$"
)
TS_PREFIX = re.compile(r"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})")


def ts_of(line):
    m = TS_PREFIX.match(line)
    if not m:
        return None
    try:
        return datetime.strptime(m.group(1), "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return None


def runlog_dt(path):
    m = re.search(r"reauth_(\d{4}-\d{2}-\d{2})_(\d{6})\.log$", os.path.basename(path))
    if not m:
        return None
    try:
        return datetime.strptime(m.group(1) + " " + m.group(2), "%Y-%m-%d %H%M%S")
    except ValueError:
        return None


def repair_log(path):
    """Strip headers injected by old view-logs.py and restore chronological order."""
    try:
        with open(path, encoding="utf-8", errors="replace") as f:
            original = f.read().splitlines()
    except OSError:
        return False
    kept = [ln for ln in original if not SYNTH_LINE_RE.match(ln)]
    if not any(ts_of(ln) for ln in kept):
        return False
    ordered = sorted(kept, key=lambda ln: ts_of(ln) or datetime.max)
    if ordered == original:
        return False
    backup = path + ".orig"
    if not os.path.exists(backup):
        try:
            with open(backup, "w", encoding="utf-8") as f:
                f.write("\n".join(original) + "\n")
        except OSError:
            return False
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(ordered) + "\n")
    except OSError:
        return False
    return True


if args.repair:
    repaired = 0
    for p in log_files:
        try:
            if repair_log(p):
                repaired += 1
                print(f"Repaired: {os.path.relpath(p, project_root)}")
        except Exception as e:
            print(f"Repair failed: {p}: {e}", file=sys.stderr)
    if repaired:
        print(f"Repaired {repaired} log file(s) (backups written as *.orig)")
        log_files = []
        for p in glob.glob(os.path.join(log_dir, "*")):
            if os.path.isfile(p) and p.lower().endswith((".log", ".txt")):
                log_files.append(p)
        log_files.sort()

reauth_logs = [
    p
    for p in log_files
    if re.search(r"reauth_\d{4}-\d{2}-\d{2}_\d{6}\.log$", os.path.basename(p), re.IGNORECASE)
]

# Parse run logs
run_summaries: list[dict] = []
for path in reauth_logs:
    lines = read_lines(path)
    if lines is None:
        continue
    name = os.path.basename(path)
    text = "\n".join(lines)
    outcomes = [
        ln
        for ln in lines
        if "Final outcome:" in ln or "| Outcome:" in ln or ln.strip().startswith("Outcome:")
    ]
    started = any(
        "=== Re-Auth report run started ===" in ln
        or "Auto source:" in ln
        or "Workbook opened" in ln
        for ln in lines
    )
    outcome = outcomes[-1].split("|", 1)[-1].strip() if outcomes else ""
    summary = {"file": name, "path": path, "outcome": outcome, "started": started}
    run_summaries.append(summary)

    if outcomes:
        if "SUCCESS" not in outcome:
            add("error", "run_failed", path, outcome)
    elif started:
        add("error", "run_incomplete", path, "run log has no final outcome")
    for i, ln in enumerate(lines, 1):
        if "WATCHDOG: exceeded" in ln:
            add("error", "watchdog_timeout", path, ln.strip(), i)
        if "Excel.Run exception" in ln or "Traceback (most recent call last)" in ln:
            add("error", "run_exception", path, ln.strip(), i)
        if "Closing workbook/Excel" in ln and not outcomes and started:
            # incomplete/aborted mid-run context, already captured as run_incomplete
            pass
        if "~$" in ln and "new report file(s) detected" in ln.lower():
            add("warning", "lockfile_detected", path, ln.strip(), i)
        # real VBA errors (exclude spurious Err=0 lines)
        if "| ERROR |" in ln:
            m = re.search(r"Err=(\-?\d+)", ln)
            if m and m.group(1) != "0":
                add("error", "vba_error", path, ln.strip(), i)
        if "CreateReAuthReport failed" in ln and "Err=0" not in ln:
            add("error", "macro_failed", path, ln.strip(), i)

# compile-test and watcher
compile_log = os.path.join(log_dir, "compile-test.log")
if os.path.exists(compile_log):
    lines = read_lines(compile_log) or []
    if lines:
        last_result = [ln for ln in lines if "RESULT:" in ln]
        if last_result and "COMPILE FAILED" in last_result[-1]:
            add("error", "compile_failed", compile_log, last_result[-1].strip())
        # unresolved failures: last compile event failed
        comp_events = [ln for ln in lines if "Compile test started" in ln or "RESULT:" in ln]
        if comp_events and "COMPILE FAILED" in comp_events[-1]:
            add("error", "compile_failed_latest", compile_log, comp_events[-1].strip())
watcher_log = os.path.join(log_dir, "watcher.log")
if os.path.exists(watcher_log):
    for i, ln in enumerate(read_lines(watcher_log) or [], 1):
        if "DIALOG:" in ln and "Microsoft Excel" in ln and "error" in ln.lower():
            add("warning", "excel_dialog", watcher_log, ln.strip(), i)
        if "KILLED EXCEL" in ln:
            add("warning", "excel_killed", watcher_log, ln.strip(), i)

ans_log = os.path.join(log_dir, "answer-dialogs.log")
if os.path.exists(ans_log):
    for i, ln in enumerate(read_lines(ans_log) or [], 1):
        if "answerer error" in ln.lower():
            add("warning", "answerer_error", ans_log, ln.strip(), i)

# Harness-level failures that should never pass silently


def _line_dt(line):
    m = TS_PREFIX.match(line)
    if not m:
        return None
    try:
        return datetime.strptime(m.group(1), "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return None


HARNESS_PATTERNS = [
    ("PREFLIGHT FAIL", "error", "preflight_failed"),
    ("THREAD ERROR", "error", "thread_error"),
    ("Output verify FAILED", "error", "output_invalid"),
    ("Output verify error", "error", "output_verify_error"),
]
for path in reauth_logs:
    lines = read_lines(path) or []
    for i, ln in enumerate(lines, 1):
        for needle, sev, kind in HARNESS_PATTERNS:
            if needle in ln:
                # cutoff applied later via the generic historical pass on file mtime,
                # but multiples lines in old files are all historical anyway.
                add(sev, kind, path, ln.strip(), i)

# VBA debug log errors (macro-debug.log may be current even when its old lines are not)
debug_log_path = os.path.join(log_dir, "macro-debug.log")
if os.path.exists(debug_log_path):
    for i, ln in enumerate(read_lines(debug_log_path) or [], 1):
        if "| ERROR |" in ln:
            m = re.search(r"Err=(-?\d+)", ln)
            if m and m.group(1) != "0":
                add("error", "vba_error_debug", debug_log_path, ln.strip(), i)
        if "LOG WRITE FAILED" in ln or "COMPLETION MARKER WRITE FAILED" in ln:
            add("error", "log_write_failed", debug_log_path, ln.strip(), i)

# sync-template.log
sync_log = os.path.join(log_dir, "sync-template.log")
if os.path.exists(sync_log):
    slines = read_lines(sync_log) or []
    if slines:
        finish = [ln for ln in slines if "Sync finished OK" in ln or "ERROR" in ln]
        if finish and "ERROR" in finish[-1]:
            add("error", "sync_failed", sync_log, finish[-1].strip())

# Python syntax check
if not args.skip_syntax:
    import py_compile
    import tempfile

    pyc_target = os.path.join(tempfile.gettempdir(), "_reauth_scan_check.pyc")
    for p in sorted(glob.glob(os.path.join(project_root, "scripts", "*.py"))):
        try:
            py_compile.compile(p, cfile=pyc_target, doraise=True)
        except Exception as e:
            msg = str(e).strip().splitlines()
            add("error", "python_syntax", p, msg[-1] if msg else repr(e))

# VBA lint via pyvbaanalysis (a real VBA parser); best-effort structural checks
# are kept for things the parser does not flag.


def _offset_to_line(text: str, offset: int) -> int:
    return text.count("\n", 0, offset) + 1


if not args.skip_vba_lint:
    try:
        import pyvbaanalysis  # noqa: F401  (availability marker)
        from pyvbaanalysis import (
            DiagnosticSeverity,
            ModuleInput,
            ModuleSymbolKind,
            analyze_project,
        )
    except ImportError:
        add(
            "warning",
            "vba_lint_unavailable",
            project_root,
            "install pyvbaanalysis for VBA syntax checks",
        )
    else:
        SEVERITY_MAP = {
            DiagnosticSeverity.ERROR: "error",
            DiagnosticSeverity.WARNING: "warning",
            DiagnosticSeverity.INFORMATION: "info",
        }

        modules: list[ModuleInput] = []
        sources: dict[str, str] = {}
        path_by_module: dict[str, str] = {}
        for p in sorted(glob.glob(os.path.join(project_root, "vba", "mod*.bas"))):
            try:
                with open(p, encoding="utf-8", errors="replace") as f:
                    text = f.read()
            except OSError as e:
                add("error", "vba_read_fail", p, str(e))
                continue
            module_name = os.path.splitext(os.path.basename(p))[0]
            modules.append(ModuleInput(module_name, ModuleSymbolKind.STANDARD, text))
            sources[module_name] = text
            path_by_module[module_name] = p

            if "Option Explicit" not in text:
                add("error", "vba_lint", p, "missing 'Option Explicit'")
            if "Attribute VB_Name" not in text:
                add("error", "vba_lint", p, "missing 'Attribute VB_Name'")
            bad = [
                (i, ch) for i, ln in enumerate(text.splitlines(), 1) for ch in ln if ord(ch) > 127
            ]
            if bad:
                i, ch = bad[0]
                nbad = sum(1 for ln in text.splitlines() for ch in ln if ord(ch) > 127)
                add(
                    "warning",
                    "vba_non_ascii",
                    p,
                    f"{nbad} non-ASCII char(s), first {ch!r} at line {i}",
                )

        try:
            results = analyze_project(modules, whole_project=True)
        except Exception as e:
            add("error", "vba_lint_error", project_root, f"pyvbaanalysis failed: {e}")
            results = {}
        for module_name, diagnostics in results.items():
            p = path_by_module[module_name]
            text = sources[module_name]
            for d in diagnostics:
                sev = SEVERITY_MAP.get(d.severity, "warning")
                line = _offset_to_line(text, d.span.start)
                detail = d.message
                if d.spec_reference:
                    detail = f"{detail} (MS-VBAL {d.spec_reference})"
                add(sev, f"vba_{d.code}", p, detail, line)


# Run consistency and data-quality checks (ignore duplicated "VBA " tail lines)
TOTAL_RE = re.compile(
    r"Counts: Total=(\d+) PastDue=(\d+) Upcoming=(\d+) Denied=(\d+) Other=(\d+) Managers=(\d+)"
)
RAWROWS_RE = re.compile(r"Raw Data done rows=(\d+)")
SAVED_RE = re.compile(r"Saved: (.+?\.xlsx)")
saved_uses: dict[str, set[str]] = {}
for path in reauth_logs:
    lines = read_lines(path) or []
    plain = [ln for ln in lines if "| VBA " not in ln]
    totals = [tuple(map(int, m.groups())) for ln in plain for m in [TOTAL_RE.search(ln)] if m]
    rawrows = [int(m.group(1)) for ln in plain for m in [RAWROWS_RE.search(ln)] if m]
    if totals and rawrows:
        t, r = totals[-1][0], rawrows[-1]
        if r not in (t, t + 1):
            add("warning", "count_mismatch", path, f"Total={t} but Raw Data rows={r}")
    tabs = sum(1 for ln in plain if "Tab done:" in ln)
    fmt = sum(1 for ln in plain if "FormatManagerSheet done" in ln)
    if totals and tabs and totals[-1][5] != tabs:
        add(
            "warning",
            "manager_tab_mismatch",
            path,
            f"Managers={totals[-1][5]} but tabs={tabs}",
        )
    if fmt < tabs:
        add(
            "warning",
            "format_incomplete",
            path,
            f"FormatManagerSheet done={fmt} of {tabs} tabs",
        )
    stamps = [ts for ln in lines for ts in [ts_of(ln)] if ts]
    inversions = sum(1 for a, b in zip(stamps, stamps[1:]) if b < a)
    if inversions:
        add(
            "warning",
            "log_out_of_order",
            path,
            f"{inversions} out-of-order timestamp(s)",
        )
    if any("Display order: newest first" in ln for ln in lines):
        add(
            "warning",
            "log_header_corrupt",
            path,
            "contains synthetic view-logs header (run with --repair)",
        )
    for ln in plain:
        m = SAVED_RE.search(ln)
        if m:
            saved_uses.setdefault(os.path.basename(m.group(1).strip()), set()).add(
                os.path.basename(path)
            )

for name, users in sorted(saved_uses.items()):
    if len(users) > 1:
        target = os.path.join(log_dir, next(iter(users)))
        add(
            "warning",
            "report_name_collision",
            target,
            f"{name} referenced by {len(users)} runs: {', '.join(sorted(users))}",
        )

# Stale temp / lock files
for folder in (project_root, reports_dir):
    if not os.path.isdir(folder):
        continue
    for pat in ("~$*", "_run_temp.xlsm", "_compile_test.xlsm"):
        for p in glob.glob(os.path.join(folder, pat)):
            add("warning", "stale_temp_file", p, f"leftover file: {os.path.basename(p)}")

# Newest harness run: auto-source parity and Excel cleanup
harness_logs = []
for path in reauth_logs:
    lines = read_lines(path) or []
    if any("Final outcome:" in ln for ln in lines):
        harness_logs.append((path, lines))
if harness_logs:
    newest_harness, newest_lines = max(
        harness_logs, key=lambda item: runlog_dt(item[0]) or datetime.min
    )
    src_line = next((ln for ln in newest_lines if re.search(r"\| Source: ", ln)), "")
    auto_file = os.path.join(log_dir, "auto-source.txt")
    if src_line and os.path.exists(auto_file):
        logged_src = src_line.split("| Source:", 1)[1].strip()
        with open(auto_file, encoding="utf-8", errors="replace") as f:
            auto_src = f.read().strip()
        if logged_src != auto_src:
            add(
                "warning",
                "auto_source_mismatch",
                auto_file,
                f"auto-source points to {auto_src!r} but newest run used {logged_src!r}",
            )
    for ln in newest_lines:
        if "Excel cleanup:" in ln:
            m = re.search(r"(?:orphans|leftover)=\[([^\]]*)\]", ln)
            if m and m.group(1).strip():
                add("warning", "excel_orphan", newest_harness, ln.strip())


# completion marker sanity
marker = os.path.join(log_dir, "reauth-complete.txt")
if os.path.exists(marker):
    mlines = read_lines(marker) or []
    mtext = "\n".join(mlines)
    saved = ""
    for ln in mlines:
        if ln.startswith("saved="):
            saved = ln.split("=", 1)[1].strip()
    if "(failed)" in mtext.lower():
        add("error", "completion_failed", marker, saved or "(failed)")
    elif saved and os.path.isabs(saved) and not os.path.exists(saved):
        add("warning", "completion_missing_output", marker, saved)

# Reports sanity
report_files: list[str] = []
if os.path.isdir(reports_dir):
    for p in glob.glob(os.path.join(reports_dir, "*.xlsx")):
        if not os.path.basename(p).startswith("~$"):
            report_files.append(p)
if not report_files:
    add("warning", "no_reports", reports_dir, "No report .xlsx files found in Reports/")

# duplicate per-run pair heuristic (harness log and VBA run log share timestamp prefix)
prefixes: dict[str, int] = {}
for p in reauth_logs:
    m = re.match(r"(reauth_\d{4}-\d{2}-\d{2}_\d{4})", os.path.basename(p))
    if m:
        prefixes.setdefault(m.group(1), 0)
        prefixes[m.group(1)] += 1
# not an error, informational only in text mode

# newest-run quick status
outcome_runs = [s for s in run_summaries if s["outcome"]]
if outcome_runs:
    newest = max(outcome_runs, key=lambda s: runlog_dt(s["path"]) or datetime.min)["path"]
else:
    newest = max(reauth_logs, key=lambda p: runlog_dt(p) or datetime.min) if reauth_logs else None
newest_outcome = ""
if newest:
    nlines = read_lines(newest) or []
    nouts = [ln for ln in nlines if "Final outcome:" in ln or "Outcome:" in ln]
    newest_outcome = nouts[-1].split("|", 1)[-1].strip() if nouts else "(no outcome)"

if os.path.exists(marker) and "SUCCESS" in newest_outcome:
    try:
        with open(marker, encoding="utf-8", errors="replace") as f:
            marker_text = f.read()
        if "(failed)" in marker_text.lower():
            add(
                "warning",
                "marker_inconsistent",
                marker,
                "completion marker says failed but newest run outcome is SUCCESS",
            )
    except OSError:
        pass


# Anything older than the newest successful run is historical; report it
# separately so it does not mask problems in the current build. Use the run log
# filename timestamp (stable across repairs) and line timestamps where available.
def _runlog_dt(path):
    return runlog_dt(path)


success_dts = [
    t for s in run_summaries if "SUCCESS" in s["outcome"] for t in [_runlog_dt(s["path"])] if t
]
cutoff_dt = max(success_dts) if success_dts else None
for it in issues:
    if cutoff_dt is None:
        break
    abs_path = os.path.join(project_root, it["file"]) if it["file"] else ""
    event_dt = _runlog_dt(abs_path)
    if event_dt is None:
        event_dt = _line_dt(it["detail"] or "")
    if event_dt is None and it["mtime"]:
        event_dt = datetime.fromtimestamp(it["mtime"])
    if event_dt is not None and event_dt < cutoff_dt:
        it["historical"] = True

errors = [i for i in issues if i["severity"] == "error" and not i["historical"]]
warnings = [i for i in issues if i["severity"] == "warning" and not i["historical"]]
historical = [i for i in issues if i["historical"]]

summary = {
    "log_dir": log_dir,
    "reports_dir": reports_dir,
    "log_files_scanned": len(log_files),
    "reauth_run_logs": len(reauth_logs),
    "report_files": len(report_files),
    "errors": len(errors),
    "warnings": len(warnings),
    "historical_issues": len(historical),
    "newest_run_log": os.path.basename(newest) if newest else None,
    "newest_run_outcome": newest_outcome,
}

if args.json:
    print(json.dumps({"summary": summary, "issues": issues}, indent=2))
    sys.exit(1 if errors else 0)

print(
    f"Scanned {len(log_files)} log files ({len(reauth_logs)} run logs), {len(report_files)} report(s)"
)
if newest:
    print(f"Newest run: {os.path.basename(newest)} -> {newest_outcome}")
if not errors and not warnings:
    print("No current issues found.")
    if historical and not args.history:
        print(
            f"({len(historical)} historical issue(s) from logs older than the last successful run -- use --history to list)"
        )
    if not (args.history and historical):
        sys.exit(0)

# group by kind for compact output


def show_grouped(items, title):
    by_kind = defaultdict(list)
    for it in items:
        by_kind[(it["severity"], it["kind"])].append(it)
    for (sev, kind), group in sorted(by_kind.items()):
        print(f"\n{title}[{sev.upper()}] {kind} ({len(group)})")
        for it in group[:10]:
            loc = f":{it['line']}" if it.get("line") else ""
            detail = (it["detail"] or "").replace("\n", " ")
            if len(detail) > 220:
                detail = detail[:217] + "..."
            print(f"  - {it['file']}{loc} :: {detail}")
        if len(group) > 10:
            print(f"  ... and {len(group) - 10} more")


show_grouped(errors + warnings, "")
if historical and args.history:
    show_grouped(historical, "HISTORICAL ")

print(
    f"\nSummary: {len(errors)} current error(s), {len(warnings)} current warning(s), {len(historical)} historical"
)
sys.exit(1 if errors else 0)
