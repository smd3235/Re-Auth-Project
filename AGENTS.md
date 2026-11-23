# AGENTS.md

Guidance for AI agents working in this repository.

## Project overview

Re-Auth report generator: a VBA macro (`CreateReAuthReport`) embedded in `Re-Auth Template.xlsm`
builds a Dashboard + per-manager tabs + Raw Data workbook from `reauth raw.csv`. Python scripts
drive Excel through COM for unattended runs, testing, verification, and diagnostics.

## Layout

- `Re-Auth Template.xlsm` - master macro-enabled template (START HERE + embedded VBA).
- `Re-Auth Template (No Macros).xlsx` - macro-free base for local builds.
- `reauth raw.csv` - sample/source input (20 columns, 145 data rows).
- `vba/modCore.bas` - `CreateReAuthReport` main workflow.
- `vba/modDeclarations.bas` - constants, headers, sheet names, shared runtime state.
- `vba/modFormatting.bas` - Dashboard/manager/Raw Data formatting.
- `vba/modLogging.bas` - `LogMsg`/`LogErr`/`LogStep`, completion marker, log fallback.
- `vba/modUtils.bas` - header/date/array helpers, paths, output naming.
- `scripts/` - Python tooling (see table below).
- `documentation/screenshots/` - committed README screenshots (do not delete unrelated assets).
- `Reports/`, `log/` - generated runtime artifacts (gitignored; archives under `archive/`).

## Prerequisites

Windows, desktop Excel with COM automation, Python 3, `pywin32`, `psutil`. Excel must allow
programmatic VBA access ("Trust access to the VBA project object model").

## Standard workflow

Run from the repository root:

```powershell
python .\scripts\sync_template.py      # import vba/mod*.bas into the template
python .\scripts\test-compile.py       # VBA compile probe -> COMPILE_OK=1
python .\scripts\run-report.py         # unattended report run -> OUTCOME=..., NEW_REPORT=...
python .\scripts\verify-report.py      # structure/content checks -> 0 errors
python .\scripts\scan-logs.py          # log/lint/syntax scan -> 0 current issues
```

Optional tools:

```powershell
python .\scripts\monitor-logs.py [--follow]
python .\scripts\view-logs.py            # newest-first copies in log/views
python .\scripts\snapshot-report.py      # export screenshots
python .\scripts\cleanup.py [--apply] [--archive]
python .\scripts\scan-logs.py --history  # include pre-fix runs
python .\scripts\scan-logs.py --repair   # fix corrupt/reordered logs (.orig backups)
```

## VBA editing rules

- Edit `vba/mod*.bas` only; never edit embedded VBA directly inside the `.xlsm`.
- After any `.bas` change, run `sync_template.py`, then `test-compile.py`, then `run-report.py`.
- Keep `Option Explicit` and `Attribute VB_Name` intact.
- Use centralized logging: `LogMsg`, `LogErr`, `LogStep`; never leave silent `On Error Resume Next`
  without an `If Err.Number <> 0 Then LogErr ...` guard.
- Error handlers must pass captured values or be guarded; do not emit `Err=0` error lines.
- Log messages should include useful identifiers (manager, sheet, column, step).

## Python rules

- All scripts must byte-compile (`python -m py_compile scripts\*.py` individually; wildcards are not
  expanded).
- Lint/format: `python -m ruff check scripts` and `python -m ruff format --check scripts` (config in
  `pyproject.toml`).
- Type check:
  `python -m mypy scripts\scan-logs.py scripts\cleanup.py scripts\view-logs.py scripts\monitor-logs.py`
  (COM-heavy scripts are excluded from mypy config).
- Tests: `python -m pytest` (`tests/`).
- Do not use a global `taskkill /IM excel.exe`. Release COM references
  (`wb = None; xl = None; gc.collect()`), wait, then kill only PIDs the script started.
  `--force-kill-excel` is destructive opt-in.
- Non-interactive by default; cleanup is a dry run unless `--apply`.
- Write logs under `log/` with timestamps; keep formats parseable (`timestamp | message`).

## Logs and artifacts

Expected files after a clean pipeline: `log/sync-template.log`, `log/compile-test.log`,
`log/watcher.log`, `log/macro-debug.log`, `log/answer-dialogs.log`, `log/verify-report.log`,
`log/reauth_<stamp>.log` (harness and VBA), `log/auto-source.txt`, `log/reauth-complete.txt`,
`log/probe_ran.txt`.

- Run outputs: `Reports/ReAuth_Report_YYYY-mm-dd_HHMMSS.xlsx` (seconds prevent same-minute
  overwrite).
- Completion marker contains `completed`, `saved`, `run_log`, `step`, `version`.
- Retention: 20 newest run logs and 5 newest reports; older artifacts archived by `cleanup.py`.

## Validation checklist before claiming success

1. `test-compile.py` prints `COMPILE_OK=1`.
1. `run-report.py` prints `OUTCOME=SUCCESS` and a `NEW_REPORT=` path.
1. `verify-report.py` reports 0 errors/warnings.
1. `scan-logs.py` reports 0 current issues.
1. Confirm no `Err=0` lines and no `| ERROR |` lines in the latest run logs.

## Known pitfalls

- VBA creates its own per-run log (`reauth_<stamp>.log`) separate from the harness log; harness logs
  duplicate the VBA tail with a `VBA ` prefix. Ignore prefixed lines when counting tabs/steps.
- `view-logs.py` is non-destructive by default; `--in-place` rewrites originals with `.orig`
  backups.
- `scan-logs.py --repair` restores chronological order and strips old synthetic headers.
- `snapshot-report.py` only prunes its own `report-*.png` outputs; do not delete other screenshots.
- Back-to-back runs are supported; names use seconds. Avoid editing files while a run is active.
- Do not commit generated `Reports/` or `log/` content. Commit only when explicitly requested.
