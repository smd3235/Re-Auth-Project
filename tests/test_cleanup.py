"""Tests for cleanup.py retention behavior."""

from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]
CLEANUP = ROOT / "scripts" / "cleanup.py"


def _touch(path: Path) -> None:
    path.touch()


def test_cleanup_dry_run_deletes_nothing(tmp_path: Path) -> None:
    log_dir = tmp_path / "log"
    reports_dir = tmp_path / "Reports"
    log_dir.mkdir()
    reports_dir.mkdir()

    for i in range(25):
        _touch(log_dir / f"reauth_2026-11-{i + 1:02d}_12{i:02d}00.log")
    for i in range(8):
        _touch(reports_dir / f"ReAuth_Report_2026-11-20_12{i:02d}00.xlsx")

    before = {p.name for p in log_dir.iterdir()} | {p.name for p in reports_dir.iterdir()}

    result = subprocess.run(
        [
            sys.executable,
            str(CLEANUP),
            "--log-dir",
            str(log_dir),
            "--reports-dir",
            str(reports_dir),
            "--keep-logs",
            "20",
            "--keep-reports",
            "5",
        ],
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert "DRY RUN" in result.stdout
    after = {p.name for p in log_dir.iterdir()} | {p.name for p in reports_dir.iterdir()}
    assert before == after


def test_cleanup_apply_prunes_to_retention(tmp_path: Path) -> None:
    log_dir = tmp_path / "log"
    reports_dir = tmp_path / "Reports"
    log_dir.mkdir()
    reports_dir.mkdir()

    for i in range(25):
        _touch(log_dir / f"reauth_2026-11-{i + 1:02d}_12{i:02d}00.log")
    for i in range(8):
        _touch(reports_dir / f"ReAuth_Report_2026-11-20_12{i:02d}00.xlsx")

    result = subprocess.run(
        [
            sys.executable,
            str(CLEANUP),
            "--apply",
            "--log-dir",
            str(log_dir),
            "--reports-dir",
            str(reports_dir),
            "--keep-logs",
            "20",
            "--keep-reports",
            "5",
        ],
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    remaining_logs = [p for p in log_dir.iterdir() if p.name.startswith("reauth_")]
    remaining_reports = [p for p in reports_dir.iterdir() if p.suffix == ".xlsx"]
    assert len(remaining_logs) == 20
    assert len(remaining_reports) == 5


def test_cleanup_apply_archives(tmp_path: Path) -> None:
    log_dir = tmp_path / "log"
    reports_dir = tmp_path / "Reports"
    log_dir.mkdir()
    reports_dir.mkdir()

    for i in range(25):
        _touch(log_dir / f"reauth_2026-11-{i + 1:02d}_12{i:02d}00.log")

    result = subprocess.run(
        [
            sys.executable,
            str(CLEANUP),
            "--apply",
            "--archive",
            "--log-dir",
            str(log_dir),
            "--keep-logs",
            "20",
        ],
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    archives = list((log_dir / "archive").iterdir())
    assert len(archives) == 1
    remaining_logs = [p for p in log_dir.iterdir() if p.name.startswith("reauth_")]
    assert len(remaining_logs) == 20
