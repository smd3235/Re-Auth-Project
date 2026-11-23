"""Tests for scan-logs.py using temporary log directories."""

from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]
SCAN = ROOT / "scripts" / "scan-logs.py"


def _write(path: Path, text: str) -> None:
    path.write_text(text, encoding="utf-8")


def test_clean_run_logs_no_issues(tmp_path: Path) -> None:
    log_dir = tmp_path / "log"
    log_dir.mkdir()
    stamp = "2026-11-20 14:18:00"
    _write(
        log_dir / "reauth_2026-11-20_141800.log",
        f"{stamp} | === Re-Auth report run started ===\n{stamp} | Final outcome: SUCCESS\n",
    )
    result = subprocess.run(
        [sys.executable, str(SCAN), "--log-dir", str(log_dir)],
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert "No current issues found" in result.stdout


def test_failed_run_log_is_flagged(tmp_path: Path) -> None:
    log_dir = tmp_path / "log"
    log_dir.mkdir()
    stamp = "2026-11-20 14:18:00"
    _write(
        log_dir / "reauth_2026-11-20_141800.log",
        f"{stamp} | === Re-Auth report run started ===\n{stamp} | Final outcome: FAILED\n",
    )
    result = subprocess.run(
        [sys.executable, str(SCAN), "--log-dir", str(log_dir)],
        capture_output=True,
        text=True,
    )
    assert result.returncode != 0
    assert "run_failed" in result.stdout


def test_json_output_is_valid(tmp_path: Path) -> None:
    log_dir = tmp_path / "log"
    log_dir.mkdir()
    stamp = "2026-11-20 14:18:00"
    _write(
        log_dir / "reauth_2026-11-20_141800.log",
        f"{stamp} | === Re-Auth report run started ===\n{stamp} | Final outcome: SUCCESS\n",
    )
    result = subprocess.run(
        [sys.executable, str(SCAN), "--log-dir", str(log_dir), "--json"],
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    import json

    payload = json.loads(result.stdout)
    assert payload["summary"]["errors"] == 0
