"""Tests for view-logs.py non-destructive view generation."""

from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]
VIEW = ROOT / "scripts" / "view-logs.py"


def test_view_logs_creates_views_and_leaves_originals(tmp_path: Path) -> None:
    log_dir = tmp_path / "log"
    log_dir.mkdir()
    (log_dir / "test.log").write_text(
        "2026-11-20 14:00:00 | first\n2026-11-20 14:00:01 | second\n",
        encoding="utf-8",
    )

    result = subprocess.run(
        [sys.executable, str(VIEW), "--log-dir", str(log_dir)],
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr

    view = log_dir / "views" / "test.log"
    assert view.exists()
    viewed = view.read_text(encoding="utf-8").splitlines()
    assert viewed[3:] == ["2026-11-20 14:00:01 | second", "2026-11-20 14:00:00 | first"]

    # Original is untouched (oldest first)
    original = (log_dir / "test.log").read_text(encoding="utf-8").splitlines()
    assert original == ["2026-11-20 14:00:00 | first", "2026-11-20 14:00:01 | second"]
