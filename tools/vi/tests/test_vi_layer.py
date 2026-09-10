"""Kiểm tra tính nhất quán của lớp Việt hoá với upstream."""

import fnmatch
import re
import shutil
import subprocess
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
SKILL_DIR = REPO_ROOT / "skills" / "ppt-master"

ALLOWED_CHANGES = (
    "README.md",
    "CLAUDE.md",
    "NOTICE",
    ".gitattributes",
    "AGENTS.vi.md",
    "CHANGELOG-VI.md",
    "CAI-DAT.bat",
    "KIEM-TRA.bat",
    "CAP-NHAT.bat",
    ".cursor/rules/ppt-master-vi.mdc",
    ".agents/rules/ppt-master-vi.md",
    "tools/vi/*",
    "docs/vi/*",
)


def read(relative: str) -> str:
    return (REPO_ROOT / relative).read_text(encoding="utf-8")


def git(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        ["git", "-C", str(REPO_ROOT), *args],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )


def upstream_base_tag():
    if shutil.which("git") is None or not (REPO_ROOT / ".git").exists():
        return None
    proc = git("describe", "--tags", "--abbrev=0", "--match", "v[0-9]*", "--exclude", "*-vi*")
    tag = proc.stdout.strip()
    return tag if proc.returncode == 0 and tag else None


class AttributionTest(unittest.TestCase):
    def test_notice_credits_upstream_and_vietnamese_edition(self):
        notice = read("NOTICE")
        for expected in ("Hugo He", "https://github.com/hugohe3/ppt-master", "MIT", "Lương Hải Anh"):
            self.assertIn(expected, notice)

    def test_gitattributes_keeps_vietnamese_readme_on_merge(self):
        lines = [line.strip() for line in read(".gitattributes").splitlines()]
        self.assertIn("README.md merge=ours", lines)


class UpstreamBoundaryTest(unittest.TestCase):
    def test_changes_limited_to_vietnamese_layer(self):
        tag = upstream_base_tag()
        if tag is None:
            self.skipTest("Không có git hoặc tag upstream")
        proc = git("diff", "--name-only", tag)
        self.assertEqual(proc.returncode, 0, proc.stderr)
        changed = [path for path in proc.stdout.splitlines() if path]
        outside = [
            path for path in changed
            if not any(fnmatch.fnmatch(path, pattern) for pattern in ALLOWED_CHANGES)
        ]
        self.assertEqual(outside, [], "File upstream bị sửa ngoài danh sách cho phép")


if __name__ == "__main__":
    unittest.main()
