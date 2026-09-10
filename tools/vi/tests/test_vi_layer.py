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


def canvas_format_keys() -> set:
    text = (SKILL_DIR / "scripts" / "config.py").read_text(encoding="utf-8")
    block = re.search(r"^CANVAS_FORMATS\b[^\n]*\{(.*?)^\}", text, re.S | re.M)
    if block is None:
        raise AssertionError("Không tìm thấy CANVAS_FORMATS trong config.py")
    return set(re.findall(r"^\s{4}'([a-z0-9_]+)'\s*:", block.group(1), re.M))


def format_map_keys() -> list:
    block = re.search(r"<!-- format-map:start -->(.*?)<!-- format-map:end -->", read("AGENTS.vi.md"), re.S)
    if block is None:
        raise AssertionError("AGENTS.vi.md thiếu bảng format-map")
    return re.findall(r"`([a-z0-9_]+)`\s*\|\s*$", block.group(1), re.M)


class EditorWiringTest(unittest.TestCase):
    def test_claude_md_imports_upstream_and_vietnamese_rules(self):
        lines = [line.strip() for line in read("CLAUDE.md").splitlines()]
        self.assertIn("@AGENTS.md", lines)
        self.assertIn("@AGENTS.vi.md", lines)

    def test_format_map_uses_existing_canvas_keys(self):
        keys = format_map_keys()
        self.assertEqual(len(keys), 6)
        self.assertEqual(sorted(set(keys) - canvas_format_keys()), [])

    def test_cursor_rule_always_applies_vietnamese_rules(self):
        rule = read(".cursor/rules/ppt-master-vi.mdc")
        self.assertTrue(rule.startswith("---\n"))
        self.assertIn("alwaysApply: true", rule)
        self.assertIn("@AGENTS.vi.md", rule)

    def test_antigravity_rule_references_both_agent_files(self):
        rule = read(".agents/rules/ppt-master-vi.md")
        self.assertIn("trigger: always_on", rule)
        self.assertIn("@../../AGENTS.md", rule)
        self.assertIn("@../../AGENTS.vi.md", rule)

    def test_rule_files_tracked_by_git(self):
        if shutil.which("git") is None or not (REPO_ROOT / ".git").exists():
            self.skipTest("Không có git")
        for path in (".cursor/rules/ppt-master-vi.mdc", ".agents/rules/ppt-master-vi.md"):
            self.assertEqual(git("ls-files", "--error-unmatch", path).returncode, 0, path)


class ScriptEncodingTest(unittest.TestCase):
    def test_powershell_scripts_have_utf8_bom(self):
        scripts = sorted((REPO_ROOT / "tools" / "vi").glob("*.ps1"))
        self.assertTrue(scripts, "Chưa có script PowerShell")
        for path in scripts:
            self.assertTrue(path.read_bytes().startswith(b"\xef\xbb\xbf"), path.name)

    def test_batch_files_are_ascii_and_call_launcher(self):
        for name, action in (("CAI-DAT.bat", "setup"), ("KIEM-TRA.bat", "check"), ("CAP-NHAT.bat", "update")):
            text = (REPO_ROOT / name).read_bytes().decode("ascii")
            self.assertIn(r"tools\vi\pptmaster.ps1", text)
            self.assertIn(f"-Action {action}", text)


LINK_RE = re.compile(r"\]\(([^)\s]+)\)")
REQUIRED_DOCS = (
    "bat-dau-nhanh.md",
    "cai-dat-windows.md",
    "cau-lenh-mau.md",
    "xu-ly-loi.md",
    "lay-api-key.md",
    "phat-trien/bao-tri.md",
)


class DocsTest(unittest.TestCase):
    def test_required_vietnamese_docs_exist(self):
        for name in REQUIRED_DOCS:
            self.assertTrue((REPO_ROOT / "docs" / "vi" / name).is_file(), name)

    def test_troubleshooting_has_sections_referenced_by_launcher(self):
        text = read("docs/vi/xu-ly-loi.md")
        for heading in ("## Cài thư viện thất bại", "## Cập nhật thất bại", "## Đã cài Python nhưng bộ cài báo không tìm thấy"):
            self.assertIn(heading, text)

    def test_relative_markdown_links_resolve(self):
        files = [REPO_ROOT / "README.md", REPO_ROOT / "AGENTS.vi.md", REPO_ROOT / "CHANGELOG-VI.md"]
        files += [path for path in sorted((REPO_ROOT / "docs" / "vi").rglob("*.md")) if not path.name.startswith("20")]
        broken = []
        for markdown in files:
            for target in LINK_RE.findall(markdown.read_text(encoding="utf-8")):
                if target.startswith(("http://", "https://", "mailto:", "#")):
                    continue
                path = target.split("#", 1)[0]
                if path and not (markdown.parent / path).exists():
                    broken.append(f"{markdown.relative_to(REPO_ROOT)} -> {target}")
        self.assertEqual(broken, [])

    def test_readme_credits_upstream(self):
        readme = read("README.md")
        self.assertIn("Hugo He", readme)
        self.assertIn("https://github.com/hugohe3/ppt-master", readme)


if __name__ == "__main__":
    unittest.main()
