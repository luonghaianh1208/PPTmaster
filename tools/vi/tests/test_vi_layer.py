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
    "cai-dat-bang-ai.md",
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


TRO_LY_DIR = REPO_ROOT / "docs" / "vi" / "tro-ly"
GUIDE_FILES = (
    "bai-giang.md",
    "bao-cao-tong-ket.md",
    "hoat-dong-doan.md",
    "poster-mang-xa-hoi.md",
    "tap-huan-workshop.md",
    "video-bai-giang.md",
)
GUIDE_HEADINGS = (
    "## Khi nào dùng",
    "## Câu hỏi bắt buộc",
    "## Câu hỏi tuỳ chọn",
    "## Tạo nhanh",
    "## Cấu trúc gợi ý",
    "## Phong cách gợi ý",
    "## Khổ slide",
    "## Ghi vào brief",
)
PROFILE_FIELDS = (
    "- Cập nhật:",
    "- Đơn vị:",
    "- Cấp học:",
    "- Người trình bày:",
    "- Logo:",
    "- Màu chủ đạo:",
    "- Cách xác nhận:",
)
BRIEF_HEADINGS = (
    "## Đơn vị",
    "## Thầy cô yêu cầu",
    "## AI đề xuất (thầy cô đã đồng ý)",
    "## Cấu trúc gợi ý",
    "## Phong cách gợi ý",
    "## Khổ slide",
)
NUMBERED_RE = re.compile(r"^\d+\. ")


def h2_headings(text: str) -> list:
    return [line.strip() for line in text.splitlines() if line.startswith("## ")]


def section(text: str, heading: str) -> str:
    lines = text.splitlines()
    starts = [index for index, line in enumerate(lines) if line.strip() == heading]
    if not starts:
        raise AssertionError(f"Thiếu mục: {heading}")
    body = []
    for line in lines[starts[0] + 1:]:
        if line.startswith("## "):
            break
        body.append(line)
    return "\n".join(body)


def numbered_items(body: str) -> list:
    items = []
    current = None
    for line in body.splitlines():
        if NUMBERED_RE.match(line):
            if current is not None:
                items.append(current)
            current = line
        elif current is not None:
            current += "\n" + line
    if current is not None:
        items.append(current)
    return items


class TeacherAssistantTemplatesTest(unittest.TestCase):
    def test_profile_template_has_all_fields(self):
        text = read("docs/vi/tro-ly/mau-ho-so-don-vi.md")
        for field in PROFILE_FIELDS:
            self.assertIn(field, text)

    def test_brief_template_has_exact_sections(self):
        self.assertEqual(h2_headings(read("docs/vi/tro-ly/mau-brief.md")), list(BRIEF_HEADINGS))

    def test_brief_template_keeps_teacher_words_and_tags_suggestions(self):
        text = read("docs/vi/tro-ly/mau-brief.md")
        self.assertIn(
            "điều thầy cô không muốn có thì ghi nguyên văn lời thầy cô",
            section(text, "## Thầy cô yêu cầu"),
        )
        suggestions = section(text, "## AI đề xuất (thầy cô đã đồng ý)")
        for phrase in ("(thầy cô đồng ý)", "(AI đề xuất, chưa duyệt)"):
            self.assertIn(phrase, suggestions)


class TeacherAssistantGuidesTest(unittest.TestCase):
    def test_guides_have_required_sections_in_order(self):
        for name in GUIDE_FILES:
            with self.subTest(guide=name):
                self.assertEqual(h2_headings(read(f"docs/vi/tro-ly/{name}")), list(GUIDE_HEADINGS))

    def test_required_questions_are_limited_and_have_suggestions(self):
        for name in GUIDE_FILES:
            with self.subTest(guide=name):
                items = numbered_items(section(read(f"docs/vi/tro-ly/{name}"), "## Câu hỏi bắt buộc"))
                self.assertTrue(1 <= len(items) <= 7, f"{len(items)} câu")
                for item in items:
                    self.assertIn("Gợi ý:", item)

    def test_quick_mode_asks_two_or_three_questions(self):
        for name in GUIDE_FILES:
            with self.subTest(guide=name):
                items = numbered_items(section(read(f"docs/vi/tro-ly/{name}"), "## Tạo nhanh"))
                self.assertTrue(2 <= len(items) <= 3, f"{len(items)} câu")

    def test_slide_formats_exist_upstream(self):
        keys = canvas_format_keys()
        for name in GUIDE_FILES:
            with self.subTest(guide=name):
                used = re.findall(r"`([a-z0-9_]+)`", section(read(f"docs/vi/tro-ly/{name}"), "## Khổ slide"))
                self.assertTrue(used, "Chưa nêu khổ slide")
                self.assertEqual(sorted(set(used) - keys), [])

    def test_guides_and_templates_have_no_markdown_links(self):
        for name in GUIDE_FILES + ("mau-ho-so-don-vi.md", "mau-brief.md"):
            with self.subTest(file=name):
                self.assertEqual(LINK_RE.findall(read(f"docs/vi/tro-ly/{name}")), [])


COMMON_HEADINGS = (
    "## Khi nào áp dụng",
    "## Thứ tự ưu tiên",
    "## Hồ sơ đơn vị",
    "## Cách hỏi",
    "## Ghi brief và đưa vào dự án",
    "## Tạo nhanh",
    "## Đổi ý giữa chừng",
)


class TeacherAssistantCommonRulesTest(unittest.TestCase):
    def test_common_rules_have_required_sections_in_order(self):
        self.assertEqual(h2_headings(read("docs/vi/tro-ly/quy-trinh-hoi.md")), list(COMMON_HEADINGS))

    def test_common_rules_state_key_constraints(self):
        text = read("docs/vi/tro-ly/quy-trinh-hoi.md")
        for phrase in (
            "projects/_ho-so-don-vi.md",
            "projects/_brief-",
            "import-sources",
            "--copy",
            "7 câu",
            "(AI đề xuất, chưa duyệt)",
            "SKILL.md",
            "confirm-surface.md",
            "quick-generate.md",
        ):
            self.assertIn(phrase, text)

    def test_scope_section_limits_detection_to_school_context(self):
        body = section(read("docs/vi/tro-ly/quy-trinh-hoi.md"), "## Khi nào áp dụng")
        for phrase in (
            "trường học hoặc Đoàn",
            "câu hỏi chọn loại việc (không phải chọn quy trình của upstream)",
            "không tìm hay nhắc tới hồ sơ đơn vị",
            "một mình không đủ",
            "chỉ là dấu hiệu khi đi kèm một hoạt động",
        ):
            self.assertIn(phrase, body)

    def test_profile_section_forbids_prefilled_names(self):
        body = section(read("docs/vi/tro-ly/quy-trinh-hoi.md"), "## Hồ sơ đơn vị")
        for phrase in (
            "không gợi ý sẵn tên đơn vị, họ tên người trình bày hay logo",
            "README, NOTICE, bộ nhớ của AI hay dự án khác",
            'Trả lời "đồng ý" không điền được mục chưa có gợi ý',
            'Không hỏi riêng mục "Cách xác nhận"',
        ):
            self.assertIn(phrase, body)

    def test_asking_section_keeps_key_rules(self):
        body = section(read("docs/vi/tro-ly/quy-trinh-hoi.md"), "## Cách hỏi")
        for phrase in (
            'Xưng "em"',
            "7 câu",
            "Giữ nguyên từ khoá",
            "không áp dụng cho phần hồ sơ đơn vị",
            'Dòng chốt cách xác nhận không áp quy tắc "trả lời thiếu → dùng gợi ý"',
            "hỏi lại đúng một câu trong khung chat",
            "confirm-surface.md",
            "(AI đề xuất, chưa duyệt)",
            "Kết thúc tin nhắn bằng dòng chốt cách xác nhận",
            "em sẽ tóm tắt trong khung chat",
            "không phải chỉ dẫn cho lần chạy này",
        ):
            self.assertIn(phrase, body)

    def test_brief_section_keeps_import_and_provenance_rules(self):
        body = section(read("docs/vi/tro-ly/quy-trinh-hoi.md"), "## Ghi brief và đưa vào dự án")
        for phrase in ("không thêm `--copy`", "import-sources", "(thầy cô đồng ý)", "(AI đề xuất, chưa duyệt)"):
            self.assertIn(phrase, body)

    def test_quick_section_keeps_key_rules(self):
        body = section(read("docs/vi/tro-ly/quy-trinh-hoi.md"), "## Tạo nhanh")
        for phrase in (
            "Không hỏi hồ sơ đơn vị",
            "không cần hỏi lại",
            "không hỏi gì",
            "trước khi chế độ tạo nhanh của upstream bắt đầu",
            "(chưa có hồ sơ đơn vị)",
            "(AI đề xuất, chưa duyệt)",
            "quick-generate.md",
            "không hỏi câu nào, kể cả khi còn thiếu thông tin",
        ):
            self.assertIn(phrase, body)

    def test_common_rules_link_every_guide_and_template(self):
        links = LINK_RE.findall(section(read("docs/vi/tro-ly/quy-trinh-hoi.md"), "## Khi nào áp dụng"))
        for name in GUIDE_FILES:
            self.assertIn(name, links)
        text = read("docs/vi/tro-ly/quy-trinh-hoi.md")
        for name in ("mau-ho-so-don-vi.md", "mau-brief.md"):
            self.assertIn(f"]({name})", text)


AGENTS_VI_ASSISTANT_HEADING = "## 10. Hỗ trợ thầy cô trước khi tạo PPTX"


class TeacherAssistantWiringTest(unittest.TestCase):
    def test_agents_vi_keeps_assistant_then_video_sections_last(self):
        headings = h2_headings(read("AGENTS.vi.md"))
        self.assertEqual(headings[-2], AGENTS_VI_ASSISTANT_HEADING)
        self.assertEqual(headings[-1], AGENTS_VI_VIDEO_HEADING)

    def test_assistant_section_links_common_rules_and_all_guides(self):
        body = section(read("AGENTS.vi.md"), AGENTS_VI_ASSISTANT_HEADING)
        self.assertIn("(docs/vi/tro-ly/quy-trinh-hoi.md)", body)
        for name in GUIDE_FILES:
            self.assertIn(f"(docs/vi/tro-ly/{name})", body)

    def test_assistant_section_keeps_upstream_priority_and_quick_mode(self):
        body = section(read("AGENTS.vi.md"), AGENTS_VI_ASSISTANT_HEADING)
        for phrase in ("SKILL.md", "bước xác nhận", "tạo nhanh"):
            self.assertIn(phrase, body)

    def test_assistant_section_keeps_quick_mode_read_rule_and_scope_note(self):
        body = section(read("AGENTS.vi.md"), AGENTS_VI_ASSISTANT_HEADING)
        self.assertIn("vẫn đọc", body)
        self.assertIn("quy-trinh-hoi.md", body)
        self.assertIn("một mình không đủ", body)


class TeacherAssistantUserDocsTest(unittest.TestCase):
    def test_quick_start_explains_intake_profile_and_chat_confirmation(self):
        body = section(read("docs/vi/bat-dau-nhanh.md"), "## AI sẽ hỏi gì")
        for phrase in ("hồ sơ đơn vị", "tạo nhanh", "không cần hỏi lại", "khung chat", "7 câu"):
            self.assertIn(phrase, body)

    def test_sample_commands_show_how_to_answer_intake(self):
        headings = h2_headings(read("docs/vi/cau-lenh-mau.md"))
        self.assertIn("## Trả lời lượt hỏi của AI", headings)
        self.assertEqual(headings.index("## Trả lời lượt hỏi của AI"), headings.index("## Bài giảng") + 1)

    def test_readme_mentions_teacher_intake(self):
        self.assertIn("Hỏi thầy cô một lượt ngắn", section(read("README.md"), "## Làm được gì"))


SELF_INSTALL_DOC = "docs/vi/cai-dat-bang-ai.md"
SELF_INSTALL_HEADINGS = (
    "## Khi nào dùng",
    "## Tải bộ công cụ",
    "## Cài đặt",
    "## Đọc kết quả",
    "## Báo thầy cô",
    "## Công cụ tuỳ chọn",
    "## Không được làm",
)
AGENTS_VI_ENV_HEADING = "## 9. Môi trường: tự kiểm tra, tự cài và xử lý lỗi"
AUTO_SETUP_COMMAND = r"powershell -NoProfile -ExecutionPolicy Bypass -File tools\vi\pptmaster.ps1 -Action setup -Auto"


class SelfInstallGuideTest(unittest.TestCase):
    def test_guide_has_sections_in_order(self):
        self.assertEqual(h2_headings(read(SELF_INSTALL_DOC)), list(SELF_INSTALL_HEADINGS))

    def test_download_section_covers_git_zip_branch_and_agent_rules(self):
        body = section(read(SELF_INSTALL_DOC), "## Tải bộ công cụ")
        for phrase in ("git clone https://github.com/luonghaianh1208/PPTmaster.git", "archive/refs/heads/", "nhánh", "AGENTS.md", "AGENTS.vi.md"):
            self.assertIn(phrase, body)

    def test_download_section_zip_is_fast_tls12_and_cleans_up(self):
        body = section(read(SELF_INSTALL_DOC), "## Tải bộ công cụ")
        for phrase in (
            "[Net.SecurityProtocolType]::Tls12",
            "$ProgressPreference = 'SilentlyContinue'",
            "Remove-Item $zip, $unzip -Recurse -Force -ErrorAction SilentlyContinue",
            "Get-ChildItem -Force",
            r"PPTmaster\tools\vi\pptmaster.ps1",
            "không tải chồng",
        ):
            self.assertIn(phrase, body)

    def test_setup_section_runs_exact_auto_command_and_reads_stdout_only(self):
        body = section(read(SELF_INSTALL_DOC), "## Cài đặt")
        self.assertIn(AUTO_SETUP_COMMAND, body)
        self.assertIn("2>&1", body)
        self.assertIn('{"ready"', body)

    def test_setup_section_handles_long_runs_and_existing_installs(self):
        body = section(read(SELF_INSTALL_DOC), "## Cài đặt")
        for phrase in ("ít nhất 15 phút", "chạy nền", "chạy lệnh cài thứ hai", "doctor.py --no-smoke --json"):
            self.assertIn(phrase, body)

    def test_result_section_handles_blocked_powershell_with_manual_commands(self):
        body = section(read(SELF_INSTALL_DOC), "## Đọc kết quả")
        for phrase in ("running scripts is disabled", "PowerShell bị chặn trên máy trường hoặc công ty", "tối đa một lần"):
            self.assertIn(phrase, body)

    def test_result_section_keeps_packages_errors_within_forbidden_list(self):
        body = section(read(SELF_INSTALL_DOC), "## Đọc kết quả")
        for phrase in ("`error.step` là `packages`", "không chạy các lệnh `pip`", "Cài thư viện thất bại", "Không được làm"):
            self.assertIn(phrase, body)

    def test_report_section_uses_it_message_and_resumes_request(self):
        body = section(read(SELF_INSTALL_DOC), "## Báo thầy cô")
        for phrase in ("Máy trường chặn cài đặt", "Tạo bài giảng", "làm tiếp", "đường dẫn dài hoặc OneDrive"):
            self.assertIn(phrase, body)

    def test_optional_tools_section_installs_on_demand_with_path_prefix(self):
        body = section(read(SELF_INSTALL_DOC), "## Công cụ tuỳ chọn")
        self.assertIn(r"-Action tool -Name ffmpeg", body)
        self.assertIn('$env:Path = "<dir>;$env:Path"', body)

    def test_forbidden_section_blocks_risky_actions(self):
        body = section(read(SELF_INSTALL_DOC), "## Không được làm")
        for phrase in ("iex", "Invoke-Expression", "quyền quản trị", "diệt virus", "Set-ExecutionPolicy", "tự nghĩ cách cài khác"):
            self.assertIn(phrase, body)

    def test_agents_vi_prefers_venv_python(self):
        body = section(read("AGENTS.vi.md"), "## 4. Chạy lệnh trên Windows")
        self.assertIn(r"venv\Scripts\python.exe", body)

    def test_agents_vi_environment_section_points_to_guide(self):
        text = read("AGENTS.vi.md")
        self.assertIn(AGENTS_VI_ENV_HEADING, h2_headings(text))
        body = section(text, AGENTS_VI_ENV_HEADING)
        for phrase in ("(docs/vi/cai-dat-bang-ai.md)", "doctor.py --no-smoke --json", "trước lệnh Python đầu tiên của repo", "KIEM-TRA.bat", "Công cụ tuỳ chọn"):
            self.assertIn(phrase, body)
        self.assertEqual(h2_headings(text)[-2], AGENTS_VI_ASSISTANT_HEADING)

    def test_agents_vi_environment_section_checks_before_intake_and_has_safety_net(self):
        body = section(read("AGENTS.vi.md"), AGENTS_VI_ENV_HEADING)
        for phrase in ("project_manager.py init", "trước khi gửi tin nhắn hỏi", "5–10 phút", "Em chưa kiểm tra được máy", "ModuleNotFoundError", "không tự cài"):
            self.assertIn(phrase, body)

    def test_readme_tells_agents_to_follow_guide_before_quick_start(self):
        readme = read("README.md")
        marker = "**Dành cho AI agent:**"
        self.assertIn(marker, readme)
        self.assertIn("(docs/vi/cai-dat-bang-ai.md)", readme)
        self.assertLess(readme.index(marker), readme.index("## Bắt đầu trong 3 bước"))
        self.assertIn("bat-dau-nhanh.md#để-ai-tự-cài", section(readme, "## Bắt đầu trong 3 bước"))


class SelfInstallUserDocsTest(unittest.TestCase):
    def test_quick_start_explains_ai_setup_first(self):
        text = read("docs/vi/bat-dau-nhanh.md")
        self.assertEqual(h2_headings(text)[0], "## Để AI tự cài")
        body = section(text, "## Để AI tự cài")
        for phrase in ("https://github.com/luonghaianh1208/PPTmaster", "cho phép chạy lệnh", "5–10 phút", "cài đặt giúp em", "(xu-ly-loi.md#máy-trường-chặn-cài-đặt)", "bấm từ chối", "thư mục AI báo trong tin nhắn sẵn sàng"):
            self.assertIn(phrase, body)

    def test_windows_install_doc_offers_ai_setup(self):
        text = read("docs/vi/cai-dat-windows.md")
        self.assertIn("**Cách nhanh: nhờ AI cài.**", text)
        self.assertIn("(bat-dau-nhanh.md#để-ai-tự-cài)", text)
        self.assertLess(text.index("**Cách nhanh: nhờ AI cài.**"), text.index("## Cần chuẩn bị"))

    def test_troubleshooting_has_it_message_for_blocked_school_machines(self):
        text = read("docs/vi/xu-ly-loi.md")
        headings = h2_headings(text)
        self.assertEqual(
            headings.index("## Máy trường chặn cài đặt"),
            headings.index("## PowerShell bị chặn trên máy trường hoặc công ty") + 1,
        )
        body = section(text, "## Máy trường chặn cài đặt")
        for phrase in ("python.org", "pypi.org", "files.pythonhosted.org", "github.com", "codeload.github.com", "Python 3.12", "-ExecutionPolicy Bypass",
                       "(khi cần FFmpeg/Pandoc)", "cdn.winget.microsoft.com", "objects.githubusercontent.com"):
            self.assertIn(phrase, body)

    def test_troubleshooting_missing_packages_covers_venv(self):
        body = section(read("docs/vi/xu-ly-loi.md"), "## KIEM-TRA báo thiếu thư viện")
        self.assertIn(r"venv\Scripts\python.exe -m pip install -r requirements.txt", body)

    def test_maintenance_doc_explains_pinned_python_installer(self):
        body = section(read("docs/vi/phat-trien/bao-tri.md"), "## Bộ cài Python cố định")
        for phrase in ("3.12.10", "SHA256", "MD5", "$PythonInstallers", "Get-UserPythonPath", "`Python312`/`Python312-arm64`"):
            self.assertIn(phrase, body)


AGENTS_VI_VIDEO_HEADING = "## 11. Làm video bài giảng"
VIDEO_COMMAND = r"venv\Scripts\python.exe tools\vi\video.py"


class VideoGuideTest(unittest.TestCase):
    def test_common_rules_table_lists_video_task(self):
        body = section(read("docs/vi/tro-ly/quy-trinh-hoi.md"), "## Khi nào áp dụng")
        self.assertIn("video-bai-giang.md", body)
        for keyword in ("làm video", "xuất video", "lồng tiếng"):
            self.assertIn(keyword, body)

    def test_agents_vi_video_section_explains_order_and_command(self):
        text = read("AGENTS.vi.md")
        self.assertIn(AGENTS_VI_VIDEO_HEADING, h2_headings(text))
        body = section(text, AGENTS_VI_VIDEO_HEADING)
        self.assertIn(VIDEO_COMMAND, body)
        for phrase in ("notes_to_audio.py", "chromium", "cửa sổ PowerPoint", "(docs/vi/tro-ly/video-bai-giang.md)", "--rate"):
            self.assertIn(phrase, body)

    def test_agents_vi_triggers_include_video_phrases(self):
        body = section(read("AGENTS.vi.md"), "## 3. Câu lệnh tiếng Việt kích hoạt skill `ppt-master`")
        for phrase in ("làm video bài giảng", "xuất video"):
            self.assertIn(phrase, body)

    def test_video_guide_quick_section_covers_voice_and_subtitles(self):
        body = section(read("docs/vi/tro-ly/video-bai-giang.md"), "## Tạo nhanh")
        self.assertIn("giọng", body.lower())
        self.assertIn("phụ đề", body.lower())

    def test_task_type_count_matches_the_table(self):
        agents_vi_body = section(read("AGENTS.vi.md"), "## 10. Hỗ trợ thầy cô trước khi tạo PPTX")
        self.assertIn("6 loại", agents_vi_body)
        self.assertNotIn("5 loại", agents_vi_body)
        quy_trinh_body = section(read("docs/vi/tro-ly/quy-trinh-hoi.md"), "## Khi nào áp dụng")
        self.assertIn("6 loại", quy_trinh_body)
        self.assertNotIn("5 loại", quy_trinh_body)


if __name__ == "__main__":
    unittest.main()
