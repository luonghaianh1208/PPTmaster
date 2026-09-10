"""Unit test cho tools/vi/doctor.py (chỉ dùng thư viện chuẩn)."""

import importlib.metadata
import io
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import doctor  # noqa: E402


def make_pptx(path, slide_texts):
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        for index, text in enumerate(slide_texts, start=1):
            archive.writestr(f"ppt/slides/slide{index}.xml", f"<p:sld><a:t>{text}</a:t></p:sld>".encode("utf-8"))
            archive.writestr(f"ppt/slides/_rels/slide{index}.xml.rels", "<Relationships/>")


def completed(cmd, code=0, stdout="", stderr=""):
    return subprocess.CompletedProcess(cmd, code, stdout, stderr)


class ParseRequirementNamesTest(unittest.TestCase):
    def test_strips_versions_markers_extras_comments_and_options(self):
        text = (
            "# comment\n\n"
            "PyYAML>=6.0\n"
            "python-pptx==0.6.21  # inline\n"
            "requests[socks]~=2.31 ; python_version > '3.8'\n"
            "-r other.txt\n"
            "curl_cffi\n"
        )
        self.assertEqual(
            doctor.parse_requirement_names(text),
            ["PyYAML", "python-pptx", "requests", "curl_cffi"],
        )


class CheckPythonTest(unittest.TestCase):
    def test_accepts_310(self):
        self.assertTrue(doctor.check_python((3, 10, 0)).ok)

    def test_rejects_39_as_required_failure(self):
        result = doctor.check_python((3, 9, 18))
        self.assertFalse(result.ok)
        self.assertEqual(result.level, doctor.REQUIRED)
        self.assertIn("3.10", result.detail)


class CheckPackagesTest(unittest.TestCase):
    def _requirements(self, tmp, text):
        path = Path(tmp) / "requirements.txt"
        path.write_text(text, encoding="utf-8")
        return path

    def test_lists_only_missing_packages(self):
        def fake_find(name):
            if name == "flask":
                raise importlib.metadata.PackageNotFoundError(name)
            return object()

        with tempfile.TemporaryDirectory() as tmp:
            result = doctor.check_packages(self._requirements(tmp, "PyYAML>=6.0\nflask>=3.0\n"), find_dist=fake_find)
        self.assertFalse(result.ok)
        self.assertIn("flask", result.detail)
        self.assertNotIn("PyYAML", result.detail)

    def test_all_present(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = doctor.check_packages(self._requirements(tmp, "PyYAML>=6.0\nflask>=3.0\n"), find_dist=lambda name: object())
        self.assertTrue(result.ok)
        self.assertIn("2", result.detail)

    def test_missing_requirements_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = doctor.check_packages(Path(tmp) / "absent.txt", find_dist=lambda name: object())
        self.assertFalse(result.ok)
        self.assertEqual(result.level, doctor.REQUIRED)

    def test_requirements_with_utf8_bom_parses_first_package_name(self):
        seen = []

        def fake_find(name):
            seen.append(name)
            return object()

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "requirements.txt"
            path.write_bytes(b"\xef\xbb\xbfPyYAML>=6.0\nflask>=3.0\n")
            result = doctor.check_packages(path, find_dist=fake_find)
        self.assertTrue(result.ok)
        self.assertEqual(seen, ["PyYAML", "flask"])

    def test_requirements_with_invalid_utf8_bytes_reported_gracefully(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "requirements.txt"
            path.write_bytes(b"PyYAML>=6.0\n\xe0\xff\n")
            result = doctor.check_packages(path, find_dist=lambda name: object())
        self.assertFalse(result.ok)
        self.assertEqual(result.level, doctor.REQUIRED)


class CheckIntegrityTest(unittest.TestCase):
    def test_passes_on_exit_zero_and_runs_guard(self):
        calls = []

        def fake_run(cmd, **kwargs):
            calls.append(cmd)
            return completed(cmd, 0)

        result = doctor.check_integrity(run=fake_run, python="python")
        self.assertTrue(result.ok)
        self.assertTrue(calls[0][1].endswith("attribution_guard.py"))

    def test_fails_on_guard_exit_78(self):
        result = doctor.check_integrity(run=lambda cmd, **kwargs: completed(cmd, 78), python="python")
        self.assertFalse(result.ok)
        self.assertIn("78", result.detail)

    def test_fails_when_python_cannot_start(self):
        def fake_run(cmd, **kwargs):
            raise OSError("not found")

        self.assertFalse(doctor.check_integrity(run=fake_run, python="python").ok)


class CheckToolTest(unittest.TestCase):
    def test_missing_tool_keeps_level_and_uses_purpose_as_fix(self):
        result = doctor.check_tool("pandoc", "Pandoc", doctor.OPTIONAL, "Chỉ cần khi chuyển tài liệu", which=lambda name: None)
        self.assertFalse(result.ok)
        self.assertEqual(result.level, doctor.OPTIONAL)
        self.assertEqual(result.fix, "Chỉ cần khi chuyển tài liệu")

    def test_present_tool(self):
        result = doctor.check_tool("git", "Git", doctor.RECOMMENDED, "x", which=lambda name: "C:/git.exe")
        self.assertTrue(result.ok)


class EnvTest(unittest.TestCase):
    def test_find_env_file_returns_first_existing_candidate(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            first, second, third = root / "a" / ".env", root / "b" / ".env", root / "c" / ".env"
            for path in (second, third):
                path.parent.mkdir()
                path.write_text("X=1\n", encoding="utf-8")
            self.assertEqual(doctor.find_env_file([first, second, third]), second)

    def test_find_env_file_returns_none_when_missing(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertIsNone(doctor.find_env_file([Path(tmp) / ".env"]))

    def test_default_env_candidates_follow_upstream_order(self):
        cwd, home = Path("/work"), Path("/home/u")
        self.assertEqual(
            doctor.default_env_candidates(cwd, home),
            [cwd / ".env", doctor.SKILL_DIR / ".env", doctor.REPO_ROOT / ".env", home / ".ppt-master" / ".env"],
        )

    def test_read_env_file_parses_quotes_export_and_comments(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / ".env"
            path.write_text(
                'GEMINI_API_KEY="abc"\n# HIDDEN_API_KEY=zzz\nexport OPENAI_API_KEY=def\nEMPTY_API_KEY=\nnot a pair\n',
                encoding="utf-8",
            )
            self.assertEqual(
                doctor.read_env_file(path),
                {"GEMINI_API_KEY": "abc", "OPENAI_API_KEY": "def", "EMPTY_API_KEY": ""},
            )

    def test_read_env_file_empty_value_with_inline_comment_placeholder(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / ".env"
            path.write_text("GEMINI_API_KEY= # dán key vào đây\n", encoding="utf-8")
            self.assertEqual(doctor.read_env_file(path), {"GEMINI_API_KEY": ""})
            result = doctor.check_api_keys({}, doctor.read_env_file(path))
            self.assertFalse(result.ok)

    def test_read_env_file_strips_comment_after_closing_quote(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / ".env"
            path.write_text('A="x # y" # note\n', encoding="utf-8")
            self.assertEqual(doctor.read_env_file(path), {"A": "x # y"})

    def test_read_env_file_strips_unquoted_inline_comment(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / ".env"
            path.write_text("B=abc # note\n", encoding="utf-8")
            self.assertEqual(doctor.read_env_file(path), {"B": "abc"})

    def test_check_api_keys_counts_non_empty_keys_without_leaking_values(self):
        result = doctor.check_api_keys(
            {"OPENAI_API_KEY": "def", "PATH": "x"},
            {"GEMINI_API_KEY": "abc", "EMPTY_API_KEY": ""},
        )
        self.assertTrue(result.ok)
        self.assertEqual(result.level, doctor.OPTIONAL)
        self.assertIn("2", result.detail)
        for secret in ("abc", "def"):
            self.assertNotIn(secret, result.detail + result.fix)

    def test_check_api_keys_warns_when_no_key(self):
        result = doctor.check_api_keys({}, {})
        self.assertFalse(result.ok)
        self.assertIn("lay-api-key.md", result.fix)

    def test_env_file_has_bom_true(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / ".env"
            path.write_bytes(b"\xef\xbb\xbfGEMINI_API_KEY=abc\n")
            self.assertTrue(doctor.env_file_has_bom(path))

    def test_env_file_has_bom_false(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / ".env"
            path.write_bytes(b"GEMINI_API_KEY=abc\n")
            self.assertFalse(doctor.env_file_has_bom(path))

    def test_env_file_has_bom_missing_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertFalse(doctor.env_file_has_bom(Path(tmp) / "absent.env"))

    def test_check_api_keys_reports_bom_and_never_leaks_values(self):
        result = doctor.check_api_keys(
            {}, {"GEMINI_API_KEY": "abc"}, env_has_bom=True,
        )
        self.assertFalse(result.ok)
        self.assertEqual(result.level, doctor.OPTIONAL)
        self.assertIn("BOM", result.detail)
        self.assertIn("KIEM-TRA.bat", result.fix)
        self.assertNotIn("abc", result.detail + result.fix)


class VerifyPptxTest(unittest.TestCase):
    def test_accepts_single_slide_with_vietnamese_text(self):
        with tempfile.TemporaryDirectory() as tmp:
            pptx = Path(tmp) / "a.pptx"
            make_pptx(pptx, [doctor.SMOKE_TEXT])
            self.assertTrue(doctor.verify_pptx(pptx).ok)

    def test_rejects_wrong_slide_count(self):
        with tempfile.TemporaryDirectory() as tmp:
            pptx = Path(tmp) / "a.pptx"
            make_pptx(pptx, [doctor.SMOKE_TEXT, doctor.SMOKE_TEXT])
            result = doctor.verify_pptx(pptx)
        self.assertFalse(result.ok)
        self.assertIn("2", result.detail)

    def test_rejects_missing_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertFalse(doctor.verify_pptx(Path(tmp) / "none.pptx").ok)

    def test_rejects_broken_vietnamese_text(self):
        with tempfile.TemporaryDirectory() as tmp:
            pptx = Path(tmp) / "a.pptx"
            make_pptx(pptx, ["Kiem tra tieng Viet"])
            self.assertFalse(doctor.verify_pptx(pptx).ok)

    def test_rejects_non_zip(self):
        with tempfile.TemporaryDirectory() as tmp:
            pptx = Path(tmp) / "a.pptx"
            pptx.write_bytes(b"not a zip")
            self.assertFalse(doctor.verify_pptx(pptx).ok)

    def test_rejects_directory_path_without_raising(self):
        with tempfile.TemporaryDirectory() as tmp:
            fake_pptx = Path(tmp) / "a_dir.pptx"
            fake_pptx.mkdir()
            result = doctor.verify_pptx(fake_pptx)
        self.assertFalse(result.ok)

    def test_rejects_oserror_when_opening_archive(self):
        with tempfile.TemporaryDirectory() as tmp:
            pptx = Path(tmp) / "a.pptx"
            pptx.write_bytes(b"not a zip")
            with mock.patch.object(zipfile, "ZipFile", side_effect=PermissionError("locked")):
                result = doctor.verify_pptx(pptx)
        self.assertFalse(result.ok)


class RunSmokeTest(unittest.TestCase):
    def test_success_when_export_writes_valid_pptx(self):
        seen = []

        def fake_run(cmd, **kwargs):
            script = Path(cmd[1]).name
            seen.append(script)
            project = Path(cmd[2])
            if script == "finalize_svg.py":
                self.assertTrue((project / "svg_output" / doctor.SMOKE_SVG.name).is_file())
            else:
                make_pptx(Path(cmd[cmd.index("-o") + 1]), [doctor.SMOKE_TEXT])
            self.assertEqual(kwargs["env"]["PYTHONIOENCODING"], "utf-8")
            return completed(cmd)

        result = doctor.run_smoke(run=fake_run, python="python")
        self.assertTrue(result.ok, result.detail)
        self.assertEqual(seen, ["finalize_svg.py", "svg_to_pptx.py"])

    def test_reports_failing_step_with_log_tail(self):
        result = doctor.run_smoke(run=lambda cmd, **kwargs: completed(cmd, 2, stderr="line1\nline2\nboom"), python="python")
        self.assertFalse(result.ok)
        self.assertIn("finalize_svg.py", result.detail)
        self.assertIn("boom", result.detail)

    def test_reports_timeout(self):
        def fake_run(cmd, **kwargs):
            raise subprocess.TimeoutExpired(cmd, 300)

        result = doctor.run_smoke(run=fake_run, python="python")
        self.assertFalse(result.ok)
        self.assertIn("Quá thời gian", result.detail)

    def test_reports_oserror_from_run_without_raising(self):
        def fake_run(cmd, **kwargs):
            raise OSError("no such file")

        result = doctor.run_smoke(run=fake_run, python="python")
        self.assertFalse(result.ok)
        self.assertIn("finalize_svg.py", result.detail)

    def test_missing_smoke_fixture_reported_gracefully(self):
        def fail_if_called(cmd, **kwargs):
            self.fail("run() should not be called when the smoke fixture is missing")

        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch.object(doctor, "SMOKE_SVG", Path(tmp) / "absent.svg"):
                result = doctor.run_smoke(run=fail_if_called, python="python")
        self.assertFalse(result.ok)
        self.assertIn("Không chép được file mẫu smoke test", result.detail)
        self.assertEqual(result.fix, "Tải lại bản đầy đủ của bộ công cụ")


class RenderAndExitCodeTest(unittest.TestCase):
    def test_exit_code_fails_only_on_required(self):
        required_ok = doctor.CheckResult("Python", doctor.REQUIRED, True, "3.12")
        optional_fail = doctor.CheckResult("Pandoc", doctor.OPTIONAL, False, "Chưa cài", "x")
        required_fail = doctor.CheckResult("Thư viện Python", doctor.REQUIRED, False, "Thiếu: flask", "Chạy CAI-DAT.bat")
        self.assertEqual(doctor.exit_code([required_ok, optional_fail]), 0)
        self.assertEqual(doctor.exit_code([required_ok, required_fail]), 1)

    def test_render_shows_icons_fix_and_summary(self):
        text = doctor.render([
            doctor.CheckResult("Python", doctor.REQUIRED, True, "3.12"),
            doctor.CheckResult("Pandoc", doctor.OPTIONAL, False, "Chưa cài", "Chỉ cần khi chuyển tài liệu"),
            doctor.CheckResult("Thư viện Python", doctor.REQUIRED, False, "Thiếu: flask", "Chạy CAI-DAT.bat"),
        ])
        self.assertIn("✅ Python", text)
        self.assertIn("⚠️ Pandoc", text)
        self.assertIn("❌ Thư viện Python", text)
        self.assertIn("→ Chạy CAI-DAT.bat", text)
        self.assertIn("còn lỗi bắt buộc", text)


class MainTest(unittest.TestCase):
    def _ok(self, name, level=doctor.REQUIRED):
        return doctor.CheckResult(name, level, True, "ok")

    def _patches(self, integrity_ok=True):
        integrity = self._ok("Toàn vẹn") if integrity_ok else doctor.CheckResult("Toàn vẹn", doctor.REQUIRED, False, "lỗi", "x")
        return [
            mock.patch.object(doctor, "check_python", return_value=self._ok("Python")),
            mock.patch.object(doctor, "check_packages", return_value=self._ok("Thư viện")),
            mock.patch.object(doctor, "check_integrity", return_value=integrity),
            mock.patch.object(doctor, "check_tool", side_effect=lambda command, label, level, purpose: self._ok(label, level)),
            mock.patch.object(doctor, "check_api_keys", return_value=self._ok("API", doctor.OPTIONAL)),
        ]

    def _run_main(self, argv, integrity_ok=True):
        patches = self._patches(integrity_ok)
        for patch in patches:
            patch.start()
        try:
            with mock.patch.object(doctor, "run_smoke", return_value=self._ok("Xuất thử PPTX")) as smoke, \
                    mock.patch("sys.stdout", new_callable=io.StringIO) as out:
                code = doctor.main(argv)
            return code, smoke, out.getvalue()
        finally:
            for patch in patches:
                patch.stop()

    def test_no_smoke_skips_export_and_returns_zero(self):
        code, smoke, output = self._run_main(["--no-smoke"])
        self.assertEqual(code, 0)
        smoke.assert_not_called()
        self.assertIn("sẵn sàng", output)

    def test_runs_smoke_by_default(self):
        code, smoke, _ = self._run_main([])
        self.assertEqual(code, 0)
        smoke.assert_called_once()

    def test_skips_smoke_and_fails_when_integrity_fails(self):
        code, smoke, output = self._run_main([], integrity_ok=False)
        self.assertEqual(code, 1)
        smoke.assert_not_called()
        self.assertIn("Bỏ qua", output)

    def test_collect_short_circuits_on_failed_python_check(self):
        def fail_if_called(*args, **kwargs):
            self.fail("check should not be called when Python check fails")

        failed_python = doctor.CheckResult("Python", doctor.REQUIRED, False, "3.9", "Cài Python 3.10+")
        with mock.patch.object(doctor, "check_python", return_value=failed_python), \
                mock.patch.object(doctor, "check_packages", side_effect=fail_if_called), \
                mock.patch.object(doctor, "check_integrity", side_effect=fail_if_called), \
                mock.patch.object(doctor, "run_smoke", side_effect=fail_if_called), \
                mock.patch.object(doctor, "check_tool", side_effect=fail_if_called), \
                mock.patch.object(doctor, "check_api_keys", side_effect=fail_if_called):
            results = doctor.collect(no_smoke=True)
        self.assertEqual(results, [failed_python])
        self.assertEqual(doctor.exit_code(results), 1)


if __name__ == "__main__":
    unittest.main()
