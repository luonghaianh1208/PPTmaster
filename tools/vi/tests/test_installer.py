"""Test cho tools/vi/pptmaster.ps1 ở chế độ -PlanOnly (không tải, không cài, không mở cửa sổ)."""

import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
LAUNCHER = REPO_ROOT / "tools" / "vi" / "pptmaster.ps1"
POWERSHELL = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe"


@unittest.skipUnless(sys.platform == "win32" and POWERSHELL.is_file(), "Chỉ chạy trên Windows có PowerShell")
class InstallerPlanTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory(prefix="pptmaster-plan-")
        self.tmp = Path(self._tmp.name)
        self.repo = self.make_repo(self.tmp / "repo")
        self.bin = self.tmp / "bin"
        self.bin.mkdir()
        self.localappdata = self.tmp / "localappdata"
        self.localappdata.mkdir()

    def tearDown(self):
        self._tmp.cleanup()

    def make_repo(self, root):
        (root / "tools" / "vi").mkdir(parents=True)
        shutil.copyfile(LAUNCHER, root / "tools" / "vi" / "pptmaster.ps1")
        (root / ".env.example").write_text("# mẫu\n", encoding="utf-8")
        return root

    def stub(self, name, body, folder=None):
        target = folder or self.bin
        target.mkdir(parents=True, exist_ok=True)
        (target / f"{name}.cmd").write_text(f"@echo off\r\n{body}\r\n", encoding="ascii")

    def run_plan(self, *args, repo=None, extra_path=(), env_overrides=None):
        repo = repo or self.repo
        env = {key: value for key, value in os.environ.items() if not key.upper().startswith("ONEDRIVE")}
        env["PATH"] = os.pathsep.join([str(path) for path in extra_path] + [str(self.bin)])
        env["LOCALAPPDATA"] = str(self.localappdata)
        env.update(env_overrides or {})
        proc = subprocess.run(
            [str(POWERSHELL), "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
             str(repo / "tools" / "vi" / "pptmaster.ps1"), *args, "-PlanOnly"],
            capture_output=True, env=env, timeout=120,
        )
        stdout = proc.stdout.decode("utf-8", errors="replace").strip()
        stderr = proc.stderr.decode("utf-8", errors="replace")
        self.assertEqual(proc.returncode, 0, stderr)
        self.assertEqual(len(stdout.splitlines()), 1, f"stdout phải là đúng một dòng JSON:\n{stdout}")
        return json.loads(stdout)

    def steps(self, plan):
        return [(step["step"], step["action"], step["method"]) for step in plan["steps"]]

    def test_no_python_with_winget_plans_winget_install_then_venv(self):
        self.stub("winget", "echo winget")
        plan = self.run_plan("-Action", "setup", "-Auto")
        self.assertIsNone(plan["python_found"])
        self.assertEqual(self.steps(plan), [
            ("python", "install", "winget"),
            ("venv", "create", "python -m venv"),
            ("packages", "ensure", "pip"),
            ("env", "create", "copy .env.example"),
            ("doctor", "run", "doctor.py --json"),
        ])
        self.assertEqual(plan["warnings"], [])

    def test_no_python_without_winget_plans_python_org_installer(self):
        plan = self.run_plan("-Action", "setup", "-Auto")
        self.assertEqual(self.steps(plan)[0], ("python", "install", "python-org"))

    def test_store_alias_only_counts_as_no_python(self):
        self.stub("python", "exit /b 9009", folder=self.tmp / "WindowsApps")
        plan = self.run_plan("-Action", "setup", "-Auto", extra_path=[self.tmp / "WindowsApps"])
        self.assertEqual(self.steps(plan)[0], ("python", "install", "python-org"))

    def test_old_python_counts_as_no_python(self):
        self.stub("python", "echo 3.8")
        plan = self.run_plan("-Action", "setup", "-Auto")
        self.assertEqual(self.steps(plan)[0][:2], ("python", "install"))

    def test_existing_python_312_skips_install_and_creates_venv(self):
        self.stub("python", "echo 3.12")
        plan = self.run_plan("-Action", "setup", "-Auto")
        self.assertTrue(plan["python_found"].lower().endswith("python.cmd"))
        self.assertEqual(self.steps(plan)[0], ("venv", "create", "python -m venv"))
        self.assertNotIn("python", [step["step"] for step in plan["steps"]])

    def test_user_scope_python_found_without_path(self):
        user_python = self.localappdata / "Programs" / "Python" / "Python312" / "python.exe"
        user_python.parent.mkdir(parents=True)
        user_python.write_bytes(b"")
        plan = self.run_plan("-Action", "setup", "-Auto")
        self.assertEqual(Path(plan["python_found"]), user_python)
        self.assertNotIn("python", [step["step"] for step in plan["steps"]])

    def test_already_set_up_only_ensures_packages_and_runs_doctor(self):
        venv_python = self.repo / "venv" / "Scripts" / "python.exe"
        venv_python.parent.mkdir(parents=True)
        venv_python.write_bytes(b"")
        (self.repo / ".env").write_text("", encoding="utf-8")
        plan = self.run_plan("-Action", "setup", "-Auto")
        self.assertTrue(plan["python_found"].endswith(r"venv\Scripts\python.exe"))
        self.assertEqual(self.steps(plan), [("packages", "ensure", "pip"), ("doctor", "run", "doctor.py --json")])

    def test_onedrive_folder_warns(self):
        plan = self.run_plan("-Action", "setup", "-Auto", env_overrides={"OneDrive": str(self.tmp)})
        self.assertEqual(len(plan["warnings"]), 1)
        self.assertIn("OneDrive", plan["warnings"][0])

    def test_long_path_warns(self):
        repo = self.make_repo(self.tmp / ("thu-muc-rat-dai-" * 5) / "repo")
        plan = self.run_plan("-Action", "setup", "-Auto", repo=repo)
        self.assertTrue(any("ký tự" in warning for warning in plan["warnings"]), plan["warnings"])

    def test_tool_missing_with_winget_plans_install(self):
        self.stub("winget", "echo winget")
        plan = self.run_plan("-Action", "tool", "-Name", "ffmpeg")
        self.assertEqual(plan["tool"], "ffmpeg")
        self.assertFalse(plan["found"])
        self.assertEqual(self.steps(plan), [("ffmpeg", "install", "winget")])

    def test_tool_missing_without_winget_plans_manual(self):
        plan = self.run_plan("-Action", "tool", "-Name", "pandoc")
        self.assertEqual(self.steps(plan), [("pandoc", "install", "manual")])

    def test_tool_found_in_user_install_folder(self):
        pandoc = self.localappdata / "Pandoc" / "pandoc.exe"
        pandoc.parent.mkdir(parents=True)
        pandoc.write_bytes(b"")
        plan = self.run_plan("-Action", "tool", "-Name", "pandoc")
        self.assertTrue(plan["found"])
        self.assertEqual(Path(plan["dir"]), pandoc.parent)
        self.assertEqual(plan["steps"], [])


if __name__ == "__main__":
    unittest.main()
