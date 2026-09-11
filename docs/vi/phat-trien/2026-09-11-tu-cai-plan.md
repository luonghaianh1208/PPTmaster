# AI tự cài đặt bộ công cụ (`v6.3.2-vi.3`): Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Cho AI trong Antigravity tự tải repo, kiểm tra máy, cài Python/thư viện còn thiếu và báo "sẵn sàng tạo slide", không cần thầy cô gõ lệnh hay bấm `.bat`; phát hành `v6.3.2-vi.3`.

**Architecture:** Việc cài nằm trong script cố định: `tools/vi/pptmaster.ps1` thêm `-Auto` (cài Python cho tài khoản, tạo `venv\`, cài thư viện, tạo `.env`, chạy `doctor.py --json`), `-PlanOnly` (chỉ in kế hoạch JSON để test) và `-Action tool` (cài FFmpeg/Pandoc khi cần). AI chỉ điều phối theo `docs/vi/cai-dat-bang-ai.md`; `AGENTS.vi.md` mục 4 và 9 trỏ AI tới đó. Không file nào dưới `skills/` bị sửa.

**Tech Stack:** Windows PowerShell 5.1, Python ≥ 3.10 stdlib (`unittest`, `json`, `subprocess`), winget, Markdown tiếng Việt, Claude Code CLI (kiểm tra headless), git, `gh`.

**Spec:** [2026-09-11-tu-cai-bang-ai-design.md](2026-09-11-tu-cai-bang-ai-design.md)

## Global Constraints

- Chỉ bắt đầu khi `v6.3.2-vi.2` (M2) đã được **squash** vào `main` và gắn tag. Nhánh làm việc: `feat/vi-tu-cai`, được đặt lại lên `main` ở Task 1.
- Không sửa bất kỳ file nào dưới `skills/`: `git diff v6.3.2 -- skills/` phải rỗng; `python skills/ppt-master/scripts/attribution_guard.py` exit 0.
- File được tạo/sửa: `tools/vi/pptmaster.ps1`, `tools/vi/doctor.py`, `tools/vi/tests/test_doctor.py`, `tools/vi/tests/test_installer.py` (mới), `tools/vi/tests/test_vi_layer.py`, `docs/vi/cai-dat-bang-ai.md` (mới), `AGENTS.vi.md`, `README.md`, `docs/vi/bat-dau-nhanh.md`, `docs/vi/cai-dat-windows.md`, `docs/vi/xu-ly-loi.md`, `docs/vi/phat-trien/bao-tri.md`, `docs/vi/phat-trien/2026-09-11-tu-cai-bang-ai-design.md`, `CHANGELOG-VI.md`. Không cần mở rộng `ALLOWED_CHANGES`.
- Không sửa `CAI-DAT.bat`, `KIEM-TRA.bat`, `CAP-NHAT.bat` (xem `docs/vi/phat-trien/bao-tri.md`). `setup` không có `-Auto` (từ `CAI-DAT.bat`) giữ nguyên hành vi M1.
- `tools/vi/*.ps1` phải bắt đầu bằng UTF-8 BOM (test `test_powershell_scripts_have_utf8_bom`); PowerShell 5.1 đọc file không BOM sai dấu tiếng Việt.
- Ở `setup -Auto`, `-PlanOnly` và `-Action tool`: stdout chỉ có đúng một đối tượng JSON trên một dòng (ghi bằng `[Console]::Out.WriteLine`); mọi dòng tiến trình, kể cả đầu ra của winget/pip, đi ra stderr bằng `[Console]::Error.WriteLine`. Không dùng `Write-Host`, `Out-Host`, `Write-Output` trong các nhánh này (trong PowerShell 5.1 chạy bằng `-File`, `Write-Host` ra stdout).
- Python cố định: `3.12.10`; `amd64` SHA256 `67B5635E80EA51072B87941312D00EC8927C4DB9BA18938F7AD2D27B328B95FB`, `arm64` SHA256 `377AC8FD478987940088E879441E702A71B53164D2A1E6F1D51FF77A7E470258` (đã đối chiếu MD5 với trang python.org ngày 2026-09-11).
- `doctor.py` và test chỉ dùng thư viện chuẩn Python.
- Test không tải, không cài, không mở cửa sổ; `test_installer.py` bỏ qua trên máy không phải Windows.
- File Markdown viết hoàn toàn bằng tiếng Việt; chuỗi không phải tiếng Việt chỉ ở lệnh, đường dẫn, key JSON, tên nút/menu và câu lệnh mẫu.
- Test: `& $PY -m unittest discover -s tools/vi/tests -v` từ thư mục gốc repo. `$PY` = `C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe` (định nghĩa lại trong mỗi lần gọi PowerShell). `$SCRATCH` = `C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad`.
- Commit theo Conventional Commits; message kết thúc bằng một dòng `Co-Authored-By: Claude … <noreply@anthropic.com>` theo hướng dẫn ghi công của phiên đang thực hiện.
- Không push trước Task 7; không bao giờ force push; mọi lệnh `gh` đều có `--repo luonghaianh1208/PPTmaster`. Xoá nhánh (cục bộ hay trên `origin`) phải hỏi chủ repo.

## File Structure

| File | Trách nhiệm |
|---|---|
| `tools/vi/doctor.py` | Thêm `render_json()` và cờ `--json` |
| `tools/vi/pptmaster.ps1` | Thêm `-Auto`, `-PlanOnly`, `-Action tool -Name`; `check`/`update` ưu tiên `venv` |
| `tools/vi/tests/test_doctor.py` | Test `--json` |
| `tools/vi/tests/test_installer.py` | Test `-PlanOnly` bằng PATH và `LOCALAPPDATA` giả |
| `docs/vi/cai-dat-bang-ai.md` | Hướng dẫn cài dành cho AI |
| `AGENTS.vi.md` | Mục 4 (Python của `venv`), mục 9 (tự kiểm tra, tự cài) |
| `README.md` | Khối "Dành cho AI agent", dòng dẫn tới cách dán link |
| `docs/vi/bat-dau-nhanh.md`, `docs/vi/cai-dat-windows.md`, `docs/vi/xu-ly-loi.md` | Tài liệu cho thầy cô |
| `docs/vi/phat-trien/bao-tri.md` | Cách đổi phiên bản Python cố định |
| `tools/vi/tests/test_vi_layer.py` | Test nội dung hướng dẫn và quy tắc |
| `CHANGELOG-VI.md` | Mục `6.3.2-vi.3` (Task 8) |

Thứ tự task: 1 đặt lại nhánh → 2 `doctor --json` → 3 bộ cài tự động → 4 hướng dẫn cho AI + `AGENTS.vi.md` + README → 5 tài liệu cho thầy cô → 6 kiểm tra headless → 7 nghiệm thu máy thật → 8 phát hành.

---

### Task 1: Đặt lại nhánh lên `main` sau khi M2 phát hành

**Files:** không sửa file.

**Interfaces:**
- Consumes: `main` đã có commit squash của M2 và tag `v6.3.2-vi.2`; nhánh `feat/vi-tu-cai` hiện tách từ commit `73e9cec7` (đầu nhánh `feat/vi-tro-ly` trước khi phát hành).
- Produces: `feat/vi-tu-cai` nằm trên `main`, chỉ gồm các commit spec/plan của tính năng này.

- [ ] **Step 1: Kiểm tra điều kiện**

```powershell
git fetch origin
git tag -l v6.3.2-vi.2
git log -1 --format="%h %s" main
git status --porcelain; "(status end)"
git diff --stat 73e9cec7 main
```
Expected: in ra `v6.3.2-vi.2`; status rỗng; `git diff --stat` chỉ liệt kê file của đợt phát hành M2 (thường là `CHANGELOG-VI.md`, `README.md`). Không có tag → dừng, báo BLOCKED (M2 chưa phát hành). Có thêm file khác → vẫn tiếp tục, ghi danh sách vào báo cáo.

- [ ] **Step 2: Đặt lại nhánh**

```powershell
git switch feat/vi-tu-cai
git rebase --onto main 73e9cec7 feat/vi-tu-cai
git log --oneline main..feat/vi-tu-cai
```
Expected: rebase thành công; log chỉ có các commit `docs(vi): …` của spec và plan tính năng này. Xung đột → `git rebase --abort`, báo BLOCKED kèm tên file.

- [ ] **Step 3: Chạy test nền**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests
& $PY skills/ppt-master/scripts/attribution_guard.py; "guard=$LASTEXITCODE"
git diff --stat v6.3.2 HEAD -- skills/; "(skills diff end)"
```
Expected: test OK; `guard=0`; diff `skills/` rỗng. Không có commit ở task này.

---

### Task 2: `doctor.py --json`

**Files:**
- Modify: `tools/vi/doctor.py` (docstring đầu file, import, thêm `render_json`, `main`)
- Test: `tools/vi/tests/test_doctor.py`

**Interfaces:**
- Consumes: `CheckResult(name, level, ok, detail, fix)`, `collect(no_smoke) -> list[CheckResult]`, `exit_code(results) -> int`, `render(results) -> str` (đã có).
- Produces: `render_json(results: Sequence[CheckResult], python: str = sys.executable) -> str` trả chuỗi JSON `{"ready": bool, "python": str, "checks": [{"name", "level", "ok", "detail", "fix"}]}` với `ensure_ascii=False`; CLI `python tools/vi/doctor.py --json [--no-smoke]` in đúng chuỗi đó, mã thoát không đổi. Task 3 dựa vào tên mục `"Thư viện Python"` trong `checks`.

- [ ] **Step 1: Viết test lỗi**

Trong `tools/vi/tests/test_doctor.py`, thêm `import json` vào khối import (sau `import io`). Thêm class mới ngay trước `class MainTest(unittest.TestCase):`:

```python
class RenderJsonTest(unittest.TestCase):
    def test_render_json_reports_ready_python_and_all_fields(self):
        results = [
            doctor.CheckResult("Python", doctor.REQUIRED, True, "Phiên bản 3.12"),
            doctor.CheckResult("Pandoc", doctor.OPTIONAL, False, "Chưa cài", "Chỉ cần khi chuyển tài liệu"),
        ]
        data = json.loads(doctor.render_json(results, python="C:/py/python.exe"))
        self.assertTrue(data["ready"])
        self.assertEqual(data["python"], "C:/py/python.exe")
        self.assertEqual(data["checks"][1], {
            "name": "Pandoc", "level": doctor.OPTIONAL, "ok": False,
            "detail": "Chưa cài", "fix": "Chỉ cần khi chuyển tài liệu",
        })

    def test_render_json_not_ready_when_required_fails_and_keeps_vietnamese(self):
        results = [doctor.CheckResult("Thư viện Python", doctor.REQUIRED, False, "Thiếu: flask", "Chạy CAI-DAT.bat")]
        text = doctor.render_json(results)
        self.assertIn("Thư viện Python", text)
        self.assertFalse(json.loads(text)["ready"])

    def test_package_check_name_is_installer_contract(self):
        with tempfile.TemporaryDirectory() as tmp:
            requirements = Path(tmp) / "requirements.txt"
            requirements.write_text("", encoding="utf-8")
            self.assertEqual(doctor.check_packages(requirements).name, "Thư viện Python")
```

Trong `class MainTest`, thêm hai test sau `test_skips_smoke_and_fails_when_integrity_fails`:

```python
    def test_json_flag_prints_only_json_with_smoke(self):
        code, smoke, output = self._run_main(["--json"])
        data = json.loads(output)
        self.assertEqual(code, 0)
        self.assertTrue(data["ready"])
        self.assertEqual(data["python"], sys.executable)
        self.assertIn("Xuất thử PPTX", [check["name"] for check in data["checks"]])
        self.assertNotIn("Kết quả:", output)
        smoke.assert_called_once()

    def test_json_flag_with_no_smoke_keeps_exit_code(self):
        code, smoke, output = self._run_main(["--json", "--no-smoke"], integrity_ok=False)
        data = json.loads(output)
        self.assertEqual(code, 1)
        self.assertFalse(data["ready"])
        self.assertNotIn("Xuất thử PPTX", [check["name"] for check in data["checks"]])
        smoke.assert_not_called()
```

- [ ] **Step 2: Chạy test, xác nhận lỗi**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -p test_doctor.py -v
```
Expected: FAIL/ERROR ở `RenderJsonTest` (`AttributeError: module 'doctor' has no attribute 'render_json'`) và hai test `--json` (`SystemExit: 2`, argparse không biết `--json`). `test_package_check_name_is_installer_contract` có thể đã PASS.

- [ ] **Step 3: Viết code**

Trong `tools/vi/doctor.py`:

Docstring đầu file, thay khối "Cách dùng" bằng:

```python
"""Kiểm tra môi trường PPT Master (bản Việt).

Cách dùng:
    python tools/vi/doctor.py             # kiểm tra đầy đủ, có xuất thử 1 file PPTX
    python tools/vi/doctor.py --no-smoke  # bỏ bước xuất thử
    python tools/vi/doctor.py --json      # in kết quả dạng JSON cho AI đọc (dùng được với --no-smoke)

Mã thoát: 0 khi mọi mục bắt buộc đạt, 1 khi còn mục bắt buộc lỗi.
Chỉ dùng thư viện chuẩn Python.
"""
```

Import: thêm `import json` sau `import importlib.metadata`; đổi `from dataclasses import dataclass` thành `from dataclasses import asdict, dataclass`.

Thêm ngay sau hàm `render`:

```python
def render_json(results: Sequence[CheckResult], python: str = sys.executable) -> str:
    payload = {
        "ready": exit_code(results) == 0,
        "python": python,
        "checks": [asdict(result) for result in results],
    }
    return json.dumps(payload, ensure_ascii=False)
```

Thay hàm `main` bằng:

```python
def main(argv: Optional[Sequence[str]] = None) -> int:
    _configure_utf8()
    parser = argparse.ArgumentParser(description="Kiểm tra môi trường PPT Master (bản Việt)")
    parser.add_argument("--no-smoke", action="store_true", help="Bỏ bước xuất thử PPTX")
    parser.add_argument("--json", action="store_true", help="In kết quả dạng JSON cho AI đọc")
    args = parser.parse_args(argv)
    results = collect(args.no_smoke)
    print(render_json(results) if args.json else render(results))
    return exit_code(results)
```

- [ ] **Step 4: Chạy test, xác nhận đạt**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -v
& $PY tools/vi/doctor.py --no-smoke --json; "exit=$LASTEXITCODE"
```
Expected: toàn bộ test OK; lệnh thứ hai in một dòng JSON có `"ready": true` và `exit=0`.

- [ ] **Step 5: Commit**

```powershell
git add tools/vi/doctor.py tools/vi/tests/test_doctor.py
git commit -m "feat(vi): add JSON output to doctor" -m "Co-Authored-By: Claude <noreply@anthropic.com>"
```
(Thay dòng `Co-Authored-By` bằng dòng ghi công đúng của phiên đang thực hiện.)

---

### Task 3: Bộ cài tự động trong `pptmaster.ps1`

**Files:**
- Modify: `tools/vi/pptmaster.ps1` (thay toàn bộ nội dung)
- Create: `tools/vi/tests/test_installer.py`

**Interfaces:**
- Consumes: `doctor.py --json [--no-smoke]` (Task 2), mục `"Thư viện Python"` trong `checks`.
- Produces (Task 4 và 6 dựa vào):
  - `powershell -NoProfile -ExecutionPolicy Bypass -File tools\vi\pptmaster.ps1 -Action setup -Auto` → stdout một dòng JSON `{"ready": bool, "python": str|null, "installed": [str], "warnings": [str], "checks": [...], "error": null|{"step","message","fix"}}`; mã thoát 0 khi `ready`, ngược lại 1. Giá trị `installed`: `python`, `venv`, `packages`, `env`. Giá trị `error.step`: `python`, `venv`, `packages`, `doctor`.
  - `… -Action setup -Auto -PlanOnly` → `{"python_found": str|null, "steps": [{"step","action","method"}], "warnings": [str]}`, mã thoát 0. Bước: `python/install/winget|python-org`, `venv/create/python -m venv`, `packages/ensure/pip`, `env/create/copy .env.example`, `doctor/run/doctor.py --json`.
  - `… -Action tool -Name ffmpeg|pandoc` → `{"tool", "found": bool, "installed": bool, "dir": str|null, "error": null|{...}}`, mã thoát 0 khi `found`. Với `-PlanOnly` → `{"tool", "found", "dir", "steps": [{"step": "<tên>", "action": "install", "method": "winget|manual"}]}`.
  - `check`/`update`: có `venv\Scripts\python.exe` thì dùng nó.

- [ ] **Step 1: Viết test lỗi**

Tạo `tools/vi/tests/test_installer.py`:

```python
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
```

- [ ] **Step 2: Chạy test, xác nhận lỗi**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -p test_installer.py -v
```
Expected: FAIL ở mọi test (script hiện tại không có `-PlanOnly`/`-Auto`/`tool` nên PowerShell báo lỗi tham số và mã thoát khác 0).

- [ ] **Step 3: Viết code**

Thay toàn bộ nội dung `tools/vi/pptmaster.ps1` bằng đoạn dưới đây, rồi đảm bảo file có UTF-8 BOM:

```powershell
<#
.SYNOPSIS
  Trình khởi chạy bản Việt của PPT Master: cài đặt, kiểm tra, cập nhật, cài công cụ tuỳ chọn.
.PARAMETER Action
  setup  - kiểm tra Python, cài thư viện, tạo .env, công cụ tuỳ chọn, chạy doctor
  check  - chạy doctor
  update - cập nhật bằng git (update_repo.py của upstream) rồi kiểm tra nhanh
  tool   - cài công cụ tuỳ chọn cho tài khoản hiện tại (dùng với -Name), in JSON
.PARAMETER Auto
  Dùng với setup: tự cài không hỏi (Python cho tài khoản, venv, thư viện, .env), in một đối tượng JSON ra stdout.
.PARAMETER PlanOnly
  Dùng với setup -Auto hoặc tool: chỉ in kế hoạch dạng JSON; không tải, không cài, không tạo file.
.PARAMETER Name
  Dùng với tool: ffmpeg hoặc pandoc.
.PARAMETER NonInteractive
  Không hỏi Y/N và không tự cài phần mềm (dùng khi kiểm thử).
#>
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('setup', 'check', 'update', 'tool')]
    [string]$Action,
    [switch]$Auto,
    [switch]$PlanOnly,
    [ValidateSet('ffmpeg', 'pandoc')]
    [string]$Name,
    [switch]$NonInteractive
)

$ErrorActionPreference = 'Continue'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch { }
$env:PYTHONIOENCODING = 'utf-8'

$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$Doctor = Join-Path $RepoRoot 'tools\vi\doctor.py'
$FixDoc = Join-Path $RepoRoot 'docs\vi\xu-ly-loi.md'
$InstallDoc = Join-Path $RepoRoot 'docs\vi\cai-dat-windows.md'
$VenvDir = Join-Path $RepoRoot 'venv'
$VenvPython = Join-Path $VenvDir 'Scripts\python.exe'
$MaxRepoPathLength = 80

# Bộ cài python.org dự phòng khi không có winget. Đổi phiên bản: xem docs/vi/phat-trien/bao-tri.md.
$PythonVersion = '3.12.10'
$PythonInstallers = @{
    amd64 = @{ Url = 'https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe'; Sha256 = '67B5635E80EA51072B87941312D00EC8927C4DB9BA18938F7AD2D27B328B95FB' }
    arm64 = @{ Url = 'https://www.python.org/ftp/python/3.12.10/python-3.12.10-arm64.exe'; Sha256 = '377AC8FD478987940088E879441E702A71B53164D2A1E6F1D51FF77A7E470258' }
}
$OptionalTools = @{
    ffmpeg = @{ Id = 'Gyan.FFmpeg'; Exe = 'ffmpeg.exe'; Manual = 'https://ffmpeg.org/download.html' }
    pandoc = @{ Id = 'JohnMacFarlane.Pandoc'; Exe = 'pandoc.exe'; Manual = 'https://pandoc.org/installing.html' }
}
$script:SetupError = $null

function Write-Step([string]$Text) { Write-Host ''; Write-Host "==> $Text" -ForegroundColor Cyan }
function Write-Fail([string]$Text) { Write-Host "[LỖI] $Text" -ForegroundColor Red }
function Write-Ok([string]$Text) { Write-Host "[OK] $Text" -ForegroundColor Green }

# Chế độ cho AI (setup -Auto, -PlanOnly, tool): stdout chỉ có một dòng JSON, mọi dòng khác ra stderr.
function Write-Log([string]$Text) { [Console]::Error.WriteLine($Text) }
function Write-Json($Object) { [Console]::Out.WriteLine(($Object | ConvertTo-Json -Depth 6 -Compress)) }
function Invoke-Logged([scriptblock]$Command) { & $Command 2>&1 | ForEach-Object { Write-Log "$_" } }
function Set-SetupError([string]$Step, [string]$Message, [string]$Fix) {
    $script:SetupError = [pscustomobject]@{ step = $Step; message = $Message; fix = $Fix }
}

function Test-Winget { return [bool](Get-Command winget -ErrorAction SilentlyContinue) }

function Confirm-Choice([string]$Question) {
    if ($NonInteractive) { return $false }
    $answer = Read-Host "$Question (Y/N)"
    return ($answer -match '^(y|yes|c|co|có)$')
}

function Get-PythonInfo {
    $cmd = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $cmd) { return $null }
    $version = $null
    try {
        $raw = & $cmd.Source -c "import sys; print('%d.%d' % sys.version_info[:2])" 2>$null
        if ($LASTEXITCODE -eq 0 -and $raw) { $version = [version]("$raw".Trim()) }
    } catch { $version = $null }
    return [pscustomobject]@{
        Path         = $cmd.Source
        Version      = $version
        IsStoreAlias = ($cmd.Source -like '*\WindowsApps\*')
    }
}

function Get-LauncherPython {
    $launcher = Get-Command py -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $launcher) { return $null }
    try {
        $raw = @(& $launcher.Source -3 -c "import sys; print(sys.executable); print('%d.%d' % sys.version_info[:2])" 2>$null)
        if ($LASTEXITCODE -ne 0 -or $raw.Count -lt 2) { return $null }
        $version = [version]("$($raw[$raw.Count - 1])".Trim())
    } catch { return $null }
    if ($version -lt [version]'3.10') { return $null }
    return [pscustomobject]@{
        Path    = "$($raw[$raw.Count - 2])".Trim()
        Version = $version
    }
}

function Get-UserPythonPath {
    if (-not $env:LOCALAPPDATA) { return $null }
    $candidate = Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe'
    if (Test-Path $candidate) { return $candidate }
    return $null
}

function Find-BasePython {
    $py = Get-PythonInfo
    if ($py -and $py.Version -and $py.Version -ge [version]'3.10') { return $py.Path }
    $launcher = Get-LauncherPython
    if ($launcher) { return $launcher.Path }
    return (Get-UserPythonPath)
}

function Resolve-Python([bool]$OfferInstall) {
    $py = Get-PythonInfo
    if ($py -and $py.Version -and $py.Version -ge [version]'3.10') {
        Write-Ok "Python $($py.Version) tại $($py.Path)"
        return $py
    }
    $installed = Get-LauncherPython
    if ($installed) {
        Write-Fail "Đã cài Python $($installed.Version) tại $($installed.Path) nhưng bản này chưa có trong PATH."
        Write-Host 'Cách sửa: Settings → Apps → Installed apps → Python 3.x → Modify → tick "Add Python to environment variables" (hoặc chạy lại bộ cài tải từ python.org và chọn Modify).'
        Write-Host 'Sau đó đóng cửa sổ này, mở lại rồi bấm lại CAI-DAT.bat.'
        Write-Host "Hướng dẫn chi tiết: $FixDoc (mục Đã cài Python nhưng bộ cài báo không tìm thấy)"
        return $null
    }
    if (-not $py) {
        Write-Fail 'Chưa tìm thấy Python trong PATH.'
    } elseif ((-not $py.Version) -and $py.IsStoreAlias) {
        Write-Fail 'Lệnh python đang trỏ tới lối tắt của Microsoft Store, chưa phải Python thật.'
        Write-Host 'Tắt lối tắt: Settings → Apps → Advanced app settings → App execution aliases → tắt "python.exe" và "python3.exe".'
    } elseif (-not $py.Version) {
        Write-Fail "Không chạy được Python tại $($py.Path)."
    } else {
        Write-Fail "Python $($py.Version) quá cũ, cần 3.10 trở lên."
    }
    if ($OfferInstall -and (Test-Winget) -and (Confirm-Choice 'Cài Python 3.12 bằng winget?')) {
        winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements | Out-Host
        Write-Host ''
        Write-Host 'Đã chạy trình cài Python. Hãy ĐÓNG cửa sổ này rồi bấm lại CAI-DAT.bat để PATH mới có hiệu lực.' -ForegroundColor Yellow
    } else {
        Write-Host 'Cài Python 3.12 tại https://www.python.org/downloads/ và tick "Add python.exe to PATH".'
        Write-Host "Hướng dẫn chi tiết: $InstallDoc"
    }
    return $null
}

function Resolve-RunPython {
    if (Test-Path $VenvPython) {
        Write-Ok "Python của venv tại $VenvPython"
        return [pscustomobject]@{ Path = $VenvPython }
    }
    return (Resolve-Python $false)
}

function Invoke-Doctor($Py, [string[]]$DoctorArgs) {
    & $Py.Path $Doctor @DoctorArgs | Out-Host
    return $LASTEXITCODE
}

function Invoke-Setup {
    Write-Step 'Bước 1/4: Kiểm tra Python'
    $py = Resolve-Python $true
    if (-not $py) { return 1 }

    Write-Step 'Bước 2/4: Cài thư viện Python (lần đầu có thể mất vài phút)'
    & $py.Path -m pip install --upgrade pip | Out-Host
    & $py.Path -m pip install -r (Join-Path $RepoRoot 'requirements.txt') | Out-Host
    if ($LASTEXITCODE -ne 0) {
        Write-Fail 'Cài thư viện thất bại. Xem thông báo phía trên.'
        Write-Host "Cách xử lý: $FixDoc (mục Cài thư viện thất bại)"
        return 1
    }
    Write-Ok 'Đã cài thư viện.'

    Write-Step 'Bước 3/4: Tạo file cấu hình .env'
    $envFile = Join-Path $RepoRoot '.env'
    if (Test-Path $envFile) {
        Write-Ok 'Đã có .env, giữ nguyên.'
    } else {
        Copy-Item (Join-Path $RepoRoot '.env.example') $envFile
        Write-Ok 'Đã tạo .env từ .env.example.'
    }

    Write-Step 'Bước 4/4: Công cụ tuỳ chọn'
    $optional = @(
        @{ Command = 'git'; Id = 'Git.Git'; Label = 'Git (để cập nhật bằng CAP-NHAT.bat)' },
        @{ Command = 'pandoc'; Id = 'JohnMacFarlane.Pandoc'; Label = 'Pandoc (chuyển tài liệu định dạng cũ)' },
        @{ Command = 'ffmpeg'; Id = 'Gyan.FFmpeg'; Label = 'FFmpeg (thuyết minh, video)' }
    )
    $installedAny = $false
    foreach ($tool in $optional) {
        if (Get-Command $tool.Command -ErrorAction SilentlyContinue) {
            Write-Ok "$($tool.Label): đã có"
            continue
        }
        if ((Test-Winget) -and (Confirm-Choice "Cài $($tool.Label) bằng winget?")) {
            winget install -e --id $tool.Id --accept-package-agreements --accept-source-agreements | Out-Host
            $installedAny = $true
        } else {
            Write-Host "Bỏ qua: $($tool.Label)"
        }
    }
    if ($installedAny) {
        $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [Environment]::GetEnvironmentVariable('Path', 'User')
    }

    Write-Step 'Kiểm tra lại toàn bộ'
    return (Invoke-Doctor $py @())
}

function Get-FolderWarnings {
    $warnings = @()
    foreach ($variable in 'OneDrive', 'OneDriveCommercial', 'OneDriveConsumer') {
        $root = [Environment]::GetEnvironmentVariable($variable)
        if ($root -and $RepoRoot.StartsWith($root.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
            $warnings += "Thư mục bộ công cụ nằm trong OneDrive ($root). Nên chuyển sang đường dẫn ngắn như D:\PPTmaster để tránh lỗi khoá file khi cài và khi xuất PPTX."
            break
        }
    }
    if ($RepoRoot.Length -gt $MaxRepoPathLength) {
        $warnings += "Đường dẫn thư mục bộ công cụ dài $($RepoRoot.Length) ký tự (nên dưới $MaxRepoPathLength). Nên chuyển sang đường dẫn ngắn như D:\PPTmaster."
    }
    if ($RepoRoot -like '*PPTmaster-main\PPTmaster-main*') {
        $warnings += 'Thư mục bị lồng PPTmaster-main\PPTmaster-main. Nên chuyển nội dung ra một thư mục ngắn như D:\PPTmaster.'
    }
    return $warnings
}

function Get-SetupPlan {
    $steps = @()
    $venvReady = Test-Path $VenvPython
    $basePython = $null
    if (-not $venvReady) {
        $basePython = Find-BasePython
        if (-not $basePython) {
            $method = if (Test-Winget) { 'winget' } else { 'python-org' }
            $steps += [pscustomobject]@{ step = 'python'; action = 'install'; method = $method }
        }
        $steps += [pscustomobject]@{ step = 'venv'; action = 'create'; method = 'python -m venv' }
    }
    $steps += [pscustomobject]@{ step = 'packages'; action = 'ensure'; method = 'pip' }
    if (-not (Test-Path (Join-Path $RepoRoot '.env'))) {
        $steps += [pscustomobject]@{ step = 'env'; action = 'create'; method = 'copy .env.example' }
    }
    $steps += [pscustomobject]@{ step = 'doctor'; action = 'run'; method = 'doctor.py --json' }
    return [pscustomobject]@{ BasePython = $basePython; VenvReady = $venvReady; Steps = $steps }
}

function Install-UserPython {
    if (Test-Winget) {
        Write-Log "Cài Python $PythonVersion cho tài khoản này bằng winget..."
        Invoke-Logged { winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements --disable-interactivity }
        $found = Find-BasePython
        if ($found) { return $found }
        Write-Log 'winget chưa cài được Python, chuyển sang bộ cài của python.org.'
    }
    $arch = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { 'arm64' } else { 'amd64' }
    $installer = $PythonInstallers[$arch]
    $file = Join-Path $env:TEMP "python-$PythonVersion-$arch.exe"
    Write-Log "Tải bộ cài Python $PythonVersion ($arch) từ python.org..."
    try {
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $installer.Url -OutFile $file -UseBasicParsing
    } catch {
        Set-SetupError 'python' "Không tải được bộ cài Python: $($_.Exception.Message)" 'Kiểm tra kết nối mạng. Máy trường có thể cần mở truy cập python.org (xem mục "Máy trường chặn cài đặt" trong docs/vi/xu-ly-loi.md).'
        return $null
    }
    $hash = (Get-FileHash -Path $file -Algorithm SHA256).Hash
    if ($hash -ne $installer.Sha256) {
        Remove-Item -Path $file -Force -ErrorAction SilentlyContinue
        Set-SetupError 'python' 'Bộ cài Python tải về không khớp mã SHA256 nên đã bị huỷ.' 'Thử lại sau; nếu vẫn lỗi, mạng có thể đang chặn hoặc sửa nội dung tải về (xem docs/vi/xu-ly-loi.md).'
        return $null
    }
    Write-Log 'Chạy bộ cài Python ở chế độ ngầm, chỉ cho tài khoản này...'
    try {
        $proc = Start-Process -FilePath $file -ArgumentList '/quiet', 'InstallAllUsers=0', 'PrependPath=1', 'Include_launcher=1', 'InstallLauncherAllUsers=0', 'Include_test=0' -Wait -PassThru -WindowStyle Hidden
    } catch {
        Set-SetupError 'python' "Không chạy được bộ cài Python: $($_.Exception.Message)" 'Máy có thể đang chặn chạy bộ cài; xem mục "Máy trường chặn cài đặt" trong docs/vi/xu-ly-loi.md.'
        return $null
    }
    $found = Find-BasePython
    if (-not $found) {
        Set-SetupError 'python' "Bộ cài Python kết thúc (mã $($proc.ExitCode)) nhưng không tìm thấy Python." 'Cài Python 3.12 thủ công theo docs/vi/cai-dat-windows.md, hoặc nhờ bộ phận IT (xem mục "Máy trường chặn cài đặt" trong docs/vi/xu-ly-loi.md).'
        return $null
    }
    return $found
}

function Get-DoctorReport([string[]]$DoctorArgs) {
    $raw = & $VenvPython $Doctor --json @DoctorArgs
    $text = ($raw | Out-String).Trim()
    if (-not $text) { return $null }
    try { return ($text | ConvertFrom-Json) } catch { return $null }
}

function Test-PackagesOk($Report) {
    if (-not $Report) { return $false }
    foreach ($check in $Report.checks) {
        if ($check.name -eq 'Thư viện Python') { return [bool]$check.ok }
    }
    return $false
}

function Write-SetupResult([bool]$Ready, $Installed, $Warnings, $Checks) {
    $python = if (Test-Path $VenvPython) { $VenvPython } else { $null }
    Write-Json ([pscustomobject]@{
        ready     = $Ready
        python    = $python
        installed = @($Installed)
        warnings  = @($Warnings)
        checks    = @($Checks)
        error     = $script:SetupError
    })
}

function Invoke-AutoSetup {
    $warnings = @(Get-FolderWarnings)
    if ($PlanOnly) {
        $plan = Get-SetupPlan
        $found = if ($plan.VenvReady) { $VenvPython } else { $plan.BasePython }
        Write-Json ([pscustomobject]@{ python_found = $found; steps = @($plan.Steps); warnings = $warnings })
        return 0
    }
    foreach ($warning in $warnings) { Write-Log "[CẢNH BÁO] $warning" }
    $installed = @()

    if (-not (Test-Path $VenvPython)) {
        $base = Find-BasePython
        if (-not $base) {
            $base = Install-UserPython
            if (-not $base) { Write-SetupResult $false $installed $warnings @(); return 1 }
            $installed += 'python'
        }
        Write-Log "Tạo môi trường Python riêng (venv) bằng $base..."
        Invoke-Logged { & $base -m venv $VenvDir }
        if (-not (Test-Path $VenvPython)) {
            Set-SetupError 'venv' 'Không tạo được thư mục venv.' 'Xoá thư mục venv trong bộ công cụ (nếu có) rồi chạy lại lệnh cài.'
            Write-SetupResult $false $installed $warnings @()
            return 1
        }
        $installed += 'venv'
    }

    if (-not (Test-PackagesOk (Get-DoctorReport @('--no-smoke')))) {
        $pipCode = 1
        for ($attempt = 1; $attempt -le 2 -and $pipCode -ne 0; $attempt++) {
            Write-Log "Cài thư viện Python (lần $attempt, có thể mất vài phút)..."
            Invoke-Logged { & $VenvPython -m pip install --upgrade pip }
            Invoke-Logged { & $VenvPython -m pip install -r (Join-Path $RepoRoot 'requirements.txt') }
            $pipCode = $LASTEXITCODE
        }
        if ($pipCode -ne 0) {
            Set-SetupError 'packages' 'Cài thư viện Python thất bại.' 'Xem mục "Cài thư viện thất bại" trong docs/vi/xu-ly-loi.md; mạng trường có thể cần mở truy cập pypi.org và files.pythonhosted.org.'
            Write-SetupResult $false $installed $warnings @()
            return 1
        }
        $installed += 'packages'
    }

    $envFile = Join-Path $RepoRoot '.env'
    if (-not (Test-Path $envFile)) {
        Copy-Item (Join-Path $RepoRoot '.env.example') $envFile
        $installed += 'env'
    }

    Write-Log 'Kiểm tra lại toàn bộ (có xuất thử một file PPTX)...'
    $report = Get-DoctorReport @()
    if (-not $report) {
        Set-SetupError 'doctor' 'Không đọc được kết quả kiểm tra môi trường.' 'Chạy KIEM-TRA.bat để xem chi tiết.'
        Write-SetupResult $false $installed $warnings @()
        return 1
    }
    $ready = [bool]$report.ready
    Write-SetupResult $ready $installed $warnings @($report.checks)
    if ($ready) { return 0 }
    return 1
}

function Find-ToolDir([string]$ToolName) {
    $exe = $OptionalTools[$ToolName].Exe
    $cmd = Get-Command $exe -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($cmd) { return (Split-Path -Path $cmd.Source -Parent) }
    if (-not $env:LOCALAPPDATA) { return $null }
    $candidates = @()
    if ($ToolName -eq 'ffmpeg') {
        $candidates += Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Links\ffmpeg.exe'
        $packages = Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Packages'
        if (Test-Path $packages) {
            foreach ($folder in @(Get-ChildItem -Path $packages -Directory -Filter 'Gyan.FFmpeg*' -ErrorAction SilentlyContinue)) {
                $candidates += @(Get-ChildItem -Path $folder.FullName -Filter 'ffmpeg.exe' -Recurse -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })
            }
        }
    } else {
        $candidates += Join-Path $env:LOCALAPPDATA 'Pandoc\pandoc.exe'
    }
    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) { return (Split-Path -Path $candidate -Parent) }
    }
    return $null
}

function Invoke-Tool {
    if (-not $Name) {
        Write-Json ([pscustomobject]@{ tool = $null; found = $false; installed = $false; dir = $null; error = [pscustomobject]@{ step = 'tool'; message = 'Thiếu tham số -Name.'; fix = 'Chạy lại với -Name ffmpeg hoặc -Name pandoc.' } })
        return 1
    }
    $tool = $OptionalTools[$Name]
    $dir = Find-ToolDir $Name
    if ($PlanOnly) {
        $steps = @()
        if (-not $dir) {
            $method = if (Test-Winget) { 'winget' } else { 'manual' }
            $steps += [pscustomobject]@{ step = $Name; action = 'install'; method = $method }
        }
        Write-Json ([pscustomobject]@{ tool = $Name; found = [bool]$dir; dir = $dir; steps = @($steps) })
        return 0
    }
    if ($dir) {
        Write-Json ([pscustomobject]@{ tool = $Name; found = $true; installed = $false; dir = $dir; error = $null })
        return 0
    }
    if (-not (Test-Winget)) {
        Write-Json ([pscustomobject]@{ tool = $Name; found = $false; installed = $false; dir = $null; error = [pscustomobject]@{ step = 'tool'; message = "Máy không có winget nên không tự cài được $Name."; fix = "Tải và cài thủ công tại $($tool.Manual)" } })
        return 1
    }
    $id = $tool.Id
    Write-Log "Cài $Name cho tài khoản này bằng winget..."
    Invoke-Logged { winget install -e --id $id --scope user --silent --accept-package-agreements --accept-source-agreements --disable-interactivity }
    $dir = Find-ToolDir $Name
    if (-not $dir) {
        Write-Json ([pscustomobject]@{ tool = $Name; found = $false; installed = $false; dir = $null; error = [pscustomobject]@{ step = 'tool'; message = "winget không cài được $Name."; fix = "Tải và cài thủ công tại $($tool.Manual)" } })
        return 1
    }
    Write-Json ([pscustomobject]@{ tool = $Name; found = $true; installed = $true; dir = $dir; error = $null })
    return 0
}

function Invoke-Check {
    $py = Resolve-RunPython
    if (-not $py) { Write-Host 'Hãy bấm CAI-DAT.bat trước.'; return 1 }
    return (Invoke-Doctor $py @())
}

function Invoke-Update {
    if (-not (Test-Path (Join-Path $RepoRoot '.git'))) {
        Write-Fail 'Thư mục này được tải dạng ZIP nên không tự cập nhật được.'
        Write-Host 'Tải bản mới tại https://github.com/luonghaianh1208/PPTmaster rồi chép thư mục projects\ và file .env của bạn sang.'
        return 1
    }
    $py = Resolve-RunPython
    if (-not $py) { Write-Host 'Hãy bấm CAI-DAT.bat trước.'; return 1 }
    Write-Step 'Tải bản mới nhất'
    & $py.Path (Join-Path $RepoRoot 'skills\ppt-master\scripts\update_repo.py') | Out-Host
    if ($LASTEXITCODE -ne 0) {
        Write-Fail 'Cập nhật thất bại.'
        Write-Host "Nếu thông báo có 'Tracked local changes' là bạn đã sửa file của bộ công cụ. Cách xử lý: $FixDoc (mục Cập nhật thất bại)"
        return 1
    }
    Write-Step 'Kiểm tra nhanh sau cập nhật'
    return (Invoke-Doctor $py @('--no-smoke'))
}

Push-Location $RepoRoot
try {
    switch ($Action) {
        'setup' { if ($Auto -or $PlanOnly) { $code = Invoke-AutoSetup } else { $code = Invoke-Setup } }
        'check' { $code = Invoke-Check }
        'update' { $code = Invoke-Update }
        'tool' { $code = Invoke-Tool }
    }
} finally {
    Pop-Location
}
exit ([int]($code | Select-Object -Last 1))
```

Sau khi ghi file, thêm BOM nếu thiếu:

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -c "import pathlib; p = pathlib.Path('tools/vi/pptmaster.ps1'); b = p.read_bytes(); p.write_bytes(b if b.startswith(b'\xef\xbb\xbf') else b'\xef\xbb\xbf' + b); print(p.read_bytes()[:3])"
```
Expected: in `b'\xef\xbb\xbf'`.

- [ ] **Step 4: Chạy test, xác nhận đạt**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -v
```
Expected: toàn bộ test OK, gồm 12 test của `InstallerPlanTest` và `test_powershell_scripts_have_utf8_bom`. Test lỗi → sửa script (không nới test), chạy lại.

- [ ] **Step 5: Kiểm chứng chạy thật trên bản sao tạm (máy này đã có Python; không cài phần mềm hệ thống, không mở cửa sổ)**

Việc này tạo `venv` và cài thư viện trong một bản sao tạm, cần mạng, có thể mất vài phút.

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
$SCRATCH = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad'
$env:PYTHONIOENCODING = 'utf-8'
$repo = Join-Path $SCRATCH 'tu-cai-auto'
git clone --local --no-hardlinks . $repo
Copy-Item tools\vi\pptmaster.ps1 (Join-Path $repo 'tools\vi\pptmaster.ps1') -Force
Copy-Item tools\vi\doctor.py (Join-Path $repo 'tools\vi\doctor.py') -Force
Push-Location $repo
$run = "import json, subprocess, sys; p = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', r'tools\vi\pptmaster.ps1'] + sys.argv[1:], capture_output=True, timeout=1800); out = p.stdout.decode('utf-8'); print('exit', p.returncode, 'stdout_lines', len(out.strip().splitlines())); d = json.loads(out); print({k: d[k] for k in d if k != 'checks'}); print([(c['name'], c['ok']) for c in d.get('checks', [])])"
& $PY -c $run -Action setup -Auto
& $PY -c $run -Action setup -Auto
& $PY -c $run -Action tool -Name pandoc
Pop-Location
```
Expected:
- Lần chạy thứ nhất: `exit 0 stdout_lines 1`; `ready` True; `installed` chứa `venv`, `packages`, `env` (không có `python` vì máy đã có Python); `error` None; mục `Xuất thử PPTX` True.
- Lần thứ hai: `exit 0`; `installed` rỗng (`[]`).
- `tool -Name pandoc`: nếu máy có Pandoc → `found` True, `installed` False; nếu không có → ghi kết quả vào báo cáo (không bắt buộc cài).
- Ghi toàn bộ đầu ra vào báo cáo. Nếu lần một lỗi do mạng, ghi lại và chạy lại một lần.

Sau đó chạy `check` bằng venv của bản sao:

```powershell
$SCRATCH = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad'
Push-Location (Join-Path $SCRATCH 'tu-cai-auto')
powershell -NoProfile -ExecutionPolicy Bypass -File tools\vi\pptmaster.ps1 -Action check
"check exit=$LASTEXITCODE"
Pop-Location
```
Expected: dòng đầu có `Python của venv tại`; `check exit=0`. Không xoá thư mục `tu-cai-auto` trong task này (Task 6 dùng lại bản sao sạch khác).

- [ ] **Step 6: Commit**

```powershell
git status --porcelain
git add tools/vi/pptmaster.ps1 tools/vi/tests/test_installer.py
git commit -m "feat(vi): add unattended setup mode for AI agents" -m "Co-Authored-By: Claude <noreply@anthropic.com>"
```
Expected: `git status` trước commit chỉ có hai file trên. (Thay dòng `Co-Authored-By` bằng dòng ghi công đúng của phiên.)

---

### Task 4: Hướng dẫn cài cho AI, `AGENTS.vi.md` và README

**Files:**
- Create: `docs/vi/cai-dat-bang-ai.md`
- Modify: `AGENTS.vi.md` (mục 4 và mục 9), `README.md` (khối "Dành cho AI agent", một dòng trong "Bắt đầu trong 3 bước")
- Test: `tools/vi/tests/test_vi_layer.py`

**Interfaces:**
- Consumes: lệnh và JSON của Task 3; trong `test_vi_layer.py` đã có `read()`, `h2_headings()`, `section()`, `REQUIRED_DOCS`, `AGENTS_VI_ASSISTANT_HEADING`.
- Produces: hằng `SELF_INSTALL_DOC`, `SELF_INSTALL_HEADINGS`, `AGENTS_VI_ENV_HEADING`, class `SelfInstallGuideTest`; Task 5 thêm test vào cùng file, trước `if __name__ == "__main__":`.

- [ ] **Step 1: Viết test lỗi**

Trong `tools/vi/tests/test_vi_layer.py`, thêm `"cai-dat-bang-ai.md",` vào `REQUIRED_DOCS` (sau `"cai-dat-windows.md",`). Thêm trước dòng `if __name__ == "__main__":`:

```python
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

    def test_setup_section_runs_exact_auto_command_and_reads_stdout_only(self):
        body = section(read(SELF_INSTALL_DOC), "## Cài đặt")
        self.assertIn(AUTO_SETUP_COMMAND, body)
        self.assertIn("2>&1", body)

    def test_result_section_handles_blocked_powershell_with_manual_commands(self):
        body = section(read(SELF_INSTALL_DOC), "## Đọc kết quả")
        for phrase in ("running scripts is disabled", "PowerShell bị chặn trên máy trường hoặc công ty", "tối đa một lần"):
            self.assertIn(phrase, body)

    def test_report_section_uses_it_message_and_resumes_request(self):
        body = section(read(SELF_INSTALL_DOC), "## Báo thầy cô")
        for phrase in ("Máy trường chặn cài đặt", "Tạo bài giảng", "làm tiếp"):
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
        for phrase in ("(docs/vi/cai-dat-bang-ai.md)", "doctor.py --no-smoke --json", "trước lần tạo slide đầu tiên", "KIEM-TRA.bat", "Công cụ tuỳ chọn"):
            self.assertIn(phrase, body)
        self.assertEqual(h2_headings(text)[-1], AGENTS_VI_ASSISTANT_HEADING)

    def test_readme_tells_agents_to_follow_guide_before_quick_start(self):
        readme = read("README.md")
        marker = "**Dành cho AI agent:**"
        self.assertIn(marker, readme)
        self.assertIn("(docs/vi/cai-dat-bang-ai.md)", readme)
        self.assertLess(readme.index(marker), readme.index("## Bắt đầu trong 3 bước"))
        self.assertIn("bat-dau-nhanh.md#để-ai-tự-cài", section(readme, "## Bắt đầu trong 3 bước"))
```

- [ ] **Step 2: Chạy test, xác nhận lỗi**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -p test_vi_layer.py -v
```
Expected: FAIL/ERROR ở `test_required_vietnamese_docs_exist` và các test của `SelfInstallGuideTest` (thiếu file, thiếu mục). Link `bat-dau-nhanh.md#để-ai-tự-cài` chỉ kiểm phần đường dẫn nên `test_relative_markdown_links_resolve` vẫn đạt.

- [ ] **Step 3: Viết nội dung**

Tạo `docs/vi/cai-dat-bang-ai.md` với nội dung sau (giữ nguyên, mỗi khối lệnh là một khối code riêng):

````markdown
# Cài đặt bộ công cụ bằng AI

File dành cho AI agent (Antigravity, Claude Code, Cursor) trên Windows. Làm đúng các bước dưới đây, không tự nghĩ cách cài khác. Xưng "em", gọi người dùng là "thầy cô".

## Khi nào dùng

- Người dùng nhờ cài PPT Master, dán link `https://github.com/luonghaianh1208/PPTmaster`, hoặc nhờ "cài đặt", "kiểm tra máy".
- `AGENTS.vi.md` mục 9 báo môi trường chưa sẵn sàng trước khi tạo slide.
- Máy không phải Windows: không dùng file này; chạy `sh tools/vi/setup.sh` như README.

## Tải bộ công cụ

Bỏ qua mục này nếu thư mục đang mở đã có `tools/vi/pptmaster.ps1`.

1. Chọn thư mục đích: thư mục đang mở còn trống thì dùng chính nó; đã có file khác thì dùng thư mục con `PPTmaster` bên trong (tạo bằng `New-Item -ItemType Directory -Force <thư_mục_đích>`).
2. Chọn nhánh: mặc định là `main`. Người dùng nêu rõ một nhánh khác (ví dụ "nhánh feat/vi-tu-cai") thì dùng nhánh đó.
3. Có Git (`git --version` chạy được):

   ```
   git clone -b <nhánh> https://github.com/luonghaianh1208/PPTmaster.git <thư_mục_đích>
   ```

   Với nhánh `main` có thể bỏ `-b <nhánh>`: `git clone https://github.com/luonghaianh1208/PPTmaster.git <thư_mục_đích>`.

4. Không có Git: tải ZIP của nhánh rồi giải nén bằng PowerShell:

   ```
   $zip = Join-Path $env:TEMP 'PPTmaster.zip'
   $unzip = Join-Path $env:TEMP ('PPTmaster-zip-' + [guid]::NewGuid())
   Invoke-WebRequest -Uri 'https://github.com/luonghaianh1208/PPTmaster/archive/refs/heads/<nhánh>.zip' -OutFile $zip -UseBasicParsing
   Expand-Archive -Path $zip -DestinationPath $unzip -Force
   $src = Get-ChildItem -Path $unzip -Directory | Select-Object -First 1
   Get-ChildItem -Path $src.FullName -Force | Move-Item -Destination '<thư_mục_đích>'
   ```

5. Đọc `AGENTS.md` và `AGENTS.vi.md` trong thư mục đích và áp dụng từ đây. Mọi lệnh sau chạy từ thư mục đích.

## Cài đặt

1. Gửi thầy cô đúng một tin nhắn, không hỏi lại: "Em sẽ cài Python và thư viện cho PPT Master, mất khoảng 5–10 phút. Nếu Antigravity hỏi cho phép chạy lệnh, thầy cô bấm đồng ý giúp em."
2. Chạy từ thư mục đích:

   ```
   powershell -NoProfile -ExecutionPolicy Bypass -File tools\vi\pptmaster.ps1 -Action setup -Auto
   ```

3. Tiến trình cài hiện ở stderr. Kết quả là đúng một dòng JSON ở stdout; đọc riêng stdout, không gộp stderr vào (không dùng `2>&1`).

## Đọc kết quả

- `ready` là `true`: sang mục "Báo thầy cô".
- `error` khác `null`, hoặc `ready` là `false`: đọc `error.fix` và trường `fix` của các mục bắt buộc chưa đạt trong `checks`. Việc nào AI làm được ngay trong thư mục bộ công cụ thì làm, rồi chạy lại đúng lệnh ở mục "Cài đặt", tối đa một lần. Vẫn chưa sẵn sàng thì báo bị chặn.
- Lệnh cài không chạy được vì PowerShell bị chặn chạy script (thông báo có "running scripts is disabled" hoặc "execution policy"): nếu máy đã có Python 3.10 trở lên, chạy các lệnh tay trong mục "PowerShell bị chặn trên máy trường hoặc công ty" của `docs/vi/xu-ly-loi.md`, rồi kiểm tra bằng `python tools/vi/doctor.py --json`. Chưa có Python thì báo bị chặn.
- `warnings` có nội dung: chép lại cho thầy cô trong tin nhắn cuối.

## Báo thầy cô

- Sẵn sàng: nói ngắn gọn đã cài gì theo `installed` (`python` = Python 3.12, `venv` = môi trường Python riêng, `packages` = thư viện, `env` = file cấu hình; danh sách rỗng thì nói máy đã cài sẵn từ trước), kết quả mục "Xuất thử PPTX", các cảnh báo nếu có, và một câu lệnh mẫu như `Tạo bài giảng Toán 6 bài Phân số`. Nếu vừa tải bộ công cụ về, thêm: "Lần sau thầy cô mở thư mục <thư_mục_đích> trong Antigravity là dùng được ngay."
- Bị chặn: nêu bước lỗi và `error.message` bằng lời dễ hiểu, cách tự xử lý nếu có, rồi chép nguyên đoạn gửi bộ phận IT trong mục "Máy trường chặn cài đặt" của `docs/vi/xu-ly-loi.md`.
- Thầy cô đã gửi yêu cầu tạo slide trước đó: khi đã sẵn sàng thì làm tiếp yêu cầu đó.

## Công cụ tuỳ chọn

- Chỉ cài khi cần: thuyết minh và video cần FFmpeg; tài liệu định dạng cũ (`.doc`, `.odt`, `.rtf`…) cần Pandoc.
- Chạy (thay `ffmpeg` bằng `pandoc` khi cần Pandoc):

   ```
   powershell -NoProfile -ExecutionPolicy Bypass -File tools\vi\pptmaster.ps1 -Action tool -Name ffmpeg
   ```

- JSON có `found` là `true`: lấy `dir`, rồi chạy lệnh cần công cụ đó trong cùng một lệnh PowerShell, thêm thư mục vào PATH trước: `$env:Path = "<dir>;$env:Path"; <lệnh>`.
- `found` là `false`: báo thầy cô `error.message` và `error.fix`; làm tiếp những phần không cần công cụ đó.

## Không được làm

- Không chạy script tải thẳng từ mạng (`iex`, `Invoke-Expression`, `irm … | iex`).
- Không xin quyền quản trị, không chạy lệnh ở chế độ "Run as administrator".
- Không tắt phần mềm diệt virus hay tường lửa.
- Không đổi ExecutionPolicy vĩnh viễn (`Set-ExecutionPolicy`); chỉ dùng `-ExecutionPolicy Bypass` cho từng lệnh như trên.
- Không cài phần mềm nào khác ngoài Python 3.12, thư viện trong `requirements.txt`, Pandoc và FFmpeg; chỉ cài Git khi người dùng yêu cầu.
- Không tự nghĩ cách cài khác thay cho lệnh ở mục "Cài đặt" (ví dụ tự gõ `pip install`, tự tải Python từ nơi khác), trừ các lệnh tay được nêu ở mục "Đọc kết quả".
````

Trong `AGENTS.vi.md`:

1. Trong mục `## 4. Chạy lệnh trên Windows`, thêm dòng đầu tiên của danh sách (trước dòng "Nếu lệnh `python3 ...` báo không tìm thấy…"):

```markdown
- Có `venv\Scripts\python.exe` ở thư mục gốc repo thì dùng nó cho mọi lệnh `python3 …` hoặc `python …` của repo, ví dụ `venv\Scripts\python.exe skills/ppt-master/scripts/project_manager.py init <tên_dự_án>`.
```

2. Thay toàn bộ mục 9 (từ dòng `## 9. Khi người dùng gặp lỗi môi trường` đến trước `## 10.`) bằng:

```markdown
## 9. Môi trường: tự kiểm tra, tự cài và xử lý lỗi

- Mỗi cuộc trò chuyện, trước lần tạo slide đầu tiên: chạy `tools/vi/doctor.py --no-smoke --json` (bằng Python của `venv` nếu có). Kết quả có `"ready": true` thì làm tiếp; chưa sẵn sàng, hoặc không chạy được Python, thì làm theo [docs/vi/cai-dat-bang-ai.md](docs/vi/cai-dat-bang-ai.md) rồi mới tạo slide.
- Người dùng nhờ cài đặt, kiểm tra máy, hoặc dán link repo để cài: làm theo [docs/vi/cai-dat-bang-ai.md](docs/vi/cai-dat-bang-ai.md).
- Cần FFmpeg (thuyết minh, video) hoặc Pandoc (tài liệu định dạng cũ như `.doc`, `.odt`, `.rtf`) mà máy chưa có: làm theo mục "Công cụ tuỳ chọn" của file đó.
- Người dùng gặp lỗi môi trường: đề nghị chạy `KIEM-TRA.bat` (Windows) hoặc `python tools/vi/doctor.py`, rồi đối chiếu với [docs/vi/xu-ly-loi.md](docs/vi/xu-ly-loi.md).
```

Trong `README.md`:

1. Chèn giữa đoạn giới thiệu (dòng bắt đầu "Đây là bản Việt hoá…") và dòng `Phiên bản:` một khối, mỗi phần cách một dòng trống:

```markdown
> **Dành cho AI agent:** người dùng nhờ cài bộ công cụ này (hoặc chỉ dán link repo) thì làm theo [docs/vi/cai-dat-bang-ai.md](docs/vi/cai-dat-bang-ai.md), không tự nghĩ cách cài khác.
```

2. Trong mục `## Bắt đầu trong 3 bước`, chèn trước dòng `macOS/Linux: chạy \`sh tools/vi/setup.sh\`.` (cách một dòng trống):

```markdown
Dùng Antigravity và muốn AI làm hết: mở một thư mục trống rồi dán câu lệnh mẫu trong [Bắt đầu nhanh](docs/vi/bat-dau-nhanh.md#để-ai-tự-cài), AI tự tải, tự cài và báo khi sẵn sàng.
```

- [ ] **Step 4: Chạy test, xác nhận đạt**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -v
```
Expected: toàn bộ test OK (kể cả `test_agents_vi_has_assistant_section_last` và `test_relative_markdown_links_resolve`).

- [ ] **Step 5: Commit**

```powershell
git add docs/vi/cai-dat-bang-ai.md AGENTS.vi.md README.md tools/vi/tests/test_vi_layer.py
git commit -m "docs(vi): guide AI agents through unattended setup" -m "Co-Authored-By: Claude <noreply@anthropic.com>"
```

---

### Task 5: Tài liệu cho thầy cô và bảo trì

**Files:**
- Modify: `docs/vi/bat-dau-nhanh.md` (thêm mục `## Để AI tự cài` ngay dưới tiêu đề `# Bắt đầu nhanh`)
- Modify: `docs/vi/cai-dat-windows.md` (khối "Cách nhanh" ngay dưới tiêu đề `# Cài đặt trên Windows`)
- Modify: `docs/vi/xu-ly-loi.md` (mục `## Máy trường chặn cài đặt` ngay sau mục `## PowerShell bị chặn trên máy trường hoặc công ty`)
- Modify: `docs/vi/phat-trien/bao-tri.md` (mục `## Bộ cài Python cố định` ngay sau mục `## Chạy test`)
- Test: `tools/vi/tests/test_vi_layer.py`

**Interfaces:**
- Consumes: `section()`, `h2_headings()`, `read()`; mục "Máy trường chặn cài đặt" được `docs/vi/cai-dat-bang-ai.md` (Task 4) nhắc tới; anchor `#để-ai-tự-cài` được README (Task 4) trỏ tới.
- Produces: class `SelfInstallUserDocsTest`.

- [ ] **Step 1: Viết test lỗi**

Thêm vào `tools/vi/tests/test_vi_layer.py`, trước `if __name__ == "__main__":`:

```python
class SelfInstallUserDocsTest(unittest.TestCase):
    def test_quick_start_explains_ai_setup_first(self):
        text = read("docs/vi/bat-dau-nhanh.md")
        self.assertEqual(h2_headings(text)[0], "## Để AI tự cài")
        body = section(text, "## Để AI tự cài")
        for phrase in ("https://github.com/luonghaianh1208/PPTmaster", "cho phép chạy lệnh", "5–10 phút", "cài đặt giúp em", "(xu-ly-loi.md#máy-trường-chặn-cài-đặt)"):
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
        for phrase in ("python.org", "pypi.org", "files.pythonhosted.org", "github.com", "codeload.github.com", "Python 3.12", "-ExecutionPolicy Bypass"):
            self.assertIn(phrase, body)

    def test_maintenance_doc_explains_pinned_python_installer(self):
        body = section(read("docs/vi/phat-trien/bao-tri.md"), "## Bộ cài Python cố định")
        for phrase in ("3.12.10", "SHA256", "MD5", "$PythonInstallers", "Get-UserPythonPath"):
            self.assertIn(phrase, body)
```

- [ ] **Step 2: Chạy test, xác nhận lỗi**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -p test_vi_layer.py -v
```
Expected: 4 test mới FAIL/ERROR (thiếu mục).

- [ ] **Step 3: Viết nội dung**

`docs/vi/bat-dau-nhanh.md` — chèn ngay sau dòng `# Bắt đầu nhanh` (cách một dòng trống, trước `## Mở thư mục trong AI editor`):

````markdown
## Để AI tự cài

Dùng Antigravity và máy chưa cài gì: không cần tải ZIP hay bấm `CAI-DAT.bat`.

1. Mở Antigravity → **Open Folder** → chọn một thư mục trống, ví dụ `D:\PPTmaster` (tránh thư mục đang đồng bộ OneDrive).
2. Mở khung **Agent**, dán câu lệnh:

   ```
   Cài PPT Master từ https://github.com/luonghaianh1208/PPTmaster vào thư mục này rồi báo khi sẵn sàng tạo slide
   ```

3. Chờ khoảng 5–10 phút. Khi Antigravity hỏi cho phép chạy lệnh, bấm đồng ý.
4. AI báo "sẵn sàng" kèm một câu lệnh mẫu là dùng được. Lần sau chỉ cần mở lại thư mục này.

Đã tải ZIP và mở thư mục rồi: nhắn "cài đặt giúp em", hoặc gửi luôn yêu cầu tạo slide. AI tự kiểm tra và cài những gì còn thiếu trước khi làm.

Máy trường chặn cài đặt: AI sẽ báo lý do và đưa một đoạn để gửi bộ phận IT (xem [Xử lý lỗi](xu-ly-loi.md#máy-trường-chặn-cài-đặt)).
````

`docs/vi/cai-dat-windows.md` — chèn ngay sau dòng `# Cài đặt trên Windows` (cách một dòng trống, trước `## Cần chuẩn bị`):

```markdown
> **Cách nhanh: nhờ AI cài.** Dùng Antigravity thì chỉ cần mở một thư mục trống và dán câu lệnh mẫu trong [Bắt đầu nhanh](bat-dau-nhanh.md#để-ai-tự-cài); AI tự tải bộ công cụ, tự cài Python và thư viện. Các bước bấm đúp dưới đây dành cho ai muốn tự cài.
```

`docs/vi/xu-ly-loi.md` — chèn ngay trước dòng `## Cài thư viện thất bại`:

```markdown
## Máy trường chặn cài đặt

Dấu hiệu: AI báo không tải được bộ công cụ hoặc Python, không chạy được bộ cài Python, không cài được thư viện vì mạng, hoặc PowerShell bị chặn chạy script. Máy do nhà trường quản lý có thể chặn các việc này; AI không tìm cách vượt qua.

Gửi bộ phận IT đoạn sau:

> Nhờ anh/chị hỗ trợ để tôi dùng bộ công cụ PPT Master trên máy này:
> 1. Cho tài khoản Windows của tôi truy cập: python.org, pypi.org, files.pythonhosted.org, github.com, codeload.github.com.
> 2. Cho phép cài Python 3.12 cho riêng tài khoản của tôi (không cần quyền quản trị).
> 3. Cho phép chạy PowerShell với tuỳ chọn `-ExecutionPolicy Bypass` cho từng lệnh.

Nếu chỉ bị chặn chạy script PowerShell mà máy đã có Python, xem mục **PowerShell bị chặn trên máy trường hoặc công ty** ở trên.

```

`docs/vi/phat-trien/bao-tri.md` — chèn ngay trước dòng `## Đánh số phiên bản`:

```markdown
## Bộ cài Python cố định

Chế độ `-Auto` của `tools/vi/pptmaster.ps1` cài Python bằng winget; khi máy không có winget hoặc winget lỗi, script dùng bộ cài python.org với phiên bản `3.12.10` (bản 3.12 cuối cùng có bộ cài cho Windows), đường dẫn và mã SHA256 ghi cứng trong biến `$PythonVersion` và `$PythonInstallers` cho `amd64` và `arm64`.

Khi đổi phiên bản:

1. Tải hai file cài từ `https://www.python.org/ftp/python/<phiên bản>/` và tính mã bằng `Get-FileHash -Algorithm SHA256`.
2. Tính thêm `Get-FileHash -Algorithm MD5` và đối chiếu với trang phát hành trên python.org trước khi ghi SHA256 vào script.
3. Cập nhật `$PythonVersion`, `$PythonInstallers`, mã gói winget `Python.Python.3.12` và thư mục `Python312` trong `Get-UserPythonPath` nếu đổi nhánh phiên bản (ví dụ lên 3.13).
4. Chạy lại test, rồi thử `-Action setup -Auto` trên một máy chưa có Python.

```

- [ ] **Step 4: Chạy test, xác nhận đạt**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -v
```
Expected: toàn bộ test OK, gồm `test_troubleshooting_has_sections_referenced_by_launcher`, `test_quick_start_explains_intake_profile_and_chat_confirmation` và `test_relative_markdown_links_resolve`.

- [ ] **Step 5: Commit**

```powershell
git add docs/vi/bat-dau-nhanh.md docs/vi/cai-dat-windows.md docs/vi/xu-ly-loi.md docs/vi/phat-trien/bao-tri.md tools/vi/tests/test_vi_layer.py
git commit -m "docs(vi): explain AI-driven setup to teachers" -m "Co-Authored-By: Claude <noreply@anthropic.com>"
```

---

### Task 6: Kiểm tra hành vi headless

**Files:** không sửa file trong repo (trừ khi cần sửa câu chữ ở Step 4). Kết quả ghi vào báo cáo của task.

**Interfaces:**
- Consumes: `docs/vi/cai-dat-bang-ai.md`, `AGENTS.vi.md` mục 9 (Task 4); Claude Code CLI tại `/c/Users/ADMIN/.local/bin/claude`.
- Produces: kết quả 2 ca kiểm tra cho Task 7.

- [ ] **Step 1: Tạo bản sao sạch không có `venv`**

```bash
SCRATCH=/c/Users/ADMIN/AppData/Local/Temp/claude/c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster/c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f/scratchpad
REPO="/c/Users/ADMIN/Downloads/VIBE CODING/PPTmaster"
rm -rf "$SCRATCH/tu-cai-headless-repo" "$SCRATCH/tu-cai-headless"
git clone --local --no-hardlinks --branch feat/vi-tu-cai "$REPO" "$SCRATCH/tu-cai-headless-repo"
mkdir -p "$SCRATCH/tu-cai-headless"
ls "$SCRATCH/tu-cai-headless-repo/venv" 2>/dev/null; echo "venv-absent=$?"
```
Expected: `venv-absent=2` (không có `venv`). Nếu `rm -rf` bị chặn bởi kiểm tra an toàn, dùng tên thư mục mới có hậu tố `-2`.

- [ ] **Step 2: Chạy 2 ca (tối đa 2 ca mỗi lần gọi Bash, timeout 600000)**

```bash
SCRATCH=/c/Users/ADMIN/AppData/Local/Temp/claude/c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster/c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f/scratchpad
OUT="$SCRATCH/tu-cai-headless"
cd "$SCRATCH/tu-cai-headless-repo"
run_case() { timeout 300 /c/Users/ADMIN/.local/bin/claude -p "$2" --disallowedTools Bash Write Edit NotebookEdit PowerShell Agent Monitor Skill > "$OUT/case$1.txt" 2>&1; echo "case$1 exit=$?"; }
run_case 1 "cài đặt giúp em"
run_case 2 "Tạo bài giảng Toán 6 bài Phân số"
```
Expected: mỗi ca in `exit=0` (hoặc `124` khi hết giờ; ghi lại).

- [ ] **Step 3: Chấm**

```bash
SCRATCH=/c/Users/ADMIN/AppData/Local/Temp/claude/c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster/c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f/scratchpad
OUT="$SCRATCH/tu-cai-headless"
has() { grep -q -- "$2" "$OUT/case$1.txt"; }
first_line() { grep -n -m1 -- "$2" "$OUT/case$1.txt" | cut -d: -f1; }
{ has 1 "-Action setup -Auto" || has 1 "cai-dat-bang-ai.md"; } && ! grep -qiE "\biex\b|Invoke-Expression" "$OUT/case1.txt" && echo "case1 PASS" || echo "case1 FAIL"
D=$(first_line 2 "doctor.py\|cai-dat-bang-ai.md\|-Action setup -Auto"); I=$(first_line 2 "project_manager.py init")
{ [ -n "$D" ] && { [ -z "$I" ] || [ "$D" -lt "$I" ]; }; } && echo "case2 PASS" || echo "case2 FAIL"
git -C "$SCRATCH/tu-cai-headless-repo" status --porcelain; echo "(status end)"
```
Expected: `case1 PASS`, `case2 PASS`, status rỗng. Đọc thêm từng file đầu ra để xác nhận AI không tự nghĩ cách cài khác (ví dụ tự gõ `pip install` như bước cài chính); ghi 10–20 dòng trích đoạn mỗi ca vào báo cáo. Nếu `grep` bỏ sót vì chữ hoa/thường tiếng Việt, chấm thủ công và ghi rõ.

- [ ] **Step 4: Nếu một ca FAIL vì câu chữ hướng dẫn**

Sửa tối thiểu câu chữ ở `docs/vi/cai-dat-bang-ai.md` hoặc `AGENTS.vi.md` mục 9 (giữ nguyên các cụm mà test ở Task 4 yêu cầu), chạy lại toàn bộ test, commit `fix(vi): clarify self-install guidance`, tạo lại bản sao ở Step 1 và chạy lại cả 2 ca. Tối đa 2 vòng; vẫn FAIL thì ghi rõ trong báo cáo, không nới tiêu chí chấm.

---

### Task 7: Nghiệm thu trên máy thật (điểm dừng)

**Files:** không có.

**Interfaces:**
- Consumes: nhánh `feat/vi-tu-cai` sau Task 6.
- Produces: xác nhận "đạt" của chủ repo cho tiêu chí thành công #1, #2, #3, #4 của spec.

- [ ] **Step 1: Kiểm tra toàn nhánh và đẩy nhánh lên để nghiệm thu**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -v
& $PY skills/ppt-master/scripts/attribution_guard.py; "guard=$LASTEXITCODE"
git diff --stat v6.3.2 HEAD -- skills/; "(skills diff end)"
git status --porcelain; "(status end)"
git push -u origin feat/vi-tu-cai
```
Expected: test OK; `guard=0`; diff `skills/` rỗng; status rỗng; push thường thành công (không `--force`).

- [ ] **Step 2: Gửi chủ repo danh sách nghiệm thu và chờ trả lời**

Nhờ chủ repo (sẽ mở Antigravity, trình duyệt tải Antigravity nếu cần, và PowerPoint — báo trước):
1. Dùng một máy Windows chưa có Python (hoặc một tài khoản Windows mới trên máy chưa cài Python cho mọi người dùng); cài Antigravity.
2. Mở một thư mục trống ngắn (ví dụ `D:\PPTmaster`), mở khung Agent, dán: `Cài PPT Master từ https://github.com/luonghaianh1208/PPTmaster (nhánh feat/vi-tu-cai) vào thư mục này rồi báo khi sẵn sàng tạo slide`.
3. Kiểm: AI không hỏi Y/N; chỉ cần bấm cho phép chạy lệnh; AI báo sẵn sàng, mục "Xuất thử PPTX" đạt; không phải mở lại Antigravity.
4. Nhắn tạo một bài giảng thật → nhận PPTX, mở bằng PowerPoint.
5. Đóng Antigravity, mở lại thư mục, tạo bài thứ hai → AI kiểm tra nhanh, không cài lại.
6. (Tuỳ chọn) Nhắn thêm thuyết minh cho bài → AI cài FFmpeg cho tài khoản rồi làm tiếp.

Không sang Task 8 khi chưa có xác nhận "đạt". Nếu không đạt: ghi mô tả lỗi, sửa tối thiểu qua một vòng sửa có review, chạy lại Task 6 và Step 1 của task này, rồi nhờ nghiệm thu lại. Không xoá nhánh `feat/vi-tu-cai` trên `origin` (phải hỏi chủ repo).

---

### Task 8: Phát hành `v6.3.2-vi.3`

**Files:**
- Modify: `CHANGELOG-VI.md` (thêm mục mới ngay dưới tiêu đề `# Nhật ký thay đổi — Bản Việt`)
- Modify: `README.md` (dòng `Phiên bản:`)

**Interfaces:**
- Consumes: xác nhận "đạt" (Task 7); `gh` đã đăng nhập tài khoản `luonghaianh1208`.
- Produces: `origin/main` chứa tính năng, tag `v6.3.2-vi.3`, GitHub Release `v6.3.2-vi.3`.

- [ ] **Step 1: Cập nhật CHANGELOG và dòng phiên bản**

Lấy ngày: `Get-Date -Format yyyy-MM-dd`. Thêm vào `CHANGELOG-VI.md` ngay dưới dòng `# Nhật ký thay đổi — Bản Việt` (trước mục `## 6.3.2-vi.2 — …`), thay `NGAY_PHAT_HANH` bằng ngày vừa lấy:

```markdown
## 6.3.2-vi.3 — NGAY_PHAT_HANH

AI tự cài đặt: thầy cô chỉ cần dán link repo vào Antigravity (hoặc mở thư mục đã tải), AI tự kiểm tra máy, cài những gì còn thiếu rồi báo khi sẵn sàng tạo slide.

### Thêm
- `tools/vi/pptmaster.ps1 -Action setup -Auto`: cài Python 3.12 cho riêng tài khoản (không cần quyền quản trị), tạo môi trường Python riêng `venv\`, cài thư viện, tạo `.env`, kiểm tra có xuất thử PPTX, trả kết quả JSON cho AI.
- `-Action tool -Name ffmpeg|pandoc`: chỉ cài FFmpeg hoặc Pandoc khi cần.
- `tools/vi/doctor.py --json`.
- Hướng dẫn cài cho AI `docs/vi/cai-dat-bang-ai.md`; `AGENTS.vi.md` mục 9 cho AI tự kiểm tra môi trường trước lần tạo slide đầu tiên.
- Tài liệu: mục "Để AI tự cài" trong Bắt đầu nhanh, mục "Máy trường chặn cài đặt" kèm đoạn gửi bộ phận IT.

### Không thay đổi
- `CAI-DAT.bat`, `KIEM-TRA.bat`, `CAP-NHAT.bat` vẫn dùng như cũ (`KIEM-TRA.bat` và `CAP-NHAT.bat` ưu tiên `venv` nếu có).
- Lõi PPT Master v6.3.2 của Hugo He giữ nguyên.
```

Trong `README.md`, sửa dòng `Phiên bản: **6.3.2-vi.2** · [Nhật ký thay đổi](CHANGELOG-VI.md)` thành `Phiên bản: **6.3.2-vi.3** · [Nhật ký thay đổi](CHANGELOG-VI.md)`.

- [ ] **Step 2: Chạy test và commit**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -v
git add CHANGELOG-VI.md README.md
git commit -m "docs(vi): release 6.3.2-vi.3" -m "Co-Authored-By: Claude <noreply@anthropic.com>"
```
Expected: test OK. (Thay dòng `Co-Authored-By` bằng dòng ghi công đúng của phiên.)

- [ ] **Step 3: Đưa vào `main`, gắn tag, push**

```powershell
git switch main
git pull --ff-only origin main
git merge --ff-only feat/vi-tu-cai
git tag -a v6.3.2-vi.3 -m "PPT Master ban Viet 6.3.2-vi.3"
git push origin main
git push origin v6.3.2-vi.3
git ls-remote origin refs/heads/main refs/tags/v6.3.2-vi.3
"local main=" + (git rev-parse main)
```
Expected: fast-forward thành công (không phải fast-forward: dừng, báo BLOCKED); push thường thành công (không `--force`, không `--tags`); `ls-remote` in đủ 2 ref, `main` từ xa trùng `main` trên máy.

- [ ] **Step 4: Tạo GitHub Release**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
$notes = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\release-notes-vi3.md'
& $PY -c "import re, pathlib, sys; t = pathlib.Path('CHANGELOG-VI.md').read_text(encoding='utf-8'); m = re.search(r'^## 6\.3\.2-vi\.3.*?(?=^## |\Z)', t, re.S | re.M); pathlib.Path(sys.argv[1]).write_text(m.group(0), encoding='utf-8')" $notes
gh release create v6.3.2-vi.3 --repo luonghaianh1208/PPTmaster --title "PPT Master bản Việt 6.3.2-vi.3" --notes-file $notes
gh release view v6.3.2-vi.3 --repo luonghaianh1208/PPTmaster --json url,tagName,isDraft
gh release view v6.3.2-vi.3 --repo hugohe3/ppt-master --json url
```
Expected: in URL release, `tagName` = `v6.3.2-vi.3`, `isDraft` = false; lệnh cuối báo `release not found`.

- [ ] **Step 5: Xác minh sau phát hành**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
$env:PYTHONIOENCODING = 'utf-8'
$fresh = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\freshclone-vi3'
git clone --depth 1 https://github.com/luonghaianh1208/PPTmaster.git $fresh
Push-Location $fresh
& $PY skills/ppt-master/scripts/attribution_guard.py; "guard=$LASTEXITCODE"
& $PY -m unittest discover -s tools/vi/tests
& $PY -c "import json, subprocess; p = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', r'tools\vi\pptmaster.ps1', '-Action', 'setup', '-Auto', '-PlanOnly'], capture_output=True, timeout=120); print('plan exit', p.returncode, [s['step'] for s in json.loads(p.stdout.decode('utf-8'))['steps']])"
Test-Path docs/vi/cai-dat-bang-ai.md
Pop-Location
git status --porcelain; "(status end)"
```
Expected: `guard=0`; test OK (test ranh giới upstream được bỏ qua trong bản clone nông); `plan exit 0` kèm danh sách bước; `True`; status rỗng. Sau đó xoá `$fresh`, `release-notes-vi3.md`, `tu-cai-auto`, `tu-cai-headless-repo`, `tu-cai-headless` trong thư mục tạm (mỗi đường dẫn một lệnh; bị kiểm tra an toàn chặn thì bỏ qua và ghi vào báo cáo). Không xoá nhánh `feat/vi-tu-cai` (cục bộ hay trên `origin`) — phải hỏi chủ repo.
