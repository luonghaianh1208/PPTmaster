"""Test cho công cụ dòng lệnh thi_nghiem.py."""

import contextlib
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from thi_nghiem_mau import HOOKE_JS, HOOKE_JSON, HOOKE_MD, PENDULUM  # noqa: E402
from thi_nghiem_parts import build_html, kiem_so, thu_vien  # noqa: E402

import thi_nghiem  # noqa: E402

HAS_NODE = kiem_so.find_node() is not None
NEED_NODE = "máy không có Node nên không chạy được JavaScript"


class CliCase(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.folder = Path(self._tmp.name) / "con-lac"
        self.folder.mkdir()
        (self.folder / "thi-nghiem.md").write_text(PENDULUM, encoding="utf-8")

    def tearDown(self):
        self._tmp.cleanup()

    def run_tool(self, *argv):
        out = io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(io.StringIO()):
            code = thi_nghiem.main([str(arg) for arg in argv])
        lines = out.getvalue().splitlines()
        self.assertEqual(len(lines), 1, out.getvalue())
        return code, json.loads(lines[0])

    def new_model(self, js=HOOKE_JS, kb=None):
        (self.folder / "thi-nghiem.md").write_text(HOOKE_MD, encoding="utf-8")
        (self.folder / "mo-hinh.json").write_text(json.dumps(kb or HOOKE_JSON, ensure_ascii=False), encoding="utf-8")
        (self.folder / "mo-hinh.js").write_text(js, encoding="utf-8")


class CliTest(CliCase):
    def test_builds_page_worksheet_and_review_file(self):
        code, data = self.run_tool(self.folder)
        self.assertEqual(code, 0, data)
        self.assertTrue(data["ready"])
        self.assertEqual([Path(path).name for path in data["files"]], ["thi-nghiem.html", "phieu-hoc-tap.docx", "can-soat.md"])
        self.assertEqual((data["mau"], data["tham_so"], data["so_lan_do"]), ("li-con-lac-don", ["chieu-dai"], 5))
        review = (self.folder / "can-soat.md").read_text(encoding="utf-8")
        self.assertIn("T = 2π√(l/g)", review)
        self.assertNotIn("do AI viết", review)

    def test_plan_only_and_part_selection(self):
        code, data = self.run_tool(self.folder, "--plan-only")
        self.assertEqual((code, data["files"]), (0, []))
        self.assertFalse((self.folder / "thi-nghiem.html").exists())
        code, data = self.run_tool(self.folder, "--phan", "html")
        self.assertEqual([Path(path).name for path in data["files"]], ["thi-nghiem.html", "can-soat.md"])
        code, data = self.run_tool(self.folder, "--phan", "word")
        self.assertEqual((code, data["error"]["step"]), (1, "input"))

    def test_input_errors(self):
        code, data = self.run_tool(self.folder / "khong-co")
        self.assertEqual((code, data["error"]["step"]), (1, "input"))
        (self.folder / "thi-nghiem.md").unlink()
        code, data = self.run_tool(self.folder)
        self.assertEqual((code, data["error"]["step"]), (1, "input"))
        code, data = self.run_tool()
        self.assertEqual((code, data["error"]["step"]), (1, "input"))

    def test_parse_error_names_the_line(self):
        (self.folder / "thi-nghiem.md").write_text(PENDULUM.replace("so-lan-do: 5", "so-lan-do: 99"), encoding="utf-8")
        code, data = self.run_tool(self.folder)
        self.assertEqual((code, data["error"]["step"]), (1, "parse"))
        self.assertIn("Dòng 22", data["error"]["message"])

    def test_unknown_model_is_a_model_error(self):
        (self.folder / "thi-nghiem.md").write_text(PENDULUM.replace("mau: li-con-lac-don", "mau: li-con-lac"), encoding="utf-8")
        code, data = self.run_tool(self.folder)
        self.assertEqual((code, data["error"]["step"]), (1, "model"))

    def test_new_model_is_flagged_for_the_teacher(self):
        self.new_model()
        code, data = self.run_tool(self.folder)
        self.assertEqual(code, 0, data)
        self.assertTrue(any("do AI viết" in warning for warning in data["warnings"]))
        review = (self.folder / "can-soat.md").read_text(encoding="utf-8")
        for needle in ("chưa có người duyệt", "F = k·x", "Bảng số kiểm do AI viết", "máy tính cầm tay"):
            self.assertIn(needle, review)

    def test_new_model_without_conditions_or_with_network_code_is_blocked(self):
        kb = json.loads(json.dumps(HOOKE_JSON))
        kb["congThuc"]["dieuKien"] = ""
        self.new_model(kb=kb)
        code, data = self.run_tool(self.folder)
        self.assertEqual((code, data["error"]["step"]), (1, "model"))
        self.assertIn("dieuKien", data["error"]["message"])
        self.new_model(js=HOOKE_JS + "\nfetch('x');")
        code, data = self.run_tool(self.folder)
        self.assertEqual((code, data["error"]["step"]), (1, "model"))

    @unittest.skipUnless(HAS_NODE, NEED_NODE)
    def test_wrong_physics_is_a_check_error(self):
        self.new_model(js=HOOKE_JS.replace("p['do-cung'] * p['do-gian']", "p['do-cung'] * p['do-gian'] * 2"))
        code, data = self.run_tool(self.folder)
        self.assertEqual((code, data["error"]["step"]), (1, "check"))
        self.assertIn("dòng 1 `luc`", data["error"]["message"])
        self.assertFalse((self.folder / "thi-nghiem.html").exists())

    def test_machine_without_node_still_builds_and_says_so(self):
        with mock.patch.object(kiem_so, "find_node", return_value=None):
            code, data = self.run_tool(self.folder, "--phan", "html")
        self.assertEqual(code, 0, data)
        self.assertEqual(data["kiem_so"], {"chay": False, "dat": 0, "tong": 6})
        self.assertTrue(any("không có Node" in warning for warning in data["warnings"]))
        self.assertIn("chưa chạy trên máy này", (self.folder / "can-soat.md").read_text(encoding="utf-8"))

    def test_missing_python_docx_is_a_docx_error(self):
        error = ImportError("No module named 'docx'", name="docx")
        with mock.patch.object(thi_nghiem, "load_phieu", side_effect=error):
            code, data = self.run_tool(self.folder)
        self.assertEqual((code, data["error"]["step"]), (1, "docx"))
        self.assertIn("requirements-vi.txt", data["error"]["fix"])

    def test_write_and_internal_errors(self):
        with mock.patch.object(build_html, "write", side_effect=PermissionError("đang mở")):
            code, data = self.run_tool(self.folder)
        self.assertEqual((code, data["error"]["step"]), (1, "write"))
        with mock.patch.object(thu_vien, "load", side_effect=RuntimeError("boom")):
            code, data = self.run_tool(self.folder)
        self.assertEqual((code, data["error"]["step"]), (1, "internal"))

if __name__ == "__main__":
    unittest.main()
