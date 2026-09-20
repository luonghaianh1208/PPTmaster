"""Test cho thư viện mô hình thí nghiệm ảo: khuôn, mã, và đối chiếu JavaScript với bản Python."""

import json
import random
import sys
import tempfile
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from thi_nghiem_mau import HOOKE_JS, HOOKE_JSON, grid  # noqa: E402
from thi_nghiem_parts import kiem_so, tham_chieu, thu_vien  # noqa: E402

HAS_NODE = kiem_so.find_node() is not None
NEED_NODE = "máy không có Node nên không chạy được JavaScript"
MODELS = ("hoa-can-bang-no2", "hoa-chuan-do", "hoa-toc-do", "li-con-lac-don", "li-mach-ohm", "li-nem-xien")


class ModelLibraryTest(unittest.TestCase):
    def test_library_has_the_models(self):
        self.assertEqual(tuple(thu_vien.list_models()), MODELS)

    def test_every_library_model_loads_and_follows_the_contract(self):
        for ma in MODELS:
            with self.subTest(model=ma):
                model = thu_vien.load(ma, Path("."))
                self.assertFalse(model.moi)
                self.assertEqual(model.khai_bao["ma"], ma)
                self.assertGreaterEqual(len(model.khai_bao["bangKiem"]), thu_vien.MIN_CHECK_ROWS)
                self.assertIn(ma, tham_chieu.THAM_CHIEU)

    def test_unknown_model_lists_what_exists(self):
        with self.assertRaises(thu_vien.ModelError) as caught:
            thu_vien.load("li-khong-co", Path("."))
        self.assertIn("li-con-lac-don", str(caught.exception))
        self.assertIn("moi", str(caught.exception))

    def test_declaration_errors_are_named(self):
        cases = {
            "thiếu `congThuc.dieuKien`": lambda kb: kb["congThuc"].update(dieuKien=" "),
            "`bangKiem` phải có ít nhất 5 dòng": lambda kb: kb.update(bangKiem=kb["bangKiem"][:4]),
            "`hoatHinh` phải là một trong": lambda kb: kb.update(hoatHinh="nhanh"),
            "cần min < max": lambda kb: kb["thamSo"][0].update(min=200),
            "`vao` có tham số lạ `la`": lambda kb: kb["bangKiem"][0]["vao"].update(la=1),
            "`ra` có đại lượng lạ `la`": lambda kb: kb["bangKiem"][0]["ra"].update(la=1),
            "`saiSo` phải là số không âm": lambda kb: kb["daiLuongDo"][0].update(saiSo=-1),
            "`daiLuongDo` phải có ít nhất một": lambda kb: kb.update(daiLuongDo=[]),
        }
        for message, damage in cases.items():
            with self.subTest(message=message):
                kb = json.loads(json.dumps(HOOKE_JSON))
                damage(kb)
                self.assertTrue(any(message in error for error in thu_vien.check_declaration(kb)),
                                thu_vien.check_declaration(kb))
        self.assertEqual(thu_vien.check_declaration(HOOKE_JSON), [])

    def test_model_code_may_not_reach_the_network_or_the_page(self):
        for banned in ("fetch('x')", "https://cdn", "document.title", "eval('1')", "require('fs')", "</script>"):
            with self.subTest(banned=banned):
                self.assertTrue(thu_vien.check_js(HOOKE_JS + "// " + banned, "khong"))
        self.assertEqual(thu_vien.check_js(HOOKE_JS, "khong"), [])
        self.assertTrue(thu_vien.check_js(HOOKE_JS, "mot-lan"))

    def test_library_code_passes_its_own_rules(self):
        for ma in MODELS:
            model = thu_vien.load(ma, Path("."))
            self.assertEqual(thu_vien.check_js(model.js, model.khai_bao["hoatHinh"]), [], ma)

    def test_new_model_needs_both_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            (Path(tmp) / "mo-hinh.json").write_text(json.dumps(HOOKE_JSON), encoding="utf-8")
            with self.assertRaises(thu_vien.ModelError) as caught:
                thu_vien.load("moi", Path(tmp))
            self.assertIn("mo-hinh.js", str(caught.exception))


@unittest.skipUnless(HAS_NODE, NEED_NODE)
class JavaScriptAgreementTest(unittest.TestCase):
    def test_check_tables_pass(self):
        for ma in MODELS:
            with self.subTest(model=ma):
                check = kiem_so.bang_kiem(thu_vien.load(ma, Path(".")))
                self.assertEqual((check["dat"], check["truot"]), (check["tong"], []))

    def test_javascript_matches_the_python_reference_on_a_grid(self):
        rng = random.Random(20260920)
        for ma in MODELS:
            with self.subTest(model=ma):
                model = thu_vien.load(ma, Path("."))
                points = grid(model, 40 if ma == "toan-xac-suat" else 300, rng)
                for point, got in zip(points, kiem_so.luoi(model, points)):
                    expected = tham_chieu.THAM_CHIEU[ma](point)
                    for measure in model.khai_bao["daiLuongDo"]:
                        a, b = expected[measure["ma"]], got[measure["ma"]]
                        if a is None or b is None:
                            self.assertIsNone(a, point)
                            self.assertIsNone(b, point)
                        else:
                            self.assertLessEqual(abs(a - b), 1e-9 * max(1.0, abs(a)), (point, measure["ma"]))

    def test_a_wrong_model_fails_its_check_table(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            (folder / "mo-hinh.json").write_text(json.dumps(HOOKE_JSON), encoding="utf-8")
            (folder / "mo-hinh.js").write_text(HOOKE_JS.replace("p['do-cung'] * p['do-gian']", "p['do-cung'] + p['do-gian']"),
                                                encoding="utf-8")
            check = kiem_so.bang_kiem(thu_vien.load("moi", folder))
            self.assertLess(check["dat"], check["tong"])

    def test_a_crashing_model_is_a_check_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            (folder / "mo-hinh.json").write_text(json.dumps(HOOKE_JSON), encoding="utf-8")
            (folder / "mo-hinh.js").write_text(HOOKE_JS + "\nthis is not javascript(", encoding="utf-8")
            with self.assertRaises(kiem_so.CheckError):
                kiem_so.bang_kiem(thu_vien.load("moi", folder))

if __name__ == "__main__":
    unittest.main()
