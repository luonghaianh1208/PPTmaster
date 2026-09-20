"""Test cho bộ đọc thi-nghiem.md."""

import sys
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from thi_nghiem_mau import PENDULUM  # noqa: E402
from thi_nghiem_parts import parse, thu_vien  # noqa: E402

class ParseTest(unittest.TestCase):
    def setUp(self):
        self.kb = thu_vien.load("li-con-lac-don", Path(".")).khai_bao

    def parse(self, text):
        return parse.parse_experiment(text, self.kb)

    def assertLine(self, text, line_no, fragment):
        with self.assertRaises(parse.ParseError) as caught:
            self.parse(text)
        self.assertEqual(caught.exception.line_no, line_no, str(caught.exception))
        self.assertIn(fragment, caught.exception.message)

    def test_valid_file_becomes_the_page_config(self):
        experiment = self.parse(PENDULUM)
        self.assertEqual(experiment.sliders(), ["chieu-dai"])
        config = experiment.config()
        self.assertEqual(config["thamSo"]["chieu-dai"], {"kieu": "truot", "min": 0.4, "max": 1.6, "buoc": 0.2, "macDinh": 1.0})
        self.assertEqual(config["thamSo"]["g"], {"kieu": "co-dinh", "giaTri": 9.8})
        self.assertEqual(config["thamSo"]["khoi-luong"], {"kieu": "co-dinh", "giaTri": 0.2})
        self.assertEqual(list(config["thamSo"]), ["chieu-dai", "g", "goc-lech", "khoi-luong"])
        self.assertEqual(config["duDoan"]["dapAn"], "B")
        self.assertEqual(config["quanSat"], {"soLanDo": 5, "cot": ["chieu-dai", "chu-ki"], "doThi": {
            "tung": {"ma": "chu-ki", "phep": "binh-phuong"}, "hoanh": {"ma": "chieu-dai", "phep": "khong"}}})
        self.assertTrue(config["saiSo"])
        self.assertEqual(config["nguoiThaoTac"], "nhom")

    def test_defaults_are_teacher_mode_and_no_noise(self):
        text = PENDULUM.replace("nguoi-thao-tac: nhom\n", "").replace("sai-so: bat\n", "")
        config = self.parse(text).config()
        self.assertEqual((config["nguoiThaoTac"], config["saiSo"]), ("giao-vien", False))

    def test_open_prediction_has_no_key(self):
        text = PENDULUM.replace("A: Tăng 4 lần\nB: Tăng 2 lần\nC: Không đổi\ndap-an: B\n", "")
        config = self.parse(text).config()
        self.assertEqual((config["duDoan"]["luaChon"], config["duDoan"]["dapAn"]), ([], None))

    def test_comma_decimals_are_accepted(self):
        text = PENDULUM.replace("0.4..1.6 buoc 0.2 mac-dinh 1.0", "0,4..1,6 buoc 0,2 mac-dinh 1,0")
        self.assertEqual(self.parse(text).tham_so["chieu-dai"]["min"], 0.4)

    def test_errors_name_the_line(self):
        self.assertLine(PENDULUM.replace("chieu-dai: 0.4..1.6", "chieu-dai: 0.1..1.6"), 11, "công thức của mẫu không còn đúng")
        self.assertLine(PENDULUM.replace("g: co-dinh 9.8", "g: co-dinh 50"), 12, "1.6..24.8")
        self.assertLine(PENDULUM.replace("g: co-dinh 9.8", "luc-can: co-dinh 1"), 12, "mẫu không có tham số `luc-can`")
        self.assertLine(PENDULUM.replace("buoc 0.2", "buoc 0"), 11, "`buoc`")
        self.assertLine(PENDULUM.replace("mac-dinh 1.0", "mac-dinh 3"), 11, "`mac-dinh`")
        self.assertLine(PENDULUM.replace("chieu-dai: 0.4..1.6 buoc 0.2 mac-dinh 1.0", "chieu-dai: tu 0.4 den 1.6"), 11, "cần dạng")
        self.assertLine(PENDULUM.replace("dap-an: B", "dap-an: D"), 19, "`dap-an`")
        self.assertLine(PENDULUM.replace("so-lan-do: 5", "so-lan-do: 2"), 22, "từ 3 đến 20")
        self.assertLine(PENDULUM.replace("do: chieu-dai, chu-ki", "do: chieu-dai, tan-so"), 23, "mẫu không có `tan-so`")
        self.assertLine(PENDULUM.replace("do: chieu-dai, chu-ki", "do: chieu-dai"), 23, "ít nhất một đại lượng đo")
        self.assertLine(PENDULUM.replace("chu-ki^2 theo chieu-dai", "chu-ki^3 theo chieu-dai"), 24, "không hiểu biểu thức")
        self.assertLine(PENDULUM.replace("chu-ki^2 theo chieu-dai", "chu-ki^2 theo g"), 24, "dòng `do:` không có")
        self.assertLine(PENDULUM.replace("chu-ki^2 theo chieu-dai", "chu-ki^2"), 24, "theo")
        self.assertLine(PENDULUM.replace("cau: Từ đồ thị", "xem https://example.com Từ đồ thị"), 27, "địa chỉ web")

    def test_structure_errors(self):
        for removed in ("## Giải thích\n", "## Kết luận\nChu kì con lắc đơn chỉ phụ thuộc chiều dài dây và g.\n"):
            text = PENDULUM.replace(removed, "")
            self.assertLine(text, len(text.splitlines()), "thiếu mục `" + removed.splitlines()[0] + "`")
        self.assertLine(PENDULUM.replace("## Quan sát", "## Đo đạc"), 21, "mục lạ")
        self.assertLine(PENDULUM.replace("goi-y-dap-an: T^2^ tỉ lệ thuận với l; T = 2π√(l/g).\n", ""), 26, "`goi-y-dap-an:`")
        self.assertLine(PENDULUM.replace("chieu-dai: 0.4..1.6 buoc 0.2 mac-dinh 1.0\n", ""), 10, "ít nhất một tham số thay đổi được")
        self.assertLine("## Tham số\n", 1, "`---`")

    def test_table_without_a_changing_parameter_warns(self):
        text = PENDULUM.replace("do: chieu-dai, chu-ki\ndo-thi: chu-ki^2 theo chieu-dai", "do: chu-ki")
        self.assertIn("không có tham số nào thay đổi được", self.parse(text).warnings[0])

    def test_meta_errors(self):
        for text, fragment in (
            (PENDULUM.replace("mau: li-con-lac-don\n", ""), "thiếu `mau`"),
            (PENDULUM.replace("sai-so: bat", "sai-so: co"), "`sai-so`"),
            (PENDULUM.replace("nguoi-thao-tac: nhom", "nguoi-thao-tac: hoc-sinh"), "`nguoi-thao-tac`"),
            (PENDULUM.replace("lop: 11", "khoi: 11"), "khoá lạ `khoi`"),
        ):
            with self.subTest(fragment=fragment):
                with self.assertRaises(parse.ParseError) as caught:
                    parse.read_meta(text)
                self.assertIn(fragment, caught.exception.message)

    def test_choice_parameters(self):
        kb = thu_vien.load("li-mach-ohm", Path(".")).khai_bao
        text = (PENDULUM.replace("mau: li-con-lac-don", "mau: li-mach-ohm")
                .replace("chieu-dai: 0.4..1.6 buoc 0.2 mac-dinh 1.0\ng: co-dinh 9.8", "kieu-mac: chon song-song, noi-tiep\ndien-tro-1: co-dinh 30")
                .replace("do: chieu-dai, chu-ki\ndo-thi: chu-ki^2 theo chieu-dai", "do: kieu-mac, cuong-do-mach-chinh"))
        experiment = parse.parse_experiment(text, kb)
        self.assertEqual(experiment.tham_so["kieu-mac"], {"kieu": "chon", "luaChon": ["song-song", "noi-tiep"], "macDinh": "song-song"})
        self.assertEqual(experiment.warnings, [])
        for bad, fragment in (("kieu-mac: chon song-song", "ít nhất hai lựa chọn"), ("kieu-mac: 1..2 buoc 1 mac-dinh 1", "tham số lựa chọn"),
                              ("kieu-mac: co-dinh cheo", "chỉ nhận"), ("dien-tro-2: chon 1, 2", "là tham số số")):
            with self.subTest(bad=bad):
                with self.assertRaises(parse.ParseError) as caught:
                    parse.parse_experiment(text.replace("kieu-mac: chon song-song, noi-tiep", bad), kb)
                self.assertIn(fragment, caught.exception.message)

if __name__ == "__main__":
    unittest.main()
