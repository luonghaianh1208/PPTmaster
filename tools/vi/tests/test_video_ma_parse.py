"""Test bộ đọc và bộ kiểm giới hạn của video.md (không cần Chromium, FFmpeg, mạng)."""

import sys
import tempfile
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))

from video_ma_parts import kiem, parse  # noqa: E402

MAU = TOOLS_VI / "fixtures" / "video-mau" / "video.md"
META = "tieu-de: T\nmon: Toán\nlop: 8\n"


def doc(canh: str, meta: str = META) -> str:
    return f"---\n{meta}---\n\n{canh}"


def line_of(text: str, needle: str) -> int:
    for number, line in enumerate(text.splitlines(), 1):
        if needle in line:
            return number
    raise AssertionError(needle)


class FixtureTest(unittest.TestCase):
    def test_sample_parses_all_eight_scene_types_in_order(self):
        video = parse.parse(MAU.read_text(encoding="utf-8"))
        self.assertEqual([c.loai for c in video.canh], list(parse.SCENE_TYPES[:8]))
        self.assertEqual([c.so for c in video.canh], list(range(1, 9)))
        self.assertEqual(video.meta["giong"], "nu")
        self.assertEqual(video.meta["toc-do"], "vua")
        self.assertEqual(video.meta["phu-de"], "hinh")
        self.assertEqual(video.meta["phong-cach"], "viet-tay")

    def test_sample_passes_the_limits_without_warnings(self):
        video = parse.parse(MAU.read_text(encoding="utf-8"))
        self.assertEqual(kiem.kiem(video, Path(".")), [])

    def test_repeated_fields_keep_order_and_line_numbers(self):
        text = MAU.read_text(encoding="utf-8")
        video = parse.parse(text)
        y = video.canh[3]
        self.assertEqual(len(y.truong["y"]), 3)
        self.assertEqual(y.dong_truong["y"][0], line_of(text, "Chiều dài dây l: dây dài"))


class ParseErrorTest(unittest.TestCase):
    def assert_error(self, text: str, line: int, fragment: str):
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse(text)
        self.assertEqual(caught.exception.line_no, line, str(caught.exception))
        self.assertIn(fragment, str(caught.exception))

    def test_missing_front_matter(self):
        self.assert_error("## Cảnh 1\nloai: tieu-de\n", 1, "---")

    def test_missing_required_meta_points_at_the_closing_line(self):
        self.assert_error("---\ntieu-de: T\nmon: Toán\n---\n", 4, "lop")

    def test_unknown_meta_key(self):
        text = doc("## Cảnh 1\n", META + "mau-nen: do\n")
        self.assert_error(text, 5, "mau-nen")

    def test_bad_meta_value(self):
        text = doc("## Cảnh 1\n", META + "giong: tre-em\n")
        self.assert_error(text, 5, "giong")

    def test_no_scene(self):
        self.assert_error(doc(""), 6, "Cảnh 1")

    def test_scene_numbers_must_be_sequential(self):
        text = doc("## Cảnh 2\nloai: tieu-de\nchu: A\nloi: Xin chào.\n")
        self.assert_error(text, 7, "Cảnh 1")

    def test_missing_loi(self):
        text = doc("## Cảnh 1\nloai: tieu-de\nchu: A\n")
        self.assert_error(text, 7, "loi")

    def test_unknown_scene_type(self):
        text = doc("## Cảnh 1\nloai: hoat-hinh\nloi: Xin chào.\n")
        self.assert_error(text, 8, "loai")

    def test_unknown_field_for_the_type(self):
        text = doc("## Cảnh 1\nloai: tieu-de\nchu: A\nbuoc: B\nloi: Xin chào.\n")
        self.assert_error(text, 10, "buoc")

    def test_duplicate_single_field(self):
        text = doc("## Cảnh 1\nloai: tieu-de\nchu: A\nchu: B\nloi: Xin chào.\n")
        self.assert_error(text, 10, "chu")

    def test_too_many_repeated_items(self):
        ys = "".join(f"y: ý {i}\n" for i in range(7))
        text = doc(f"## Cảnh 1\nloai: y-tung-y\ntieu-de: A\n{ys}loi: Xin chào.\n")
        self.assert_error(text, line_of(text, "ý 6"), "tối đa 6")

    def test_too_few_repeated_items(self):
        text = doc("## Cảnh 1\nloai: quy-trinh\ntieu-de: A\nbuoc: B\nloi: Xin chào.\n")
        self.assert_error(text, 7, "buoc")

    def test_point_needs_two_numbers_with_dot_decimals(self):
        text = doc("## Cảnh 1\nloai: do-thi\ntieu-de: A\ntruc-ngang: x\ntruc-doc: y\ndiem: 1;2\ndiem: 2, 3\nloi: Xin chào.\n")
        self.assert_error(text, line_of(text, "1;2"), "x, y")

    def test_parameter_line_needs_three_parts(self):
        text = doc("## Cảnh 1\nloai: thi-nghiem\nmau: li-con-lac-don\ntham-so: 0 chieu-dai\nloi: Xin chào.\n")
        self.assert_error(text, line_of(text, "tham-so"), "<giây> <mã> <giá trị>")

    def test_web_addresses_are_refused(self):
        text = doc("## Cảnh 1\nloai: tieu-de\nchu: xem https://a.vn\nloi: Xin chào.\n")
        self.assert_error(text, line_of(text, "https"), "địa chỉ web")

    def test_empty_value(self):
        text = doc("## Cảnh 1\nloai: tieu-de\nchu:\nloi: Xin chào.\n")
        self.assert_error(text, line_of(text, "chu:"), "trống")

    def test_line_that_is_not_a_field(self):
        text = doc("## Cảnh 1\nloai: tieu-de\nĐây là một dòng lạ\nloi: Xin chào.\n")
        self.assert_error(text, line_of(text, "dòng lạ"), "khoá: giá trị")


class LimitTest(unittest.TestCase):
    def video(self, canh: str):
        return parse.parse(doc(canh))

    def test_title_over_limit_names_scene_and_line(self):
        text = doc(f"## Cảnh 1\nloai: tieu-de\nchu: {'a' * 91}\nloi: Xin chào.\n")
        with self.assertRaises(kiem.CanhError) as caught:
            kiem.kiem(parse.parse(text), Path("."))
        self.assertEqual(caught.exception.so, 1)
        self.assertIn("91", str(caught.exception))
        self.assertIn(f"dòng {line_of(text, 'aaa')}", str(caught.exception))

    def test_markup_characters_are_not_counted(self):
        video = self.video(f"## Cảnh 1\nloai: tieu-de\nchu: **{'a' * 90}**\nloi: Xin chào.\n")
        self.assertEqual(kiem.kiem(video, Path(".")), [])
        self.assertEqual(kiem.hien_thi("H~2~SO~4~ m/s^2^"), len("H2SO4 m/s2"))

    def test_long_narration_is_a_warning_not_an_error(self):
        video = self.video(f"## Cảnh 1\nloai: tieu-de\nchu: A\nloi: {'Câu này dài. ' * 60}\n")
        warnings = kiem.kiem(video, Path("."))
        self.assertEqual(len(warnings), 1)
        self.assertIn("Cảnh 1", warnings[0])
        self.assertIn("700", warnings[0])


class ExperimentSceneTest(unittest.TestCase):
    def check(self, extra: str):
        text = doc(f"## Cảnh 1\nloai: thi-nghiem\nmau: li-con-lac-don\n{extra}loi: Xin chào.\n")
        return text, lambda: kiem.kiem(parse.parse(text), Path("."))

    def test_new_model_is_not_allowed_in_video(self):
        text = doc("## Cảnh 1\nloai: thi-nghiem\nmau: moi\nloi: Xin chào.\n")
        with self.assertRaises(kiem.CanhError) as caught:
            kiem.kiem(parse.parse(text), Path("."))
        self.assertIn("thư viện", str(caught.exception))

    def test_unknown_model_lists_the_library(self):
        for mau in ("li-khong-co", "../../x"):
            with self.subTest(mau=mau):
                text = doc(f"## Cảnh 1\nloai: thi-nghiem\nmau: {mau}\nloi: Xin chào.\n")
                with self.assertRaises(kiem.CanhError) as caught:
                    kiem.kiem(parse.parse(text), Path("."))
                self.assertIn("li-con-lac-don", str(caught.exception))
                self.assertNotIn("moi", str(caught.exception))

    def test_unknown_parameter(self):
        _, run = self.check("tham-so: 0 toc-do 3\n")
        with self.assertRaises(kiem.CanhError) as caught:
            run()
        self.assertIn("toc-do", str(caught.exception))

    def test_value_outside_the_allowed_range_names_the_line(self):
        text, run = self.check("tham-so: 0 chieu-dai 5\n")
        with self.assertRaises(parse.ParseError) as caught:
            run()
        self.assertEqual(caught.exception.line_no, line_of(text, "tham-so"))
        self.assertIn("0.2", str(caught.exception))

    def test_more_than_three_parameters(self):
        _, run = self.check("tham-so: 0 chieu-dai 1\ntham-so: 0 g 9.8\ntham-so: 0 goc-lech 8\ntham-so: 0 khoi-luong 0.2\n")
        with self.assertRaises(kiem.CanhError):
            run()

    def test_unknown_measured_quantity(self):
        _, run = self.check("do: van-toc\n")
        with self.assertRaises(kiem.CanhError) as caught:
            run()
        self.assertIn("chu-ki", str(caught.exception))

    def test_schedule_is_sorted_and_default_measures_are_the_first_two(self):
        from thi_nghiem_parts import thu_vien

        text = doc("## Cảnh 1\nloai: thi-nghiem\nmau: li-con-lac-don\ntham-so: 6 chieu-dai 1.6\ntham-so: 0 chieu-dai 0.4\nloi: Xin chào.\n")
        scene = parse.parse(text).canh[0]
        self.assertEqual(kiem.tham_so_theo_thoi_gian(scene), {"chieu-dai": [(0.0, 0.4), (6.0, 1.6)]})
        model = thu_vien.load("li-con-lac-don", Path("."))
        self.assertEqual(kiem.ma_do(scene, model), ["chu-ki", "thoi-gian-10-dao-dong"])


class NewFieldsTest(unittest.TestCase):
    def test_meta_choices_have_new_keys_with_defaults(self):
        for key, choices, default in (
            ("ban-tay", ("co", "khong"), "co"),
            ("may-quay", ("co", "khong"), "co"),
            ("chuyen-canh", ("lau-bang", "khong"), "lau-bang"),
        ):
            self.assertEqual(parse.META_CHOICES[key], choices)
            self.assertEqual(parse.META_DEFAULTS[key], default)

    def test_new_meta_keys_default_when_video_omits_them(self):
        video = parse.parse(doc("## Cảnh 1\nloai: tieu-de\nchu: A\nloi: Xin chào.\n"))
        self.assertEqual(video.meta["ban-tay"], "co")
        self.assertEqual(video.meta["may-quay"], "co")
        self.assertEqual(video.meta["chuyen-canh"], "lau-bang")

    def test_new_scene_types_are_registered(self):
        self.assertIn("minh-hoa", parse.SCENE_SPEC)
        self.assertIn("anh", parse.SCENE_SPEC)
        self.assertEqual(parse.SCENE_SPEC["minh-hoa"], (("tieu-de",), (), {"hinh": (1, 3)}))
        self.assertEqual(parse.SCENE_SPEC["anh"], (("anh", "chu-thich"), ("nguon",), {}))

    def test_hinh_and_anh_are_optional_fields_of_the_four_text_scenes(self):
        for loai in ("tieu-de", "khai-niem", "cong-thuc", "y-tung-y"):
            _, optional, _ = parse.SCENE_SPEC[loai]
            self.assertIn("hinh", optional, loai)
            self.assertIn("anh", optional, loai)

    def test_scene_with_both_hinh_and_anh_is_a_parse_error_at_the_second_line(self):
        text = doc("## Cảnh 1\nloai: tieu-de\nchu: A\nhinh: flask\nanh: x.png\nloi: Xin chào.\n")
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse(text)
        self.assertEqual(caught.exception.line_no, line_of(text, "anh: x.png"))

    def test_scene_with_anh_then_hinh_points_at_the_hinh_line(self):
        text = doc("## Cảnh 1\nloai: tieu-de\nchu: A\nanh: x.png\nhinh: flask\nloi: Xin chào.\n")
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse(text)
        self.assertEqual(caught.exception.line_no, line_of(text, "hinh: flask"))

    def test_minh_hoa_needs_at_least_one_picture(self):
        text = doc("## Cảnh 1\nloai: minh-hoa\ntieu-de: A\nloi: Xin chào.\n")
        with self.assertRaises(parse.ParseError):
            parse.parse(text)

    def test_minh_hoa_allows_at_most_three_pictures(self):
        hinhs = "".join(f"hinh: clock | Nhãn {i}\n" for i in range(4))
        text = doc(f"## Cảnh 1\nloai: minh-hoa\ntieu-de: A\n{hinhs}loi: Xin chào.\n")
        with self.assertRaises(parse.ParseError):
            parse.parse(text)


if __name__ == "__main__":
    unittest.main()
