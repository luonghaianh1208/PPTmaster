"""Test lịch thời gian của video giải thích: câu, mốc hiện ý, thời lượng cảnh, dữ liệu cảnh."""

import sys
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))

from thi_nghiem_parts import thu_vien  # noqa: E402
from video_ma_parts import kiem, lich, parse  # noqa: E402

META = "tieu-de: T\nmon: Toán\nlop: 8\n"


def canh_dau(noi_dung: str, loi: str):
    text = f"---\n{META}---\n\n## Cảnh 1\n{noi_dung}loi: {loi}\n"
    return parse.parse(text).canh[0]


def giong(giay: float, moc=None, uoc=False):
    return lich.GiongInfo(mp3=None, giay=giay, moc_cau=moc or [], uoc_luong=uoc, nguon="may")


class SentenceTest(unittest.TestCase):
    def test_split_on_terminators_and_keep_ellipsis(self):
        self.assertEqual(lich.tach_cau("Một. Hai! Ba? Bốn… Năm."), ["Một.", "Hai!", "Ba?", "Bốn…", "Năm."])

    def test_no_terminator_is_one_sentence(self):
        self.assertEqual(lich.tach_cau("Không có dấu chấm cuối"), ["Không có dấu chấm cuối"])

    def test_estimated_marks_follow_character_ratio(self):
        marks = lich.moc_uoc_luong(["aaaa", "bb", "cccccc"], 12.0)
        self.assertEqual(marks, [0.0, 4.0, 6.0])


class DurationTest(unittest.TestCase):
    def test_minimum_and_frame_multiple(self):
        toi_thieu = lich.thoi_luong_canh(0.1)
        self.assertGreaterEqual(toi_thieu, 2.5)
        self.assertLess(toi_thieu, 2.5 + 1 / lich.FPS)
        long = lich.thoi_luong_canh(4.13)
        self.assertGreaterEqual(long, lich.DAN_DAU + 4.13 + 0.6)
        self.assertAlmostEqual(long * lich.FPS, round(long * lich.FPS), places=6)
        self.assertLess(long, lich.DAN_DAU + 4.13 + 0.6 + 1 / lich.FPS)


class RevealTimesTest(unittest.TestCase):
    def test_item_k_appears_at_sentence_k(self):
        self.assertEqual(lich.moc_hien(3, [0.0, 2.0, 5.0, 7.0], 9.0), [0.0, 2.0, 5.0])

    def test_fewer_sentences_than_items_spreads_evenly(self):
        self.assertEqual(lich.moc_hien(4, [0.0], 8.0), [0.0, 2.0, 4.0, 6.0])

    def test_loi_mot_cau_nhieu_y(self):
        scene = canh_dau("loai: y-tung-y\ntieu-de: A\ny: a\ny: b\ny: c\ny: d\ny: e\ny: g\n",
                         "Một câu duy nhất không có dấu chấm cuối")
        plan, _ = lich.dung_lich([scene], [giong(12.0)])
        du = lich.du_lieu_canh(scene, plan[0])
        self.assertEqual(len(du["moc"]), 6)
        gaps = [b - a for a, b in zip(du["moc"], du["moc"][1:])]
        self.assertTrue(all(abs(g - 2.0) < 1e-6 for g in gaps), gaps)
        self.assertAlmostEqual(du["moc"][0], lich.DAN_DAU)

    def test_items_per_scene_type(self):
        text = f"---\n{META}---\n\n## Cảnh 1\nloai: so-sanh\ntieu-de: A\ntrai: T\nphai: P\ny-trai: a\ny-trai: b\ny-phai: c\nloi: Ok.\n"
        self.assertEqual(lich.so_muc(parse.parse(text).canh[0]), 3)


class PlanTest(unittest.TestCase):
    def scenes(self):
        text = (f"---\n{META}---\n\n## Cảnh 1\nloai: tieu-de\nchu: A\nloi: Xin chào các em. Hôm nay học bài mới.\n\n"
                "## Cảnh 2\nloai: tieu-de\nchu: B\nloi: Tiếp theo.\n")
        return parse.parse(text).canh

    def test_offsets_accumulate_and_frames_match(self):
        plan, warnings = lich.dung_lich(self.scenes(), [giong(4.0, [0.0, 1.9]), giong(0.5, [0.0])])
        self.assertEqual(warnings, [])
        self.assertEqual(plan[0].bat_dau, 0.0)
        self.assertAlmostEqual(plan[1].bat_dau, plan[0].thoi_luong)
        self.assertEqual(plan[1].thoi_luong, lich.thoi_luong_canh(0.5))
        self.assertGreaterEqual(plan[1].thoi_luong, 2.5)
        for cl in plan:
            self.assertEqual(cl.so_khung, round(cl.thoi_luong * lich.FPS))
        self.assertEqual(plan[0].cau, ["Xin chào các em.", "Hôm nay học bài mới."])
        self.assertEqual(plan[0].moc_cau_giong, [0.0, 1.9])
        self.assertAlmostEqual(plan[0].moc_cau[1], 1.9 + lich.DAN_DAU)

    def test_missing_marks_are_estimated_with_a_warning(self):
        plan, warnings = lich.dung_lich(self.scenes(), [giong(4.0, [], True), giong(1.0, [0.0])])
        self.assertTrue(plan[0].uoc_luong)
        self.assertEqual(len(warnings), 1)
        self.assertIn("Cảnh 1", warnings[0])
        self.assertIn("ước lượng", warnings[0])

    def test_mark_count_mismatch_falls_back_to_estimate(self):
        plan, warnings = lich.dung_lich(self.scenes(), [giong(4.0, [0.0]), giong(1.0, [0.0])])
        self.assertTrue(plan[0].uoc_luong)
        self.assertEqual(len(warnings), 1)

    def test_long_scene_and_long_video_warn(self):
        plan, warnings = lich.dung_lich(self.scenes(), [giong(45.0, [0.0, 20.0]), giong(1.0, [0.0])])
        self.assertTrue(any("40 giây" in w for w in warnings))
        scenes = self.scenes() * 1
        _, warnings = lich.dung_lich(scenes, [giong(300.0, [0.0, 100.0]), giong(300.0, [0.0])])
        self.assertTrue(any("8 phút" in w for w in warnings))


class ExperimentDataTest(unittest.TestCase):
    def scene(self, extra: str):
        return canh_dau(f"loai: thi-nghiem\nmau: li-con-lac-don\n{extra}", "Hãy quan sát.")

    def test_mark_beyond_scene_duration_is_a_canh_error(self):
        scene = self.scene("tham-so: 30 chieu-dai 1.6\n")
        with self.assertRaises(kiem.CanhError) as caught:
            lich.dung_lich([scene], [giong(4.0, [0.0])])
        self.assertIn("30", str(caught.exception))
        self.assertIn("dòng", str(caught.exception))

    def test_preview_skips_the_mark_check(self):
        scene = self.scene("tham-so: 30 chieu-dai 1.6\n")
        plan, _ = lich.dung_lich([scene], [giong(4.0, [0.0])], kiem_moc=False)
        self.assertEqual(len(plan), 1)

    def test_scene_data_carries_model_schedule_and_measures(self):
        scene = self.scene("tham-so: 6 chieu-dai 1.6\ntham-so: 0 chieu-dai 0.4\ndo: chu-ki\n")
        plan, _ = lich.dung_lich([scene], [giong(8.0, [0.0])])
        model = thu_vien.load("li-con-lac-don", Path("."))
        du = lich.du_lieu_canh(scene, plan[0], model)
        self.assertEqual(du["loai"], "thi-nghiem")
        self.assertEqual(du["thamSo"], {"chieu-dai": [[0.0, 0.4], [6.0, 1.6]]})
        self.assertEqual(du["do"], ["chu-ki"])
        self.assertEqual(du["khaiBao"]["ma"], "li-con-lac-don")
        self.assertEqual(du["thoiLuong"], plan[0].thoi_luong)
        self.assertEqual(du["danDau"], lich.DAN_DAU)
        self.assertEqual(du["giayLauBang"], lich.LAU_BANG)

    def test_graph_scene_data_has_numeric_points(self):
        scene = canh_dau("loai: do-thi\ntieu-de: A\ntruc-ngang: x\ntruc-doc: y\ndiem: 0.5, 1\ndiem: 2, 3.5\n", "Ok.")
        plan, _ = lich.dung_lich([scene], [giong(3.0, [0.0])])
        du = lich.du_lieu_canh(scene, plan[0])
        self.assertEqual(du["diem"], [[0.5, 1.0], [2.0, 3.5]])
        self.assertEqual(len(du["moc"]), 2)


class ResourceDataTest(unittest.TestCase):
    def test_minh_hoa_reveal_marks_match_picture_count(self):
        scene = canh_dau("loai: minh-hoa\ntieu-de: A\nhinh: clock | Đồng hồ\nhinh: atom | Nguyên tử\n",
                         "Đầu tiên xem đồng hồ. Sau đó xem nguyên tử.")
        plan, _ = lich.dung_lich([scene], [giong(6.0, [0.0, 3.0])])
        du = lich.du_lieu_canh(scene, plan[0])
        self.assertEqual(len(du["moc"]), 2)

    def test_hinhs_keep_order_and_labels(self):
        scene = canh_dau("loai: minh-hoa\ntieu-de: A\nhinh: clock | Đồng hồ\nhinh: atom | Nguyên tử\n", "Ok.")
        plan, _ = lich.dung_lich([scene], [giong(3.0, [0.0])])
        tai_nguyen = {"hinh": None, "anh": None,
                      "hinhs": [{"ten": "clock", "nhan": "Đồng hồ"}, {"ten": "atom", "nhan": "Nguyên tử"}]}
        du = lich.du_lieu_canh(scene, plan[0], tai_nguyen=tai_nguyen)
        self.assertEqual([h["ten"] for h in du["hinhs"]], ["clock", "atom"])
        self.assertEqual([h["nhan"] for h in du["hinhs"]], ["Đồng hồ", "Nguyên tử"])

    def test_co_flags_for_scene_one_and_two(self):
        scene1 = canh_dau("loai: tieu-de\nchu: A\n", "Xin chào.")
        scene2 = canh_dau("loai: tieu-de\nchu: B\n", "Tiếp theo.")
        scene2.so = 2
        meta = {"ban-tay": "co", "may-quay": "khong", "chuyen-canh": "lau-bang"}
        plan1, _ = lich.dung_lich([scene1], [giong(2.0, [0.0])])
        du1 = lich.du_lieu_canh(scene1, plan1[0], tai_nguyen={"meta": meta})
        self.assertEqual(du1["co"], {"banTay": True, "mayQuay": False, "lauBang": False})
        plan2, _ = lich.dung_lich([scene2], [giong(2.0, [0.0])])
        du2 = lich.du_lieu_canh(scene2, plan2[0], tai_nguyen={"meta": meta})
        self.assertEqual(du2["co"], {"banTay": True, "mayQuay": False, "lauBang": True})

    def test_lau_bang_off_when_meta_says_khong(self):
        scene2 = canh_dau("loai: tieu-de\nchu: B\n", "Tiếp theo.")
        scene2.so = 2
        meta = {"ban-tay": "co", "may-quay": "co", "chuyen-canh": "khong"}
        plan2, _ = lich.dung_lich([scene2], [giong(2.0, [0.0])])
        du2 = lich.du_lieu_canh(scene2, plan2[0], tai_nguyen={"meta": meta})
        self.assertFalse(du2["co"]["lauBang"])


class WipeLengthTest(unittest.TestCase):
    def test_runtime_defaults_match_the_python_wipe_length(self):
        import re

        runtime = TOOLS_VI / "video_ma_parts" / "runtime"
        for ten in ("khung-video.js", "ban-tay.js"):
            with self.subTest(file=ten):
                so = re.search(r"var LAU_BANG = ([0-9.]+);", (runtime / ten).read_text(encoding="utf-8"))
                self.assertIsNotNone(so)
                self.assertEqual(float(so.group(1)), lich.LAU_BANG)


if __name__ == "__main__":
    unittest.main()
