"""Test công cụ dòng lệnh video_ma.py: mọi nhánh ra đúng một dòng JSON, không gọi giọng thật, Chromium/FFmpeg giả lập."""

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

import video_ma  # noqa: E402
from video_ma_parts import lich  # noqa: E402
from video_parts import media  # noqa: E402

MAU = (TOOLS_VI / "fixtures" / "video-mau" / "video.md").read_text(encoding="utf-8")
MOT_CANH = "---\ntieu-de: T\nmon: Toán\nlop: 8\n---\n\n## Cảnh 1\nloai: tieu-de\nchu: Xin chào\nloi: Xin chào các em.\n"


def chay(args):
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        code = video_ma.main(args)
    lines = [l for l in out.getvalue().splitlines() if l.strip()]
    return code, lines, err.getvalue()


class CliTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dir = Path(self.tmp.name) / "bai"
        self.dir.mkdir()

    def viet(self, text):
        (self.dir / "video.md").write_text(text, encoding="utf-8")

    def one_json(self, args):
        code, lines, _ = chay(args)
        self.assertEqual(len(lines), 1, lines)
        return code, json.loads(lines[0])

    def test_missing_folder_and_missing_file_are_input_errors(self):
        code, data = self.one_json([str(self.dir / "khong-co")])
        self.assertEqual((code, data["error"]["step"], data["ready"]), (1, "input", False))
        code, data = self.one_json([str(self.dir)])
        self.assertEqual((code, data["error"]["step"]), (1, "input"))
        self.assertIn("video-giai-thich.md", data["error"]["fix"])

    def test_parse_error_names_the_line(self):
        self.viet("---\ntieu-de: T\nmon: Toán\n---\n")
        code, data = self.one_json([str(self.dir)])
        self.assertEqual((code, data["error"]["step"]), (1, "parse"))
        self.assertIn("Dòng 4", data["error"]["message"])

    def test_scene_content_error_is_canh(self):
        self.viet(MOT_CANH.replace("chu: Xin chào", "chu: " + "a" * 95))
        code, data = self.one_json([str(self.dir)])
        self.assertEqual((code, data["error"]["step"]), (1, "canh"))
        self.assertIn("Cảnh 1", data["error"]["message"])

    def test_plan_only_checks_and_writes_nothing(self):
        self.viet(MAU)
        code, data = self.one_json([str(self.dir), "--plan-only"])
        self.assertEqual(code, 0)
        self.assertTrue(data["ready"])
        self.assertEqual(data["so_canh"], 8)
        self.assertEqual(data["files"], [])
        self.assertIsNone(data["error"])
        self.assertEqual(sorted(p.name for p in self.dir.iterdir()), ["video.md"])

    def test_missing_ffmpeg_and_chromium_are_reported_before_any_voice(self):
        self.viet(MOT_CANH)
        with mock.patch.object(video_ma, "co_ffmpeg", return_value=False), \
                mock.patch.object(video_ma.giong, "lay_giong", side_effect=AssertionError("không được gọi giọng")):
            code, data = self.one_json([str(self.dir)])
        self.assertEqual((code, data["error"]["step"]), (1, "ffmpeg"))
        with mock.patch.object(video_ma, "co_ffmpeg", return_value=True), \
                mock.patch.object(video_ma, "co_chromium", return_value=False), \
                mock.patch.object(video_ma.giong, "lay_giong", side_effect=AssertionError("không được gọi giọng")):
            code, data = self.one_json([str(self.dir)])
        self.assertEqual((code, data["error"]["step"]), (1, "chromium"))
        self.assertIn("pptmaster.ps1", data["error"]["fix"])

    def test_voice_error_is_reported_as_giong(self):
        self.viet(MOT_CANH)
        boom = media.MediaError("giong", "mất mạng", "Có mạng rồi chạy lại.")
        with mock.patch.object(video_ma, "co_ffmpeg", return_value=True), \
                mock.patch.object(video_ma, "co_chromium", return_value=True), \
                mock.patch.object(video_ma.giong, "lay_giong", side_effect=boom):
            code, data = self.one_json([str(self.dir)])
        self.assertEqual((code, data["error"]["step"]), (1, "giong"))

    def test_unexpected_error_is_internal(self):
        self.viet(MOT_CANH)
        with mock.patch.object(video_ma, "co_ffmpeg", side_effect=RuntimeError("bất ngờ")):
            code, data = self.one_json([str(self.dir)])
        self.assertEqual((code, data["error"]["step"]), (1, "internal"))
        self.assertIn("bất ngờ", data["error"]["message"])

    def test_success_payload_and_temp_folder_is_cleaned(self):
        self.viet(MOT_CANH)
        giong = lich.GiongInfo(mp3=self.dir / "giong" / "canh-1.mp3", giay=3.0, moc_cau=[0.0], uoc_luong=False, nguon="may")

        class FakeBrowser:
            def __enter__(self_inner):
                return object()

            def __exit__(self_inner, *a):
                return False

        (self.dir / ".khung" / "anh").mkdir(parents=True)
        (self.dir / ".khung" / "anh" / "f000000.png").write_bytes(b"khung-cu")

        def fake_chup(page, html, so_khung, fps, thu_muc, so_dau, ghi_log=None):
            self.assertFalse((thu_muc / "f000000.png").exists(), "khung cũ phải bị dọn trước khi chụp")
            for i in range(so_khung):
                (thu_muc / f"f{so_dau + i:06d}.png").write_bytes(b"x")
            return so_dau + so_khung

        def fake_ghep(thu_muc, cac_lich, cac_giong, phu_de, **kw):
            (thu_muc / "video.mp4").write_bytes(b"mp4")
            return ["video.mp4"]

        with mock.patch.object(video_ma, "co_ffmpeg", return_value=True), \
                mock.patch.object(video_ma, "co_chromium", return_value=True), \
                mock.patch.object(video_ma.giong, "lay_giong", return_value=giong), \
                mock.patch.object(video_ma.chup, "trinh_duyet", return_value=FakeBrowser()), \
                mock.patch.object(video_ma.chup, "trang_moi", return_value=object()), \
                mock.patch.object(video_ma.chup, "kiem_tran", return_value=[]), \
                mock.patch.object(video_ma.chup, "chup_canh", side_effect=fake_chup), \
                mock.patch.object(video_ma.ghep, "ghep_video", side_effect=fake_ghep):
            code, data = self.one_json([str(self.dir)])
        self.assertEqual(code, 0, data)
        self.assertTrue(data["ready"])
        self.assertEqual(data["files"], ["video.mp4"])
        self.assertEqual(data["so_canh"], 1)
        self.assertEqual(data["giong"], "may")
        self.assertEqual(data["phong_cach"], "viet-tay")
        self.assertAlmostEqual(data["thoi_luong_giay"], lich.thoi_luong_canh(3.0), delta=0.01)
        self.assertFalse((self.dir / ".khung").exists())

    def test_thu_muc_khung_tam_duoc_don(self):
        self.test_success_payload_and_temp_folder_is_cleaned()

    def test_overflow_is_a_canh_error_before_capturing(self):
        self.viet(MOT_CANH)
        giong = lich.GiongInfo(mp3=self.dir / "x.mp3", giay=3.0, moc_cau=[0.0], uoc_luong=False, nguon="co-san")

        class FakeBrowser:
            def __enter__(self_inner):
                return object()

            def __exit__(self_inner, *a):
                return False

        with mock.patch.object(video_ma, "co_ffmpeg", return_value=True), \
                mock.patch.object(video_ma, "co_chromium", return_value=True), \
                mock.patch.object(video_ma.giong, "lay_giong", return_value=giong), \
                mock.patch.object(video_ma.chup, "trinh_duyet", return_value=FakeBrowser()), \
                mock.patch.object(video_ma.chup, "trang_moi", return_value=object()), \
                mock.patch.object(video_ma.chup, "kiem_tran", return_value=["chu"]), \
                mock.patch.object(video_ma.chup, "chup_canh", side_effect=AssertionError("không được chụp")):
            code, data = self.one_json([str(self.dir)])
        self.assertEqual((code, data["error"]["step"]), (1, "canh"))
        self.assertIn("chu", data["error"]["message"])
        self.assertIn("Cảnh 1", data["error"]["message"])
        self.assertFalse((self.dir / ".khung").exists())

    def test_preview_needs_no_voice_and_no_ffmpeg(self):
        self.viet(MAU)

        class FakeBrowser:
            def __enter__(self_inner):
                return object()

            def __exit__(self_inner, *a):
                return False

        def fake_cuoi(page, html, duong_dan):
            duong_dan.parent.mkdir(parents=True, exist_ok=True)
            duong_dan.write_bytes(b"png")

        with mock.patch.object(video_ma, "co_ffmpeg", return_value=False), \
                mock.patch.object(video_ma, "co_chromium", return_value=True), \
                mock.patch.object(video_ma.giong, "lay_giong", side_effect=AssertionError("không được gọi giọng")), \
                mock.patch.object(video_ma.chup, "trinh_duyet", return_value=FakeBrowser()), \
                mock.patch.object(video_ma.chup, "trang_moi", return_value=object()), \
                mock.patch.object(video_ma.chup, "kiem_tran", return_value=[]), \
                mock.patch.object(video_ma.chup, "chup_cuoi", side_effect=fake_cuoi):
            code, data = self.one_json([str(self.dir), "--xem-truoc"])
        self.assertEqual(code, 0, data)
        self.assertEqual(data["files"], [f"xem-truoc/canh-{i}.png" for i in range(1, 9)])
        self.assertTrue((self.dir / "xem-truoc" / "canh-8.png").is_file())

    def test_json_line_is_valid_even_with_vietnamese_text(self):
        self.viet(MOT_CANH.replace("Xin chào các em.", "Nhờ ướt nhẫm quyết định."))
        code, lines, _ = chay([str(self.dir), "--plan-only"])
        self.assertEqual(len(lines), 1)
        json.loads(lines[0])


if __name__ == "__main__":
    unittest.main()
