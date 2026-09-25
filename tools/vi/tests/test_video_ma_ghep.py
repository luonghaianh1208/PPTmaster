"""Test ghép video: lệnh FFmpeg, phụ đề, danh sách nối tiếng. Không chạy FFmpeg thật."""

import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))

from video_ma_parts import ghep, lich  # noqa: E402
from video_parts import media, srt  # noqa: E402


def canh_lich(so, bat_dau, thoi_luong, giay, cau, moc_giong):
    return lich.CanhLich(so=so, bat_dau=bat_dau, thoi_luong=thoi_luong, so_khung=round(thoi_luong * lich.FPS), giay_giong=giay,
                         cau=cau, moc_cau=[lich.DAN_DAU + m for m in moc_giong], moc_cau_giong=moc_giong, uoc_luong=False)


PLAN = [
    canh_lich(1, 0.0, 6.0, 4.0, ["Xin chào **các** em.", "Học H~2~SO~4~ nhé."], [0.0, 2.0]),
    canh_lich(2, 6.0, 3.0, 1.5, ["Tạm biệt."], [0.0]),
]


class SubtitleTest(unittest.TestCase):
    def test_cues_follow_sentence_marks_with_scene_offsets_and_strip_markup(self):
        cues = ghep.cues_phu_de(PLAN)
        self.assertEqual([c.index for c in cues], [1, 2, 3])
        self.assertEqual(cues[0].text, "Xin chào các em.")
        self.assertEqual(cues[1].text, "Học H2SO4 nhé.")
        self.assertAlmostEqual(cues[0].start, lich.DAN_DAU)
        self.assertAlmostEqual(cues[0].end, lich.DAN_DAU + 2.0)
        self.assertAlmostEqual(cues[1].start, lich.DAN_DAU + 2.0)
        self.assertAlmostEqual(cues[1].end, lich.DAN_DAU + 4.0)
        self.assertAlmostEqual(cues[2].start, 6.0 + lich.DAN_DAU)
        self.assertAlmostEqual(cues[2].end, 6.0 + lich.DAN_DAU + 1.5)

    def test_cues_render_to_valid_srt(self):
        text = srt.render_srt(ghep.cues_phu_de(PLAN))
        self.assertEqual(len(srt.parse_srt(text)), 3)


class CommandTest(unittest.TestCase):
    def test_audio_command_delays_pads_and_cuts_to_scene_length(self):
        cmd = ghep.lenh_am_canh(Path("giong/canh-1.mp3"), Path("am-1.wav"), 6.0)
        joined = " ".join(cmd)
        self.assertIn(f"adelay={int(round(lich.DAN_DAU * 1000))}:all=1", joined)
        self.assertIn("apad=whole_dur=6.000", joined)
        self.assertEqual(cmd[cmd.index("-t") + 1], "6.000")

    def test_video_command_uses_scene_fps_frames_and_30_fps_output(self):
        cmd = ghep.lenh_video(Path("am.txt"), Path(".khung/video.mp4"), lich.FPS, ".khung/phu-de.srt")
        self.assertEqual(cmd[cmd.index("-framerate") + 1], str(lich.FPS))
        self.assertEqual(cmd[cmd.index("-framerate") + 3], ".khung/anh/f%06d.png")
        self.assertEqual(cmd[cmd.index("-r") + 1], "30")
        self.assertIn("libx264", cmd)
        vf = cmd[cmd.index("-vf") + 1]
        self.assertTrue(vf.startswith("subtitles=.khung/phu-de.srt:fontsdir=.khung/fonts:force_style="), vf)
        self.assertIn("FontName=Itim", vf)

    def test_video_command_without_burned_subtitles_has_no_filter(self):
        cmd = ghep.lenh_video(Path("am.txt"), Path(".khung/video.mp4"), lich.FPS, None)
        self.assertNotIn("-vf", cmd)


class AssembleTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)

    def project(self, name):
        thu_muc = Path(self.tmp.name) / name
        (thu_muc / ".khung" / "anh").mkdir(parents=True)
        return thu_muc

    def giong(self, thu_muc):
        out = []
        for cl in PLAN:
            mp3 = thu_muc / "giong" / f"canh-{cl.so}.mp3"
            mp3.parent.mkdir(exist_ok=True)
            mp3.write_bytes(b"ID3")
            out.append(lich.GiongInfo(mp3=mp3, giay=cl.giay_giong, moc_cau=[], uoc_luong=False, nguon="may"))
        return out

    def fake_run(self, thu_muc, calls, fail=None):
        def run(cmd, **kwargs):
            calls.append((cmd, kwargs.get("cwd")))
            if cmd[-1].endswith("video.mp4"):
                Path(cmd[-1] if Path(cmd[-1]).is_absolute() else Path(kwargs["cwd"]) / cmd[-1]).write_bytes(b"mp4")
            if fail:
                return subprocess.CompletedProcess(cmd, 1, "", "loi ffmpeg")
            return subprocess.CompletedProcess(cmd, 0, "", "")
        return run

    def test_folder_with_spaces_diacritics_and_apostrophe(self):
        thu_muc = self.project("Bài 5 – Sulfur dioxide (thử) 'a'")
        calls = []
        files = ghep.ghep_video(thu_muc, PLAN, self.giong(thu_muc), "hinh", run=self.fake_run(thu_muc, calls))
        self.assertEqual(files, ["video.mp4"])
        self.assertTrue((thu_muc / "video.mp4").is_file())
        listing = (thu_muc / ".khung" / "am.txt").read_text(encoding="utf-8")
        self.assertEqual(listing.count("file '"), 2)
        self.assertIn("Bài 5 – Sulfur dioxide (thử) '\\''a'\\''", listing)
        self.assertTrue(all(cwd == thu_muc for _, cwd in calls))
        self.assertEqual(len(calls), 3)

    def test_burned_subtitles_bundle_the_itim_font(self):
        thu_muc = self.project("p-font")
        calls = []
        ghep.ghep_video(thu_muc, PLAN, self.giong(thu_muc), "hinh", run=self.fake_run(thu_muc, calls))
        self.assertTrue((thu_muc / ".khung" / "fonts" / "Itim-Regular.ttf").is_file())
        joined = " ".join(calls[-1][0])
        self.assertIn("fontsdir=.khung/fonts", joined)
        self.assertIn("FontName=Itim", joined)

    def test_default_assembly_reads_and_writes_30_frames_per_second(self):
        thu_muc = self.project("p-fps")
        calls = []
        ghep.ghep_video(thu_muc, PLAN, self.giong(thu_muc), "khong", run=self.fake_run(thu_muc, calls))
        cmd = calls[-1][0]
        self.assertEqual(cmd[cmd.index("-framerate") + 1], "30")
        self.assertEqual(cmd[cmd.index("-r") + 1], "30")

    def test_output_rate_follows_the_capture_rate(self):
        cmd = ghep.lenh_video(Path("am.txt"), Path(".khung/video.mp4"), 25, None)
        self.assertEqual(cmd[cmd.index("-framerate") + 1], "25")
        self.assertEqual(cmd[cmd.index("-r") + 1], "25")

    def test_subtitle_modes(self):
        for mode, expect_srt, expect_burn in (("hinh", False, True), ("file", True, False), ("khong", False, False)):
            with self.subTest(mode=mode):
                thu_muc = self.project(f"p-{mode}")
                calls = []
                files = ghep.ghep_video(thu_muc, PLAN, self.giong(thu_muc), mode, run=self.fake_run(thu_muc, calls))
                self.assertEqual((thu_muc / "phu-de.srt").is_file(), expect_srt)
                self.assertEqual("phu-de.srt" in files, expect_srt)
                joined = " ".join(calls[-1][0])
                self.assertEqual("subtitles=" in joined, expect_burn)

    def test_braces_are_escaped_only_in_burned_subtitles(self):
        plan = [canh_lich(1, 0.0, 6.0, 4.0, ["Tập hợp A = {1; 2; 3}."], [0.0])]
        hinh = self.project("p-ngoac-hinh")
        ghep.ghep_video(hinh, plan, self.giong(hinh)[:1], "hinh", run=self.fake_run(hinh, []))
        burned = (hinh / ".khung" / "phu-de.srt").read_text(encoding="utf-8")
        self.assertIn(r"Tập hợp A = \{1; 2; 3\}.", burned)
        tep = self.project("p-ngoac-file")
        ghep.ghep_video(tep, plan, self.giong(tep)[:1], "file", run=self.fake_run(tep, []))
        plain = (tep / "phu-de.srt").read_text(encoding="utf-8")
        self.assertIn("Tập hợp A = {1; 2; 3}.", plain)
        self.assertNotIn("\\{", plain)

    def test_ffmpeg_failure_is_a_dung_error(self):
        thu_muc = self.project("loi")
        with self.assertRaises(media.MediaError) as caught:
            ghep.ghep_video(thu_muc, PLAN, self.giong(thu_muc), "khong", run=self.fake_run(thu_muc, [], fail=True))
        self.assertEqual(caught.exception.step, "dung")
        self.assertIn("loi ffmpeg", str(caught.exception))

    def test_missing_ffmpeg_is_an_ffmpeg_error(self):
        thu_muc = self.project("khong-co")

        def run(cmd, **kwargs):
            raise FileNotFoundError("ffmpeg")

        with self.assertRaises(media.MediaError) as caught:
            ghep.ghep_video(thu_muc, PLAN, self.giong(thu_muc), "khong", run=run)
        self.assertEqual(caught.exception.step, "ffmpeg")

    def test_video_open_in_a_player_is_a_write_error_with_advice(self):
        thu_muc = self.project("dang-mo")
        (thu_muc / "video.mp4").write_bytes(b"cu")
        import os
        real_replace = os.replace

        def deny(src, dst):
            raise PermissionError("dang mo")

        os.replace = deny
        try:
            with self.assertRaises(media.MediaError) as caught:
                ghep.ghep_video(thu_muc, PLAN, self.giong(thu_muc), "khong", run=self.fake_run(thu_muc, []))
        finally:
            os.replace = real_replace
        self.assertEqual(caught.exception.step, "write")
        self.assertIn("trình phát", caught.exception.fix)


if __name__ == "__main__":
    unittest.main()
