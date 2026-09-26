"""Test phụ đề karaoke: file .ass với \\kf theo mốc từng từ. Không chạy FFmpeg thật trừ lớp cuối."""

import re
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))

from video_ma_parts import karaoke, lich  # noqa: E402

_KF_RE = re.compile(r"\\kf(\d+)")
_KF_TAG_RE = re.compile(r"\{\\kf\d+\}")
# Dialogue: Layer,Start,End,Style,Name,MarginL,MarginR,MarginV,Effect,Text
_DIALOGUE_RE = re.compile(r"^Dialogue: \d+,(\d+:\d{2}:\d{2}\.\d{2}),(\d+:\d{2}:\d{2}\.\d{2}),"
                          r"[^,]*,[^,]*,[^,]*,[^,]*,[^,]*,[^,]*,(.*)$")


def _tu(t, d, chu):
    return {"t": t, "d": d, "chu": chu, "khoa": chu.lower()}


def canh(so, bat_dau, giay_giong, cau, moc_cau, tu):
    return lich.CanhLich(so=so, bat_dau=bat_dau, thoi_luong=lich.DAN_DAU + giay_giong + 0.6, so_khung=1,
                         giay_giong=giay_giong, cau=cau, moc_cau=moc_cau, moc_cau_giong=[m - lich.DAN_DAU for m in moc_cau],
                         uoc_luong=False, moc_tu=tu)


def _thoi_gian_giay(nhan: str) -> float:
    gio, phut, giay = nhan.split(":")
    s, cs = giay.split(".")
    return int(gio) * 3600 + int(phut) * 60 + int(s) + int(cs) / 100


class CauTrucTest(unittest.TestCase):
    def test_has_all_three_required_sections(self):
        text = karaoke.tao_ass([])
        for section in ("[Script Info]", "[V4+ Styles]", "[Events]"):
            self.assertIn(section, text)
        self.assertIn("PlayResX: 1280", text)
        self.assertIn("PlayResY: 720", text)
        self.assertIn("WrapStyle: 0", text)

    def test_style_uses_itim_yellow_spoken_and_white_unspoken(self):
        text = karaoke.tao_ass([])
        style_line = next(l for l in text.splitlines() if l.startswith("Style:"))
        self.assertIn("Itim", style_line)
        self.assertIn("&H0000D7FF", style_line)
        self.assertIn("&H00FFFFFF", style_line)
        self.assertIn("MarginV", text)
        fields = style_line.split(",")
        # Format: Name,Fontname,Fontsize,Primary,Secondary,Outline,Back,Bold,Italic,Underline,
        # StrikeOut,ScaleX,ScaleY,Spacing,Angle,BorderStyle,Outline,Shadow,Alignment,MarginL,MarginR,MarginV,Encoding
        self.assertEqual(fields[-2], "22")  # MarginV
        self.assertEqual(fields[16], "2")  # Outline width


class KfSumTest(unittest.TestCase):
    def test_kf_total_matches_sentence_duration_in_centiseconds(self):
        cl = canh(1, 0.0, 4.0, ["Xin chao cac em."],
                  [lich.DAN_DAU], [
                      _tu(0.0, 0.3, "Xin"), _tu(0.4, 0.3, "chao"),
                      _tu(0.9, 0.3, "cac"), _tu(1.3, 0.3, "em."),
                  ])
        # moc_tu trong CanhLich đã cộng DAN_DAU theo lich.dung_lich thật
        cl.moc_tu = [{"t": lich.DAN_DAU + w["t"], "d": w["d"], "chu": w["chu"], "khoa": w["chu"].lower()} for w in cl.moc_tu]
        text = karaoke.tao_ass([cl])
        dialogues = [l for l in text.splitlines() if l.startswith("Dialogue:")]
        self.assertEqual(len(dialogues), 1)
        match = _DIALOGUE_RE.match(dialogues[0])
        self.assertIsNotNone(match)
        start, end = _thoi_gian_giay(match.group(1)), _thoi_gian_giay(match.group(2))
        du_kf = sum(int(n) for n in _KF_RE.findall(match.group(3)))
        du_that = round((end - start) * 100)
        self.assertLessEqual(abs(du_kf - du_that), 1)

    def test_multiple_sentences_each_get_their_own_dialogue_with_matching_totals(self):
        cl = canh(1, 10.0, 5.0, ["Cau mot day.", "Cau hai ngan hon."],
                  [lich.DAN_DAU, lich.DAN_DAU + 2.0],
                  [
                      {"t": lich.DAN_DAU + 0.0, "d": 0.3, "chu": "Cau", "khoa": "cau"},
                      {"t": lich.DAN_DAU + 0.4, "d": 0.3, "chu": "mot", "khoa": "mot"},
                      {"t": lich.DAN_DAU + 0.8, "d": 0.3, "chu": "day.", "khoa": "day"},
                      {"t": lich.DAN_DAU + 2.0, "d": 0.3, "chu": "Cau", "khoa": "cau"},
                      {"t": lich.DAN_DAU + 2.4, "d": 0.3, "chu": "hai", "khoa": "hai"},
                      {"t": lich.DAN_DAU + 2.8, "d": 0.3, "chu": "ngan", "khoa": "ngan"},
                      {"t": lich.DAN_DAU + 3.2, "d": 0.3, "chu": "hon.", "khoa": "hon"},
                  ])
        text = karaoke.tao_ass([cl])
        dialogues = [l for l in text.splitlines() if l.startswith("Dialogue:")]
        self.assertEqual(len(dialogues), 2)
        for line in dialogues:
            match = _DIALOGUE_RE.match(line)
            start, end = _thoi_gian_giay(match.group(1)), _thoi_gian_giay(match.group(2))
            du_kf = sum(int(n) for n in _KF_RE.findall(match.group(3)))
            self.assertLessEqual(abs(du_kf - round((end - start) * 100)), 1)
        # mốc toàn cục cộng thêm bat_dau của cảnh (10.0 giây)
        first_start = _thoi_gian_giay(_DIALOGUE_RE.match(dialogues[0]).group(1))
        self.assertAlmostEqual(first_start, 10.0 + lich.DAN_DAU, delta=0.02)


class ThoatKyTuTest(unittest.TestCase):
    def test_braces_and_backslash_in_words_are_escaped(self):
        cl = canh(1, 0.0, 1.0, ["Tap hop A = {1}."],
                  [lich.DAN_DAU],
                  [
                      {"t": lich.DAN_DAU + 0.0, "d": 0.2, "chu": "Tap", "khoa": "tap"},
                      {"t": lich.DAN_DAU + 0.3, "d": 0.2, "chu": "hop", "khoa": "hop"},
                      {"t": lich.DAN_DAU + 0.6, "d": 0.2, "chu": "A", "khoa": "a"},
                      {"t": lich.DAN_DAU + 0.8, "d": 0.2, "chu": "={1}.", "khoa": "1"},
                  ])
        text = karaoke.tao_ass([cl])
        dialogue = next(l for l in text.splitlines() if l.startswith("Dialogue:"))
        self.assertIn(r"\{1\}", dialogue)
        self.assertNotIn("={1}.", dialogue)


class BocDongTest(unittest.TestCase):
    def test_long_sentence_wraps_to_at_most_two_lines(self):
        tu = [{"t": lich.DAN_DAU + i * 0.3, "d": 0.25, "chu": f"tuso{i}", "khoa": f"tuso{i}"} for i in range(10)]
        cl = canh(1, 0.0, 3.0, [" ".join(w["chu"] for w in tu)], [lich.DAN_DAU], tu)
        text = karaoke.tao_ass([cl])
        dialogues = [l for l in text.splitlines() if l.startswith("Dialogue:")]
        for line in dialogues:
            match = _DIALOGUE_RE.match(line)
            for sub in match.group(3).split("\\N"):
                hien = _KF_TAG_RE.sub("", sub)
                self.assertLessEqual(len(hien), 50, sub)

    def test_very_long_sentence_splits_into_consecutive_dialogues(self):
        tu = [{"t": lich.DAN_DAU + i * 0.3, "d": 0.25, "chu": "tuvuadai" * 2 + str(i), "khoa": "x"} for i in range(12)]
        cl = canh(1, 0.0, 5.0, [" ".join(w["chu"] for w in tu)], [lich.DAN_DAU], tu)
        text = karaoke.tao_ass([cl])
        dialogues = [l for l in text.splitlines() if l.startswith("Dialogue:")]
        self.assertGreater(len(dialogues), 1)


class FfmpegBurnTest(unittest.TestCase):
    @unittest.skipUnless(shutil.which("ffmpeg"), "máy không có FFmpeg")
    def test_real_ffmpeg_burns_vietnamese_karaoke_into_a_colour_video(self):
        from video_ma_parts.phong import FONT as ITIM_FONT, TEN as ITIM_TEN

        cl = canh(1, 0.0, 1.4, ["Chu ki con lac don."],
                  [lich.DAN_DAU],
                  [
                      {"t": lich.DAN_DAU + 0.0, "d": 0.3, "chu": "Chu", "khoa": "chu"},
                      {"t": lich.DAN_DAU + 0.3, "d": 0.3, "chu": "ki", "khoa": "ki"},
                      {"t": lich.DAN_DAU + 0.6, "d": 0.3, "chu": "con", "khoa": "con"},
                      {"t": lich.DAN_DAU + 0.9, "d": 0.3, "chu": "lac", "khoa": "lac"},
                      {"t": lich.DAN_DAU + 1.2, "d": 0.2, "chu": "don.", "khoa": "don"},
                  ])
        with tempfile.TemporaryDirectory() as tmp:
            thu_muc = Path(tmp)
            (thu_muc / "phu-de.ass").write_text(karaoke.tao_ass([cl]), encoding="utf-8")
            fonts_dir = thu_muc / "fonts"
            fonts_dir.mkdir()
            shutil.copy2(ITIM_FONT, fonts_dir / ITIM_FONT.name)
            out = thu_muc / "out.mp4"
            cmd = ["ffmpeg", "-y", "-hide_banner", "-loglevel", "error", "-f", "lavfi",
                   "-i", "color=c=blue:s=1280x720:d=2", "-vf",
                   f"subtitles=phu-de.ass:fontsdir=fonts", "-r", "30", "-pix_fmt", "yuv420p", str(out)]
            proc = subprocess.run(cmd, cwd=thu_muc, capture_output=True, text=True)
            self.assertEqual(proc.returncode, 0, proc.stderr)
            self.assertTrue(out.is_file())


if __name__ == "__main__":
    unittest.main()
