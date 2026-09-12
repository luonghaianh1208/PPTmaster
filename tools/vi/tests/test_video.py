"""Test cho lớp làm video của bản Việt (không chạy FFmpeg, không cần mạng)."""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT / "tools" / "vi"))

from video_parts import srt  # noqa: E402


SAMPLE_SRT = (
    "1\n"
    "00:00:00,000 --> 00:00:02,500\n"
    "Chào các em.\n"
    "\n"
    "2\n"
    "00:00:02,500 --> 00:00:05,000\n"
    "Hôm nay ta học phân số.\n"
    "\n"
)


class TimestampTest(unittest.TestCase):
    def test_format_timestamp_uses_srt_form(self):
        self.assertEqual(srt.format_timestamp(0), "00:00:00,000")
        self.assertEqual(srt.format_timestamp(62.5), "00:01:02,500")
        self.assertEqual(srt.format_timestamp(3723.004), "01:02:03,004")

    def test_parse_timestamp_round_trips(self):
        self.assertAlmostEqual(srt.parse_timestamp("01:02:03,004"), 3723.004, places=3)
        self.assertEqual(srt.format_timestamp(srt.parse_timestamp("00:01:02,500")), "00:01:02,500")


class ParseRenderTest(unittest.TestCase):
    def test_parse_srt_reads_cues(self):
        cues = srt.parse_srt(SAMPLE_SRT)
        self.assertEqual(len(cues), 2)
        self.assertAlmostEqual(cues[1].start, 2.5, places=3)
        self.assertAlmostEqual(cues[1].end, 5.0, places=3)
        self.assertEqual(cues[1].text, "Hôm nay ta học phân số.")

    def test_parse_srt_accepts_crlf_and_bom(self):
        cues = srt.parse_srt("﻿" + SAMPLE_SRT.replace("\n", "\r\n"))
        self.assertEqual(len(cues), 2)
        self.assertEqual(cues[0].text, "Chào các em.")

    def test_parse_srt_keeps_multiline_text(self):
        text = "1\n00:00:00,000 --> 00:00:01,000\ndòng một\ndòng hai\n\n"
        self.assertEqual(srt.parse_srt(text)[0].text, "dòng một\ndòng hai")

    def test_render_srt_round_trips(self):
        self.assertEqual(srt.render_srt(srt.parse_srt(SAMPLE_SRT)), SAMPLE_SRT)


class MergeTest(unittest.TestCase):
    def test_merge_shifts_and_renumbers(self):
        merged = srt.merge_srt([(SAMPLE_SRT, 0.0), (SAMPLE_SRT, 5.0)])
        cues = srt.parse_srt(merged)
        self.assertEqual([cue.index for cue in cues], [1, 2, 3, 4])
        self.assertAlmostEqual(cues[2].start, 5.0, places=3)
        self.assertAlmostEqual(cues[3].end, 10.0, places=3)

    def test_merge_skips_empty_sources(self):
        merged = srt.merge_srt([("", 0.0), (SAMPLE_SRT, 3.0)])
        cues = srt.parse_srt(merged)
        self.assertEqual(len(cues), 2)
        self.assertAlmostEqual(cues[0].start, 3.0, places=3)


from video_parts import media  # noqa: E402


def fake_run(stdout="", returncode=0):
    def run(cmd, **kwargs):
        return subprocess.CompletedProcess(cmd, returncode, stdout, "")
    return run


class ProbeDurationTest(unittest.TestCase):
    def test_reads_duration_from_ffprobe(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "a.mp3"
            path.write_bytes(b"")
            self.assertAlmostEqual(media.probe_duration(path, run=fake_run("12.340000\n")), 12.34, places=3)

    def test_raises_media_error_when_ffprobe_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "a.mp3"
            path.write_bytes(b"")
            with self.assertRaises(media.MediaError) as ctx:
                media.probe_duration(path, run=fake_run("", returncode=1))
            self.assertEqual(ctx.exception.step, "ffmpeg")

    def test_raises_media_error_on_unparsable_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "a.mp3"
            path.write_bytes(b"")
            with self.assertRaises(media.MediaError):
                media.probe_duration(path, run=fake_run("N/A\n"))


class ConcatTest(unittest.TestCase):
    def test_image_concat_repeats_last_file(self):
        entries = [(Path("C:/p/01.png"), 2.0), (Path("C:/p/02.png"), 3.5)]
        lines = media.build_concat_text(entries).splitlines()
        self.assertEqual(lines[0], "file 'C:/p/01.png'")
        self.assertEqual(lines[1], "duration 2.000")
        self.assertEqual(lines[2], "file 'C:/p/02.png'")
        self.assertEqual(lines[3], "duration 3.500")
        self.assertEqual(lines[4], "file 'C:/p/02.png'")

    def test_audio_concat_lists_every_file(self):
        text = media.build_audio_concat_text([Path("C:/p/01.mp3"), Path("C:/p/02.mp3")])
        self.assertEqual(text.splitlines(), ["file 'C:/p/01.mp3'", "file 'C:/p/02.mp3'"])

    def test_paths_with_quote_are_escaped(self):
        text = media.build_audio_concat_text([Path("C:/p/it's.mp3")])
        self.assertIn("'\\''", text)


class RenderCommandTest(unittest.TestCase):
    def test_command_uses_concat_inputs_and_h264(self):
        cmd = media.build_render_command(
            Path("C:/p/images.txt"), Path("C:/p/audio.txt"), Path("C:/p/out.mp4"), fps=30, height=1080,
        )
        self.assertEqual(cmd[0], "ffmpeg")
        self.assertIn("-c:v", cmd)
        self.assertIn("libx264", cmd)
        self.assertIn("-c:a", cmd)
        self.assertIn("aac", cmd)
        self.assertIn("-shortest", cmd)
        self.assertEqual(cmd[-1], str(Path("C:/p/out.mp4")))
        self.assertIn("scale=-2:1080", " ".join(cmd))

    def test_burned_subtitles_add_filter(self):
        cmd = media.build_render_command(
            Path("C:/p/images.txt"), Path("C:/p/audio.txt"), Path("C:/p/out.mp4"),
            fps=30, height=720, burn_srt=Path("C:/p/phu de.srt"),
        )
        joined = " ".join(cmd)
        self.assertIn("subtitles=", joined)
        self.assertIn("scale=-2:720", joined)

    def test_escape_subtitles_filter_escapes_drive_and_backslash(self):
        escaped = media.escape_subtitles_filter(Path(r"C:\du an\phu de.srt"))
        self.assertNotIn("\\d", escaped.replace("\\\\", ""))
        self.assertIn(r"C\:", escaped)


if __name__ == "__main__":
    unittest.main()
