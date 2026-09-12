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


if __name__ == "__main__":
    unittest.main()
