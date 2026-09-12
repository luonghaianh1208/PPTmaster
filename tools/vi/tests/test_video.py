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


from video_parts import selection  # noqa: E402

VIDEO_CLI = REPO_ROOT / "tools" / "vi" / "video.py"


def state(**kwargs):
    base = dict(
        slides=["01_mo_dau", "02_noi_dung"],
        audio=["01_mo_dau", "02_noi_dung"],
        narrated_pptx="exports/bai_narrated.pptx",
        has_powerpoint=True,
        has_chromium=True,
        previews=[],
    )
    base.update(kwargs)
    return selection.ProjectState(**base)


class SelectBackendTest(unittest.TestCase):
    def test_auto_prefers_powerpoint_when_available(self):
        backend, warnings = selection.select_backend(state(), "auto")
        self.assertEqual(backend, "powerpoint")
        self.assertEqual(warnings, [])

    def test_auto_falls_back_to_ffmpeg_without_powerpoint(self):
        backend, warnings = selection.select_backend(state(has_powerpoint=False), "auto")
        self.assertEqual(backend, "ffmpeg")
        self.assertTrue(any("PowerPoint" in warning for warning in warnings))

    def test_auto_falls_back_to_ffmpeg_without_narrated_pptx(self):
        backend, _ = selection.select_backend(state(narrated_pptx=None), "auto")
        self.assertEqual(backend, "ffmpeg")

    def test_explicit_powerpoint_without_powerpoint_is_an_error(self):
        with self.assertRaises(selection.SelectionError) as ctx:
            selection.select_backend(state(has_powerpoint=False), "powerpoint")
        self.assertEqual(ctx.exception.step, "powerpoint")

    def test_missing_audio_is_an_error_for_every_backend(self):
        for requested in ("auto", "powerpoint", "ffmpeg"):
            with self.subTest(requested=requested):
                with self.assertRaises(selection.SelectionError) as ctx:
                    selection.select_backend(state(audio=[]), requested)
                self.assertEqual(ctx.exception.step, "audio")

    def test_audio_missing_for_one_slide_is_an_error(self):
        with self.assertRaises(selection.SelectionError) as ctx:
            selection.select_backend(state(audio=["01_mo_dau"]), "ffmpeg")
        self.assertEqual(ctx.exception.step, "audio")
        self.assertIn("02_noi_dung", ctx.exception.message)


class PlanStepsTest(unittest.TestCase):
    def test_ffmpeg_plan_lists_capture_and_render(self):
        steps = [(s["step"], s["action"], s["method"]) for s in selection.plan_steps(state(), "ffmpeg", "file")]
        self.assertEqual(steps, [
            ("preview", "capture", "visual_review.py"),
            ("subtitle", "merge", "srt"),
            ("render", "run", "ffmpeg"),
        ])

    def test_ffmpeg_plan_adds_chromium_install_when_missing(self):
        steps = selection.plan_steps(state(has_chromium=False), "ffmpeg", "file")
        self.assertEqual(steps[0]["step"], "chromium")
        self.assertEqual(steps[0]["method"], "pip+playwright")

    def test_ffmpeg_plan_skips_capture_when_previews_exist(self):
        steps = [s["step"] for s in selection.plan_steps(state(previews=["01_mo_dau", "02_noi_dung"]), "ffmpeg", "file")]
        self.assertNotIn("preview", steps)

    def test_powerpoint_plan_uses_upstream_exporter(self):
        steps = [(s["step"], s["method"]) for s in selection.plan_steps(state(), "powerpoint", "hinh")]
        self.assertIn(("render", "powerpoint_video.py"), steps)
        self.assertIn(("subtitle", "srt"), steps)
        self.assertIn(("burn", "ffmpeg"), steps)

    def test_no_subtitle_mode_drops_subtitle_steps(self):
        steps = [s["step"] for s in selection.plan_steps(state(), "ffmpeg", "khong")]
        self.assertNotIn("subtitle", steps)
        self.assertNotIn("burn", steps)


class VideoCliPlanTest(unittest.TestCase):
    def build_project(self, root, with_audio=True):
        (root / "svg_output").mkdir(parents=True)
        (root / "audio").mkdir(parents=True)
        (root / "exports").mkdir(parents=True)
        for stem in ("01_mo_dau", "02_noi_dung"):
            (root / "svg_output" / f"{stem}.svg").write_text("<svg/>", encoding="utf-8")
            if with_audio:
                (root / "audio" / f"{stem}.mp3").write_bytes(b"")
                (root / "audio" / f"{stem}.srt").write_text(SAMPLE_SRT, encoding="utf-8")
        return root

    def run_cli(self, *args):
        proc = subprocess.run(
            [sys.executable, str(VIDEO_CLI), *args],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=120,
            env={**os.environ, "PYTHONIOENCODING": "utf-8"},
        )
        stdout = proc.stdout.strip()
        self.assertEqual(len(stdout.splitlines()), 1, f"stdout phải là một dòng JSON:\n{proc.stdout}\n{proc.stderr}")
        return proc.returncode, json.loads(stdout)

    def test_plan_only_reports_backend_and_steps(self):
        with tempfile.TemporaryDirectory() as tmp:
            project = self.build_project(Path(tmp) / "du_an")
            code, data = self.run_cli(str(project), "--cach", "ffmpeg", "--plan-only")
            self.assertEqual(code, 0)
            self.assertEqual(data["backend"], "ffmpeg")
            self.assertIn("render", [step["step"] for step in data["steps"]])

    def test_missing_audio_reports_error_json(self):
        with tempfile.TemporaryDirectory() as tmp:
            project = self.build_project(Path(tmp) / "du_an", with_audio=False)
            code, data = self.run_cli(str(project), "--plan-only")
            self.assertEqual(code, 1)
            self.assertEqual(data["error"]["step"], "audio")
            self.assertIn("notes_to_audio.py", data["error"]["fix"])

    def test_unknown_project_reports_error_json(self):
        with tempfile.TemporaryDirectory() as tmp:
            code, data = self.run_cli(str(Path(tmp) / "khong_co"), "--plan-only")
            self.assertEqual(code, 1)
            self.assertEqual(data["error"]["step"], "project")


if __name__ == "__main__":
    unittest.main()
