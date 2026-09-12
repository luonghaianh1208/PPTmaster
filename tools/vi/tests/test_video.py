"""Test cho lớp làm video của bản Việt (không chạy FFmpeg, không cần mạng)."""

import contextlib
import io
import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

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


class CumulativeOffsetsTest(unittest.TestCase):
    def test_offsets_start_at_zero_and_sum_previous_durations(self):
        self.assertEqual(srt.cumulative_offsets([2.0, 3.5, 1.0]), [0.0, 2.0, 5.5])

    def test_empty_durations_give_no_offsets(self):
        self.assertEqual(srt.cumulative_offsets([]), [])


class DriftWarningTest(unittest.TestCase):
    """Số học của cảnh báo lệch phụ đề (trước đây không có test nào)."""

    def test_no_warning_when_every_slide_is_inside_the_tolerance(self):
        self.assertIsNone(srt.drift_warning([0.0, 30.0, 61.0], [0.0, 30.2, 61.4]))

    def test_warning_names_the_worst_slide_and_the_ffmpeg_remedy(self):
        message = srt.drift_warning([0.0, 28.4, 58.6], [0.0, 30.5, 62.0])
        self.assertIsNotNone(message)
        self.assertIn("slide 3", message)
        self.assertIn("3.4 giây", message)
        self.assertIn("--cach ffmpeg", message)

    def test_a_progressive_drift_is_caught_even_when_totals_match(self):
        # Tổng bằng nhau (90 s) nhưng slide giữa lệch 5 s: phép so tổng thời
        # lượng cũ bỏ qua đúng trường hợp này.
        self.assertIsNotNone(srt.drift_warning([0.0, 30.0, 60.0], [0.0, 35.0, 60.0]))

    def test_tolerance_is_configurable(self):
        self.assertIsNone(srt.drift_warning([0.0, 30.0], [0.0, 31.0], tolerance=1.5))


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
            Path("C:/p/images.txt"), Path("C:/p/audio.txt"), Path("C:/p/out.mp4"),
            fps=30, height=1080, total_seconds=90.38400001,
        )
        self.assertEqual(cmd[0], "ffmpeg")
        self.assertIn("-c:v", cmd)
        self.assertIn("libx264", cmd)
        self.assertIn("-c:a", cmd)
        self.assertIn("aac", cmd)
        self.assertEqual(cmd[-1], str(Path("C:/p/out.mp4")))
        # Cắt đúng tổng thời lượng tiếng: ảnh cuối bị lặp trong file concat nên
        # không cắt thì video dài hơn tiếng. `-shortest` không làm được việc này.
        self.assertEqual(cmd[cmd.index("-t") + 1], "90.384")
        self.assertNotIn("-shortest", cmd)
        # Không phóng to: ảnh slide chỉ 1280x720 nên 1080 phải là trần, không phải đích.
        self.assertIn("scale=-2:'min(ih,1080)'", " ".join(cmd))
        self.assertNotIn("scale=-2:1080", " ".join(cmd))
        self.assertNotIn("-r", cmd, "forcing -r after an image concat duplicates frames on some ffmpeg builds; use the fps= filter instead")
        video_filter = cmd[cmd.index("-vf") + 1]
        self.assertTrue(video_filter.startswith("fps=30,"), video_filter)

    def test_burned_subtitles_add_filter(self):
        cmd = media.build_render_command(
            Path("C:/p/images.txt"), Path("C:/p/audio.txt"), Path("C:/p/out.mp4"),
            fps=30, height=720, total_seconds=12.0, burn_srt=Path("C:/p/phu de.srt"),
        )
        joined = " ".join(cmd)
        self.assertIn("subtitles=", joined)
        self.assertIn("scale=-2:'min(ih,720)'", joined)
        self.assertNotIn("-r", cmd)
        video_filter = cmd[cmd.index("-vf") + 1]
        self.assertTrue(video_filter.startswith("fps=30,"), video_filter)
        # Order must stay fps -> subtitles -> scale/format: fps first turns the
        # slideshow into real CFR frames before burn-in and scaling touch them.
        self.assertLess(video_filter.index("fps=30,"), video_filter.index("subtitles="))
        self.assertLess(video_filter.index("subtitles="), video_filter.index("scale="))

    def test_scale_filter_clamps_instead_of_upscaling(self):
        self.assertEqual(media.scale_filter(720), "scale=-2:'min(ih,720)'")
        self.assertEqual(media.scale_filter(1080), "scale=-2:'min(ih,1080)'")

    def test_escape_subtitles_filter_escapes_drive_and_backslash(self):
        escaped = media.escape_subtitles_filter(Path(r"C:\du an\phu de.srt"))
        self.assertNotIn("\\d", escaped.replace("\\\\", ""))
        self.assertIn(r"C\:", escaped)


from video_parts import selection  # noqa: E402
import video  # noqa: E402

VIDEO_CLI = REPO_ROOT / "tools" / "vi" / "video.py"


def state(**kwargs):
    base = dict(
        slides=["01_mo_dau", "02_noi_dung"],
        audio=["01_mo_dau", "02_noi_dung"],
        narrated_pptx="exports/bai_narrated.pptx",
        has_powerpoint=True,
        has_chromium=True,
        previews=[],
        notes=["01_mo_dau", "02_noi_dung"],
    )
    base.update(kwargs)
    return selection.ProjectState(**base)


def fresh_state(**kwargs):
    """Trạng thái có ảnh xem trước mới hơn file SVG."""
    slides = kwargs.pop("slides", ["01_mo_dau", "02_noi_dung"])
    base = dict(
        slides=slides,
        previews=list(slides),
        slide_mtimes={stem: 100.0 for stem in slides},
        preview_mtimes={stem: 200.0 for stem in slides},
    )
    base.update(kwargs)
    return state(**base)


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
        self.assertIn("notes_to_audio.py", ctx.exception.fix)

    def test_missing_notes_keeps_the_audio_step_but_names_the_notes_step(self):
        with self.assertRaises(selection.SelectionError) as ctx:
            selection.select_backend(state(audio=[], notes=[]), "ffmpeg")
        self.assertEqual(ctx.exception.step, "audio")
        self.assertIn("notes/", ctx.exception.message)
        self.assertIn("bước ghi chú", ctx.exception.fix)
        self.assertLess(ctx.exception.fix.index("ghi chú"), ctx.exception.fix.index("thuyết minh"))

    def test_one_slide_without_notes_points_at_the_notes_step(self):
        with self.assertRaises(selection.SelectionError) as ctx:
            selection.select_backend(state(audio=["01_mo_dau"], notes=["01_mo_dau"]), "ffmpeg")
        self.assertEqual(ctx.exception.step, "audio")
        self.assertIn("ghi chú", ctx.exception.message)
        self.assertIn("02_noi_dung", ctx.exception.message)


class PreviewsFreshTest(unittest.TestCase):
    def test_previews_newer_than_the_slides_are_fresh(self):
        self.assertTrue(selection.previews_fresh(fresh_state()))

    def test_an_edited_slide_makes_its_preview_stale(self):
        stale = fresh_state(slide_mtimes={"01_mo_dau": 100.0, "02_noi_dung": 300.0})
        self.assertFalse(selection.previews_fresh(stale))

    def test_missing_preview_names_are_stale(self):
        self.assertFalse(selection.previews_fresh(fresh_state(previews=["01_mo_dau"])))

    def test_previews_without_mtimes_are_treated_as_stale(self):
        self.assertFalse(selection.previews_fresh(state(previews=["01_mo_dau", "02_noi_dung"])))


class PlanStepsTest(unittest.TestCase):
    def test_ffmpeg_plan_lists_capture_and_render(self):
        steps = [(s["step"], s["action"], s["method"]) for s in selection.plan_steps(state(), "ffmpeg", "file")]
        self.assertEqual(steps, [
            ("preview", "capture", "visual_review.py"),
            ("subtitle", "merge", "srt"),
            ("render", "run", "ffmpeg"),
        ])

    def test_ffmpeg_plan_requires_chromium_when_missing(self):
        steps = selection.plan_steps(state(has_chromium=False), "ffmpeg", "file")
        self.assertEqual(steps[0]["step"], "chromium")
        self.assertEqual(steps[0]["action"], "require")
        self.assertEqual(steps[0]["method"], "pip+playwright")

    def test_ffmpeg_plan_skips_capture_when_previews_are_fresh(self):
        steps = [s["step"] for s in selection.plan_steps(fresh_state(), "ffmpeg", "file")]
        self.assertNotIn("preview", steps)

    def test_ffmpeg_plan_skips_chromium_when_previews_are_fresh(self):
        # Ruling R23: không có bước chụp ảnh nào trong kế hoạch thì cũng
        # không cần đòi Chromium, dù has_chromium là False.
        steps = [s["step"] for s in selection.plan_steps(fresh_state(has_chromium=False), "ffmpeg", "file")]
        self.assertNotIn("chromium", steps)
        self.assertNotIn("preview", steps)

    def test_ffmpeg_plan_recaptures_when_a_slide_is_newer_than_its_preview(self):
        stale = fresh_state(slide_mtimes={"01_mo_dau": 100.0, "02_noi_dung": 300.0})
        steps = [s["step"] for s in selection.plan_steps(stale, "ffmpeg", "file")]
        self.assertIn("preview", steps)

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
            self.assertNotIn("ký tự", data["error"]["message"])

    def test_error_payload_reports_the_slide_count(self):
        with tempfile.TemporaryDirectory() as tmp:
            project = self.build_project(Path(tmp) / "du_an", with_audio=False)
            _, data = self.run_cli(str(project), "--plan-only")
            self.assertEqual(data["slides"], 2)

    def test_a_very_long_project_path_names_the_length_limit(self):
        with tempfile.TemporaryDirectory() as tmp:
            deep = Path(tmp) / ("d" * 120) / ("e" * 120) / "du_an"
            code, data = self.run_cli(str(deep), "--plan-only")
            self.assertEqual(code, 1)
            self.assertEqual(data["error"]["step"], "project")
            self.assertIn("260 ký tự", data["error"]["message"])
            self.assertIn("D:\\PPTmaster", data["error"]["fix"])

    def test_plan_only_never_instantiates_powerpoint(self):
        """Ruling R1/R17: không đường `--plan-only` nào được gọi COM."""
        def explode():
            raise AssertionError("--plan-only đã gọi has_powerpoint()")

        with tempfile.TemporaryDirectory() as tmp:
            project = self.build_project(Path(tmp) / "du_an")
            (project / "exports" / "bai_narrated.pptx").write_bytes(b"")
            with mock.patch.object(video, "has_powerpoint", side_effect=explode), \
                    mock.patch.object(video, "has_powerpoint_installed", return_value=True), \
                    mock.patch.object(video, "has_chromium", return_value=True):
                buf = io.StringIO()
                with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(io.StringIO()):
                    code = video.main([str(project), "--cach", "auto", "--plan-only"])
            data = json.loads(buf.getvalue().strip())
            self.assertEqual(code, 0)
            self.assertEqual(data["backend"], "powerpoint")

    def test_existing_video_names_keep_their_own_timestamp(self):
        with tempfile.TemporaryDirectory() as tmp:
            project = self.build_project(Path(tmp) / "du_an")
            older = project / "exports" / f"{project.name}_video_20260101_000000.mp4"
            older.write_bytes(b"cu")

            def render(_project, _stems, _durations, out_path, _height, _burn):
                out_path.write_bytes(b"moi")

            with mock.patch.object(video, "has_chromium", return_value=True), \
                    mock.patch.object(video.shutil, "which", return_value="ffmpeg"), \
                    mock.patch.object(video.media, "probe_duration", return_value=2.0), \
                    mock.patch.object(video, "capture_previews"), \
                    mock.patch.object(video, "render_ffmpeg", side_effect=render):
                buf = io.StringIO()
                with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(io.StringIO()):
                    code = video.main([str(project), "--cach", "ffmpeg"])
            data = json.loads(buf.getvalue().strip())
            self.assertEqual(code, 0)
            self.assertTrue(older.is_file(), "video cũ bị ghi đè")
            self.assertNotEqual(Path(data["video"]).name, older.name)
            self.assertRegex(Path(data["video"]).name, r"_video_\d{8}_\d{6}\.mp4$")


    def test_no_subtitles_means_no_pptx_timeline_warning(self):
        """`--phu-de khong` không được cảnh báo về phụ đề nó không tạo."""
        with tempfile.TemporaryDirectory() as tmp:
            project = self.build_project(Path(tmp) / "du_an")
            (project / "exports" / "bai_narrated.pptx").write_bytes(b"khong phai zip")

            def render(_pptx, out_path, _height):
                out_path.write_bytes(b"video")

            with mock.patch.object(video, "has_powerpoint", return_value=True), \
                    mock.patch.object(video, "has_chromium", return_value=True), \
                    mock.patch.object(video.shutil, "which", return_value="ffmpeg"), \
                    mock.patch.object(video.media, "probe_duration", return_value=2.0), \
                    mock.patch.object(video, "render_powerpoint", side_effect=render):
                buf = io.StringIO()
                with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(io.StringIO()):
                    code = video.main([str(project), "--cach", "powerpoint", "--phu-de", "khong"])
            data = json.loads(buf.getvalue().strip())
            self.assertEqual(code, 0)
            self.assertEqual(data["backend"], "powerpoint")
            self.assertIsNone(data["subtitle"])
            self.assertEqual(data["warnings"], [])


class HasChromiumTest(unittest.TestCase):
    def test_browser_folder_plus_importable_playwright_is_found(self):
        with tempfile.TemporaryDirectory() as tmp:
            (Path(tmp) / "ms-playwright" / "chromium-1234").mkdir(parents=True)
            with mock.patch.dict(os.environ, {"LOCALAPPDATA": tmp}), \
                    mock.patch.object(video.subprocess, "run", return_value=subprocess.CompletedProcess([], 0)):
                self.assertTrue(video.has_chromium())

    def test_browser_folder_without_playwright_package_is_not_found(self):
        with tempfile.TemporaryDirectory() as tmp:
            (Path(tmp) / "ms-playwright" / "chromium-1234").mkdir(parents=True)
            with mock.patch.dict(os.environ, {"LOCALAPPDATA": tmp}), \
                    mock.patch.object(video.subprocess, "run", return_value=subprocess.CompletedProcess([], 1)):
                self.assertFalse(video.has_chromium())

    def test_the_check_asks_the_interpreter_that_will_capture_the_slides(self):
        with tempfile.TemporaryDirectory() as tmp:
            (Path(tmp) / "ms-playwright" / "chromium-1234").mkdir(parents=True)
            with mock.patch.dict(os.environ, {"LOCALAPPDATA": tmp}), \
                    mock.patch.object(video.subprocess, "run",
                                      return_value=subprocess.CompletedProcess([], 0)) as run:
                video.has_chromium()
            self.assertEqual(run.call_args.args[0], [video.python_exe(), "-c", "import playwright"])

    def test_a_broken_interpreter_is_not_found_instead_of_raising(self):
        with tempfile.TemporaryDirectory() as tmp:
            (Path(tmp) / "ms-playwright" / "chromium-1234").mkdir(parents=True)
            with mock.patch.dict(os.environ, {"LOCALAPPDATA": tmp}), \
                    mock.patch.object(video.subprocess, "run", side_effect=OSError("khong chay duoc")):
                self.assertFalse(video.has_chromium())

    def test_no_browser_folder_is_not_found(self):
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch.dict(os.environ, {"LOCALAPPDATA": tmp}):
                self.assertFalse(video.has_chromium())


class PreviewServerRunningTest(unittest.TestCase):
    def test_no_lock_reports_not_running(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertFalse(video.preview_server_running(Path(tmp)))

    def test_live_preview_lock_json_reports_running(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "live_preview").mkdir()
            (root / "live_preview" / "lock.json").write_text("{}", encoding="utf-8")
            self.assertTrue(video.preview_server_running(root))

    def test_legacy_dot_lock_reports_running(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / ".live_preview.lock").write_text("", encoding="utf-8")
            self.assertTrue(video.preview_server_running(root))


class MainUnexpectedErrorTest(unittest.TestCase):
    """video.main() must still emit exactly one JSON line and exit 1 when an
    unguarded exception happens (here: a .srt that is not valid UTF-8).

    This calls video.main() in-process instead of spawning tools/vi/video.py
    as a subprocess, and patches has_chromium/shutil.which/probe_duration so
    the test never depends on (or is skipped by) whether FFmpeg or Playwright
    Chromium happen to be installed on the machine running the suite, and
    never spawns a real ffprobe/ffmpeg process.
    """

    def build_project(self, root):
        (root / "svg_output").mkdir(parents=True)
        (root / "audio").mkdir(parents=True)
        (root / "exports").mkdir(parents=True)
        for stem in ("01_mo_dau", "02_noi_dung"):
            (root / "svg_output" / f"{stem}.svg").write_text("<svg/>", encoding="utf-8")
            (root / "audio" / f"{stem}.mp3").write_bytes(b"")
        (root / "audio" / "01_mo_dau.srt").write_bytes(b"\xff\xfe khong phai utf-8")
        (root / "audio" / "02_noi_dung.srt").write_text(SAMPLE_SRT, encoding="utf-8")
        return root

    def test_invalid_utf8_srt_is_reported_as_json_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            project = self.build_project(Path(tmp) / "du_an")
            with mock.patch.object(video, "has_chromium", return_value=True), \
                    mock.patch.object(video.shutil, "which", return_value="ffmpeg"), \
                    mock.patch.object(video.media, "probe_duration", return_value=2.0):
                buf = io.StringIO()
                with contextlib.redirect_stdout(buf):
                    code = video.main([str(project), "--cach", "ffmpeg"])
            stdout = buf.getvalue().strip()
            self.assertEqual(len(stdout.splitlines()), 1, f"stdout phải là một dòng JSON:\n{stdout}")
            data = json.loads(stdout)
            self.assertEqual(code, 1)
            self.assertIsNotNone(data["error"])
            self.assertEqual(data["error"]["step"], "render")


from video_parts import pptx_timeline  # noqa: E402

SLIDE_XML = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
    ' xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">'
    '<p:transition p14:dur="{transition}" advClick="0" advTm="{advance}"/>'
    "<p:timing><p:tnLst><p:par><p:cTn><p:childTnLst><p:audio><p:cMediaNode><p:cTn>"
    '<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>'
    "</p:cTn></p:cMediaNode></p:audio></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>"
    "</p:sld>"
)


def write_narrated_pptx(path: Path, slides, with_order=True):
    """Một zip hình dáng PPTX với các mốc advTm/dur/delay đã biết trước."""
    with zipfile.ZipFile(path, "w") as package:
        for index, (transition, delay, advance) in enumerate(slides, start=1):
            package.writestr(
                f"ppt/slides/slide{index}.xml",
                SLIDE_XML.format(transition=transition, delay=delay, advance=advance),
            )
        if not with_order:
            return
        slide_ids = "".join(
            f'<p:sldId id="{255 + index}" r:id="rId{index}"/>'
            for index in range(1, len(slides) + 1)
        )
        package.writestr(
            "ppt/presentation.xml",
            '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f"<p:sldIdLst>{slide_ids}</p:sldIdLst></p:presentation>",
        )
        relationships = "".join(
            f'<Relationship Id="rId{index}" Target="slides/slide{index}.xml"/>'
            for index in range(1, len(slides) + 1)
        )
        package.writestr(
            "ppt/_rels/presentation.xml.rels",
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f"{relationships}</Relationships>",
        )


class PptxTimelineTest(unittest.TestCase):
    """Mốc phụ đề của đường PowerPoint phải đọc từ advTm, không cộng dồn tiếng."""

    REAL_SLIDES = [(400, 400, 29267), (400, 400, 31163), (400, 400, 32652)]

    def test_starts_follow_the_pptx_clock_not_the_audio_sums(self):
        with tempfile.TemporaryDirectory() as tmp:
            pptx = Path(tmp) / "bai_narrated.pptx"
            write_narrated_pptx(pptx, self.REAL_SLIDES)
            starts, timeline = pptx_timeline.narration_starts(pptx, 3)
        # slide n bắt đầu sau tổng (dur + advTm) của các slide trước; tiếng
        # bắt đầu muộn thêm dur + delay của chính slide đó.
        self.assertEqual([round(value, 3) for value in starts], [0.8, 30.467, 62.03])
        self.assertAlmostEqual(timeline, 94.282, places=3)

    def test_zero_transition_and_missing_audio_delay_still_read(self):
        with tempfile.TemporaryDirectory() as tmp:
            pptx = Path(tmp) / "bai_narrated.pptx"
            write_narrated_pptx(pptx, [(0, 0, 10000), (0, 0, 5000)])
            starts, timeline = pptx_timeline.narration_starts(pptx, 2)
        self.assertEqual(starts, [0.0, 10.0])
        self.assertAlmostEqual(timeline, 15.0, places=3)

    def test_slide_count_mismatch_is_a_timeline_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            pptx = Path(tmp) / "bai_narrated.pptx"
            write_narrated_pptx(pptx, self.REAL_SLIDES)
            with self.assertRaises(pptx_timeline.TimelineError):
                pptx_timeline.narration_starts(pptx, 2)

    def test_missing_slide_order_is_a_timeline_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            pptx = Path(tmp) / "bai_narrated.pptx"
            write_narrated_pptx(pptx, self.REAL_SLIDES, with_order=False)
            with self.assertRaises(pptx_timeline.TimelineError):
                pptx_timeline.narration_starts(pptx, 3)

    def test_slide_without_advtm_is_a_timeline_error(self):
        xml = (
            '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            '<p:transition advClick="0"/></p:sld>'
        )
        with self.assertRaises(pptx_timeline.TimelineError):
            pptx_timeline.slide_timing(xml.encode("utf-8"))

    def test_unreadable_package_is_a_timeline_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            pptx = Path(tmp) / "khong_phai_zip.pptx"
            pptx.write_bytes(b"khong phai zip")
            with self.assertRaises(pptx_timeline.TimelineError):
                pptx_timeline.narration_starts(pptx, 1)


class SubtitleOffsetsTest(unittest.TestCase):
    def test_ffmpeg_route_uses_cumulative_audio_durations(self):
        warnings = []
        offsets, timeline = video.subtitle_offsets(state(), "ffmpeg", [28.368, 30.264], warnings)
        self.assertEqual(offsets, [0.0, 28.368])
        self.assertIsNone(timeline)
        self.assertEqual(warnings, [])

    def test_powerpoint_route_uses_the_pptx_timeline(self):
        with tempfile.TemporaryDirectory() as tmp:
            pptx = Path(tmp) / "bai_narrated.pptx"
            write_narrated_pptx(pptx, PptxTimelineTest.REAL_SLIDES)
            warnings = []
            offsets, timeline = video.subtitle_offsets(
                state(narrated_pptx=str(pptx)), "powerpoint", [28.368, 30.264, 31.752], warnings,
            )
        self.assertEqual([round(value, 3) for value in offsets], [0.8, 30.467, 62.03])
        self.assertAlmostEqual(timeline, 94.282, places=3)
        # Ruling R24/N1: đặt tên đúng bản PPTX đã dùng để tính mốc, để một bản
        # PPTX sai (nhưng cùng số slide) không âm thầm lọt qua.
        self.assertEqual(len(warnings), 1)
        self.assertIn(pptx.name, warnings[0])

    def test_unreadable_pptx_falls_back_to_sums_with_an_actionable_warning(self):
        with tempfile.TemporaryDirectory() as tmp:
            pptx = Path(tmp) / "bai_narrated.pptx"
            pptx.write_bytes(b"khong phai zip")
            warnings = []
            with contextlib.redirect_stderr(io.StringIO()):
                offsets, timeline = video.subtitle_offsets(
                    state(narrated_pptx=str(pptx)), "powerpoint", [28.368, 30.264], warnings,
                )
        self.assertEqual(offsets, [0.0, 28.368])
        self.assertIsNone(timeline)
        self.assertEqual(len(warnings), 1)
        self.assertIn("--cach ffmpeg", warnings[0])


class ReadStateNotesFilterTest(unittest.TestCase):
    """Ruling R22/N5: `notes/total.md` một mình không được tính là đã có ghi
    chú theo slide, nếu không dự án ở trạng thái trước khi tách ghi chú sẽ
    nhận nhầm cách sửa "chạy bước thuyết minh" thay vì "tách ghi chú trước".
    """

    def build_project(self, root):
        (root / "svg_output").mkdir(parents=True)
        (root / "audio").mkdir(parents=True)
        (root / "exports").mkdir(parents=True)
        (root / "notes").mkdir(parents=True)
        for stem in ("01_mo_dau", "02_noi_dung"):
            (root / "svg_output" / f"{stem}.svg").write_text("<svg/>", encoding="utf-8")
        return root

    def test_total_notes_only_gets_the_notes_first_remedy(self):
        with tempfile.TemporaryDirectory() as tmp:
            project = self.build_project(Path(tmp) / "du_an")
            (project / "notes" / "total.md").write_text("noi dung", encoding="utf-8")
            state_obj = video.read_state(project)
            self.assertEqual(state_obj.notes, [])
            with self.assertRaises(selection.SelectionError) as ctx:
                selection.check_audio(state_obj)
            self.assertEqual(ctx.exception.step, "audio")
            self.assertIn("bước ghi chú", ctx.exception.fix)

    def test_per_slide_notes_get_the_narration_remedy(self):
        with tempfile.TemporaryDirectory() as tmp:
            project = self.build_project(Path(tmp) / "du_an")
            for stem in ("01_mo_dau", "02_noi_dung"):
                (project / "notes" / f"{stem}.md").write_text("noi dung", encoding="utf-8")
            state_obj = video.read_state(project)
            self.assertEqual(state_obj.notes, ["01_mo_dau", "02_noi_dung"])
            with self.assertRaises(selection.SelectionError) as ctx:
                selection.check_audio(state_obj)
            self.assertEqual(ctx.exception.step, "audio")
            self.assertIn("notes_to_audio.py", ctx.exception.fix)
            self.assertNotIn("bước ghi chú", ctx.exception.fix)


class ChromiumGateTest(unittest.TestCase):
    """Ruling R23: đường FFmpeg chỉ đòi Chromium khi kế hoạch thật sự có bước
    chụp ảnh — dùng lại đúng phép so mới/cũ mà `plan_steps` đã dùng.
    """

    def build_project(self, root):
        (root / "svg_output").mkdir(parents=True)
        (root / "audio").mkdir(parents=True)
        (root / "exports").mkdir(parents=True)
        (root / ".preview").mkdir(parents=True)
        for stem in ("01_mo_dau", "02_noi_dung"):
            svg = root / "svg_output" / f"{stem}.svg"
            svg.write_text("<svg/>", encoding="utf-8")
            os.utime(svg, (100, 100))
            (root / "audio" / f"{stem}.mp3").write_bytes(b"")
        return root

    def make_previews_fresh(self, root):
        for stem in ("01_mo_dau", "02_noi_dung"):
            png = root / ".preview" / f"{stem}.png"
            png.write_bytes(b"")
            os.utime(png, (200, 200))

    def test_fresh_previews_skip_the_chromium_gate_and_reach_render(self):
        with tempfile.TemporaryDirectory() as tmp:
            project = self.build_project(Path(tmp) / "du_an")
            self.make_previews_fresh(project)

            def render(_project, _stems, _durations, out_path, _height, _burn):
                out_path.write_bytes(b"video")

            with mock.patch.object(video, "has_chromium", return_value=False), \
                    mock.patch.object(video.shutil, "which", return_value="ffmpeg"), \
                    mock.patch.object(video.media, "probe_duration", return_value=2.0), \
                    mock.patch.object(video, "capture_previews") as capture, \
                    mock.patch.object(video, "render_ffmpeg", side_effect=render) as render_mock:
                buf = io.StringIO()
                with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(io.StringIO()):
                    code = video.main([str(project), "--cach", "ffmpeg"])
            data = json.loads(buf.getvalue().strip())
            self.assertEqual(code, 0, data)
            self.assertIsNone(data["error"])
            render_mock.assert_called_once()
            capture.assert_not_called()

    def test_stale_previews_still_raise_the_chromium_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            project = self.build_project(Path(tmp) / "du_an")
            # Không chụp ảnh xem trước nào -> coi như cũ, vẫn cần Chromium.
            with mock.patch.object(video, "has_chromium", return_value=False):
                buf = io.StringIO()
                with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(io.StringIO()):
                    code = video.main([str(project), "--cach", "ffmpeg"])
            data = json.loads(buf.getvalue().strip())
            self.assertEqual(code, 1)
            self.assertEqual(data["error"]["step"], "chromium")


class NarratedPptxWarningCliTest(unittest.TestCase):
    """Ruling R24/N1: cảnh báo nêu tên bản PPTX đã dùng, chỉ ở đường
    PowerPoint (đường FFmpeg không đọc bản PPTX này nên không thể cảnh báo).
    """

    def build_project(self, root):
        (root / "svg_output").mkdir(parents=True)
        (root / "audio").mkdir(parents=True)
        (root / "exports").mkdir(parents=True)
        for stem in ("01_mo_dau", "02_noi_dung"):
            (root / "svg_output" / f"{stem}.svg").write_text("<svg/>", encoding="utf-8")
            (root / "audio" / f"{stem}.mp3").write_bytes(b"")
        return root

    def test_powerpoint_route_names_the_narrated_pptx_in_warnings(self):
        with tempfile.TemporaryDirectory() as tmp:
            project = self.build_project(Path(tmp) / "du_an")
            pptx = project / "exports" / "bai_narrated.pptx"
            write_narrated_pptx(pptx, [(400, 400, 29267), (400, 400, 31163)])

            def render(_pptx, out_path, _height):
                out_path.write_bytes(b"video")

            with mock.patch.object(video, "has_powerpoint", return_value=True), \
                    mock.patch.object(video, "has_chromium", return_value=True), \
                    mock.patch.object(video.shutil, "which", return_value="ffmpeg"), \
                    mock.patch.object(video.media, "probe_duration", return_value=2.0), \
                    mock.patch.object(video, "render_powerpoint", side_effect=render):
                buf = io.StringIO()
                with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(io.StringIO()):
                    code = video.main([str(project), "--cach", "powerpoint"])
            data = json.loads(buf.getvalue().strip())
            self.assertEqual(code, 0, data)
            self.assertTrue(any("bai_narrated.pptx" in warning for warning in data["warnings"]))

    def test_ffmpeg_route_has_no_narrated_pptx_warning(self):
        with tempfile.TemporaryDirectory() as tmp:
            project = self.build_project(Path(tmp) / "du_an")
            (project / "exports" / "bai_narrated.pptx").write_bytes(b"khong dung toi")

            def render(_project, _stems, _durations, out_path, _height, _burn):
                out_path.write_bytes(b"video")

            with mock.patch.object(video, "has_chromium", return_value=True), \
                    mock.patch.object(video.shutil, "which", return_value="ffmpeg"), \
                    mock.patch.object(video.media, "probe_duration", return_value=2.0), \
                    mock.patch.object(video, "capture_previews"), \
                    mock.patch.object(video, "render_ffmpeg", side_effect=render):
                buf = io.StringIO()
                with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(io.StringIO()):
                    code = video.main([str(project), "--cach", "ffmpeg"])
            data = json.loads(buf.getvalue().strip())
            self.assertEqual(code, 0, data)
            self.assertFalse(any("narrated" in warning for warning in data["warnings"]))


if __name__ == "__main__":
    unittest.main()
