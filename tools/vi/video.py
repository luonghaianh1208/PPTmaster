#!/usr/bin/env python3
"""Làm video bài giảng có lời giảng từ một dự án PPT Master.

Cách dùng:
    python tools/vi/video.py <đường_dẫn_dự_án> [--cach auto|powerpoint|ffmpeg]
        [--phu-de file|hinh|khong] [--do-phan-giai 1080|720] [--plan-only]

stdout: đúng một dòng JSON. Tiến trình và log của lệnh ngoài đi ra stderr.
Mã thoát: 0 khi dựng xong, 1 khi lỗi.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from video_parts import media, pptx_timeline, selection, srt  # noqa: E402

REPO_ROOT = Path(__file__).resolve().parents[2]
SCRIPTS = REPO_ROOT / "skills" / "ppt-master" / "scripts"
FIX_CHROMIUM = (
    "Cài Chromium bằng: powershell -NoProfile -ExecutionPolicy Bypass -File tools\\vi\\pptmaster.ps1 "
    "-Action tool -Name chromium"
)
WARN_PPTX_TIMELINE = (
    "Không đọc được mốc thời gian trong bản PPTX gắn tiếng nên phụ đề phải dựng theo tổng "
    "thời lượng tiếng; phụ đề sẽ lệch dần khoảng 1 giây mỗi slide. Dùng --cach ffmpeg nếu "
    "cần phụ đề chính xác."
)


def log(text: str) -> None:
    print(text, file=sys.stderr, flush=True)


def emit(payload: dict) -> None:
    sys.stdout.write(json.dumps(payload, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def python_exe() -> str:
    venv = REPO_ROOT / "venv" / "Scripts" / "python.exe"
    return str(venv) if venv.is_file() else sys.executable


def has_powerpoint() -> bool:
    try:
        proc = subprocess.run(
            [python_exe(), str(SCRIPTS / "powerpoint_video.py"), "--check"],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180,
        )
    except (OSError, subprocess.SubprocessError):
        return False
    return proc.returncode == 0


def has_powerpoint_installed() -> bool:
    """Dò PowerPoint mà không gọi COM, để `--plan-only` không mở cửa sổ nào."""
    try:
        import winreg
    except ImportError:
        winreg = None
    if winreg is not None:
        for root, key in (
            (winreg.HKEY_CLASSES_ROOT, r"PowerPoint.Application\CLSID"),
            (winreg.HKEY_CLASSES_ROOT, r"PowerPoint.Application\CurVer"),
        ):
            try:
                with winreg.OpenKey(root, key):
                    return True
            except OSError:
                continue
    for variable in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432"):
        base = os.environ.get(variable)
        if not base:
            continue
        office = Path(base) / "Microsoft Office"
        for pattern in ("root/Office*/POWERPNT.EXE", "Office*/POWERPNT.EXE"):
            if any(office.glob(pattern)):
                return True
    return False


def has_chromium() -> bool:
    root = os.environ.get("LOCALAPPDATA")
    if not root:
        return False
    browsers = Path(root) / "ms-playwright"
    if not (browsers.is_dir() and any(browsers.glob("chromium-*"))):
        return False
    # Phải hỏi đúng trình thông dịch sẽ chạy visual_review.py, không phải
    # trình thông dịch đang chạy file này — hai cái có thể khác nhau.
    try:
        proc = subprocess.run(
            [python_exe(), "-c", "import playwright"],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=120,
        )
    except (OSError, subprocess.SubprocessError):
        return False
    return proc.returncode == 0


def read_state(project: Path) -> selection.ProjectState:
    slide_paths = sorted((project / "svg_output").glob("*.svg"))
    preview_paths = sorted((project / ".preview").glob("*.png"))
    audio = sorted(path.stem for path in (project / "audio").glob("*.mp3"))
    notes = sorted(path.stem for path in (project / "notes").glob("*.md"))
    narrated = sorted((project / "exports").glob("*_narrated.pptx"), key=lambda p: p.stat().st_mtime)
    return selection.ProjectState(
        slides=[path.stem for path in slide_paths],
        audio=audio,
        narrated_pptx=str(narrated[-1]) if narrated else None,
        has_powerpoint=False,
        has_chromium=has_chromium(),
        previews=[path.stem for path in preview_paths],
        notes=notes,
        slide_mtimes=_mtimes(slide_paths),
        preview_mtimes=_mtimes(preview_paths),
    )


def _mtimes(paths: list[Path]) -> dict[str, float]:
    mtimes = {}
    for path in paths:
        try:
            mtimes[path.stem] = path.stat().st_mtime
        except OSError:
            continue
    return mtimes


def preview_server_running(project: Path) -> bool:
    return (project / "live_preview" / "lock.json").is_file() or (project / ".live_preview.lock").is_file()


def capture_previews(project: Path) -> None:
    server = str(SCRIPTS / "svg_editor" / "server.py")
    already_running = preview_server_running(project)
    if already_running:
        log("Máy chủ xem trước đang chạy sẵn, dùng lại.")
    else:
        log("Bật máy chủ xem trước để chụp ảnh slide...")
        started = subprocess.run(
            [python_exe(), server, str(project), "--daemon", "--live", "--no-browser"],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=300,
        )
        if started.returncode != 0:
            raise media.MediaError("render", "Không bật được máy chủ xem trước.", "Xem docs/vi/xu-ly-loi.md, mục Dựng video thất bại.")
    try:
        log("Chụp ảnh từng slide...")
        proc = subprocess.run(
            [python_exe(), str(SCRIPTS / "visual_review.py"), str(project)],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=1800,
        )
        for line in (proc.stderr or "").splitlines():
            log(line)
        if proc.returncode == 3:
            raise media.MediaError("chromium", "Chưa cài Chromium để chụp ảnh slide.", FIX_CHROMIUM)
        if proc.returncode != 0:
            raise media.MediaError("render", f"Chụp ảnh slide thất bại (mã {proc.returncode}).", "Xem docs/vi/xu-ly-loi.md, mục Dựng video thất bại.")
    finally:
        if not already_running:
            subprocess.run(
                [python_exe(), server, str(project), "--shutdown"],
                capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=120,
            )
            log("Đã tắt máy chủ xem trước.")


def build_subtitle(project: Path, stems: list[str], offsets: list[float], out_path: Path) -> Path:
    """Ghép phụ đề từng slide vào một file, mỗi slide dịch theo mốc của nó.

    `offsets` phải là mốc bắt đầu thật của từng slide trên đường đang dùng:
    cộng dồn thời lượng tiếng với đường FFmpeg, mốc đọc từ bản PPTX gắn tiếng
    với đường PowerPoint.
    """
    sources = []
    for stem, offset in zip(stems, offsets):
        srt_path = project / "audio" / f"{stem}.srt"
        text = srt_path.read_text(encoding="utf-8") if srt_path.is_file() else ""
        sources.append((text, offset))
    out_path.write_text(srt.merge_srt(sources), encoding="utf-8")
    return out_path


def subtitle_offsets(state: selection.ProjectState, backend: str, durations: list[float],
                     warnings: list[str]) -> tuple[list[float], float | None]:
    """Mốc bắt đầu của từng slide cho phụ đề, theo đúng đường dựng đang dùng.

    Trả về (mốc từng slide, tổng thời lượng theo bản PPTX hoặc None).
    """
    sums = srt.cumulative_offsets(durations)
    if backend != "powerpoint" or not state.narrated_pptx:
        return sums, None
    try:
        starts, timeline = pptx_timeline.narration_starts(Path(state.narrated_pptx), len(durations))
    except pptx_timeline.TimelineError as exc:
        log(f"Không đọc được mốc thời gian của bản PPTX gắn tiếng: {exc}")
        warnings.append(WARN_PPTX_TIMELINE)
        return sums, None
    return starts, timeline


def render_ffmpeg(project: Path, stems: list[str], durations: list[float], out_path: Path,
                  height: int, burn_srt: Path | None) -> None:
    entries = [(project / ".preview" / f"{stem}.png", duration) for stem, duration in zip(stems, durations)]
    audio_files = [project / "audio" / f"{stem}.mp3" for stem in stems]
    with tempfile.TemporaryDirectory(prefix="pptmaster-vi-video-") as tmp:
        images_concat = Path(tmp) / "images.txt"
        audio_concat = Path(tmp) / "audio.txt"
        images_concat.write_text(media.build_concat_text(entries), encoding="utf-8")
        audio_concat.write_text(media.build_audio_concat_text(audio_files), encoding="utf-8")
        cmd = media.build_render_command(
            images_concat, audio_concat, out_path,
            fps=30, height=height, total_seconds=sum(durations), burn_srt=burn_srt,
        )
        log("Dựng video bằng FFmpeg (có thể mất vài phút)...")
        proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=7200)
        for line in (proc.stderr or "").splitlines():
            log(line)
        if proc.returncode != 0:
            raise media.MediaError("render", f"FFmpeg lỗi (mã {proc.returncode}).", "Xem docs/vi/xu-ly-loi.md, mục Dựng video thất bại.")


def render_powerpoint(pptx: Path, out_path: Path, height: int) -> None:
    log("Xuất video bằng PowerPoint (cửa sổ PowerPoint sẽ hiện lên)...")
    cmd = [python_exe(), str(SCRIPTS / "powerpoint_video.py"), str(pptx), "-o", str(out_path), "--resolution", str(height)]
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=7200)
    for line in (proc.stderr or "").splitlines():
        log(line)
    if proc.returncode != 0:
        raise media.MediaError("powerpoint", f"PowerPoint không xuất được video (mã {proc.returncode}).", "Chạy lại với --cach ffmpeg.")


def burn_subtitles(video: Path, subtitle: Path, height: int) -> Path:
    burned = video.with_name(video.stem + "_phude.mp4")
    cmd = [
        "ffmpeg", "-y", "-hide_banner", "-loglevel", "error", "-i", str(video),
        "-vf", f"subtitles='{media.escape_subtitles_filter(subtitle)}',{media.scale_filter(height)},format=yuv420p",
        "-c:v", "libx264", "-preset", "medium", "-crf", "20", "-c:a", "copy", str(burned),
    ]
    log("In phụ đề lên video...")
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=7200)
    for line in (proc.stderr or "").splitlines():
        log(line)
    if proc.returncode != 0:
        raise media.MediaError("render", f"Không in được phụ đề (mã {proc.returncode}).", "Dùng --phu-de file để lấy phụ đề rời.")
    return burned


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Làm video bài giảng có lời giảng")
    parser.add_argument("project_path", type=Path)
    parser.add_argument("--cach", choices=["auto", "powerpoint", "ffmpeg"], default="auto")
    parser.add_argument("--phu-de", dest="phu_de", choices=["file", "hinh", "khong"], default="file")
    parser.add_argument("--do-phan-giai", dest="height", type=int, choices=[720, 1080], default=1080)
    parser.add_argument("--plan-only", action="store_true")
    args = parser.parse_args(argv)

    project = args.project_path.resolve()
    payload = {
        "ready": False, "backend": None, "video": None, "subtitle": None,
        "duration_seconds": None, "size_mb": None, "slides": 0,
        "installed": [], "warnings": [], "error": None,
    }
    if not (project / "svg_output").is_dir():
        message = f"Không thấy thư mục dự án hợp lệ: {project}"
        fix = "Kiểm tra lại đường dẫn dự án trong projects/."
        if len(str(project)) > 200:
            message += (
                f" Đường dẫn dài {len(str(project))} ký tự, trong khi Windows chỉ cho "
                "khoảng 260 ký tự."
            )
            fix = (
                "Chuyển bộ công cụ sang đường dẫn ngắn như D:\\PPTmaster rồi làm lại; "
                "nếu không phải do đường dẫn dài thì kiểm tra lại tên dự án trong projects/."
            )
        payload["error"] = {"step": "project", "message": message, "fix": fix}
        emit(payload)
        return 1

    try:
        state = read_state(project)
        payload["slides"] = len(state.slides)
        selection.check_audio(state)
        if args.cach in ("auto", "powerpoint"):
            # `--plan-only` không được gọi COM: dò PowerPoint bằng registry và
            # đường dẫn cài đặt, không mở tiến trình PowerPoint nào.
            state.has_powerpoint = has_powerpoint_installed() if args.plan_only else has_powerpoint()
        backend, warnings = selection.select_backend(state, args.cach)
        payload["backend"] = backend
        payload["warnings"] = warnings
        try:
            free_gb = shutil.disk_usage(project.anchor).free / (1024 ** 3)
        except OSError:
            free_gb = None
        if free_gb is not None and free_gb < 2:
            payload["warnings"].append(f"Ổ đĩa chỉ còn {free_gb:.1f} GB; video có thể nặng 100-300 MB.")
        if len(str(project)) > 120:
            payload["warnings"].append("Đường dẫn dự án dài; nên chuyển bộ công cụ sang D:\\PPTmaster nếu dựng video lỗi.")
        steps = selection.plan_steps(state, backend, args.phu_de)
        if args.plan_only:
            payload["steps"] = steps
            payload["ready"] = True
            emit(payload)
            return 0
        if backend == "ffmpeg" and not state.has_chromium:
            raise media.MediaError("chromium", "Chưa cài Chromium để chụp ảnh slide.", FIX_CHROMIUM)
        if shutil.which("ffprobe") is None or shutil.which("ffmpeg") is None:
            raise media.MediaError("ffmpeg", "Máy chưa có FFmpeg.", media.FIX_FFMPEG)

        stems = state.slides
        durations = [media.probe_duration(project / "audio" / f"{stem}.mp3") for stem in stems]
        exports = project / "exports"
        exports.mkdir(parents=True, exist_ok=True)
        stamp = time.strftime("%Y%m%d_%H%M%S")
        video_path = exports / f"{project.name}_video_{stamp}.mp4"

        offsets: list[float] = []
        pptx_timeline_seconds = None
        subtitle_path = None
        if args.phu_de != "khong":
            offsets, pptx_timeline_seconds = subtitle_offsets(state, backend, durations, payload["warnings"])
            subtitle_path = build_subtitle(project, stems, offsets, video_path.with_suffix(".srt"))

        if backend == "ffmpeg":
            if not selection.previews_fresh(state):
                capture_previews(project)
            burn = subtitle_path if args.phu_de == "hinh" else None
            render_ffmpeg(project, stems, durations, video_path, args.height, burn)
        else:
            render_powerpoint(Path(state.narrated_pptx), video_path, args.height)
            if args.phu_de == "hinh" and subtitle_path is not None:
                raw_path = video_path
                video_path = burn_subtitles(raw_path, subtitle_path, args.height)
                # Thầy cô đã chọn in phụ đề lên hình, nên bản chưa in chỉ là
                # file trung gian (khoảng 17 MB cho 90 giây); xoá để thư mục
                # exports không có hai video gần giống nhau.
                try:
                    raw_path.unlink()
                except OSError as exc:
                    log(f"Không xoá được video trung gian {raw_path.name}: {exc}")

        actual = media.probe_duration(video_path)
        if args.phu_de != "khong":
            # Mốc thật của từng slide trong video đã dựng: mốc dự kiến của
            # đường đang dùng, co giãn theo thời lượng video thực tế.
            planned_total = pptx_timeline_seconds if pptx_timeline_seconds else sum(durations)
            scale = actual / planned_total if planned_total > 0 else 1.0
            drift = srt.drift_warning(offsets, [offset * scale for offset in offsets])
            if drift:
                payload["warnings"].append(drift)
        payload["video"] = str(video_path)
        payload["subtitle"] = str(subtitle_path) if subtitle_path else None
        payload["duration_seconds"] = round(actual, 1)
        payload["size_mb"] = round(video_path.stat().st_size / (1024 * 1024), 1)
        payload["ready"] = True
        emit(payload)
        return 0
    except (selection.SelectionError, media.MediaError) as exc:
        payload["error"] = {"step": exc.step, "message": exc.message, "fix": exc.fix}
        emit(payload)
        return 1
    except Exception as exc:
        payload["error"] = {
            "step": "render",
            "message": f"Lỗi không lường trước: {exc}",
            "fix": "Xem docs/vi/xu-ly-loi.md, mục Dựng video thất bại.",
        }
        emit(payload)
        return 1


if __name__ == "__main__":
    sys.exit(main())
