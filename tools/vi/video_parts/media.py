#!/usr/bin/env python3
"""Đo thời lượng và dựng câu lệnh FFmpeg cho video bài giảng.

Chỉ dùng thư viện chuẩn Python. Không tự chạy FFmpeg ở mức module này.
"""

from __future__ import annotations

import subprocess
from pathlib import Path
from typing import Callable, Optional, Sequence

FIX_FFMPEG = 'Cài FFmpeg bằng: powershell -NoProfile -ExecutionPolicy Bypass -File tools\\vi\\pptmaster.ps1 -Action tool -Name ffmpeg'


class MediaError(Exception):
    def __init__(self, step: str, message: str, fix: str) -> None:
        super().__init__(message)
        self.step = step
        self.message = message
        self.fix = fix


def probe_duration(path: Path, run: Callable = subprocess.run) -> float:
    cmd = [
        "ffprobe", "-v", "error",
        "-show_entries", "format=duration",
        "-of", "default=noprint_wrappers=1:nokey=1",
        str(path),
    ]
    try:
        proc = run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=120)
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise MediaError("ffmpeg", f"Không chạy được ffprobe: {exc}", FIX_FFMPEG) from exc
    if proc.returncode != 0:
        raise MediaError("ffmpeg", f"ffprobe lỗi khi đọc {path.name} (mã {proc.returncode}).", FIX_FFMPEG)
    try:
        return float((proc.stdout or "").strip())
    except ValueError as exc:
        raise MediaError("ffmpeg", f"Không đọc được thời lượng của {path.name}.", FIX_FFMPEG) from exc


def _concat_path(path: Path) -> str:
    return str(path).replace("\\", "/").replace("'", "'\\''")


def build_concat_text(entries: Sequence[tuple[Path, float]]) -> str:
    lines: list[str] = []
    for path, duration in entries:
        lines.append(f"file '{_concat_path(path)}'")
        lines.append(f"duration {duration:.3f}")
    if entries:
        lines.append(f"file '{_concat_path(entries[-1][0])}'")
    return "\n".join(lines) + ("\n" if lines else "")


def build_audio_concat_text(paths: Sequence[Path]) -> str:
    lines = [f"file '{_concat_path(path)}'" for path in paths]
    return "\n".join(lines) + ("\n" if lines else "")


def escape_subtitles_filter(path: Path) -> str:
    text = str(path).replace("\\", "\\\\").replace(":", "\\:").replace("'", "\\'")
    return text


def build_render_command(
    images_concat: Path,
    audio_concat: Path,
    out_path: Path,
    fps: int,
    height: int,
    burn_srt: Optional[Path] = None,
) -> list[str]:
    video_filter = f"scale=-2:{height},format=yuv420p"
    if burn_srt is not None:
        video_filter = f"subtitles='{escape_subtitles_filter(burn_srt)}',{video_filter}"
    return [
        "ffmpeg", "-y", "-hide_banner", "-loglevel", "error",
        "-f", "concat", "-safe", "0", "-i", str(images_concat),
        "-f", "concat", "-safe", "0", "-i", str(audio_concat),
        "-vf", video_filter,
        "-r", str(fps),
        "-c:v", "libx264", "-preset", "medium", "-crf", "20",
        "-c:a", "aac", "-b:a", "160k",
        "-movflags", "+faststart",
        "-shortest",
        str(out_path),
    ]
