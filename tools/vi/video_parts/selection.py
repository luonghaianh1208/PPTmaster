#!/usr/bin/env python3
"""Chọn cách dựng video và liệt kê các bước sẽ chạy."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

FIX_AUDIO = (
    "Chạy bước thuyết minh của dự án gốc trước: "
    "python skills/ppt-master/scripts/notes_to_audio.py <đường_dẫn_dự_án> --voice vi-VN-HoaiMyNeural"
)
FIX_NOTES = (
    "Chạy bước ghi chú của dự án gốc trước (mỗi slide một file notes/<tên_slide>.md), "
    "rồi mới chạy bước thuyết minh: "
    "python skills/ppt-master/scripts/notes_to_audio.py <đường_dẫn_dự_án> --voice vi-VN-HoaiMyNeural"
)
FIX_POWERPOINT = "Máy không có PowerPoint. Chạy lại với --cach ffmpeg."
FIX_NARRATED = (
    "Chưa có bản PPTX đã gắn tiếng. Chạy lại bước xuất PPTX có thuyết minh của dự án gốc, "
    "hoặc chạy lại với --cach ffmpeg."
)


class SelectionError(Exception):
    def __init__(self, step: str, message: str, fix: str) -> None:
        super().__init__(message)
        self.step = step
        self.message = message
        self.fix = fix


@dataclass
class ProjectState:
    slides: list[str]
    audio: list[str]
    narrated_pptx: Optional[str] = None
    has_powerpoint: bool = False
    has_chromium: bool = False
    previews: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    slide_mtimes: dict[str, float] = field(default_factory=dict)
    preview_mtimes: dict[str, float] = field(default_factory=dict)


def check_audio(state: ProjectState) -> None:
    if not state.slides:
        raise SelectionError("project", "Dự án chưa có slide nào trong svg_output/.", "Tạo slide trước khi làm video.")
    if not state.audio:
        # Thiếu cả ghi chú thì bước thuyết minh cũng không chạy được, nên phải
        # gọi tên bước ghi chú trước.
        if not state.notes:
            raise SelectionError(
                "audio",
                "Dự án chưa có ghi chú lời giảng trong notes/ nên chưa tạo được tiếng thuyết minh.",
                FIX_NOTES,
            )
        raise SelectionError("audio", "Dự án chưa có file tiếng thuyết minh nào trong audio/.", FIX_AUDIO)
    missing = [stem for stem in state.slides if stem not in set(state.audio)]
    if missing:
        missing_notes = [stem for stem in missing if stem not in set(state.notes)]
        if missing_notes:
            raise SelectionError(
                "audio",
                "Thiếu ghi chú lời giảng cho: " + ", ".join(missing_notes),
                FIX_NOTES,
            )
        raise SelectionError("audio", "Thiếu tiếng thuyết minh cho: " + ", ".join(missing), FIX_AUDIO)


def select_backend(state: ProjectState, requested: str) -> tuple[str, list[str]]:
    check_audio(state)
    warnings: list[str] = []
    if requested == "powerpoint":
        if not state.has_powerpoint:
            raise SelectionError("powerpoint", "Máy không có PowerPoint để xuất video.", FIX_POWERPOINT)
        if not state.narrated_pptx:
            raise SelectionError("narrated_pptx", "Chưa có bản PPTX đã gắn tiếng thuyết minh.", FIX_NARRATED)
        return "powerpoint", warnings
    if requested == "ffmpeg":
        return "ffmpeg", warnings
    if state.has_powerpoint and state.narrated_pptx:
        return "powerpoint", warnings
    if not state.has_powerpoint:
        warnings.append("Máy không có PowerPoint nên em dựng video bằng FFmpeg.")
    else:
        warnings.append("Chưa có bản PPTX gắn tiếng nên em dựng video bằng FFmpeg.")
    return "ffmpeg", warnings


def previews_fresh(state: ProjectState) -> bool:
    """Ảnh chụp slide còn dùng được không.

    Đủ tên là chưa đủ: thầy cô sửa slide rồi dựng lại thì ảnh cũ vẫn còn đó
    và video sẽ là bài giảng cũ. Ảnh nào cũ hơn file SVG của nó thì phải
    chụp lại; không có mốc thời gian để so thì coi như cũ.
    """
    if sorted(state.previews) != sorted(state.slides):
        return False
    for stem in state.slides:
        svg_mtime = state.slide_mtimes.get(stem)
        png_mtime = state.preview_mtimes.get(stem)
        if svg_mtime is None or png_mtime is None or svg_mtime > png_mtime:
            return False
    return True


def plan_steps(state: ProjectState, backend: str, subtitle_mode: str) -> list[dict]:
    steps: list[dict] = []
    if backend == "ffmpeg":
        needs_capture = not previews_fresh(state)
        if needs_capture and not state.has_chromium:
            steps.append({"step": "chromium", "action": "require", "method": "pip+playwright"})
        if needs_capture:
            steps.append({"step": "preview", "action": "capture", "method": "visual_review.py"})
        if subtitle_mode != "khong":
            steps.append({"step": "subtitle", "action": "merge", "method": "srt"})
        steps.append({"step": "render", "action": "run", "method": "ffmpeg"})
        return steps
    if subtitle_mode != "khong":
        steps.append({"step": "subtitle", "action": "merge", "method": "srt"})
    steps.append({"step": "render", "action": "run", "method": "powerpoint_video.py"})
    if subtitle_mode == "hinh":
        steps.append({"step": "burn", "action": "run", "method": "ffmpeg"})
    return steps
