#!/usr/bin/env python3
"""Đọc, dịch mốc thời gian và ghép phụ đề SRT.

Chỉ dùng thư viện chuẩn Python.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Sequence

TIMESTAMP_RE = re.compile(r"(\d{2}):(\d{2}):(\d{2}),(\d{3})")
ARROW_RE = re.compile(r"(\d{2}:\d{2}:\d{2},\d{3})\s*-->\s*(\d{2}:\d{2}:\d{2},\d{3})")


@dataclass
class Cue:
    index: int
    start: float
    end: float
    text: str


def format_timestamp(seconds: float) -> str:
    if seconds < 0:
        seconds = 0.0
    total_ms = int(round(seconds * 1000))
    hours, rest = divmod(total_ms, 3_600_000)
    minutes, rest = divmod(rest, 60_000)
    secs, millis = divmod(rest, 1000)
    return f"{hours:02d}:{minutes:02d}:{secs:02d},{millis:03d}"


def parse_timestamp(text: str) -> float:
    match = TIMESTAMP_RE.fullmatch(text.strip())
    if match is None:
        raise ValueError(f"Mốc thời gian SRT không hợp lệ: {text!r}")
    hours, minutes, seconds, millis = (int(part) for part in match.groups())
    return hours * 3600 + minutes * 60 + seconds + millis / 1000


def parse_srt(text: str) -> list[Cue]:
    cleaned = text.lstrip("﻿").replace("\r\n", "\n").replace("\r", "\n")
    cues: list[Cue] = []
    for block in cleaned.split("\n\n"):
        lines = [line for line in block.split("\n") if line.strip()]
        if not lines:
            continue
        arrow_at = next((i for i, line in enumerate(lines) if ARROW_RE.search(line)), None)
        if arrow_at is None:
            continue
        match = ARROW_RE.search(lines[arrow_at])
        index = len(cues) + 1
        if arrow_at > 0:
            try:
                index = int(lines[arrow_at - 1].strip())
            except ValueError:
                index = len(cues) + 1
        cues.append(
            Cue(
                index=index,
                start=parse_timestamp(match.group(1)),
                end=parse_timestamp(match.group(2)),
                text="\n".join(lines[arrow_at + 1:]).strip(),
            )
        )
    return cues


def render_srt(cues: Sequence[Cue]) -> str:
    blocks = []
    for number, cue in enumerate(cues, start=1):
        blocks.append(
            f"{number}\n"
            f"{format_timestamp(cue.start)} --> {format_timestamp(cue.end)}\n"
            f"{cue.text}\n"
        )
    return "\n".join(blocks) + ("\n" if blocks else "")


def merge_srt(sources: Sequence[tuple[str, float]]) -> str:
    merged: list[Cue] = []
    for text, offset in sources:
        for cue in parse_srt(text):
            merged.append(Cue(index=0, start=cue.start + offset, end=cue.end + offset, text=cue.text))
    return render_srt(merged)
