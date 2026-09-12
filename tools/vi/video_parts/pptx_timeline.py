#!/usr/bin/env python3
"""Đọc mốc thời gian của bản PPTX đã gắn tiếng thuyết minh.

Chỉ dùng thư viện chuẩn Python (`zipfile` + `xml.etree`), không gọi tiến
trình ngoài, không mở PowerPoint.

PowerPoint không phát tiếng liên tục như FFmpeg: mỗi slide có thời gian
chuyển cảnh (`p14:dur`), một khoảng chờ trước khi tiếng bắt đầu
(`p:cond/@delay`, tức `narration_start_floor` của upstream) và thời gian ở
lại slide (`advTm`, bằng thời lượng tiếng cộng `narration_padding`). Vì vậy
phụ đề của đường PowerPoint phải theo các mốc này; cộng dồn thời lượng file
tiếng sẽ lệch dần khoảng 1 giây mỗi slide.
"""

from __future__ import annotations

import posixpath
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Sequence

_PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"


class TimelineError(Exception):
    """Không đọc được mốc thời gian từ bản PPTX đã gắn tiếng."""


def _qn(namespace: str, tag: str) -> str:
    return f"{{{namespace}}}{tag}"


def _int_attr(element: ET.Element, name: str, label: str) -> int:
    raw = element.get(name)
    try:
        return int(raw)
    except (TypeError, ValueError) as exc:
        raise TimelineError(f"{label} không phải số nguyên: {raw!r}") from exc


def slide_members(package: zipfile.ZipFile) -> list[str]:
    """Trả về danh sách file slide theo đúng thứ tự trình chiếu."""
    try:
        presentation = ET.fromstring(package.read("ppt/presentation.xml"))
        relationships = ET.fromstring(package.read("ppt/_rels/presentation.xml.rels"))
    except KeyError as exc:
        raise TimelineError(f"Bản PPTX thiếu dữ liệu thứ tự slide: {exc}") from exc
    targets = {
        relationship.get("Id"): (relationship.get("Target") or "").replace("\\", "/")
        for relationship in relationships.iter(_qn(_REL_NS, "Relationship"))
        if relationship.get("TargetMode", "Internal") != "External"
    }
    slide_list = presentation.find(_qn(_PML_NS, "sldIdLst"))
    if slide_list is None:
        raise TimelineError("Bản PPTX không có thứ tự slide.")
    members: list[str] = []
    names = set(package.namelist())
    for slide_id in slide_list.findall(_qn(_PML_NS, "sldId")):
        target = targets.get(slide_id.get(_qn(_DOC_REL_NS, "id")) or "")
        if not target:
            raise TimelineError("Thứ tự slide trỏ tới một quan hệ không có thật.")
        member = posixpath.normpath(
            target.lstrip("/") if target.startswith("/") else posixpath.join("ppt", target)
        )
        if member not in names:
            raise TimelineError(f"Bản PPTX thiếu file slide: {member}")
        members.append(member)
    if not members:
        raise TimelineError("Bản PPTX không có slide nào.")
    return members


def slide_timing(slide_xml: bytes) -> tuple[int, int, int]:
    """Trả về (chuyển cảnh, chờ trước khi đọc, thời gian ở lại) theo mili giây."""
    try:
        root = ET.fromstring(slide_xml)
    except ET.ParseError as exc:
        raise TimelineError(f"Không đọc được XML của slide: {exc}") from exc
    carriers = [
        element
        for element in root.iter(_qn(_PML_NS, "transition"))
        if element.get("advTm") is not None
    ]
    if not carriers:
        raise TimelineError("Slide chưa ghi thời gian ở lại (advTm) của bản gắn tiếng.")
    carrier = carriers[0]
    advance_ms = _int_attr(carrier, "advTm", "advTm")
    if advance_ms <= 0:
        raise TimelineError(f"advTm không hợp lệ: {advance_ms}")
    duration_attr = next(
        (name for name in (_qn(_P14_NS, "dur"), "dur") if carrier.get(name) is not None),
        None,
    )
    transition_ms = 0 if duration_attr is None else _int_attr(carrier, duration_attr, "dur")
    if transition_ms < 0:
        raise TimelineError(f"Thời gian chuyển cảnh không hợp lệ: {transition_ms}")
    delays = [
        _int_attr(condition, "delay", "delay")
        for audio in root.iter(_qn(_PML_NS, "audio"))
        for condition in audio.iter(_qn(_PML_NS, "cond"))
        if condition.get("delay") is not None
    ]
    delay_ms = max(delays) if delays else 0
    if delay_ms < 0 or delay_ms >= advance_ms:
        raise TimelineError(f"Khoảng chờ trước khi đọc không hợp lệ: {delay_ms}")
    return transition_ms, delay_ms, advance_ms


def starts_from_timings(timings: Sequence[tuple[int, int, int]]) -> tuple[list[float], float]:
    """Đổi các mốc mili giây của từng slide thành mốc bắt đầu tiếng (giây)."""
    starts: list[float] = []
    timeline_ms = 0
    for transition_ms, delay_ms, advance_ms in timings:
        starts.append((timeline_ms + transition_ms + delay_ms) / 1000)
        timeline_ms += transition_ms + advance_ms
    return starts, timeline_ms / 1000


def narration_starts(pptx_path: Path, slide_count: int) -> tuple[list[float], float]:
    """Đọc mốc bắt đầu tiếng của từng slide và tổng thời lượng theo bản PPTX."""
    try:
        with zipfile.ZipFile(pptx_path) as package:
            members = slide_members(package)
            if len(members) != slide_count:
                raise TimelineError(
                    f"Bản PPTX có {len(members)} slide, dự án có {slide_count} slide."
                )
            timings = [slide_timing(package.read(member)) for member in members]
    except (OSError, zipfile.BadZipFile) as exc:
        raise TimelineError(f"Không mở được bản PPTX đã gắn tiếng: {exc}") from exc
    return starts_from_timings(timings)
