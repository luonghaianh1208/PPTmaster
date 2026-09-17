#!/usr/bin/env python3
"""Kiểm hiệu ứng của file PPTX so với mức thầy cô đã chọn.

Cách dùng:
    python tools/vi/kiem_hieu_ung.py <file.pptx> --muc khong|vua|nhieu
    python tools/vi/kiem_hieu_ung.py <file_narrated.pptx> [--muc khong|vua|nhieu] --video

stdout: đúng một dòng JSON. Mã thoát: 0 khi đạt, 1 khi không đạt hoặc lỗi.
Định nghĩa từng mức: docs/vi/tro-ly/hieu-ung-lop-hoc.md.
"""

from __future__ import annotations

import argparse
import json
import math
import posixpath
import re
import sys
import unicodedata
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
P = "{%s}" % NS["p"]
MC = "{%s}" % NS["mc"]
OBJECT_CLASSES = ("entr", "emph", "exit", "path")
LEVELS = ("khong", "vua", "nhieu")
LEVEL_NAMES = {"khong": "không", "vua": "vừa", "nhieu": "nhiều"}
PLAIN_TRANSITIONS = ("none", "fade")
MIN_SHARE = {"vua": 0.30, "nhieu": 0.50}
MAX_CLICKS = {"vua": 8, "nhieu": 12}
BOUNDARY_DECK_SIZE = 8
LONG_EFFECT_MS = 2000
OPTION_RE = re.compile(r"^\s*([A-D])[.)]\s")
GUIDE = "docs/vi/tro-ly/hieu-ung-lop-hoc.md"
FIX_ARGS = "Chạy: python tools/vi/kiem_hieu_ung.py <file.pptx> --muc khong|vua|nhieu, hoặc thêm --video cho bài làm video"
FIX_PARSE = "Xuất lại file PPTX bằng svg_to_pptx.py rồi chạy lại; không sửa tay file PPTX."
FIX_INTERNAL = "Gửi nguyên dòng error.message cho người bảo trì."


class ArgumentError(Exception):
    """Tham số dòng lệnh sai; báo bằng JSON thay vì để argparse tự thoát."""


class JsonArgumentParser(argparse.ArgumentParser):
    def error(self, message):
        raise ArgumentError(message)


class DeckError(Exception):
    """File không phải PPTX đọc được."""


def emit(payload: dict) -> None:
    """In đúng một dòng JSON và không bao giờ ném lỗi."""
    text = json.dumps(payload, ensure_ascii=False) + "\n"
    try:
        try:
            sys.stdout.write(text)
        except UnicodeEncodeError:
            buffer = getattr(sys.stdout, "buffer", None)
            if buffer is not None:
                buffer.write(text.encode("utf-8", errors="replace"))
            else:
                sys.stdout.write(json.dumps(payload, ensure_ascii=True) + "\n")
        sys.stdout.flush()
    except OSError:
        pass


def configure_streams() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


def result(*, ready: bool, stats: dict | None = None, warnings=(), error: dict | None = None) -> dict:
    payload = {
        "ready": ready,
        "slides": 0,
        "content_slides": 0,
        "slides_with_effects": 0,
        "max_click_steps": 0,
        "max_click_slide": None,
        "morph_slides": 0,
        "flip_cards": 0,
        "transitions": {},
    }
    payload.update(stats or {})
    payload["warnings"] = list(warnings)
    payload["error"] = error
    return payload


def failure(step: str, message: str, fix: str, **rest) -> dict:
    return result(ready=False, error={"step": step, "message": message, "fix": fix}, **rest)


def read_xml(archive: zipfile.ZipFile, name: str) -> ET.Element:
    try:
        return ET.fromstring(archive.read(name))
    except KeyError as exc:
        raise DeckError(f"Thiếu phần {name} trong file PPTX") from exc
    except ET.ParseError as exc:
        raise DeckError(f"Phần {name} hỏng: {exc}") from exc


def slide_parts(archive: zipfile.ZipFile) -> list[str]:
    """Đường dẫn các slide theo đúng thứ tự trình chiếu."""
    presentation = read_xml(archive, "ppt/presentation.xml")
    rels = read_xml(archive, "ppt/_rels/presentation.xml.rels")
    targets = {rel.get("Id"): rel.get("Target", "") for rel in rels.findall("rel:Relationship", NS)}
    parts = []
    for slide_id in presentation.findall("p:sldIdLst/p:sldId", NS):
        target = targets.get(slide_id.get("{%s}id" % NS["r"]))
        if not target:
            raise DeckError("presentation.xml trỏ tới một slide không có trong file")
        # Target tương đối tính từ thư mục ppt/, target tuyệt đối tính từ gốc gói.
        parts.append(target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("ppt", target)))
    return parts


def transition_of(slide: ET.Element) -> ET.Element | None:
    for child in slide:
        if child.tag == P + "transition":
            return child
        if child.tag == MC + "AlternateContent":
            chosen = child.find("mc:Choice/p:transition", NS)
            if chosen is not None:
                return chosen
    return None


def transition_name(slide: ET.Element) -> str:
    transition = transition_of(slide)
    if transition is None:
        return "none"
    for child in transition:
        name = child.tag.rsplit("}", 1)[-1]
        if name not in ("sndAc", "extLst"):
            return name
    return "none"


def timing_of(slide: ET.Element) -> ET.Element | None:
    for child in slide:
        if child.tag == P + "timing":
            return child
        if child.tag == MC + "AlternateContent":
            chosen = child.find("mc:Choice/p:timing", NS)
            if chosen is not None:
                return chosen
    return None


def effect_nodes(root: ET.Element) -> list[ET.Element]:
    return [node for node in root.iter(P + "cTn") if node.get("presetClass") in OBJECT_CLASSES]


def longest_behaviour_ms(effect: ET.Element) -> int:
    longest = 0
    for node in effect.iter(P + "cTn"):
        if node is not effect and (node.get("dur") or "").isdigit():
            longest = max(longest, int(node.get("dur")))
    return longest


def slide_text_lines(slide: ET.Element) -> list[str]:
    return [unicodedata.normalize("NFC", "".join(t.text or "" for t in paragraph.iter("{%s}t" % NS["a"])))
            for paragraph in slide.iter("{%s}p" % NS["a"])]


def is_quiz(lines: list[str]) -> bool:
    """Câu trắc nghiệm: có câu hỏi và ít nhất ba lựa chọn A. B. C. D. (mục lục A. B. C. không tính)."""
    options = {match.group(1) for match in map(OPTION_RE.match, lines) if match}
    asks = any("?" in line or line.lstrip().lower().startswith("câu") for line in lines)
    return len(options) >= 3 and asks


def inspect_slide(slide: ET.Element) -> dict:
    timing = timing_of(slide)
    main_effects, click_steps, flip_cards, long_effects = [], 0, 0, 0
    if timing is not None:
        for sequence in timing.iter(P + "cTn"):
            node_type = sequence.get("nodeType")
            if node_type == "mainSeq":
                main_effects = effect_nodes(sequence)
                click_steps = sum(1 for node in main_effects if node.get("nodeType") == "clickEffect")
            elif node_type == "interactiveSeq" and effect_nodes(sequence):
                flip_cards += 1
        long_effects = sum(1 for node in effect_nodes(timing) if longest_behaviour_ms(node) > LONG_EFFECT_MS)
    lines = slide_text_lines(slide)
    return {
        "transition": transition_name(slide),
        "effects": len(main_effects) + flip_cards,
        "click_steps": click_steps,
        "flip_cards": flip_cards,
        "long_effects": long_effects,
        "quiz": is_quiz(lines),
        "mentions_answer": "đáp án" in " ".join(lines).lower(),
    }


def read_deck(path: Path) -> list[dict]:
    """Các slide đang hiện, theo thứ tự trình chiếu; slide ẩn không tính."""
    try:
        with zipfile.ZipFile(path) as archive:
            slides = [read_xml(archive, name) for name in slide_parts(archive)]
            return [inspect_slide(slide) for slide in slides if slide.get("show") != "0"]
    except zipfile.BadZipFile as exc:
        raise DeckError(f"Không phải file PPTX hợp lệ: {exc}") from exc


def summarise(slides: list[dict]) -> dict:
    content = slides[1:-1] if len(slides) >= 3 else slides
    transitions: dict[str, int] = {}
    for slide in slides:
        transitions[slide["transition"]] = transitions.get(slide["transition"], 0) + 1
    busiest = max(range(len(slides)), key=lambda i: slides[i]["click_steps"], default=None)
    max_clicks = slides[busiest]["click_steps"] if busiest is not None else 0
    return {
        "slides": len(slides),
        "content_slides": len(content),
        "slides_with_effects": sum(1 for s in content if s["effects"] or s["transition"] == "morph"),
        "max_click_steps": max_clicks,
        "max_click_slide": busiest + 1 if max_clicks else None,
        "morph_slides": transitions.get("morph", 0),
        "flip_cards": sum(s["flip_cards"] for s in slides),
        "transitions": transitions,
    }


def level_problems(level: str, slides: list[dict], stats: dict) -> tuple[list[str], list[str]]:
    errors, warnings = [], []
    if level == "khong":
        moving = [i + 1 for i, s in enumerate(slides) if s["effects"]]
        fancy = [i + 1 for i, s in enumerate(slides) if s["transition"] not in PLAIN_TRANSITIONS]
        if moving:
            errors.append(f"Mức không: trang {', '.join(map(str, moving))} còn hiệu ứng đối tượng.")
        if fancy:
            errors.append(f"Mức không: trang {', '.join(map(str, fancy))} dùng chuyển trang khác mờ dần.")
        return errors, warnings

    needed = math.ceil(MIN_SHARE[level] * stats["content_slides"])
    if stats["slides_with_effects"] < needed:
        errors.append(
            f"Mức {LEVEL_NAMES[level]}: {stats['slides_with_effects']}/{stats['content_slides']} trang nội dung có hiệu ứng, "
            f"cần ít nhất {needed} ({int(MIN_SHARE[level] * 100)}%)."
        )
    if level == "nhieu" and stats["morph_slides"] == 0:
        warnings.append("Mức nhiều: chưa có trang nào dùng Morph cho diễn biến.")
    if stats["max_click_steps"] > MAX_CLICKS[level]:
        warnings.append(
            f"Trang {stats['max_click_slide']} có {stats['max_click_steps']} bước bấm (quá {MAX_CLICKS[level]}); "
            "nên gộp bớt hoặc tách trang."
        )
    if level == "vua" and stats["slides"] >= BOUNDARY_DECK_SIZE and not any(
        s["transition"] not in PLAIN_TRANSITIONS + ("morph",) for s in slides
    ):
        warnings.append("Bài từ 8 trang trở lên chưa có chuyển trang nổi bật ở ranh giới hoạt động.")
    if level == "nhieu" and stats["flip_cards"] == 0:
        warnings.append("Mức nhiều: chưa có ô bấm hiện đáp án nào.")
    return errors, warnings


def video_problems(slides: list[dict]) -> list[str]:
    clicks = [i + 1 for i, s in enumerate(slides) if s["click_steps"]]
    flips = [i + 1 for i, s in enumerate(slides) if s["flip_cards"]]
    errors = []
    if clicks:
        errors.append(f"Video: trang {', '.join(map(str, clicks))} còn hiệu ứng chờ bấm.")
    if flips:
        errors.append(f"Video: trang {', '.join(map(str, flips))} còn ô bấm hiện đáp án.")
    return errors


def common_warnings(slides: list[dict]) -> list[str]:
    warnings = []
    for index, slide in enumerate(slides):
        next_slide = slides[index + 1] if index + 1 < len(slides) else None
        if (slide["quiz"] and not (slide["flip_cards"] or slide["click_steps"])
                and not (next_slide and next_slide["mentions_answer"])):
            warnings.append(f"Trang {index + 1} có câu trắc nghiệm nhưng chưa có cách hiện đáp án.")
        if slide["long_effects"]:
            warnings.append(f"Trang {index + 1} có hiệu ứng dài quá 2 giây.")
    return warnings


def run(args) -> int:
    level = args.muc
    if level is None and not args.video:
        emit(failure("input", "Thiếu --muc (chỉ được bỏ khi kiểm bài làm video bằng --video)", FIX_ARGS))
        return 1
    if level is not None and level not in LEVELS:
        emit(failure("input", f"--muc chỉ nhận khong, vua hoặc nhieu; gặp {level!r}", FIX_ARGS))
        return 1
    path = args.file.expanduser()
    if not path.is_file():
        emit(failure("input", f"Không có file: {path}", FIX_ARGS))
        return 1
    try:
        slides = read_deck(path)
    except DeckError as exc:
        emit(failure("parse", str(exc), FIX_PARSE))
        return 1
    if not slides:
        emit(failure("parse", "File PPTX không có slide nào", FIX_PARSE))
        return 1

    stats = summarise(slides)
    errors, warnings = level_problems(level, slides, stats) if level else ([], [])
    warnings += common_warnings(slides)
    if errors:
        emit(failure("muc", " ".join(errors),
                     f"Chỉnh animations.json theo {GUIDE} (mục Mức {LEVEL_NAMES[level]}), xuất lại rồi chạy lại một lần.",
                     stats=stats, warnings=warnings))
        return 1
    if args.video:
        errors = video_problems(slides)
        if errors:
            emit(failure("video", " ".join(errors),
                         f"Sửa animations_video.json theo {GUIDE} (mục Video): đổi sang after-previous hoặc "
                         "with-previous, bỏ trigger_shape; xuất lại bản thuyết minh rồi chạy lại một lần.",
                         stats=stats, warnings=warnings))
            return 1
    emit(result(ready=True, stats=stats, warnings=warnings))
    return 0


def main(argv: list[str] | None = None) -> int:
    configure_streams()
    parser = JsonArgumentParser(description="Kiểm hiệu ứng của file PPTX theo mức đã chọn")
    parser.add_argument("file", type=Path, help="File PPTX đã xuất")
    parser.add_argument("--muc", help="khong, vua hoặc nhieu; có thể bỏ khi dùng --video")
    parser.add_argument("--video", action="store_true", help="Áp thêm quy tắc cho bài làm video")
    try:
        args = parser.parse_args(argv)
    except ArgumentError as exc:
        emit(failure("input", f"Tham số không hợp lệ: {exc}", FIX_ARGS))
        return 1
    try:
        return run(args)
    except Exception as exc:  # noqa: BLE001 - stdout không bao giờ được để trống
        emit(failure("internal", f"Lỗi ngoài dự kiến: {exc}", FIX_INTERNAL))
        return 1


if __name__ == "__main__":
    sys.exit(main())
