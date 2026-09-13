"""Đọc file nguồn giao-an.md thành cấu trúc, và kiểm cấu trúc theo Công văn 5512.

Ngữ pháp của giao-an.md nằm trong docs/vi/tro-ly/giao-an.md.
Phần kiểm mã năng lực nằm ở frameworks.py, không ở đây.
Chỉ dùng thư viện chuẩn Python.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from .frameworks import NO_FRAMEWORK, TRACK_ALIASES, normalise

OBJECTIVES = "## MUC TIEU"
EQUIPMENT = "## THIET BI"
PROCESS = "## TIEN TRINH"
WORKSHEETS = "## PHIEU HOC TAP"
RUBRIC = "## RUBRIC"
REVIEW = "## CAN SOAT"
SECTIONS = (OBJECTIVES, EQUIPMENT, PROCESS, WORKSHEETS, RUBRIC, REVIEW)
REQUIRED_SECTIONS = (OBJECTIVES, EQUIPMENT, PROCESS, RUBRIC)
OBJECTIVE_PARTS = (
    "Kien thuc",
    "Nang luc chung",
    "Nang luc dac thu",
    "Nang luc so",
    "Nang luc AI",
    "Pham chat",
)
EQUIPMENT_PARTS = ("Giao vien", "Hoc sinh")
ACTIVITY_KEYS = (
    "thoi-luong",
    "muc-tieu",
    "noi-dung",
    "san-pham",
    "chuyen-giao",
    "thuc-hien",
    "bao-cao",
    "ket-luan",
)
ACTIVITY_OPTIONAL = ("nls", "ai")
LEVEL_PREFIXES = ("Mức 1:", "Mức 2:", "Mức 3:")
META_REQUIRED = ("school", "subject", "grade", "lesson", "periods")
META_OPTIONAL = ("department", "teacher", "track", "week", "period_numbers", "school_year")
MINUTES_PER_PERIOD = 45

_CODE_LINE_RE = re.compile(r"^(\S+)\s+—\s+")
_COMPETENCE_LINE_RE = re.compile(r"^\S+\s+—\s+\S")
_WORKSHEET_MENTION_RE = re.compile(r"(?:phiếu học tập|pht)\s*(?:số\s*)?(\d+)", re.IGNORECASE)
_NUMBER_RE = re.compile(r"(\d+)")


class ParseError(Exception):
    """Lỗi cấu trúc trong giao-an.md; luôn kèm số dòng để AI sửa đúng chỗ."""

    def __init__(self, line_no: int, message: str) -> None:
        super().__init__(f"Dòng {line_no}: {message}")
        self.line_no = line_no
        self.message = message


@dataclass
class Block:
    section: str
    title: str
    line_no: int
    body: list[tuple[int, str]] = field(default_factory=list)


@dataclass
class Activity:
    title: str
    line_no: int
    minutes: int
    muc_tieu: list[str]
    noi_dung: list[str]
    san_pham: list[str]
    chuyen_giao: list[str]
    thuc_hien: list[str]
    bao_cao: list[str]
    ket_luan: list[str]
    nls: list[str]
    ai: list[str]


@dataclass
class Worksheet:
    title: str
    items: list[str]


@dataclass
class Rubric:
    title: str
    levels: list[str]


@dataclass
class Lesson:
    meta: dict[str, str]
    knowledge: list[str]
    general: list[str]
    specific: list[str]
    nls_lines: list[str]
    ai_lines: list[str]
    qualities: list[str]
    teacher_tools: list[str]
    student_tools: list[str]
    activities: list[Activity]
    worksheets: list[Worksheet]
    rubrics: list[Rubric]
    review_notes: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    ai_no_framework: bool = False

    @property
    def nls_codes(self) -> list[str]:
        return _codes_from_lines(self.nls_lines)

    @property
    def ai_codes(self) -> list[str]:
        if self.ai_no_framework:
            return []
        return _codes_from_lines(self.ai_lines)

    @property
    def minutes(self) -> int:
        return sum(activity.minutes for activity in self.activities)

    @property
    def periods(self) -> int:
        return int(self.meta["periods"])


def _codes_from_lines(lines: list[str]) -> list[str]:
    codes: list[str] = []
    for line in lines:
        match = _CODE_LINE_RE.match(line)
        if match and match.group(1) not in codes:
            codes.append(match.group(1))
    return codes


def split_codes(values: list[str]) -> list[str]:
    codes: list[str] = []
    for value in values:
        pieces = value.split(",") if "," in value and not value.startswith("(") else [value]
        for piece in pieces:
            code = piece.strip()
            if code and code not in codes:
                codes.append(code)
    return codes


def parse_meta(lines: list[str]) -> tuple[dict[str, str], int]:
    if not lines or lines[0].strip() != "---":
        raise ParseError(1, "file phải mở đầu bằng một dòng '---' rồi tới khối thông tin bài dạy")
    allowed = META_REQUIRED + META_OPTIONAL
    meta: dict[str, str] = {}
    key_lines: dict[str, int] = {}
    for index in range(1, len(lines)):
        line_no = index + 1
        stripped = lines[index].strip()
        if stripped == "---":
            missing = [key for key in META_REQUIRED if not meta.get(key)]
            if missing:
                raise ParseError(line_no, "khối thông tin bài dạy thiếu: " + ", ".join(missing))
            if not meta["grade"].isdigit():
                raise ParseError(key_lines["grade"], f"'grade' phải là chữ số, gặp {meta['grade']!r}")
            if not meta["periods"].isdigit() or int(meta["periods"]) <= 0:
                raise ParseError(key_lines["periods"], f"'periods' phải là số tiết lớn hơn 0, gặp {meta['periods']!r}")
            if meta.get("track") and normalise(meta["track"]) not in TRACK_ALIASES:
                raise ParseError(
                    key_lines["track"],
                    f"'track' chỉ nhận để trống hoặc giá trị nghĩa là hệ chuyên (ví dụ 'chuyen'), "
                    f"gặp {meta['track']!r}",
                )
            return meta, index + 1
        if not stripped:
            continue
        key, separator, value = stripped.partition(":")
        if not separator:
            raise ParseError(line_no, f"dòng phải dạng 'khoá: giá trị', gặp {stripped!r}")
        key = key.strip()
        if key not in allowed:
            raise ParseError(
                line_no,
                f"khoá {key!r} không dùng trong khối thông tin bài dạy; các khoá hợp lệ: "
                + ", ".join(allowed),
            )
        meta[key] = value.strip()
        key_lines[key] = line_no
    raise ParseError(len(lines), "khối thông tin bài dạy chưa được đóng bằng dòng '---'")


def split_blocks(lines: list[str], start: int) -> tuple[list[Block], list[str], dict[str, int]]:
    blocks: list[Block] = []
    review_notes: list[str] = []
    section_lines: dict[str, int] = {}
    seen: list[str] = []
    section = ""
    for index in range(start, len(lines)):
        line_no = index + 1
        stripped = lines[index].strip()
        if not stripped:
            continue
        if stripped.startswith("## "):
            if stripped not in SECTIONS:
                raise ParseError(
                    line_no, f"mục {stripped!r} không hợp lệ; chỉ dùng: " + ", ".join(SECTIONS)
                )
            if stripped in seen:
                raise ParseError(line_no, f"mục {stripped!r} xuất hiện hai lần")
            if seen and SECTIONS.index(stripped) < SECTIONS.index(seen[-1]):
                raise ParseError(line_no, "các mục phải theo thứ tự: " + ", ".join(SECTIONS))
            seen.append(stripped)
            section = stripped
            section_lines[stripped] = line_no
            continue
        if not section:
            raise ParseError(
                line_no, f"dòng {stripped!r} nằm ngoài mục nào; mục đầu tiên phải là {OBJECTIVES!r}"
            )
        if section == REVIEW:
            if not stripped.startswith("- "):
                raise ParseError(
                    line_no, f"dòng trong {REVIEW!r} phải bắt đầu bằng '- ', gặp {stripped!r}"
                )
            review_notes.append(stripped[2:].strip())
            continue
        if stripped.startswith("### "):
            blocks.append(Block(section=section, title=stripped[4:].strip(), line_no=line_no))
            continue
        if not blocks or blocks[-1].section != section:
            raise ParseError(
                line_no,
                f"dòng {stripped!r} nằm ngoài mục con nào; mỗi mục con mở đầu bằng '### '",
            )
        blocks[-1].body.append((line_no, stripped))
    missing = [name for name in REQUIRED_SECTIONS if name not in seen]
    if missing:
        raise ParseError(max(len(lines), 1), "giáo án thiếu mục: " + ", ".join(missing))
    return blocks, review_notes, section_lines


def bullets(block: Block, *, required: bool = True) -> list[str]:
    items: list[str] = []
    for line_no, text in block.body:
        if not text.startswith("- "):
            raise ParseError(
                line_no, f"dòng trong mục {block.title!r} phải bắt đầu bằng '- ', gặp {text!r}"
            )
        items.append(text[2:].strip())
    if required and not items:
        raise ParseError(block.line_no, f"mục {block.title!r} không có dòng nào")
    return items


def _check_competence_lines(block: Block, items: list[str]) -> None:
    for index, item in enumerate(items):
        if not _COMPETENCE_LINE_RE.match(item):
            raise ParseError(
                block.body[index][0],
                f"dòng trong mục {block.title!r} phải theo khuôn 'mã — mô tả', dùng dấu gạch dài "
                f"'—' (không dùng '-', '–' hay ':'), gặp {item!r}",
            )


def build_activity(block: Block) -> Activity:
    values: dict[str, list[str]] = {}
    first_line: dict[str, int] = {}
    for line_no, text in block.body:
        key, separator, value = text.partition(":")
        if not separator:
            raise ParseError(line_no, f"dòng phải dạng 'khoá: giá trị', gặp {text!r}")
        key = key.strip()
        if key not in ACTIVITY_KEYS + ACTIVITY_OPTIONAL:
            raise ParseError(
                line_no,
                f"khoá {key!r} không dùng được trong hoạt động; các khoá hợp lệ: "
                + ", ".join(ACTIVITY_KEYS + ACTIVITY_OPTIONAL),
            )
        if not value.strip():
            raise ParseError(line_no, f"khoá {key!r} không có nội dung")
        values.setdefault(key, []).append(value.strip())
        first_line.setdefault(key, line_no)
    missing = [key for key in ACTIVITY_KEYS if key not in values]
    if missing:
        raise ParseError(
            block.line_no, f"hoạt động {block.title!r} thiếu khoá: " + ", ".join(missing)
        )
    minutes = values["thoi-luong"][0]
    if not minutes.isdigit() or int(minutes) <= 0:
        raise ParseError(
            first_line["thoi-luong"],
            f"'thoi-luong' phải là số phút lớn hơn 0, gặp {minutes!r}",
        )
    return Activity(
        title=block.title,
        line_no=block.line_no,
        minutes=int(minutes),
        muc_tieu=values["muc-tieu"],
        noi_dung=values["noi-dung"],
        san_pham=values["san-pham"],
        chuyen_giao=values["chuyen-giao"],
        thuc_hien=values["thuc-hien"],
        bao_cao=values["bao-cao"],
        ket_luan=values["ket-luan"],
        nls=split_codes(values.get("nls", [])),
        ai=split_codes(values.get("ai", [])),
    )


def build_rubric(block: Block) -> Rubric:
    items = bullets(block)
    if len(items) != 3:
        raise ParseError(
            block.line_no,
            f"tiêu chí rubric {block.title!r} phải có đúng ba mức, gặp {len(items)}",
        )
    levels = []
    for item, prefix in zip(items, LEVEL_PREFIXES):
        if not item.startswith(prefix):
            raise ParseError(
                block.line_no,
                f"các dòng của tiêu chí {block.title!r} phải mở đầu lần lượt bằng "
                + ", ".join(repr(value) for value in LEVEL_PREFIXES),
            )
        levels.append(item[len(prefix):].strip())
    return Rubric(title=block.title, levels=levels)


def _named_blocks(blocks: list[Block], section: str, expected: tuple[str, ...], line_no: int) -> dict[str, Block]:
    found = [block for block in blocks if block.section == section]
    if [block.title for block in found] != list(expected):
        raise ParseError(
            line_no,
            f"mục {section!r} phải có đúng các mục con theo thứ tự: " + ", ".join(expected),
        )
    return {block.title: block for block in found}


def _worksheet_warnings(activities: list[Activity], worksheets: list[Worksheet]) -> list[str]:
    warnings: list[str] = []
    if not worksheets:
        warnings.append(f"Giáo án chưa có mục {WORKSHEETS!r}.")
        return warnings
    defined = set()
    for sheet in worksheets:
        match = _NUMBER_RE.search(sheet.title)
        if match:
            defined.add(match.group(1))
    mentioned = set()
    for activity in activities:
        texts = (
            activity.noi_dung
            + activity.san_pham
            + activity.chuyen_giao
            + activity.thuc_hien
            + activity.bao_cao
        )
        for text in texts:
            for match in _WORKSHEET_MENTION_RE.finditer(text):
                mentioned.add(match.group(1))
    for number in sorted(mentioned - defined):
        warnings.append(
            f"Hoạt động nhắc Phiếu học tập {number} nhưng mục {WORKSHEETS!r} không có phiếu đó."
        )
    return warnings


def parse_lesson(text: str) -> Lesson:
    lines = text.splitlines()
    meta, start = parse_meta(lines)
    blocks, review_notes, section_lines = split_blocks(lines, start)

    objectives = _named_blocks(blocks, OBJECTIVES, OBJECTIVE_PARTS, section_lines[OBJECTIVES])
    equipment = _named_blocks(blocks, EQUIPMENT, EQUIPMENT_PARTS, section_lines[EQUIPMENT])
    ai_lines = bullets(objectives["Nang luc AI"])
    ai_no_framework = ai_lines[0] == NO_FRAMEWORK
    if not ai_no_framework:
        _check_competence_lines(objectives["Nang luc AI"], ai_lines)
    nls_lines = bullets(objectives["Nang luc so"])
    _check_competence_lines(objectives["Nang luc so"], nls_lines)

    activities = [build_activity(block) for block in blocks if block.section == PROCESS]
    if not activities:
        raise ParseError(section_lines[PROCESS], f"mục {PROCESS!r} không có hoạt động nào")

    worksheets = [
        Worksheet(title=block.title, items=bullets(block))
        for block in blocks
        if block.section == WORKSHEETS
    ]
    rubrics = [build_rubric(block) for block in blocks if block.section == RUBRIC]
    if len(rubrics) < 2:
        raise ParseError(
            section_lines[RUBRIC], f"mục {RUBRIC!r} phải có ít nhất hai tiêu chí, gặp {len(rubrics)}"
        )

    if not any(activity.nls for activity in activities):
        raise ParseError(
            section_lines[PROCESS],
            "không hoạt động nào có dòng 'nls:'; năng lực số phải nằm trong tiến trình dạy học, "
            "không chỉ ở mục tiêu",
        )
    if not any(activity.ai for activity in activities):
        raise ParseError(
            section_lines[PROCESS],
            "không hoạt động nào có dòng 'ai:'; năng lực AI phải nằm trong tiến trình dạy học, "
            "không chỉ ở mục tiêu",
        )

    lesson = Lesson(
        meta=meta,
        knowledge=bullets(objectives["Kien thuc"]),
        general=bullets(objectives["Nang luc chung"]),
        specific=bullets(objectives["Nang luc dac thu"]),
        nls_lines=nls_lines,
        ai_lines=ai_lines,
        qualities=bullets(objectives["Pham chat"]),
        teacher_tools=bullets(equipment["Giao vien"]),
        student_tools=bullets(equipment["Hoc sinh"]),
        activities=activities,
        worksheets=worksheets,
        rubrics=rubrics,
        review_notes=review_notes,
        ai_no_framework=ai_no_framework,
    )

    expected_minutes = lesson.periods * MINUTES_PER_PERIOD
    if lesson.minutes != expected_minutes:
        lesson.warnings.append(
            f"Tổng thời lượng các hoạt động là {lesson.minutes} phút, khác "
            f"{lesson.periods} tiết × {MINUTES_PER_PERIOD} = {expected_minutes} phút."
        )
    lesson.warnings.extend(_worksheet_warnings(activities, worksheets))
    return lesson


def parse_file(path: Path) -> Lesson:
    return parse_lesson(path.read_text(encoding="utf-8-sig"))
