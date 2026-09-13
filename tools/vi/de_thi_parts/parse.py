"""Đọc file nguồn de.md của đề kiểm tra thành cấu trúc dữ liệu.

Ngữ pháp của de.md nằm trong docs/vi/tro-ly/de-khtn-tieng-anh.md.
Chỉ dùng thư viện chuẩn Python.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

LEVELS = ("biet", "hieu", "vandung")
LEVEL_LABELS = {"biet": "Biết", "hieu": "Hiểu", "vandung": "Vận dụng"}
NO_TOPIC = "Không ghi chủ đề"
SOAT_HEADING = "## CAN SOAT"
PART_HEADINGS = {"## PART I": 1, "## PART II": 2, "## PART III": 3}
PART_LABELS = {1: "I", 2: "II", 3: "III"}
META_REQUIRED = ("school", "title", "subject", "time")
META_OPTIONAL = ("department", "code", "points")
POINTS = {1: 0.25, 2: 1.0, 3: 0.25}
_NUMBER_RE = re.compile(r"-?[0-9]+(?:[.,][0-9]+)?")


class ParseError(Exception):
    """Lỗi cú pháp trong de.md; luôn kèm số dòng để AI sửa được đúng chỗ."""

    def __init__(self, line_no: int, message: str) -> None:
        super().__init__(f"Dòng {line_no}: {message}")
        self.line_no = line_no
        self.message = message


def part_label(part: int) -> str:
    return PART_LABELS[part]


@dataclass
class Question:
    part: int
    number: int
    line_no: int
    en: str
    vi: str
    level: str
    topic: str
    why: str
    options: dict[str, str]
    statements: list[tuple[str, str, bool]]
    key: str
    unit: str


@dataclass
class Exam:
    meta: dict[str, str]
    questions: list[Question]
    review_notes: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def part(self, part: int) -> list[Question]:
        return [question for question in self.questions if question.part == part]

    def counts(self) -> dict[str, int]:
        return {f"part{number}": len(self.part(number)) for number in (1, 2, 3)}

    def total_points(self) -> float:
        counts = self.counts()
        total = sum(counts[f"part{number}"] * POINTS[number] for number in (1, 2, 3))
        return round(total, 2)

    def target_points(self) -> float:
        return float(self.meta.get("points", "10").replace(",", "."))

    def points_warning(self) -> str:
        if abs(self.total_points() - self.target_points()) < 0.005:
            return ""
        return (
            f"Thang điểm mặc định cho ra {self.total_points():g} điểm nhưng đề ghi "
            f"points: {self.target_points():g}. Kiểm tra lại số câu mỗi phần, hoặc sửa dòng "
            "points trong de.md."
        )

    def matrix(self) -> list[tuple[str, int, int, int, int]]:
        order: list[str] = []
        tally: dict[str, dict[str, int]] = {}
        for question in self.questions:
            topic = question.topic or NO_TOPIC
            if topic not in tally:
                order.append(topic)
                tally[topic] = {level: 0 for level in LEVELS}
            tally[topic][question.level] += 1
        rows = []
        for topic in order:
            counts = tally[topic]
            values = tuple(counts[level] for level in LEVELS)
            rows.append((topic, *values, sum(values)))
        return rows


def parse_meta(lines: list[str]) -> tuple[dict[str, str], int]:
    """Đọc khối thông tin đề giữa hai dòng '---'; trả về (meta, chỉ số dòng tiếp theo)."""
    if not lines or lines[0].strip() != "---":
        raise ParseError(1, "file phải mở đầu bằng một dòng '---' rồi tới khối thông tin đề")
    allowed = META_REQUIRED + META_OPTIONAL
    meta: dict[str, str] = {}
    for index in range(1, len(lines)):
        line_no = index + 1
        stripped = lines[index].strip()
        if stripped == "---":
            missing = [key for key in META_REQUIRED if not meta.get(key)]
            if missing:
                raise ParseError(line_no, "khối thông tin đề thiếu: " + ", ".join(missing))
            if not meta["time"].isdecimal():
                raise ParseError(line_no, f"'time' phải là số phút dạng chữ số, gặp {meta['time']!r}")
            if "points" in meta:
                try:
                    float(meta["points"].replace(",", "."))
                except ValueError:
                    raise ParseError(line_no, f"'points' phải là số, gặp {meta['points']!r}") from None
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
                f"khoá {key!r} không dùng trong khối thông tin đề; các khoá hợp lệ: "
                + ", ".join(allowed),
            )
        meta[key] = value.strip()
    raise ParseError(len(lines), "khối thông tin đề chưa được đóng bằng dòng '---'")


def build_question(
    part: int,
    number: int,
    line_no: int,
    body: list[tuple[int, str]],
    warnings: list[str],
) -> Question:
    fields: dict[str, tuple[int, str]] = {}
    for field_line, raw in body:
        key, separator, value = raw.partition(":")
        if not separator:
            raise ParseError(field_line, f"dòng phải dạng 'khoá: giá trị', gặp {raw!r}")
        key = key.strip()
        if key in fields:
            raise ParseError(field_line, f"khoá {key!r} xuất hiện hai lần trong cùng một câu")
        fields[key] = (field_line, value.strip())

    allowed = {"en", "vi", "level", "topic", "why"}
    if part == 1:
        allowed |= {"A", "B", "C", "D", "key"}
    elif part == 2:
        allowed |= {"a", "b", "c", "d"}
    else:
        allowed |= {"key", "unit"}
    for key, (field_line, _) in fields.items():
        if key not in allowed:
            raise ParseError(
                field_line,
                f"khoá {key!r} không dùng được ở PART {part_label(part)}; các khoá hợp lệ: "
                + ", ".join(sorted(allowed)),
            )

    for key in ("en", "level"):
        if not fields.get(key, (0, ""))[1]:
            raise ParseError(line_no, f"câu {number} ở PART {part_label(part)} thiếu khoá {key!r}")
    level = fields["level"][1]
    if level not in LEVELS:
        raise ParseError(
            fields["level"][0],
            f"'level' phải là một trong: {', '.join(LEVELS)}; gặp {level!r}",
        )

    options: dict[str, str] = {}
    statements: list[tuple[str, str, bool]] = []
    key_value = ""
    unit = ""
    if part == 1:
        for label in ("A", "B", "C", "D"):
            if not fields.get(label, (0, ""))[1]:
                raise ParseError(line_no, f"câu {number} ở PART I thiếu lựa chọn {label!r}")
            options[label] = fields[label][1]
        if "key" not in fields:
            raise ParseError(line_no, f"câu {number} ở PART I thiếu khoá 'key' (đáp án đúng)")
        key_value = fields["key"][1].strip().upper()
        if key_value not in options:
            raise ParseError(
                fields["key"][0],
                f"'key' phải là một chữ trong A, B, C, D; gặp {fields['key'][1]!r}",
            )
    elif part == 2:
        for label in ("a", "b", "c", "d"):
            if label not in fields:
                raise ParseError(line_no, f"câu {number} ở PART II thiếu ý {label!r}")
            field_line, value = fields[label]
            content, separator, flag = value.rpartition("|")
            flag = flag.strip().upper()
            if not separator or flag not in ("T", "F"):
                raise ParseError(field_line, f"ý {label!r} phải kết thúc bằng ' | T' hoặc ' | F'")
            if not content.strip():
                raise ParseError(field_line, f"ý {label!r} không có nội dung")
            statements.append((label, content.strip(), flag == "T"))
    else:
        if not fields.get("key", (0, ""))[1]:
            raise ParseError(line_no, f"câu {number} ở PART III thiếu khoá 'key' (đáp số)")
        key_value = fields["key"][1]
        if not _NUMBER_RE.fullmatch(key_value):
            raise ParseError(
                fields["key"][0],
                f"'key' ở PART III phải là một số (ví dụ 36, 12.5, 12,5, -0.25), gặp {key_value!r}",
            )
        unit = fields.get("unit", (0, ""))[1]

    vietnamese = fields.get("vi", (0, ""))[1]
    if not vietnamese:
        warnings.append(
            f"PART {part_label(part)} câu {number}: chưa có bản tiếng Việt (thiếu dòng 'vi:')"
        )

    return Question(
        part=part,
        number=number,
        line_no=line_no,
        en=fields["en"][1],
        vi=vietnamese,
        level=level,
        topic=fields.get("topic", (0, ""))[1],
        why=fields.get("why", (0, ""))[1],
        options=options,
        statements=statements,
        key=key_value,
        unit=unit,
    )


def check_numbering(questions: list[Question]) -> None:
    for part in (1, 2, 3):
        for position, question in enumerate([q for q in questions if q.part == part], start=1):
            if question.number != position:
                raise ParseError(
                    question.line_no,
                    f"câu ở PART {part_label(part)} phải đánh số {position}, gặp {question.number}",
                )


def parse_exam(text: str) -> Exam:
    lines = text.splitlines()
    meta, start = parse_meta(lines)
    blocks: list[tuple[int, int, int, list[tuple[int, str]]]] = []
    review_notes: list[str] = []
    part: int | None = None
    in_review = False
    for index in range(start, len(lines)):
        line_no = index + 1
        stripped = lines[index].strip()
        if not stripped:
            continue
        if stripped.startswith("## "):
            if stripped == SOAT_HEADING:
                in_review = True
                continue
            if stripped not in PART_HEADINGS:
                raise ParseError(
                    line_no,
                    f"tiêu đề {stripped!r} không hợp lệ; chỉ dùng '## PART I', '## PART II', "
                    "'## PART III' hoặc '## CAN SOAT'",
                )
            if in_review:
                raise ParseError(line_no, "'## CAN SOAT' phải là mục cuối file")
            new_part = PART_HEADINGS[stripped]
            if part is not None and new_part <= part:
                raise ParseError(
                    line_no,
                    "các phần phải theo thứ tự PART I, PART II, PART III và không lặp lại",
                )
            part = new_part
            continue
        if in_review:
            if not stripped.startswith("- "):
                raise ParseError(
                    line_no, f"dòng trong '## CAN SOAT' phải bắt đầu bằng '- ', gặp {stripped!r}"
                )
            review_notes.append(stripped[2:].strip())
            continue
        if stripped.startswith("### "):
            if part is None:
                raise ParseError(
                    line_no,
                    "câu hỏi nằm ngoài phần nào; thêm dòng '## PART I' trước câu hỏi đầu tiên",
                )
            number_text = stripped[4:].strip()
            if not number_text.isdecimal():
                raise ParseError(line_no, f"số câu phải là chữ số, gặp {number_text!r}")
            blocks.append((part, int(number_text), line_no, []))
            continue
        if not blocks:
            raise ParseError(
                line_no,
                f"dòng {stripped!r} nằm ngoài câu hỏi nào; mỗi câu mở đầu bằng '### <số>'",
            )
        blocks[-1][3].append((line_no, stripped))

    if not blocks:
        raise ParseError(max(len(lines), 1), "đề không có câu hỏi nào")

    warnings: list[str] = []
    questions = [
        build_question(block_part, number, block_line, body, warnings)
        for block_part, number, block_line, body in blocks
    ]
    check_numbering(questions)
    return Exam(meta=meta, questions=questions, review_notes=review_notes, warnings=warnings)


def parse_file(path: Path) -> Exam:
    return parse_exam(path.read_text(encoding="utf-8-sig"))
