# Kế hoạch thực thi: Soạn đề KHTN bằng tiếng Anh

> **Dành cho AI thực thi:** BẮT BUỘC dùng skill `superpowers:subagent-driven-development` (khuyến nghị) hoặc `superpowers:executing-plans` để làm từng task một. Các bước dùng ô đánh dấu `- [ ]`.

**Mục tiêu:** Thêm loại việc thứ 7 cho lớp Việt — soạn đề kiểm tra KHTN bằng tiếng Anh từ đề tiếng Việt có sẵn hoặc từ con số không — và xuất ra ba file Word in được.

**Kiến trúc:** AI viết đúng một file nguồn `de.md` theo ngữ pháp dòng `khoá: giá trị`. `tools/vi/de_thi.py` đọc file đó và sinh ba file Word bằng `python-docx`: đề tiếng Anh, đề song ngữ, đáp án kèm ma trận. Chất lượng tiếng Anh do hai file hướng dẫn cho AI bảo đảm, không có từ điển và không có script kiểm thuật ngữ.

**Công nghệ:** Python 3.10+ (thư viện chuẩn cho phần đọc, `python-docx>=1.1.0` cho phần dựng Word), `unittest`.

**Spec:** `docs/vi/phat-trien/2026-09-12-de-khtn-tieng-anh-design.md`

## Global Constraints

- Không sửa bất cứ gì trong `skills/`, `LICENSE`, `SPONSORS*.md`, và không sửa frontmatter của SKILL.md.
- Không bao giờ sửa, bỏ qua hay tìm cách "sửa chữa" `skills/ppt-master/scripts/attribution_guard.py`.
- Không sửa `requirements.txt` ở gốc repo và `skills/ppt-master/requirements.txt` — cả hai thuộc upstream. Thư viện mới chỉ khai ở `tools/vi/requirements-vi.txt`.
- Chỉ thêm đúng một thư viện: `python-docx>=1.1.0`.
- Đề của thầy cô đặt ở `projects/_de-thi/<tên_đề>/`; `projects/*` đã bị gitignore và không bao giờ được commit.
- Không tạo, không commit file `.env`.
- File `.ps1` phải giữ BOM UTF-8 ở đầu file.
- Mọi đường dẫn nhận từ dòng lệnh phải `.expanduser().resolve()` trước khi dùng.
- stdout của `de_thi.py` là **đúng một dòng JSON**; mọi tiến trình và log đi ra stderr.
- `error.step` chỉ thuộc `{input, parse, docx, write, internal}`.
- Test không được mở cửa sổ Word, PowerPoint hay trình duyệt.
- Chạy test bằng: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests`
- Commit message phải có dạng: dòng tiêu đề, một dòng trống, rồi trailer `Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>`. Kiểm lại bằng `git log -1 --format=%B` trước khi báo xong.
- Mọi chuỗi hiển thị cho thầy cô viết bằng tiếng Việt; tên hàm, tên biến, tên file viết bằng tiếng Anh.

---

### Task 1: Tách dấu đánh dấu trong câu (`inline.py`)

> **Lưu ý về vị trí file:** `inline.py` nằm ở `tools/vi/word_parts/`, **không** nằm trong `de_thi_parts/`. Gói giáo án (`docs/vi/phat-trien/2026-09-13-giao-an-nls-ai-design.md`) dùng lại đúng module này, nên nó thuộc tầng dựng Word dùng chung.

**Files:**
- Create: `tools/vi/word_parts/__init__.py`
- Create: `tools/vi/word_parts/inline.py`
- Create: `tools/vi/de_thi_parts/__init__.py`
- Create: `tools/vi/tests/test_de_thi.py`

**Interfaces:**
- Consumes: không có.
- Produces:
  - `inline.PLAIN = ""`, `inline.BOLD = "bold"`, `inline.SUB = "sub"`, `inline.SUP = "sup"`
  - `inline.split_runs(text: str) -> list[tuple[str, str]]`
  - `inline.plain_text(text: str) -> str`

- [ ] **Step 1: Viết test trước**

Tạo `tools/vi/tests/test_de_thi.py`:

```python
"""Test cho lớp soạn đề KHTN tiếng Anh của bản Việt."""

import sys
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT / "tools" / "vi"))

from word_parts import inline  # noqa: E402


class InlineTest(unittest.TestCase):
    def test_subscript_marks_become_sub_runs(self):
        self.assertEqual(
            inline.split_runs("H~2~SO~4~"),
            [("H", inline.PLAIN), ("2", inline.SUB), ("SO", inline.PLAIN), ("4", inline.SUB)],
        )

    def test_superscript_marks_become_sup_runs(self):
        self.assertEqual(
            inline.split_runs("5.0 m/s^2^"),
            [("5.0 m/s", inline.PLAIN), ("2", inline.SUP)],
        )

    def test_bold_marks_become_bold_runs(self):
        self.assertEqual(
            inline.split_runs("Which one is **not** correct?"),
            [("Which one is ", inline.PLAIN), ("not", inline.BOLD), (" correct?", inline.PLAIN)],
        )

    def test_plain_text_has_one_run(self):
        self.assertEqual(inline.split_runs("no marks here"), [("no marks here", inline.PLAIN)])

    def test_empty_text_still_returns_one_run(self):
        self.assertEqual(inline.split_runs(""), [("", inline.PLAIN)])

    def test_unpaired_marks_stay_plain(self):
        self.assertEqual(inline.split_runs("a ~ b ^ c"), [("a ~ b ^ c", inline.PLAIN)])

    def test_plain_text_strips_every_mark(self):
        self.assertEqual(inline.plain_text("H~2~O at 5 m/s^2^ is **not** ice"), "H2O at 5 m/s2 is not ice")


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Chạy test để thấy nó fail**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k InlineTest`
Kỳ vọng: FAIL với `ModuleNotFoundError: No module named 'word_parts'`

- [ ] **Step 3: Tạo hai gói con**

Tạo `tools/vi/word_parts/__init__.py` với đúng nội dung một dòng sau:

```python
"""Tầng dựng file Word dùng chung cho các gói của lớp Việt."""
```

Tạo `tools/vi/de_thi_parts/__init__.py` với đúng nội dung một dòng sau:

```python
"""Các phần của công cụ soạn đề KHTN tiếng Anh (lớp Việt)."""
```

- [ ] **Step 4: Viết `inline.py`**

Tạo `tools/vi/word_parts/inline.py`:

```python
"""Tách chuỗi có ~chỉ số dưới~, ^chỉ số trên^ và **in đậm** thành các đoạn để dựng Word."""

from __future__ import annotations

import re

PLAIN = ""
BOLD = "bold"
SUB = "sub"
SUP = "sup"

_TOKEN = re.compile(r"\*\*(.+?)\*\*|~([^~]+)~|\^([^\^]+)\^")


def split_runs(text: str) -> list[tuple[str, str]]:
    """Trả về danh sách (nội dung, kiểu); kiểu thuộc PLAIN, BOLD, SUB, SUP."""
    runs: list[tuple[str, str]] = []
    position = 0
    for match in _TOKEN.finditer(text):
        if match.start() > position:
            runs.append((text[position:match.start()], PLAIN))
        bold, sub, sup = match.group(1), match.group(2), match.group(3)
        if bold is not None:
            runs.append((bold, BOLD))
        elif sub is not None:
            runs.append((sub, SUB))
        else:
            runs.append((sup, SUP))
        position = match.end()
    if position < len(text):
        runs.append((text[position:], PLAIN))
    return runs or [("", PLAIN)]


def plain_text(text: str) -> str:
    """Bỏ hết dấu đánh dấu, chỉ còn chữ — dùng khi đo độ dài để xếp cột đáp án."""
    return "".join(chunk for chunk, _ in split_runs(text))
```

- [ ] **Step 5: Chạy test để thấy nó pass**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k InlineTest`
Kỳ vọng: PASS, 7 test.

- [ ] **Step 6: Chạy cả bộ test để chắc không làm hỏng gì**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests`
Kỳ vọng: OK.

- [ ] **Step 7: Commit**

```bash
git add tools/vi/word_parts/__init__.py tools/vi/word_parts/inline.py tools/vi/de_thi_parts/__init__.py tools/vi/tests/test_de_thi.py
git commit -m "feat(vi): parse inline subscript, superscript and bold marks"
```

Nhớ trailer theo Global Constraints.

---

### Task 2: Đọc file nguồn `de.md` (`parse.py`)

**Files:**
- Create: `tools/vi/de_thi_parts/parse.py`
- Modify: `tools/vi/tests/test_de_thi.py` (thêm `ParseTest`, `ExamMathTest`)

**Interfaces:**
- Consumes: không có (thuần thư viện chuẩn).
- Produces:
  - `parse.ParseError(line_no: int, message: str)` — thuộc tính `.line_no`, `.message`; `str(exc)` bắt đầu bằng `"Dòng <n>: "`
  - `parse.Question` — các trường `part: int`, `number: int`, `line_no: int`, `en: str`, `vi: str`, `level: str`, `topic: str`, `why: str`, `options: dict[str, str]`, `statements: list[tuple[str, str, bool]]`, `key: str`, `unit: str`
  - `parse.Exam` — các trường `meta: dict[str, str]`, `questions: list[Question]`, `review_notes: list[str]`, `warnings: list[str]`; các phương thức `part(n) -> list[Question]`, `counts() -> dict[str, int]`, `total_points() -> float`, `target_points() -> float`, `points_warning() -> str`, `matrix() -> list[tuple[str, int, int, int, int]]`
  - `parse.parse_exam(text: str) -> Exam`
  - `parse.part_label(part: int) -> str` trả `"I"`, `"II"`, `"III"`
  - `parse.LEVELS = ("biet", "hieu", "vandung")`, `parse.LEVEL_LABELS`, `parse.NO_TOPIC = "Không ghi chủ đề"`

- [ ] **Step 1: Viết test trước**

Thêm vào `tools/vi/tests/test_de_thi.py`, ngay trước khối `if __name__`:

```python
VALID_SOURCE = """---
school: TRƯỜNG THPT CHUYÊN NGUYỄN TRÃI
department: TỔ VẬT LÍ
title: ĐỀ KIỂM TRA GIỮA HỌC KÌ I
subject: PHYSICS — Grade 10
time: 45
code: 101
---

## PART I

### 1
en: An object of mass 2.0 kg is acted on by a resultant force of 10 N. Find its acceleration.
vi: Một vật khối lượng 2,0 kg chịu tác dụng của lực tổng hợp 10 N. Tính gia tốc của vật.
A: 2.0 m/s^2^
B: 5.0 m/s^2^
C: 10 m/s^2^
D: 20 m/s^2^
key: B
why: a = F/m = 10/2.0 = 5.0 m/s^2^
level: hieu
topic: Dynamics

### 2
en: Which of the following is **not** a vector quantity?
vi: Đại lượng nào sau đây **không** phải đại lượng vectơ?
A: Velocity
B: Force
C: Speed
D: Acceleration
key: C
level: biet
topic: Kinematics

## PART II

### 1
en: A trolley moves down a smooth slope.
vi: Một xe lăn đi xuống mặt phẳng nghiêng nhẵn.
a: Its speed increases. | T
b: Its acceleration is zero. | F
c: The resultant force acts down the slope. | T
d: Its kinetic energy decreases. | F
level: vandung
topic: Dynamics

## PART III

### 1
en: Calculate the work done by a force of 12 N moving an object 3.0 m in the direction of the force.
vi: Tính công của lực 12 N khi vật dịch chuyển 3,0 m theo hướng của lực.
key: 36
unit: J
level: hieu
topic: Energy

## CAN SOAT
- Thuật ngữ "resultant force" đã dùng thống nhất, thầy cô xác nhận lại giúp em.
"""


def source_without(line_prefix: str) -> str:
    kept = [line for line in VALID_SOURCE.splitlines() if not line.startswith(line_prefix)]
    return "\n".join(kept) + "\n"


class ParseTest(unittest.TestCase):
    def test_valid_source_reads_every_part(self):
        exam = parse.parse_exam(VALID_SOURCE)
        self.assertEqual(exam.counts(), {"part1": 2, "part2": 1, "part3": 1})
        self.assertEqual(exam.meta["school"], "TRƯỜNG THPT CHUYÊN NGUYỄN TRÃI")
        self.assertEqual(exam.meta["time"], "45")
        self.assertEqual(exam.meta["code"], "101")

    def test_multiple_choice_question_keeps_options_and_key(self):
        question = parse.parse_exam(VALID_SOURCE).part(1)[0]
        self.assertEqual(sorted(question.options), ["A", "B", "C", "D"])
        self.assertEqual(question.key, "B")
        self.assertEqual(question.level, "hieu")
        self.assertEqual(question.topic, "Dynamics")
        self.assertTrue(question.why)

    def test_true_false_question_keeps_four_statements_with_flags(self):
        question = parse.parse_exam(VALID_SOURCE).part(2)[0]
        self.assertEqual([label for label, _, _ in question.statements], ["a", "b", "c", "d"])
        self.assertEqual([correct for _, _, correct in question.statements], [True, False, True, False])

    def test_short_answer_question_keeps_key_and_unit(self):
        question = parse.parse_exam(VALID_SOURCE).part(3)[0]
        self.assertEqual(question.key, "36")
        self.assertEqual(question.unit, "J")

    def test_review_notes_are_collected(self):
        exam = parse.parse_exam(VALID_SOURCE)
        self.assertEqual(len(exam.review_notes), 1)
        self.assertIn("resultant force", exam.review_notes[0])

    def test_missing_vietnamese_line_is_a_warning_not_an_error(self):
        exam = parse.parse_exam(source_without("vi: Một vật khối lượng"))
        self.assertTrue(any("chưa có bản tiếng Việt" in warning for warning in exam.warnings))
        self.assertEqual(exam.part(1)[0].vi, "")

    def test_file_without_meta_block_fails_on_line_one(self):
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_exam("## PART I\n")
        self.assertEqual(caught.exception.line_no, 1)

    def test_meta_missing_required_key_names_it(self):
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_exam(source_without("subject:"))
        self.assertIn("subject", caught.exception.message)

    def test_meta_time_must_be_digits(self):
        broken = VALID_SOURCE.replace("time: 45", "time: bốn mươi lăm")
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_exam(broken)
        self.assertIn("time", caught.exception.message)

    def test_meta_rejects_unknown_key(self):
        broken = VALID_SOURCE.replace("code: 101", "mon: Vật lí")
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_exam(broken)
        self.assertIn("mon", caught.exception.message)

    def test_missing_option_names_the_label_and_line(self):
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_exam(source_without("C: 10 m/s^2^"))
        self.assertIn("'C'", caught.exception.message)
        self.assertGreater(caught.exception.line_no, 1)

    def test_key_outside_abcd_fails(self):
        broken = VALID_SOURCE.replace("key: B", "key: E", 1)
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_exam(broken)
        self.assertIn("A, B, C, D", caught.exception.message)

    def test_missing_key_in_part_one_fails(self):
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_exam(source_without("key: B"))
        self.assertIn("key", caught.exception.message)

    def test_level_must_be_one_of_three_values(self):
        broken = VALID_SOURCE.replace("level: hieu", "level: kho", 1)
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_exam(broken)
        for value in parse.LEVELS:
            self.assertIn(value, caught.exception.message)

    def test_true_false_statement_needs_a_flag(self):
        broken = VALID_SOURCE.replace("a: Its speed increases. | T", "a: Its speed increases.")
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_exam(broken)
        self.assertIn("' | T'", caught.exception.message)

    def test_parts_must_be_in_order(self):
        broken = VALID_SOURCE.replace("## PART I\n", "## PART III\n", 1)
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_exam(broken)
        self.assertIn("thứ tự", caught.exception.message)

    def test_question_numbers_must_restart_from_one_in_each_part(self):
        broken = VALID_SOURCE.replace("### 2\n", "### 5\n", 1)
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_exam(broken)
        self.assertIn("2", caught.exception.message)

    def test_exam_with_no_question_fails(self):
        header = VALID_SOURCE.split("## PART I")[0]
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_exam(header)
        self.assertIn("không có câu hỏi", caught.exception.message)

    def test_duplicate_key_in_one_question_fails(self):
        broken = VALID_SOURCE.replace("level: hieu\ntopic: Dynamics", "level: hieu\nlevel: biet", 1)
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_exam(broken)
        self.assertIn("hai lần", caught.exception.message)

    def test_review_section_must_be_last(self):
        broken = VALID_SOURCE.replace("## PART II", "## CAN SOAT\n- ghi chú\n\n## PART II", 1)
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_exam(broken)
        self.assertIn("mục cuối", caught.exception.message)

    def test_unit_key_is_rejected_in_part_one(self):
        broken = VALID_SOURCE.replace("key: B", "key: B\nunit: J", 1)
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_exam(broken)
        self.assertIn("unit", caught.exception.message)


class ExamMathTest(unittest.TestCase):
    def build(self, part1: int, part2: int, part3: int, points: str = "10") -> parse.Exam:
        lines = [
            "---",
            "school: TRƯỜNG A",
            "title: ĐỀ THI",
            "subject: PHYSICS — Grade 10",
            "time: 50",
            f"points: {points}",
            "---",
        ]
        if part1:
            lines.append("## PART I")
            for number in range(1, part1 + 1):
                lines += [f"### {number}", "en: Q", "A: 1", "B: 2", "C: 3", "D: 4",
                          "key: A", "level: biet", "topic: Motion"]
        if part2:
            lines.append("## PART II")
            for number in range(1, part2 + 1):
                lines += [f"### {number}", "en: Q",
                          "a: s | T", "b: s | F", "c: s | T", "d: s | F",
                          "level: hieu", "topic: Energy"]
        if part3:
            lines.append("## PART III")
            for number in range(1, part3 + 1):
                lines += [f"### {number}", "en: Q", "key: 1", "level: vandung"]
        return parse.parse_exam("\n".join(lines) + "\n")

    def test_standard_paper_totals_ten_points(self):
        exam = self.build(18, 4, 6)
        self.assertEqual(exam.total_points(), 10.0)
        self.assertEqual(exam.points_warning(), "")

    def test_wrong_count_produces_a_points_warning(self):
        exam = self.build(10, 4, 6)
        self.assertIn("points", exam.points_warning())

    def test_matrix_groups_by_topic_and_level(self):
        exam = self.build(2, 1, 0)
        rows = dict((row[0], row[1:]) for row in exam.matrix())
        self.assertEqual(rows["Motion"], (2, 0, 0, 2))
        self.assertEqual(rows["Energy"], (0, 1, 0, 1))

    def test_questions_without_topic_land_in_one_row(self):
        exam = self.build(0, 0, 2)
        self.assertEqual(exam.matrix(), [(parse.NO_TOPIC, 0, 0, 2, 2)])
```

Thêm `parse` vào phần import ở đầu file:

```python
from de_thi_parts import parse  # noqa: E402
from word_parts import inline  # noqa: E402
```

- [ ] **Step 2: Chạy test để thấy nó fail**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k ParseTest`
Kỳ vọng: FAIL với `ImportError: cannot import name 'parse'`

- [ ] **Step 3: Viết `parse.py`**

Tạo `tools/vi/de_thi_parts/parse.py`:

```python
"""Đọc file nguồn de.md của đề kiểm tra thành cấu trúc dữ liệu.

Ngữ pháp của de.md nằm trong docs/vi/tro-ly/de-khtn-tieng-anh.md.
Chỉ dùng thư viện chuẩn Python.
"""

from __future__ import annotations

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
            if not meta["time"].isdigit():
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
            if not number_text.isdigit():
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
```

- [ ] **Step 4: Chạy test để thấy nó pass**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k ParseTest`
rồi `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k ExamMathTest`
Kỳ vọng: PASS cả hai.

- [ ] **Step 5: Chạy cả bộ test**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests`
Kỳ vọng: OK.

- [ ] **Step 6: Commit**

```bash
git add tools/vi/de_thi_parts/parse.py tools/vi/tests/test_de_thi.py
git commit -m "feat(vi): read the exam source file into questions and a matrix"
```

---

### Task 3: Dựng ba file Word (`docx_build.py`)

**Files:**
- Create: `tools/vi/word_parts/base.py`
- Create: `tools/vi/de_thi_parts/docx_build.py`
- Modify: `tools/vi/tests/test_de_thi.py` (thêm `OptionLayoutTest`, `DocxBuildTest`)

**Interfaces:**
- Consumes: `word_parts.inline.split_runs`, `word_parts.inline.plain_text`, `parse.Exam`, `parse.Question`, `parse.LEVELS`, `parse.LEVEL_LABELS`, `parse.part_label`.
- Produces (tầng dùng chung, gói giáo án vi.6 sẽ dùng lại y nguyên):
  - `base.new_document(*, width_cm=21.0, height_cm=29.7, margins_cm=(1.8, 1.8, 2.5, 1.5), font="Times New Roman", size_pt=12, line_spacing=1.15, space_after_pt=2, page_numbers=True) -> Document` — `margins_cm` theo thứ tự trên, dưới, trái, phải
  - `base.write(paragraph, text, *, italic=False, bold=False, color=None) -> None`
  - `base.clear_borders(table) -> None`, `base.grid_borders(table) -> None`
  - `base.set_widths(table, widths) -> None`
  - `base.cell_text(cell, text, *, bold=False) -> None`
  - `base.fill_cell(cell, lines, *, bold_first=True) -> None`
- Produces (riêng gói đề thi):
  - `docx_build.FILENAMES = {"de": "de-en.docx", "song-ngu": "de-song-ngu.docx", "dap-an": "dap-an.docx"}`
  - `docx_build.option_columns(options: dict[str, str]) -> int`
  - `docx_build.build_de(exam, path: Path, *, bilingual: bool = False) -> Path`
  - `docx_build.build_dap_an(exam, path: Path, warnings: list[str]) -> Path`
  - `docx_build.build(exam, folder: Path, parts: list[str], warnings: list[str]) -> list[Path]`

Ghi chú cho người thực thi: `python-docx` được import bằng `from docx import ...` (tên gói trên PyPI là `python-docx`). Nếu máy chưa có, cài bằng `python -m pip install python-docx`.

- [ ] **Step 1: Viết test trước**

Thêm vào `tools/vi/tests/test_de_thi.py`:

```python
import tempfile
import zipfile


def document_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        return archive.read("word/document.xml").decode("utf-8")


def footer_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        names = [name for name in archive.namelist() if name.startswith("word/footer")]
        if not names:
            return ""
        return archive.read(names[0]).decode("utf-8")


class OptionLayoutTest(unittest.TestCase):
    def test_short_options_use_four_columns(self):
        options = {"A": "2.0 m/s^2^", "B": "5.0 m/s^2^", "C": "10 m/s^2^", "D": "20 m/s^2^"}
        self.assertEqual(docx_build.option_columns(options), 4)

    def test_medium_options_use_two_columns(self):
        options = {"A": "kinetic friction", "B": "static friction", "C": "air resistance", "D": "normal contact force"}
        self.assertEqual(docx_build.option_columns(options), 2)

    def test_long_options_use_one_column(self):
        long_text = "The resultant force acting on the trolley points down the slope at all times"
        options = {"A": long_text, "B": long_text, "C": long_text, "D": long_text}
        self.assertEqual(docx_build.option_columns(options), 1)

    def test_marks_do_not_count_towards_option_length(self):
        options = {"A": "H~2~SO~4~", "B": "HCl", "C": "NaOH", "D": "KOH"}
        self.assertEqual(docx_build.option_columns(options), 4)


class DocxBuildTest(unittest.TestCase):
    def setUp(self):
        self.exam = parse.parse_exam(VALID_SOURCE)
        self.tmp = tempfile.TemporaryDirectory()
        self.folder = Path(self.tmp.name)
        self.addCleanup(self.tmp.cleanup)

    def test_paper_uses_a4_and_exam_margins(self):
        from docx import Document

        path = docx_build.build_de(self.exam, self.folder / "de-en.docx")
        section = Document(str(path)).sections[0]
        # Word lưu khổ giấy theo twip nên đọc lại lệch vài trăm EMU (< 0,001 cm); so tới 0,01 cm.
        for attribute, expected_cm in (
            ("page_width", 21.0),
            ("page_height", 29.7),
            ("top_margin", 1.8),
            ("bottom_margin", 1.8),
            ("left_margin", 2.5),
            ("right_margin", 1.5),
        ):
            with self.subTest(attribute=attribute):
                self.assertAlmostEqual(getattr(section, attribute).cm, expected_cm, places=2)

    def test_body_font_is_times_new_roman_twelve(self):
        from docx import Document
        from docx.shared import Pt

        path = docx_build.build_de(self.exam, self.folder / "de-en.docx")
        normal = Document(str(path)).styles["Normal"]
        self.assertEqual(normal.font.name, "Times New Roman")
        self.assertEqual(normal.font.size, Pt(12))

    def test_footer_carries_page_number_fields(self):
        path = docx_build.build_de(self.exam, self.folder / "de-en.docx")
        xml = footer_xml(path)
        self.assertIn("PAGE", xml)
        self.assertIn("NUMPAGES", xml)
        self.assertIn("Trang", xml)

    def test_subscript_marks_become_real_subscript(self):
        source = VALID_SOURCE.replace("A: 2.0 m/s^2^", "A: H~2~SO~4~")
        exam = parse.parse_exam(source)
        path = docx_build.build_de(exam, self.folder / "de-en.docx")
        self.assertIn('w:val="subscript"', document_xml(path))

    def test_superscript_marks_become_real_superscript(self):
        path = docx_build.build_de(self.exam, self.folder / "de-en.docx")
        self.assertIn('w:val="superscript"', document_xml(path))

    def test_student_paper_hides_answers_and_explanations(self):
        source = VALID_SOURCE.replace(
            "why: a = F/m = 10/2.0 = 5.0 m/s^2^", "why: DAU-HIEU-GIAI-THICH"
        )
        exam = parse.parse_exam(source)
        path = docx_build.build_de(exam, self.folder / "de-en.docx")
        xml = document_xml(path)
        self.assertNotIn("DAU-HIEU-GIAI-THICH", xml)
        self.assertNotIn("Đáp án", xml)

    def test_student_paper_has_the_candidate_line_and_end_marker(self):
        path = docx_build.build_de(self.exam, self.folder / "de-en.docx")
        xml = document_xml(path)
        self.assertIn("Full name", xml)
        self.assertIn("THE END", xml)

    def test_bilingual_paper_carries_the_vietnamese_line(self):
        path = docx_build.build_de(self.exam, self.folder / "song-ngu.docx", bilingual=True)
        xml = document_xml(path)
        self.assertIn("Tính gia tốc của vật", xml)

    def test_student_paper_has_no_vietnamese_question_text(self):
        path = docx_build.build_de(self.exam, self.folder / "de-en.docx")
        self.assertNotIn("Tính gia tốc của vật", document_xml(path))

    def test_answer_file_has_keys_scoring_matrix_and_review(self):
        path = docx_build.build_dap_an(self.exam, self.folder / "dap-an.docx", [])
        xml = document_xml(path)
        for expected in ("Thang điểm", "Ma trận đặc tả", "Cần thầy cô soát", "Hướng dẫn giải", "Dynamics"):
            self.assertIn(expected, xml)

    def test_answer_file_repeats_tool_warnings_on_paper(self):
        path = docx_build.build_dap_an(self.exam, self.folder / "dap-an.docx", ["CANH-BAO-THU-NGHIEM"])
        self.assertIn("CANH-BAO-THU-NGHIEM", document_xml(path))

    def test_answer_file_says_so_when_nothing_needs_review(self):
        source = VALID_SOURCE.split("## CAN SOAT")[0]
        exam = parse.parse_exam(source)
        path = docx_build.build_dap_an(exam, self.folder / "dap-an.docx", [])
        self.assertIn("Không có mục nào cần soát", document_xml(path))

    def test_build_writes_every_requested_file(self):
        written = docx_build.build(self.exam, self.folder, ["de", "song-ngu", "dap-an"], [])
        self.assertEqual([path.name for path in written],
                         ["de-en.docx", "de-song-ngu.docx", "dap-an.docx"])
        for path in written:
            self.assertTrue(path.is_file(), path)

    def test_build_writes_only_the_requested_subset(self):
        written = docx_build.build(self.exam, self.folder, ["de", "dap-an"], [])
        self.assertEqual([path.name for path in written], ["de-en.docx", "dap-an.docx"])
        self.assertFalse((self.folder / "de-song-ngu.docx").exists())
```

Đổi phần import ở đầu file thành:

```python
from de_thi_parts import docx_build, parse  # noqa: E402
from word_parts import inline  # noqa: E402
```

- [ ] **Step 2: Chạy test để thấy nó fail**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k DocxBuildTest`
Kỳ vọng: FAIL với `ImportError: cannot import name 'docx_build'`

- [ ] **Step 3a: Viết tầng dùng chung `word_parts/base.py`**

Tạo `tools/vi/word_parts/base.py`:

```python
"""Tầng dựng file Word dùng chung: khổ giấy, font, bảng, chân trang.

Gói nào của lớp Việt cũng dùng module này; phần bố cục riêng thì để ở gói đó.
"""

from __future__ import annotations

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from . import inline


def field(paragraph, instruction: str) -> None:
    """Chèn một trường Word (PAGE, NUMPAGES) vào đoạn."""
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction_element = OxmlElement("w:instrText")
    instruction_element.set(qn("xml:space"), "preserve")
    instruction_element.text = f" {instruction} "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instruction_element)
    run._r.append(end)


def add_page_numbers(section) -> None:
    paragraph = section.footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.add_run("Trang ")
    field(paragraph, "PAGE")
    paragraph.add_run(" / ")
    field(paragraph, "NUMPAGES")


def new_document(
    *,
    width_cm: float = 21.0,
    height_cm: float = 29.7,
    margins_cm: tuple[float, float, float, float] = (1.8, 1.8, 2.5, 1.5),
    font: str = "Times New Roman",
    size_pt: int = 12,
    line_spacing: float = 1.15,
    space_after_pt: int = 2,
    page_numbers: bool = True,
) -> Document:
    """margins_cm theo thứ tự trên, dưới, trái, phải."""
    document = Document()
    section = document.sections[0]
    section.page_width = Cm(width_cm)
    section.page_height = Cm(height_cm)
    top, bottom, left, right = margins_cm
    section.top_margin = Cm(top)
    section.bottom_margin = Cm(bottom)
    section.left_margin = Cm(left)
    section.right_margin = Cm(right)
    normal = document.styles["Normal"]
    normal.font.name = font
    normal.font.size = Pt(size_pt)
    fonts = normal.element.get_or_add_rPr().get_or_add_rFonts()
    for attribute in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        fonts.set(qn(attribute), font)
    normal.paragraph_format.line_spacing = line_spacing
    normal.paragraph_format.space_after = Pt(space_after_pt)
    if page_numbers:
        add_page_numbers(section)
    return document


def write(paragraph, text: str, *, italic: bool = False, bold: bool = False, color=None) -> None:
    """Ghi text vào đoạn, dịch ~chỉ số dưới~, ^chỉ số trên^ và **in đậm** thành run thật."""
    for chunk, kind in inline.split_runs(text):
        run = paragraph.add_run(chunk)
        run.italic = italic
        run.bold = bold or kind == inline.BOLD
        if kind == inline.SUB:
            run.font.subscript = True
        elif kind == inline.SUP:
            run.font.superscript = True
        if color is not None:
            run.font.color.rgb = color


def _borders(table, value: str, size: str) -> None:
    element = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        edge_element = OxmlElement(f"w:{edge}")
        edge_element.set(qn("w:val"), value)
        edge_element.set(qn("w:sz"), size)
        element.append(edge_element)
    table._tbl.tblPr.append(element)


def clear_borders(table) -> None:
    _borders(table, "none", "0")


def grid_borders(table) -> None:
    _borders(table, "single", "6")


def set_widths(table, widths) -> None:
    table.autofit = False
    for index, width in enumerate(widths):
        table.columns[index].width = width
        for cell in table.columns[index].cells:
            cell.width = width


def cell_text(cell, text: str, *, bold: bool = False) -> None:
    cell.paragraphs[0].text = ""
    write(cell.paragraphs[0], text, bold=bold)


def fill_cell(cell, lines, *, bold_first: bool = True) -> None:
    """Mỗi dòng thành một đoạn căn giữa trong ô; dòng đầu có thể in đậm."""
    cell.paragraphs[0].text = ""
    for index, line in enumerate(lines):
        paragraph = cell.paragraphs[0] if index == 0 else cell.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        write(paragraph, line, bold=bold_first and index == 0)
```

- [ ] **Step 3b: Viết `docx_build.py`**

Tạo `tools/vi/de_thi_parts/docx_build.py`:

```python
"""Dựng ba file Word của đề kiểm tra: đề tiếng Anh, đề song ngữ, đáp án."""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor

from word_parts import base, inline
from word_parts.base import cell_text as _cell_text
from word_parts.base import clear_borders as _clear_borders
from word_parts.base import fill_cell as _fill_cell
from word_parts.base import grid_borders as _grid_borders
from word_parts.base import set_widths as _set_widths
from word_parts.base import write as _write

from .parse import LEVEL_LABELS, LEVELS, Exam, Question, part_label

GREY = RGBColor(0x59, 0x59, 0x59)
FILENAMES = {"de": "de-en.docx", "song-ngu": "de-song-ngu.docx", "dap-an": "dap-an.docx"}
PART_TITLES = {
    1: "PART I. MULTIPLE CHOICE",
    2: "PART II. TRUE / FALSE",
    3: "PART III. SHORT ANSWER",
}
PART_NOTES = {
    1: "Each question has only one correct answer.",
    2: "For each question, decide whether each statement is true or false.",
    3: "Write your answer in the space provided.",
}
CANDIDATE_LINE = (
    "Full name: ................................................  "
    "Class: ................  Student ID: ................"
)
ANSWER_LINE = "Answer: ............................................"
END_MARKER = "------ THE END ------"
NO_REVIEW = "Không có mục nào cần soát."
SHORT_OPTION = 12
MEDIUM_OPTION = 30
GRID_SIZE = 10


def option_columns(options: dict[str, str]) -> int:
    """Đáp án ngắn xếp 4 cột, trung bình 2 cột, dài thì mỗi đáp án một dòng."""
    longest = max(len(inline.plain_text(text)) for text in options.values())
    if longest <= SHORT_OPTION:
        return 4
    if longest <= MEDIUM_OPTION:
        return 2
    return 1


def _base_document() -> Document:
    """Thể thức đề thi: A4 dọc, lề 1,8/1,8/2,5/1,5 cm, Times New Roman 12, giãn dòng 1,15."""
    return base.new_document()


def _add_header(document: Document, exam: Exam, subtitle: str = "") -> None:
    table = document.add_table(rows=1, cols=2)
    _clear_borders(table)
    _set_widths(table, [Cm(8), Cm(8.5)])
    left, right = table.rows[0].cells
    left_lines = [exam.meta["school"]]
    if exam.meta.get("department"):
        left_lines.append(exam.meta["department"])
    _fill_cell(left, left_lines)
    right_lines = [exam.meta["title"]]
    if subtitle:
        right_lines.append(subtitle)
    right_lines.append(f"Môn: {exam.meta['subject']}")
    right_lines.append(f"Thời gian: {exam.meta['time']} phút")
    if exam.meta.get("code"):
        right_lines.append(f"Mã đề: {exam.meta['code']}")
    _fill_cell(right, right_lines)


def _add_candidate_line(document: Document) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(6)
    _write(paragraph, CANDIDATE_LINE)


def _add_part_heading(document: Document, part: int) -> None:
    heading = document.add_paragraph()
    heading.paragraph_format.space_before = Pt(8)
    _write(heading, PART_TITLES[part], bold=True)
    _write(document.add_paragraph(), PART_NOTES[part], italic=True)


def _add_options(document: Document, options: dict[str, str]) -> None:
    columns = option_columns(options)
    labels = ("A", "B", "C", "D")
    rows = (len(labels) + columns - 1) // columns
    table = document.add_table(rows=rows, cols=columns)
    _clear_borders(table)
    _set_widths(table, [Cm(16.5 / columns)] * columns)
    for index, label in enumerate(labels):
        cell = table.rows[index // columns].cells[index % columns]
        cell.paragraphs[0].text = ""
        _write(cell.paragraphs[0], f"{label}. ", bold=True)
        _write(cell.paragraphs[0], options[label])


def _add_statements(document: Document, question: Question) -> None:
    table = document.add_table(rows=len(question.statements) + 1, cols=3)
    _grid_borders(table)
    _set_widths(table, [Cm(12.5), Cm(2), Cm(2)])
    for cell, text in zip(table.rows[0].cells, ("Statement", "True", "False")):
        _cell_text(cell, text, bold=True)
    for row_index, (label, text, _) in enumerate(question.statements, start=1):
        cells = table.rows[row_index].cells
        cells[0].paragraphs[0].text = ""
        _write(cells[0].paragraphs[0], f"{label}) ")
        _write(cells[0].paragraphs[0], text)


def _add_question(document: Document, question: Question, *, bilingual: bool) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(4)
    _write(paragraph, f"Question {question.number}. ", bold=True)
    _write(paragraph, question.en)
    if bilingual:
        translation = document.add_paragraph()
        _write(translation, question.vi or "(chưa có bản tiếng Việt)", italic=True, color=GREY)
    if question.part == 1:
        _add_options(document, question.options)
    elif question.part == 2:
        _add_statements(document, question)
    else:
        _write(document.add_paragraph(), ANSWER_LINE)


def _add_end(document: Document) -> None:
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(8)
    _write(paragraph, END_MARKER, bold=True)


def build_de(exam: Exam, path: Path, *, bilingual: bool = False) -> Path:
    document = _base_document()
    _add_header(document, exam, subtitle="BẢN SONG NGỮ" if bilingual else "")
    _add_candidate_line(document)
    for part in (1, 2, 3):
        questions = exam.part(part)
        if not questions:
            continue
        _add_part_heading(document, part)
        for question in questions:
            _add_question(document, question, bilingual=bilingual)
    _add_end(document)
    document.save(str(path))
    return path


def _add_answer_grid(document: Document, exam: Exam) -> None:
    part1 = exam.part(1)
    if part1:
        _write(document.add_paragraph(), PART_TITLES[1], bold=True)
        for start in range(0, len(part1), GRID_SIZE):
            chunk = part1[start:start + GRID_SIZE]
            table = document.add_table(rows=2, cols=len(chunk))
            _grid_borders(table)
            for index, question in enumerate(chunk):
                _cell_text(table.rows[0].cells[index], str(question.number), bold=True)
                _cell_text(table.rows[1].cells[index], question.key)
    part2 = exam.part(2)
    if part2:
        _write(document.add_paragraph(), PART_TITLES[2], bold=True)
        table = document.add_table(rows=len(part2) + 1, cols=5)
        _grid_borders(table)
        for cell, text in zip(table.rows[0].cells, ("Question", "a", "b", "c", "d")):
            _cell_text(cell, text, bold=True)
        for row_index, question in enumerate(part2, start=1):
            cells = table.rows[row_index].cells
            _cell_text(cells[0], str(question.number))
            for column, (_, _, correct) in enumerate(question.statements, start=1):
                _cell_text(cells[column], "T" if correct else "F")
    part3 = exam.part(3)
    if part3:
        _write(document.add_paragraph(), PART_TITLES[3], bold=True)
        table = document.add_table(rows=len(part3) + 1, cols=3)
        _grid_borders(table)
        for cell, text in zip(table.rows[0].cells, ("Question", "Answer", "Unit")):
            _cell_text(cell, text, bold=True)
        for row_index, question in enumerate(part3, start=1):
            cells = table.rows[row_index].cells
            for cell, text in zip(cells, (str(question.number), question.key, question.unit or "—")):
                _cell_text(cell, text)


def _add_scoring(document: Document, exam: Exam) -> None:
    heading = document.add_paragraph()
    heading.paragraph_format.space_before = Pt(8)
    _write(heading, "Thang điểm", bold=True)
    counts = exam.counts()
    lines = []
    if counts["part1"]:
        lines.append(f"Phần I: {counts['part1']} câu × 0,25 = {counts['part1'] * 0.25:g} điểm.")
    if counts["part2"]:
        lines.append(
            f"Phần II: {counts['part2']} câu × 1,0 = {counts['part2'] * 1.0:g} điểm "
            "(đúng 1 ý 0,1; 2 ý 0,25; 3 ý 0,5; 4 ý 1,0)."
        )
    if counts["part3"]:
        lines.append(f"Phần III: {counts['part3']} câu × 0,25 = {counts['part3'] * 0.25:g} điểm.")
    lines.append(f"Tổng: {exam.total_points():g} điểm.")
    for line in lines:
        _write(document.add_paragraph(), line)


def _add_explanations(document: Document, exam: Exam) -> None:
    explained = [question for question in exam.questions if question.why]
    if not explained:
        return
    heading = document.add_paragraph()
    heading.paragraph_format.space_before = Pt(8)
    _write(heading, "Hướng dẫn giải", bold=True)
    for question in explained:
        paragraph = document.add_paragraph()
        _write(paragraph, f"PART {part_label(question.part)} — Question {question.number}. ", bold=True)
        _write(paragraph, question.why)


def _add_matrix(document: Document, exam: Exam) -> None:
    heading = document.add_paragraph()
    heading.paragraph_format.space_before = Pt(8)
    _write(heading, "Ma trận đặc tả", bold=True)
    rows = exam.matrix()
    table = document.add_table(rows=len(rows) + 2, cols=5)
    _grid_borders(table)
    _set_widths(table, [Cm(8.5), Cm(2), Cm(2), Cm(2), Cm(2)])
    header = ("Chủ đề", *(LEVEL_LABELS[level] for level in LEVELS), "Tổng")
    for cell, text in zip(table.rows[0].cells, header):
        _cell_text(cell, text, bold=True)
    for row_index, row in enumerate(rows, start=1):
        for cell, value in zip(table.rows[row_index].cells, row):
            _cell_text(cell, str(value))
    totals = (
        "Tổng",
        *(str(sum(row[index] for row in rows)) for index in (1, 2, 3)),
        str(sum(row[4] for row in rows)),
    )
    for cell, value in zip(table.rows[-1].cells, totals):
        _cell_text(cell, value, bold=True)


def _add_review(document: Document, exam: Exam, warnings: list[str]) -> None:
    heading = document.add_paragraph()
    heading.paragraph_format.space_before = Pt(8)
    _write(heading, "Cần thầy cô soát", bold=True)
    items = list(exam.review_notes) + list(warnings)
    if not items:
        _write(document.add_paragraph(), NO_REVIEW)
        return
    for item in items:
        _write(document.add_paragraph(), f"- {item}")


def build_dap_an(exam: Exam, path: Path, warnings: list[str]) -> Path:
    document = _base_document()
    _add_header(document, exam, subtitle="ĐÁP ÁN VÀ HƯỚNG DẪN CHẤM")
    _add_answer_grid(document, exam)
    _add_scoring(document, exam)
    _add_explanations(document, exam)
    _add_matrix(document, exam)
    _add_review(document, exam, warnings)
    document.save(str(path))
    return path


def build(exam: Exam, folder: Path, parts: list[str], warnings: list[str]) -> list[Path]:
    written: list[Path] = []
    for part in parts:
        path = folder / FILENAMES[part]
        if part == "de":
            build_de(exam, path)
        elif part == "song-ngu":
            build_de(exam, path, bilingual=True)
        else:
            build_dap_an(exam, path, warnings)
        written.append(path)
    return written
```

- [ ] **Step 4: Chạy test để thấy nó pass**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k OptionLayoutTest`
rồi `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k DocxBuildTest`
Kỳ vọng: PASS cả hai. Nếu `test_student_paper_hides_answers_and_explanations` fail vì chuỗi "Đáp án" xuất hiện, kiểm lại `build_de` không gọi nhánh đáp án nào.

- [ ] **Step 5: Chạy cả bộ test**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests`
Kỳ vọng: OK.

- [ ] **Step 6: Commit**

```bash
git add tools/vi/word_parts/base.py tools/vi/de_thi_parts/docx_build.py tools/vi/tests/test_de_thi.py
git commit -m "feat(vi): build the exam, bilingual and answer-key Word files"
```

---

### Task 4: Điểm vào dòng lệnh (`de_thi.py`)

**Files:**
- Create: `tools/vi/de_thi.py`
- Modify: `tools/vi/tests/test_de_thi.py` (thêm `SelectPartsTest`, `CliTest`)

**Interfaces:**
- Consumes: `parse.parse_exam`, `parse.ParseError`, `docx_build.FILENAMES`, `docx_build.build`.
- Produces:
  - `de_thi.select_parts(value: str) -> list[str]`
  - `de_thi.main(argv: list[str] | None = None) -> int`
  - stdout: đúng một dòng JSON với các khoá `ready`, `files`, `questions`, `points`, `warnings`, `error`.

- [ ] **Step 1: Viết test trước**

Thêm vào `tools/vi/tests/test_de_thi.py`:

```python
import contextlib
import io
import json
from unittest import mock

import de_thi  # noqa: E402


class SelectPartsTest(unittest.TestCase):
    def test_default_selects_every_part(self):
        self.assertEqual(de_thi.select_parts("tat-ca"), ["de", "song-ngu", "dap-an"])

    def test_comma_list_keeps_the_given_order(self):
        self.assertEqual(de_thi.select_parts("dap-an,de"), ["dap-an", "de"])

    def test_duplicates_are_dropped(self):
        self.assertEqual(de_thi.select_parts("de,de"), ["de"])

    def test_unknown_value_raises(self):
        with self.assertRaises(ValueError):
            de_thi.select_parts("dap_an")

    def test_empty_value_raises(self):
        with self.assertRaises(ValueError):
            de_thi.select_parts("")


class CliTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.folder = Path(self.tmp.name) / "đề thi vật lí 10"
        self.folder.mkdir(parents=True)
        self.addCleanup(self.tmp.cleanup)

    def write_source(self, text: str = VALID_SOURCE) -> None:
        (self.folder / "de.md").write_text(text, encoding="utf-8")

    def run_cli(self, *args: str) -> tuple[int, dict]:
        out = io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(io.StringIO()):
            code = de_thi.main([str(self.folder), *args])
        printed = out.getvalue().strip().splitlines()
        self.assertEqual(len(printed), 1, printed)
        return code, json.loads(printed[0])

    def test_full_run_writes_three_files(self):
        self.write_source()
        code, data = self.run_cli()
        self.assertEqual(code, 0, data)
        self.assertTrue(data["ready"])
        self.assertEqual(len(data["files"]), 3)
        self.assertEqual(data["questions"], {"part1": 2, "part2": 1, "part3": 1})
        self.assertIsNone(data["error"])

    def test_run_accepts_a_vietnamese_folder_name(self):
        self.write_source()
        code, data = self.run_cli()
        self.assertEqual(code, 0, data)
        for name in data["files"]:
            self.assertTrue(Path(name).is_file(), name)

    def test_subset_writes_only_two_files(self):
        self.write_source()
        code, data = self.run_cli("--phan", "de,dap-an")
        self.assertEqual(code, 0, data)
        self.assertEqual([Path(name).name for name in data["files"]], ["de-en.docx", "dap-an.docx"])

    def test_plan_only_writes_nothing(self):
        self.write_source()
        code, data = self.run_cli("--plan-only")
        self.assertEqual(code, 0, data)
        self.assertEqual(data["files"], [])
        self.assertEqual(list(self.folder.glob("*.docx")), [])

    def test_missing_source_reports_input_step(self):
        code, data = self.run_cli()
        self.assertEqual(code, 1)
        self.assertEqual(data["error"]["step"], "input")
        self.assertIn("de.md", data["error"]["message"])

    def test_broken_source_reports_parse_step_with_the_line(self):
        self.write_source(VALID_SOURCE.replace("key: B", "key: E", 1))
        code, data = self.run_cli()
        self.assertEqual(code, 1)
        self.assertEqual(data["error"]["step"], "parse")
        self.assertIn("Dòng", data["error"]["message"])

    def test_unknown_part_value_reports_input_step(self):
        self.write_source()
        code, data = self.run_cli("--phan", "dap_an")
        self.assertEqual(code, 1)
        self.assertEqual(data["error"]["step"], "input")

    def test_missing_python_docx_reports_docx_step_with_the_install_command(self):
        self.write_source()
        with mock.patch.object(de_thi, "load_docx_build", side_effect=ImportError("No module named 'docx'")):
            code, data = self.run_cli()
        self.assertEqual(code, 1)
        self.assertEqual(data["error"]["step"], "docx")
        self.assertIn("requirements-vi.txt", data["error"]["fix"])

    def test_write_failure_reports_write_step(self):
        self.write_source()
        with mock.patch.object(de_thi, "load_docx_build") as loader:
            loader.return_value.FILENAMES = {"de": "de-en.docx", "song-ngu": "de-song-ngu.docx", "dap-an": "dap-an.docx"}
            loader.return_value.build.side_effect = OSError("file đang mở trong Word")
            code, data = self.run_cli()
        self.assertEqual(code, 1)
        self.assertEqual(data["error"]["step"], "write")
        self.assertIn("Word", data["error"]["fix"])

    def test_unexpected_failure_still_prints_one_json_line(self):
        self.write_source()
        with mock.patch.object(de_thi, "load_docx_build") as loader:
            loader.return_value.FILENAMES = {"de": "de-en.docx", "song-ngu": "de-song-ngu.docx", "dap-an": "dap-an.docx"}
            loader.return_value.build.side_effect = ValueError("lỗi ngoài dự kiến")
            code, data = self.run_cli()
        self.assertEqual(code, 1)
        self.assertEqual(data["error"]["step"], "internal")

    def test_missing_vietnamese_line_surfaces_as_a_warning(self):
        self.write_source(source_without("vi: Một vật khối lượng"))
        code, data = self.run_cli()
        self.assertEqual(code, 0, data)
        self.assertTrue(any("bản tiếng Việt" in warning for warning in data["warnings"]))

    def test_second_run_warns_about_overwriting(self):
        self.write_source()
        self.run_cli()
        code, data = self.run_cli()
        self.assertEqual(code, 0, data)
        self.assertTrue(any("Ghi đè" in warning for warning in data["warnings"]))

    def test_emit_falls_back_to_utf8_buffer(self):
        class LegacyStdout:
            def __init__(self):
                self.buffer = io.BytesIO()

            def write(self, text):
                text.encode("cp1252")
                return len(text)

            def flush(self):
                pass

        stream = LegacyStdout()
        with mock.patch.object(de_thi.sys, "stdout", stream):
            de_thi.emit({"ready": False, "warnings": ["Thiếu bản tiếng Việt"]})
        data = json.loads(stream.buffer.getvalue().decode("utf-8").strip())
        self.assertEqual(data["warnings"], ["Thiếu bản tiếng Việt"])
```

- [ ] **Step 2: Chạy test để thấy nó fail**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k CliTest`
Kỳ vọng: FAIL với `ModuleNotFoundError: No module named 'de_thi'`

- [ ] **Step 3: Viết `de_thi.py`**

Tạo `tools/vi/de_thi.py`:

```python
#!/usr/bin/env python3
"""Xuất đề kiểm tra KHTN tiếng Anh ra file Word.

Cách dùng:
    python tools/vi/de_thi.py <thư_mục_đề> [--phan tat-ca|de,song-ngu,dap-an] [--plan-only]

stdout: đúng một dòng JSON. Tiến trình đi ra stderr.
Mã thoát: 0 khi xuất xong, 1 khi lỗi.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from de_thi_parts import parse  # noqa: E402

SOURCE_NAME = "de.md"
PART_CHOICES = ("de", "song-ngu", "dap-an")
MAX_PATH = 200
FIX_SOURCE = (
    "Viết file de.md trong thư mục đề theo docs/vi/tro-ly/de-khtn-tieng-anh.md rồi chạy lại."
)
FIX_PARSE = "Sửa đúng dòng đó trong de.md theo docs/vi/tro-ly/de-khtn-tieng-anh.md rồi chạy lại."
FIX_DOCX = (
    "Cài thư viện bằng: python -m pip install -r tools/vi/requirements-vi.txt "
    "(hoặc chạy lại CAI-DAT.bat)"
)
FIX_WRITE = "Đóng file Word đang mở rồi chạy lại; kiểm tra ổ đĩa còn trống."
FIX_INTERNAL = f"Gửi nguyên dòng error.message cho người bảo trì; xem {SOURCE_NAME} có ký tự lạ."


def log(text: str) -> None:
    print(text, file=sys.stderr, flush=True)


def emit(payload: dict) -> None:
    text = json.dumps(payload, ensure_ascii=False) + "\n"
    try:
        sys.stdout.write(text)
    except UnicodeEncodeError:
        sys.stdout.buffer.write(text.encode("utf-8", errors="replace"))
    sys.stdout.flush()


def result(
    *,
    ready: bool,
    files=(),
    counts: dict | None = None,
    points: float = 0.0,
    warnings=(),
    error: dict | None = None,
) -> dict:
    return {
        "ready": ready,
        "files": [str(path) for path in files],
        "questions": counts or {"part1": 0, "part2": 0, "part3": 0},
        "points": points,
        "warnings": list(warnings),
        "error": error,
    }


def failure(step: str, message: str, fix: str, **rest) -> dict:
    return result(ready=False, error={"step": step, "message": message, "fix": fix}, **rest)


def select_parts(value: str) -> list[str]:
    if value.strip() == "tat-ca":
        return list(PART_CHOICES)
    chosen = [item.strip() for item in value.split(",") if item.strip()]
    unknown = [item for item in chosen if item not in PART_CHOICES]
    if not chosen or unknown:
        raise ValueError(
            "--phan chỉ nhận tat-ca hoặc " + ", ".join(PART_CHOICES)
            + f" (cách nhau bằng dấu phẩy); gặp {value!r}"
        )
    ordered: list[str] = []
    for item in chosen:
        if item not in ordered:
            ordered.append(item)
    return ordered


def load_docx_build():
    """Import muộn để thiếu python-docx vẫn báo được lỗi dạng JSON."""
    from de_thi_parts import docx_build

    return docx_build


def configure_streams() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


def main(argv: list[str] | None = None) -> int:
    configure_streams()
    parser = argparse.ArgumentParser(description="Xuất đề kiểm tra KHTN tiếng Anh ra file Word")
    parser.add_argument("folder", type=Path, help="Thư mục đề, chứa file de.md")
    parser.add_argument("--phan", default="tat-ca", help="tat-ca, hoặc de,song-ngu,dap-an")
    parser.add_argument("--plan-only", action="store_true", help="Chỉ kiểm de.md, không ghi file")
    args = parser.parse_args(argv)

    try:
        parts = select_parts(args.phan)
    except ValueError as exc:
        emit(failure("input", str(exc), "Chạy lại với --phan tat-ca"))
        return 1

    folder = args.folder.expanduser().resolve()
    if not folder.is_dir():
        emit(failure("input", f"Không có thư mục đề: {folder}", FIX_SOURCE))
        return 1
    source = folder / SOURCE_NAME
    if not source.is_file():
        emit(failure("input", f"Không có file {SOURCE_NAME} trong {folder}", FIX_SOURCE))
        return 1
    try:
        text = source.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError) as exc:
        emit(failure("input", f"Không đọc được {source}: {exc}",
                     "Lưu lại de.md bằng bảng mã UTF-8 rồi chạy lại"))
        return 1

    try:
        exam = parse.parse_exam(text)
    except parse.ParseError as exc:
        emit(failure("parse", str(exc), FIX_PARSE))
        return 1

    paper_warnings = list(exam.warnings)
    points_warning = exam.points_warning()
    if points_warning:
        paper_warnings.append(points_warning)
    warnings = list(paper_warnings)
    if len(str(folder)) > MAX_PATH:
        warnings.append(
            f"Đường dẫn thư mục đề dài {len(str(folder))} ký tự; Windows có thể không ghi được "
            "file. Chuyển bộ công cụ ra ổ đĩa gần gốc, ví dụ D:\\PPTmaster."
        )

    counts = exam.counts()
    points = exam.total_points()
    if args.plan_only:
        log("Chỉ kiểm de.md, không ghi file.")
        emit(result(ready=True, counts=counts, points=points, warnings=warnings))
        return 0

    try:
        docx_build = load_docx_build()
    except ImportError as exc:
        emit(failure("docx", f"Chưa cài thư viện python-docx ({exc})", FIX_DOCX,
                     counts=counts, points=points, warnings=warnings))
        return 1

    for part in parts:
        target = folder / docx_build.FILENAMES[part]
        if target.exists():
            warnings.append(f"Ghi đè file có sẵn: {target.name}")

    try:
        written = docx_build.build(exam, folder, parts, list(paper_warnings))
    except OSError as exc:
        emit(failure("write", f"Không ghi được file Word: {exc}", FIX_WRITE,
                     counts=counts, points=points, warnings=warnings))
        return 1
    except Exception as exc:  # noqa: BLE001 - stdout không bao giờ được để trống
        emit(failure("internal", f"Lỗi ngoài dự kiến khi dựng file Word: {exc}", FIX_INTERNAL,
                     counts=counts, points=points, warnings=warnings))
        return 1

    log(f"Đã xuất {len(written)} file Word.")
    emit(result(ready=True, files=written, counts=counts, points=points, warnings=warnings))
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: Chạy test để thấy nó pass**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k SelectPartsTest`
rồi `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k CliTest`
Kỳ vọng: PASS cả hai.

- [ ] **Step 5: Chạy cả bộ test**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests`
Kỳ vọng: OK.

- [ ] **Step 6: Commit**

```bash
git add tools/vi/de_thi.py tools/vi/tests/test_de_thi.py
git commit -m "feat(vi): add the de_thi command with a one-line JSON contract"
```

---

### Task 5: Khai thư viện và nối vào bộ cài

**Files:**
- Create: `tools/vi/requirements-vi.txt`
- Modify: `tools/vi/doctor.py` (hàm `check_packages`, hàm `collect`)
- Modify: `tools/vi/pptmaster.ps1` (hai chỗ `pip install -r`: trong `Invoke-Setup` và trong nhánh `-Auto`; và hàm `Test-PackagesOk`)
- Modify: `tools/vi/setup.sh` (dòng `pip install -r`)
- Modify: `tools/vi/tests/test_doctor.py` (thêm hằng `VI_REQUIREMENTS` và lớp `ViPackagesTest`)
- Modify: `tools/vi/tests/test_installer.py` (thêm lớp `ViRequirementsWiringTest`, và một phương thức mới trong lớp có sẵn `InstallerPlanTest`)

**Interfaces:**
- Consumes: không có.
- Produces: `doctor.check_packages(requirements_path, find_dist=..., name="Thư viện Python", level=doctor.REQUIRED, fix="Chạy CAI-DAT.bat để cài thư viện") -> CheckResult`

- [ ] **Step 1: Tạo file khai thư viện**

Tạo `tools/vi/requirements-vi.txt`:

```
# Thư viện riêng của lớp Việt hoá. File này KHÔNG thuộc dự án gốc;
# requirements.txt ở gốc repo và skills/ppt-master/requirements.txt là của upstream, không sửa.
#
# Cài: python -m pip install -r tools/vi/requirements-vi.txt
#
# Dựng file Word cho đề kiểm tra (tools/vi/de_thi.py)
python-docx>=1.1.0
```

- [ ] **Step 2: Viết test trước cho doctor**

Thêm vào `tools/vi/tests/test_doctor.py`:

```python
VI_REQUIREMENTS = Path(doctor.__file__).resolve().parent / "requirements-vi.txt"


class ViPackagesTest(unittest.TestCase):
    def test_vi_requirements_file_declares_python_docx(self):
        text = VI_REQUIREMENTS.read_text(encoding="utf-8")
        self.assertIn("python-docx", doctor.parse_requirement_names(text))

    def test_check_packages_uses_the_given_name_and_level(self):
        result = doctor.check_packages(
            VI_REQUIREMENTS,
            find_dist=lambda name: object(),
            name="Thư viện lớp Việt",
            level=doctor.RECOMMENDED,
        )
        self.assertEqual(result.name, "Thư viện lớp Việt")
        self.assertEqual(result.level, doctor.RECOMMENDED)
        self.assertTrue(result.ok)

    def test_missing_vi_package_is_a_warning_not_a_blocking_error(self):
        def missing(name):
            raise doctor.importlib.metadata.PackageNotFoundError(name)

        result = doctor.check_packages(
            VI_REQUIREMENTS,
            find_dist=missing,
            name="Thư viện lớp Việt",
            level=doctor.RECOMMENDED,
            fix="Chạy: python -m pip install -r tools/vi/requirements-vi.txt",
        )
        self.assertFalse(result.ok)
        self.assertIn("python-docx", result.detail)
        self.assertIn("requirements-vi.txt", result.fix)
        self.assertEqual(doctor.exit_code([result]), 0)

    def test_collect_reports_the_vi_layer_packages(self):
        results = doctor.collect(no_smoke=True)
        self.assertIn("Thư viện lớp Việt", [item.name for item in results])
```

`test_doctor.py` **không** có hằng `REPO_ROOT`: nó thêm `tools/vi` vào `sys.path`, `import doctor`, và đã import sẵn `Path`. Đặt hằng `VI_REQUIREMENTS` ở mức mô-đun ngay trước lớp `ViPackagesTest`; không đổi các import có sẵn.

- [ ] **Step 3: Chạy test để thấy nó fail**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k ViPackagesTest`
Kỳ vọng: FAIL — `check_packages()` chưa nhận `name`.

- [ ] **Step 4: Sửa `doctor.py`**

Thay chữ ký và phần trả kết quả của `check_packages` (hiện ở `tools/vi/doctor.py:81-95`) thành:

```python
def check_packages(
    requirements_path: Path,
    find_dist: Callable[[str], object] = importlib.metadata.distribution,
    name: str = "Thư viện Python",
    level: str = REQUIRED,
    fix: str = "Chạy CAI-DAT.bat để cài thư viện",
) -> CheckResult:
    try:
        packages = parse_requirement_names(requirements_path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeDecodeError):
        return CheckResult(name, level, False, f"Không đọc được {requirements_path.name}",
                           "Tải lại bản đầy đủ của bộ công cụ")
    missing = []
    for package in packages:
        try:
            find_dist(package)
        except importlib.metadata.PackageNotFoundError:
            missing.append(package)
    if missing:
        return CheckResult(name, level, False, "Thiếu: " + ", ".join(missing), fix)
    return CheckResult(name, level, True, f"Đủ {len(packages)} gói")
```

Trong `collect` (hiện ở `tools/vi/doctor.py:316-339`), ngay sau dòng `results += [packages, integrity]`, thêm:

```python
    results.append(check_packages(
        REPO_ROOT / "tools" / "vi" / "requirements-vi.txt",
        name="Thư viện lớp Việt",
        level=RECOMMENDED,
        fix="Chạy: python -m pip install -r tools/vi/requirements-vi.txt (hoặc chạy lại CAI-DAT.bat)",
    ))
```

- [ ] **Step 5: Viết test trước cho bộ cài**

Thêm vào `tools/vi/tests/test_installer.py`:

`test_installer.py` đã có sẵn hằng `REPO_ROOT` và `LAUNCHER` (`= REPO_ROOT / "tools" / "vi" / "pptmaster.ps1"`). Thêm lớp mới này **sau** lớp `InstallerPlanTest`, trước khối `if __name__`:

```python
class ViRequirementsWiringTest(unittest.TestCase):
    def test_launcher_installs_the_vi_requirements_in_both_paths(self):
        text = LAUNCHER.read_text(encoding="utf-8-sig")
        self.assertEqual(text.count("'tools\\vi\\requirements-vi.txt'"), 2, text.count("requirements-vi"))

    def test_upstream_install_code_is_captured_before_the_vi_install(self):
        """Lệnh cài lớp Việt chạy sau không được che mất lỗi của lệnh cài thư viện upstream."""
        lines = LAUNCHER.read_text(encoding="utf-8-sig").splitlines()
        upstream = [i for i, line in enumerate(lines) if "'requirements.txt'" in line and "pip install" in line]
        vi_layer = [i for i, line in enumerate(lines) if "'tools\\vi\\requirements-vi.txt'" in line]
        self.assertEqual(len(upstream), 2, upstream)
        self.assertEqual(len(vi_layer), 2, vi_layer)
        for upstream_line, vi_line in zip(upstream, vi_layer):
            between = lines[upstream_line + 1:vi_line]
            self.assertTrue(
                any("$pipCode = $LASTEXITCODE" in line for line in between),
                f"thiếu '$pipCode = $LASTEXITCODE' giữa dòng {upstream_line + 1} và dòng {vi_line + 1}",
            )

    def test_package_check_also_watches_the_vi_layer_packages(self):
        text = LAUNCHER.read_text(encoding="utf-8-sig")
        start = text.index("function Test-PackagesOk")
        body = text[start:text.index("\n}", start)]
        self.assertIn("'Thư viện lớp Việt'", body)

    def test_setup_sh_installs_the_vi_requirements(self):
        text = (REPO_ROOT / "tools" / "vi" / "setup.sh").read_text(encoding="utf-8")
        self.assertIn("tools/vi/requirements-vi.txt", text)
```

Thêm phương thức này **vào trong lớp có sẵn `InstallerPlanTest`**, ngay sau `test_auto_setup_second_run_installs_nothing` — nó cần các tiện ích `self.repo` và `self.run_launcher` của lớp đó. Không sửa `test_auto_setup_second_run_installs_nothing`:

```python
    def test_auto_setup_installs_when_vi_layer_packages_are_missing(self):
        """Máy đã cài bản cũ (đủ thư viện upstream, thiếu python-docx) phải vào nhánh cài thư viện."""
        report = {"ready": True, "python": "x", "checks": [
            {"name": "Thư viện Python", "level": "required", "ok": True, "detail": "", "fix": ""},
            {"name": "Thư viện lớp Việt", "level": "recommended", "ok": False,
             "detail": "Thiếu: python-docx", "fix": ""},
        ]}
        payload = json.dumps(report, ensure_ascii=False)
        (self.repo / "tools" / "vi" / "doctor.py").write_text(
            f"import sys\nsys.stdout.reconfigure(encoding='utf-8')\nprint({payload!r})\n", encoding="utf-8",
        )
        subprocess.run([sys.executable, "-m", "venv", "--without-pip", str(self.repo / "venv")],
                       check=True, capture_output=True, timeout=120)
        (self.repo / ".env").write_text("", encoding="utf-8")
        returncode, result = self.run_launcher("-Action", "setup", "-Auto")
        # venv được tạo không có pip, nên nhánh cài thư viện dừng ở bước kiểm pip.
        # Kết quả này chứng tỏ bộ cài đã KHÔNG bỏ qua bước cài như trước.
        self.assertEqual(returncode, 1, result)
        self.assertEqual(result["error"]["step"], "venv", result)
```

- [ ] **Step 6: Sửa `pptmaster.ps1`**

Ba chỗ. Giữ nguyên BOM UTF-8 của file: sửa bằng công cụ sửa file, không ghi lại toàn bộ file bằng PowerShell.

**(a)** Ở `Invoke-Setup`, thay hai dòng hiện ở `tools/vi/pptmaster.ps1:180-181`:

```powershell
    & $py.Path -m pip install -r (Join-Path $RepoRoot 'requirements.txt') | Out-Host
    if ($LASTEXITCODE -ne 0) {
```

thành:

```powershell
    & $py.Path -m pip install -r (Join-Path $RepoRoot 'requirements.txt') | Out-Host
    $pipCode = $LASTEXITCODE
    & $py.Path -m pip install -r (Join-Path $RepoRoot 'tools\vi\requirements-vi.txt') | Out-Host
    if ($pipCode -ne 0 -or $LASTEXITCODE -ne 0) {
```

**(b)** Ở nhánh `-Auto`, thay hai dòng hiện ở `tools/vi/pptmaster.ps1:385-386`:

```powershell
            Invoke-Logged { & $VenvPython -m pip install -r (Join-Path $RepoRoot 'requirements.txt') }
            $pipCode = $LASTEXITCODE
```

thành:

```powershell
            Invoke-Logged { & $VenvPython -m pip install -r (Join-Path $RepoRoot 'requirements.txt') }
            $pipCode = $LASTEXITCODE
            Invoke-Logged { & $VenvPython -m pip install -r (Join-Path $RepoRoot 'tools\vi\requirements-vi.txt') }
            if ($LASTEXITCODE -ne 0) { $pipCode = $LASTEXITCODE }
```

Lý do cho (a) và (b): đọc `$LASTEXITCODE` sau hai lệnh cài liên tiếp chỉ thấy mã của lệnh sau, nên lỗi cài thư viện upstream sẽ bị che mất.

**(c)** Thay toàn bộ hàm `Test-PackagesOk` (hiện ở `tools/vi/pptmaster.ps1:319-325`) thành:

```powershell
function Test-PackagesOk($Report) {
    if (-not $Report) { return $false }
    $upstreamOk = $false
    foreach ($check in $Report.checks) {
        if ($check.name -eq 'Thư viện Python') { $upstreamOk = [bool]$check.ok }
        # Doctor cũ không có mục này thì coi như ổn; có mục mà báo thiếu thì phải cài.
        if ($check.name -eq 'Thư viện lớp Việt' -and -not [bool]$check.ok) { return $false }
    }
    return $upstreamOk
}
```

Lý do cho (c): hàm cũ chỉ nhìn mục `Thư viện Python`, nên máy đã cài bản cũ (đủ thư viện upstream, thiếu `python-docx`) chạy `-Action setup -Auto` sẽ bỏ qua bước cài.

- [ ] **Step 7: Sửa `setup.sh`**

Thay dòng 17 của `tools/vi/setup.sh`:

```sh
"$PY" -m pip install -r "$REPO_ROOT/requirements.txt"
"$PY" -m pip install -r "$REPO_ROOT/tools/vi/requirements-vi.txt"
```

- [ ] **Step 8: Chạy test**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests`
Kỳ vọng: OK. `doctor.py --no-smoke` phải in một dòng "Thư viện lớp Việt".

Chạy thêm: `python tools/vi/doctor.py --no-smoke`
Kỳ vọng: có dòng `Thư viện lớp Việt`, và kết quả cuối vẫn là "sẵn sàng sử dụng" khi máy đã cài `python-docx`.

- [ ] **Step 9: Commit**

```bash
git add tools/vi/requirements-vi.txt tools/vi/doctor.py tools/vi/pptmaster.ps1 tools/vi/setup.sh tools/vi/tests/test_doctor.py tools/vi/tests/test_installer.py
git commit -m "feat(vi): install and check the Vietnamese-layer requirements"
```

---

### Task 6: Hướng dẫn cho AI

**Files:**
- Create: `docs/vi/tro-ly/de-khtn-tieng-anh.md`
- Create: `docs/vi/tro-ly/tieng-anh-khoa-hoc.md`
- Modify: `AGENTS.vi.md` (mục 3, tiêu đề mục 10, bảng mục 10, thêm mục 12)
- Modify: `docs/vi/tro-ly/quy-trinh-hoi.md` (tiêu đề, mục "Khi nào áp dụng")
- Modify: `docs/vi/tro-ly/mau-brief.md` (thêm loại việc thứ 7)
- Modify: `tools/vi/tests/test_vi_layer.py`

**Interfaces:**
- Consumes: hợp đồng `de.md` ở Task 2, lệnh và JSON ở Task 4, lệnh cài ở Task 5.
- Produces: không có mã.

Ràng buộc riêng của task này (test đang khoá):
- `de-khtn-tieng-anh.md` **không** được thêm vào `GUIDE_FILES` — các file đó bị khoá phải có đúng tám mục gồm `## Khổ slide` và `## Phong cách gợi ý`.
- Hai file mới nằm trong `docs/vi/tro-ly/` nên **không được dùng liên kết Markdown**; nhắc file khác bằng đường dẫn chữ thường, như dòng 3 của `bai-giang.md`.
- `## Câu hỏi bắt buộc` phải có 1–7 câu đánh số, mỗi câu có chuỗi `Gợi ý:`. `## Tạo nhanh` phải có 2–3 câu đánh số.

- [ ] **Step 1: Viết test trước**

Trong `tools/vi/tests/test_vi_layer.py`, sửa hằng số và thêm lớp test.

Sửa `AGENTS_VI_ASSISTANT_HEADING` (dòng 397) thành:

```python
AGENTS_VI_ASSISTANT_HEADING = "## 10. Hỗ trợ thầy cô trước khi làm bài"
```

Sửa `test_agents_vi_keeps_assistant_then_video_sections_last` (dòng 401-404) thành:

```python
    def test_agents_vi_keeps_the_three_task_sections_last_in_order(self):
        headings = h2_headings(read("AGENTS.vi.md"))
        self.assertEqual(headings[-3], AGENTS_VI_ASSISTANT_HEADING)
        self.assertEqual(headings[-2], AGENTS_VI_VIDEO_HEADING)
        self.assertEqual(headings[-1], AGENTS_VI_EXAM_HEADING)
```

Sửa dòng cuối của `test_agents_vi_environment_section_points_to_guide` (dòng 550) thành:

```python
        self.assertEqual(h2_headings(text)[-3], AGENTS_VI_ASSISTANT_HEADING)
```

Sửa `test_task_type_count_matches_the_table` (dòng 669-675) thành:

```python
    def test_task_type_count_matches_the_table(self):
        agents_vi_body = section(read("AGENTS.vi.md"), AGENTS_VI_ASSISTANT_HEADING)
        self.assertIn("7 loại", agents_vi_body)
        for stale in ("5 loại", "6 loại"):
            self.assertNotIn(stale, agents_vi_body)
        quy_trinh_body = section(read("docs/vi/tro-ly/quy-trinh-hoi.md"), "## Khi nào áp dụng")
        self.assertIn("7 loại", quy_trinh_body)
        for stale in ("5 loại", "6 loại"):
            self.assertNotIn(stale, quy_trinh_body)
```

Thêm `"soan-de-tieng-anh.md"` vào `REQUIRED_DOCS` (dòng 135-144) — file đó do Task 7 tạo, nên bước này để Task 7 làm; **không** thêm ở task này.

Thêm vào cuối file, trước khối `if __name__`:

```python
AGENTS_VI_EXAM_HEADING = "## 12. Soạn đề KHTN bằng tiếng Anh"
EXAM_GUIDE = "docs/vi/tro-ly/de-khtn-tieng-anh.md"
ENGLISH_GUIDE = "docs/vi/tro-ly/tieng-anh-khoa-hoc.md"
EXAM_GUIDE_HEADINGS = (
    "## Khi nào dùng",
    "## Câu hỏi bắt buộc",
    "## Câu hỏi tuỳ chọn",
    "## Tạo nhanh",
    "## Cấu trúc đề",
    "## Đầu ra",
    "## Ghi vào brief",
)
EXAM_COMMAND = r"python tools\vi\de_thi.py"


class ExamGuideTest(unittest.TestCase):
    def test_exam_guide_has_its_own_sections_in_order(self):
        self.assertEqual(h2_headings(read(EXAM_GUIDE)), list(EXAM_GUIDE_HEADINGS))

    def test_exam_guide_is_not_treated_as_a_slide_guide(self):
        self.assertNotIn("de-khtn-tieng-anh.md", GUIDE_FILES)

    def test_exam_guide_questions_are_limited_and_have_suggestions(self):
        items = numbered_items(section(read(EXAM_GUIDE), "## Câu hỏi bắt buộc"))
        self.assertTrue(1 <= len(items) <= 7, f"{len(items)} câu")
        for item in items:
            self.assertIn("Gợi ý:", item)

    def test_exam_guide_quick_mode_asks_two_or_three_questions(self):
        items = numbered_items(section(read(EXAM_GUIDE), "## Tạo nhanh"))
        self.assertTrue(2 <= len(items) <= 3, f"{len(items)} câu")

    def test_exam_guide_must_ask_thcs_counts_instead_of_defaulting(self):
        body = section(read(EXAM_GUIDE), "## Câu hỏi bắt buộc")
        self.assertIn("18", body)
        self.assertIn("THCS", body)
        self.assertIn("phải hỏi", body)

    def test_exam_guide_names_the_output_files(self):
        body = section(read(EXAM_GUIDE), "## Đầu ra")
        for name in ("de-en.docx", "de-song-ngu.docx", "dap-an.docx", "projects/_de-thi/"):
            self.assertIn(name, body)

    def test_exam_guide_points_to_the_english_rules_and_the_source_grammar(self):
        text = read(EXAM_GUIDE)
        self.assertIn("tieng-anh-khoa-hoc.md", text)
        self.assertIn("## CAN SOAT", text)
        for key in ("en:", "vi:", "key:", "level:", "topic:"):
            self.assertIn(key, text)

    def test_exam_guide_forbids_changing_the_original_paper(self):
        text = read(EXAM_GUIDE)
        for phrase in ("không đổi số liệu", "không tự sửa", "Cần thầy cô soát"):
            self.assertIn(phrase, text)

    def test_tro_ly_files_have_no_markdown_links(self):
        for name in ("de-khtn-tieng-anh.md", "tieng-anh-khoa-hoc.md"):
            with self.subTest(file=name):
                self.assertEqual(LINK_RE.findall(read(f"docs/vi/tro-ly/{name}")), [])


class ScienceEnglishGuideTest(unittest.TestCase):
    def test_guide_states_every_principle(self):
        text = read(ENGLISH_GUIDE)
        for phrase in (
            "không dịch từng chữ",
            "uniformly accelerated motion",
            "kinetic friction",
            "molar mass",
            "cellular respiration",
            "State",
            "Explain",
            "Calculate",
            "sulfuric acid",
            "aluminium",
            "25.5",
            "at 0 °C and 1 atm",
            "terraced fields",
            "Cần thầy cô soát",
        ):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, text)

    def test_guide_covers_all_four_subject_frames(self):
        text = read(ENGLISH_GUIDE)
        for subject in ("Vật lí", "Hoá học", "Sinh học", "KHTN"):
            self.assertIn(subject, text)

    def test_guide_forbids_making_the_english_harder_than_the_science(self):
        text = read(ENGLISH_GUIDE)
        self.assertIn("Độ khó nằm ở khoa học", text)


class ExamWiringTest(unittest.TestCase):
    def test_common_rules_table_lists_the_exam_task(self):
        body = section(read("docs/vi/tro-ly/quy-trinh-hoi.md"), "## Khi nào áp dụng")
        self.assertIn("de-khtn-tieng-anh.md", body)
        for keyword in ("đề tiếng Anh", "đề KHTN"):
            self.assertIn(keyword, body)

    def test_agents_vi_exam_section_explains_both_use_cases_and_the_command(self):
        text = read("AGENTS.vi.md")
        self.assertIn(AGENTS_VI_EXAM_HEADING, h2_headings(text))
        body = section(text, AGENTS_VI_EXAM_HEADING)
        self.assertIn(EXAM_COMMAND, body)
        for phrase in (
            "source_to_md.py",
            "(docs/vi/tro-ly/de-khtn-tieng-anh.md)",
            "(docs/vi/tro-ly/tieng-anh-khoa-hoc.md)",
            "projects/_de-thi/",
            "de.md",
            "ảnh",
        ):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, body)

    def test_agents_vi_exam_section_maps_every_error_step(self):
        body = section(read("AGENTS.vi.md"), AGENTS_VI_EXAM_HEADING)
        for step in ("input", "parse", "docx", "write"):
            self.assertIn(f"`{step}`", body)
        self.assertIn("requirements-vi.txt", body)

    def test_agents_vi_exam_section_keeps_the_venv_conditional(self):
        body = section(read("AGENTS.vi.md"), AGENTS_VI_EXAM_HEADING)
        self.assertIn("mục 4", body)

    def test_agents_vi_triggers_include_exam_phrases(self):
        body = section(read("AGENTS.vi.md"), "## 3. Câu lệnh tiếng Việt kích hoạt skill `ppt-master`")
        for phrase in ("soạn đề", "đề tiếng Anh"):
            self.assertIn(phrase, body)

    def test_brief_template_lists_the_exam_task(self):
        self.assertIn("Soạn đề KHTN tiếng Anh", read("docs/vi/tro-ly/mau-brief.md"))
```

- [ ] **Step 2: Chạy test để thấy nó fail**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k ExamGuideTest`
Kỳ vọng: FAIL vì chưa có file hướng dẫn.

- [ ] **Step 3: Viết `docs/vi/tro-ly/de-khtn-tieng-anh.md`**

Tạo file với đúng bảy mục h2 theo thứ tự `EXAM_GUIDE_HEADINGS`. Nội dung bắt buộc:

- Dòng 3: `File dành cho AI. Luôn đọc docs/vi/tro-ly/quy-trinh-hoi.md trước file này, và đọc docs/vi/tro-ly/tieng-anh-khoa-hoc.md trước khi viết câu hỏi tiếng Anh.`
- `## Khi nào dùng`: thầy cô cần đề kiểm tra KHTN bằng tiếng Anh, cho môn KHTN 6–9 hoặc Vật lí / Hoá học / Sinh học 10–12. Hai trường hợp: có đề tiếng Việt sẵn (luồng A), và tạo đề mới (luồng B). Hai câu lệnh ví dụ: "Chuyển đề giữa kì Hoá 11 này sang tiếng Anh", "Soạn đề Vật lí 10 tiếng Anh 45 phút chương động lực học". Nêu rõ: **thầy cô đưa ảnh chụp đề thì không đọc được**, xin bản PDF hoặc Word, không tự đoán nội dung đề.
- `## Câu hỏi bắt buộc`: đúng bảy câu đánh số theo §6.4 của spec, mỗi câu kèm `Gợi ý:`. Câu 4 phải ghi: `Gợi ý: THPT dùng 18 – 4 – 6 theo đề tham khảo 2025. Môn KHTN ở THCS không có định dạng chung nên phải hỏi, không được lấy gợi ý làm mặc định.`
- `## Câu hỏi tuỳ chọn`: chỉ hỏi khi thầy cô nhắc tới — có cần mã đề không; có cần in chỗ trống cho học sinh trình bày bài tự luận không.
- `## Tạo nhanh`: hai câu — môn và lớp; số câu mỗi phần.
- `## Cấu trúc đề`: Phần I trắc nghiệm bốn lựa chọn, Phần II đúng/sai bốn ý, Phần III trả lời ngắn; thang điểm 0,25 / 1,0 / 0,25; nêu ngữ pháp `de.md` với các khoá `en:`, `vi:`, `A:`–`D:`, `a:`–`d:` kèm ` | T`/` | F`, `key:`, `unit:`, `why:`, `level:`, `topic:`, ba dấu đánh dấu `~ ~`, `^ ^`, `**`, và mục `## CAN SOAT` ở cuối.
- `## Đầu ra`: `projects/_de-thi/<tên_đề>/` chứa `de.md`, `de-en.docx`, `de-song-ngu.docx`, `dap-an.docx`.
- `## Ghi vào brief`: loại việc "Soạn đề KHTN tiếng Anh", ghi đúng lời thầy cô, ghi các mục AI đề xuất theo `mau-brief.md`.

Điều cấm phải ghi nguyên văn trong file: `Chuyển từ đề tiếng Việt thì không đổi số liệu, không đổi đáp án đúng, không đổi thứ tự câu. Câu gốc sai hoặc mơ hồ thì không tự sửa, ghi vào mục "Cần thầy cô soát".`

Không dùng liên kết Markdown trong file này.

- [ ] **Step 4: Viết `docs/vi/tro-ly/tieng-anh-khoa-hoc.md`**

Tạo file, dòng đầu `# Viết tiếng Anh cho đề khoa học tự nhiên`, dòng 3 `File dành cho AI. Đọc cùng docs/vi/tro-ly/de-khtn-tieng-anh.md.` Chín mục, mỗi mục là một nguyên tắc lấy đúng từ §3 của spec:

1. Dịch theo khái niệm, **không dịch từng chữ** — ví dụ sai `uniformly variable rectilinear motion`, đúng `uniformly accelerated motion`.
2. Thuật ngữ theo môn, kèm bảng ví dụ có đủ bốn khung môn `Vật lí`, `Hoá học`, `Sinh học`, `KHTN 6–9`; bắt buộc có các cặp: lực ma sát trượt → `kinetic friction`, suất điện động → `electromotive force (emf)`, công của lực → `work done by a force`, khối lượng mol → `molar mass`, hiệu suất phản ứng → `percentage yield`, nồng độ mol → `molar concentration`, hô hấp tế bào → `cellular respiration`, trao đổi chất → `metabolism`, cơ thể → `organism`, khối lượng riêng → `density`, tốc độ → `speed`.
3. Động từ lệnh hỏi: `State`, `Describe`, `Explain`, `Calculate`, `Determine`, `Deduce`, `Compare`, `Predict`; khuôn trắc nghiệm `Which of the following statements is correct?`; dạng phủ định in đậm chữ `not` bằng `**not**`.
4. Chính tả IUPAC/Anh-Anh: `sulfuric acid`, `sulfate`, `aluminium`; ký hiệu nguyên tố và phương trình giữ nguyên.
5. Số và đơn vị: dấu thập phân `25,5` → `25.5`; `1.000.000` → `1 000 000`; khoảng trắng giữa số và đơn vị (`5 kg`, `25 °C`); "ở đktc" viết rõ `at 0 °C and 1 atm`, **không** dịch thành `at STP`.
6. Ngữ cảnh Việt Nam giữ nguyên kèm giải thích ngắn: `terraced fields`, `fish sauce (a traditional Vietnamese condiment)`; không thay bằng ngữ cảnh nước ngoài.
7. `Độ khó nằm ở khoa học, không ở tiếng Anh`: câu dẫn ngắn, một mệnh đề chính, bốn đáp án cùng dạng ngữ pháp và xấp xỉ bằng nhau về độ dài.
8. Chuyển từ đề Việt thì không đổi nội dung; câu chơi chữ tiếng Việt thì nói rõ không chuyển được và chờ thầy cô quyết.
9. Tự soát trước khi xuất, rồi ghi từng chỗ chưa chắc vào mục `## CAN SOAT` của `de.md` — mục đó in ra thành `Cần thầy cô soát` trong file đáp án.

Không dùng liên kết Markdown trong file này.

- [ ] **Step 5: Sửa `AGENTS.vi.md`**

1. Mục 3: thêm vào danh sách câu lệnh `"soạn đề"`, `"làm đề kiểm tra"`, `"đề tiếng Anh"`.
2. Đổi tiêu đề mục 10 thành `## 10. Hỗ trợ thầy cô trước khi làm bài`. Trong thân mục có **hai** chỗ đếm, đổi cả hai: câu mở "một trong 6 loại việc dưới đây" thành "một trong 7 loại việc dưới đây", và gạch đầu dòng cuối "Yêu cầu không thuộc 6 loại" thành "Yêu cầu không thuộc 7 loại" (test kiểm cả thân mục không còn chuỗi "6 loại"). Rồi thêm dòng bảng:

```
| Soạn đề KHTN tiếng Anh | [docs/vi/tro-ly/de-khtn-tieng-anh.md](docs/vi/tro-ly/de-khtn-tieng-anh.md) |
```

3. Thêm mục 12 làm mục cuối file (sau mục 11), tiêu đề `## 12. Soạn đề KHTN bằng tiếng Anh`, nội dung:

- Câu mở: đọc `docs/vi/tro-ly/de-khtn-tieng-anh.md` và `docs/vi/tro-ly/tieng-anh-khoa-hoc.md`, hỏi một lượt theo file đó, rồi làm theo thứ tự dưới. Nhắc như mục 4: có `venv\Scripts\python.exe` thì dùng nó cho mọi lệnh Python dưới đây.
- Bước 1 (chỉ luồng A): đọc đề thầy cô đưa bằng `python skills/ppt-master/scripts/source_to_md.py <file> -o <thư_mục_tạm>`. Thầy cô đưa **ảnh** thì nói rõ không đọc được ảnh, xin PDF hoặc Word.
- Bước 2: tạo `projects/_de-thi/<tên_đề>/` và viết `de.md` theo ngữ pháp trong file hướng dẫn.
- Bước 3: `python tools\vi\de_thi.py projects\_de-thi\<tên_đề>`; thêm `--phan de,dap-an` khi thầy cô không cần bản song ngữ; thêm `--plan-only` khi chỉ muốn kiểm cú pháp.
- Bước 4: đọc dòng JSON ở stdout. `ready` là `true` thì báo thầy cô đường dẫn ba file, số câu mỗi phần, tổng điểm, và đọc nguyên văn các dòng `warnings`.
- Bảng xử lý lỗi, mỗi `error.step` một dòng: `input` → chưa có `de.md`, viết file rồi chạy lại; `parse` → sửa đúng dòng `error.message` nêu; `docx` → chạy `python -m pip install -r tools/vi/requirements-vi.txt` rồi chạy lại, tối đa một lần; `write` → xin thầy cô đóng file Word đang mở rồi chạy lại.
- Điều cấm: không tự sửa số liệu hay đáp án của đề gốc; không chạy `project_manager.py init`; không commit gì trong `projects/`.

- [ ] **Step 6: Sửa `docs/vi/tro-ly/quy-trinh-hoi.md`**

1. Đổi tiêu đề h1 thành `# Quy trình hỏi thầy cô trước khi làm bài`.
2. Trong `## Khi nào áp dụng`, đổi "6 loại việc" thành "7 loại việc" và thêm dòng bảng:

```
| Soạn đề KHTN tiếng Anh | "soạn đề", "đề kiểm tra", "đề tiếng Anh", "đề KHTN", "chuyển đề sang tiếng Anh" | [de-khtn-tieng-anh.md](de-khtn-tieng-anh.md) |
```

3. Trong cùng mục, thêm một dòng gạch đầu dòng: `Loại việc "Soạn đề KHTN tiếng Anh" không tạo PPTX; nó ghi brief như các loại khác nhưng không đi vào quy trình của upstream, xem docs/vi/tro-ly/de-khtn-tieng-anh.md.`

- [ ] **Step 7: Sửa `docs/vi/tro-ly/mau-brief.md`**

Dòng 5 của file hiện là `- Loại việc: <Bài giảng | Báo cáo – tổng kết | Hoạt động Đoàn – sự kiện | Poster/ấn phẩm Zalo – Facebook | Tập huấn/workshop | Video bài giảng>`. Thêm ` | Soạn đề KHTN tiếng Anh` ngay trước dấu `>` cuối dòng; không đổi gì khác.

- [ ] **Step 8: Chạy cả bộ test**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests`
Kỳ vọng: OK. Nếu `test_relative_markdown_links_resolve` fail, kiểm lại liên kết trong `AGENTS.vi.md` và `quy-trinh-hoi.md` trỏ đúng file đã tạo.

- [ ] **Step 9: Commit**

```bash
git add docs/vi/tro-ly/de-khtn-tieng-anh.md docs/vi/tro-ly/tieng-anh-khoa-hoc.md docs/vi/tro-ly/quy-trinh-hoi.md docs/vi/tro-ly/mau-brief.md AGENTS.vi.md tools/vi/tests/test_vi_layer.py
git commit -m "docs(vi): add the English KHTN exam task type for agents"
```

---

### Task 7: Tài liệu cho thầy cô

**Files:**
- Create: `docs/vi/soan-de-tieng-anh.md`
- Modify: `docs/vi/xu-ly-loi.md` (thêm `## Xuất đề Word thất bại`)
- Modify: `docs/vi/bat-dau-nhanh.md` (thêm `## Soạn đề tiếng Anh`)
- Modify: `docs/vi/cau-lenh-mau.md` (thêm `## Soạn đề tiếng Anh`)
- Modify: `README.md` (mục "Làm được gì", bảng tài liệu)
- Modify: `tools/vi/tests/test_vi_layer.py` (thêm `ExamUserDocsTest`, thêm file vào `REQUIRED_DOCS`)

**Interfaces:**
- Consumes: lệnh và các `error.step` ở Task 4, lệnh cài ở Task 5.
- Produces: không có mã.

- [ ] **Step 1: Viết test trước**

Thêm `"soan-de-tieng-anh.md"` vào `REQUIRED_DOCS` trong `tools/vi/tests/test_vi_layer.py`, rồi thêm lớp test:

```python
class ExamUserDocsTest(unittest.TestCase):
    def test_exam_doc_explains_inputs_outputs_and_limits(self):
        text = read("docs/vi/soan-de-tieng-anh.md")
        for phrase in ("Word", "PDF", "ảnh", "de-en.docx", "dap-an.docx", "song ngữ", "ma trận", "Cần thầy cô soát"):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, text)

    def test_quick_start_mentions_the_exam_task(self):
        text = read("docs/vi/bat-dau-nhanh.md")
        self.assertNotIn("6 loại", text)
        headings = h2_headings(text)
        self.assertIn("## Soạn đề tiếng Anh", headings)
        self.assertLess(headings.index("## Soạn đề tiếng Anh"), headings.index("## Lấy file kết quả"))
        self.assertIn("(soan-de-tieng-anh.md)", section(text, "## Soạn đề tiếng Anh"))

    def test_sample_commands_cover_both_use_cases(self):
        body = section(read("docs/vi/cau-lenh-mau.md"), "## Soạn đề tiếng Anh")
        self.assertIn("sang tiếng Anh", body)
        self.assertIn("Soạn đề", body)

    def test_troubleshooting_has_the_exam_export_section(self):
        text = read("docs/vi/xu-ly-loi.md")
        headings = h2_headings(text)
        self.assertIn("## Xuất đề Word thất bại", headings)
        # Đứng NGAY TRƯỚC mục video: test gói video khoá mục video phải ngay trước "Đường dẫn quá dài".
        self.assertEqual(
            headings.index("## Xuất đề Word thất bại"),
            headings.index("## Dựng video thất bại") - 1,
        )
        body = section(text, "## Xuất đề Word thất bại")
        for phrase in ("requirements-vi.txt", "python-docx", "đang mở trong Word", "Dòng", "de.md"):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, body)

    def test_readme_mentions_the_exam_feature(self):
        self.assertIn("đề kiểm tra", section(read("README.md"), "## Làm được gì"))
```

- [ ] **Step 2: Chạy test để thấy nó fail**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k ExamUserDocsTest`
Kỳ vọng: FAIL vì chưa có `docs/vi/soan-de-tieng-anh.md`.

- [ ] **Step 3: Viết `docs/vi/soan-de-tieng-anh.md`**

Viết cho thầy cô, không viết cho lập trình viên. Các mục:

- `# Soạn đề KHTN bằng tiếng Anh`
- `## Làm được gì`: từ đề tiếng Việt có sẵn, hoặc từ con số không; ra ba file Word — đề tiếng Anh, đề song ngữ để tổ soát, đáp án kèm ma trận đặc tả.
- `## Cách nhắn cho AI`: hai ví dụ câu lệnh, một cho mỗi trường hợp.
- `## AI sẽ hỏi gì`: bảy câu, nêu mấy câu quan trọng nhất (môn/lớp, số câu mỗi phần, đề dùng để làm gì).
- `## Thầy cô cần đưa gì`: file Word hoặc PDF. **Ảnh chụp đề thì AI không đọc được** — chụp thành PDF hoặc gõ lại.
- `## File nhận được`: bảng ba file kèm chỗ lưu `projects\_de-thi\<tên_đề>\`.
- `## Sửa lại rồi xuất lại`: sửa `de.md` rồi nhờ AI chạy lại, không phải làm lại từ đầu.
- `## Cần đọc lại trước khi in`: mục "Cần thầy cô soát" ở cuối file đáp án; máy không kiểm được thuật ngữ nên thầy cô đọc lại phần đó; kiểm cả thang điểm.
- `## Giới hạn`: chưa nạp được mẫu đầu đề riêng của trường và chưa chèn logo; chưa trộn mã đề.

- [ ] **Step 4: Sửa `docs/vi/xu-ly-loi.md`**

Thêm mục `## Xuất đề Word thất bại` **ngay trước** `## Dựng video thất bại`. Không chèn vào giữa `## Dựng video thất bại` và `## Đường dẫn quá dài`: test của gói video khoá hai mục đó phải liền nhau. Viết theo từng mã lỗi:

- `error.step` là `input`: chưa có `de.md` trong thư mục đề — nhờ AI viết file rồi chạy lại.
- `error.step` là `parse`: `error.message` nêu đúng số dòng; mở `de.md`, sửa dòng đó.
- `error.step` là `docx`: máy chưa có `python-docx`; chạy `python -m pip install -r tools/vi/requirements-vi.txt` hoặc bấm đúp `CAI-DAT.bat`.
- `error.step` là `write`: file Word đang mở trong Word, hoặc hết dung lượng, hoặc đường dẫn quá 200 ký tự.
- `error.step` là `internal`: dán nguyên dòng `error.message` gửi người bảo trì.

- [ ] **Step 5: Sửa `docs/vi/bat-dau-nhanh.md` và `docs/vi/cau-lenh-mau.md`**

`bat-dau-nhanh.md`: thêm mục `## Soạn đề tiếng Anh` trước `## Lấy file kết quả`, có liên kết `(soan-de-tieng-anh.md)`. Đồng thời đổi câu ở dòng 47 "Với 6 loại việc trên" thành "Với 7 loại việc trên" — spec bắt mọi chỗ đếm thành 7, và test đã thêm phép kiểm file không còn chuỗi "6 loại".

`cau-lenh-mau.md`: thêm mục `## Soạn đề tiếng Anh` với hai câu lệnh mẫu — một câu chuyển đề có sẵn ("Chuyển đề giữa kì Hoá 11 ở file này sang tiếng Anh"), một câu tạo đề mới ("Soạn đề Vật lí 10 tiếng Anh, 45 phút, chương động lực học").

- [ ] **Step 6: Sửa `README.md`**

1. Mục "Làm được gì": thêm một dòng nêu soạn **đề kiểm tra** KHTN bằng tiếng Anh, xuất Word.
2. Bảng tài liệu: thêm dòng `| [Soạn đề tiếng Anh](docs/vi/soan-de-tieng-anh.md) | Từ đề tiếng Việt hoặc từ đầu, ra ba file Word |`.

- [ ] **Step 7: Chạy cả bộ test**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests`
Kỳ vọng: OK.

- [ ] **Step 8: Commit**

```bash
git add docs/vi/soan-de-tieng-anh.md docs/vi/xu-ly-loi.md docs/vi/bat-dau-nhanh.md docs/vi/cau-lenh-mau.md README.md tools/vi/tests/test_vi_layer.py
git commit -m "docs(vi): explain English exam authoring for teachers"
```

---

### Task 8: Chạy thật hai đề mẫu

**Files:**
- Create: `docs/vi/phat-trien/2026-09-12-de-thi-kiem-thu.md`
- Không commit gì trong `projects/` (đã gitignore).

**Interfaces:**
- Consumes: toàn bộ Task 1–7.
- Produces: báo cáo chạy thật.

Ràng buộc: **không mở Word**. Kiểm bằng cách đọc XML trong file docx.

- [ ] **Step 1: Dựng đề mẫu luồng A (chuyển từ tiếng Việt)**

Tạo `projects/_de-thi/thu-nghiem-ly-10/de.md`: 5 câu (3 câu Phần I, 1 câu Phần II, 1 câu Phần III) môn `PHYSICS — Grade 10`, chương động lực học, có `vi:` đủ mọi câu, có `why:`, có `topic:`, có `## CAN SOAT` với một dòng. Đặt `points: 10` và chấp nhận cảnh báo lệch điểm — đề mẫu ngắn nên lệch là đúng.

- [ ] **Step 2: Chạy và kiểm**

```bash
python tools/vi/de_thi.py projects/_de-thi/thu-nghiem-ly-10
```

Kiểm: mã thoát 0; stdout đúng một dòng JSON; `files` có ba đường dẫn; ba file tồn tại; `warnings` có dòng lệch điểm.

Rồi kiểm nội dung không mở Word:

```bash
python -c "import zipfile,sys; p='projects/_de-thi/thu-nghiem-ly-10/de-en.docx'; x=zipfile.ZipFile(p).read('word/document.xml').decode('utf-8'); print('subscript' in x or 'superscript' in x, 'Full name' in x, 'THE END' in x)"
```

Kỳ vọng: `True True True`.

- [ ] **Step 3: Dựng đề mẫu luồng B (tạo mới) môn KHTN 8**

Tạo `projects/_de-thi/thu-nghiem-khtn-8/de.md`: 5 câu môn `NATURAL SCIENCES — Grade 8`, có ít nhất một công thức hoá học dùng `~ ~` (ví dụ `H~2~SO~4~`) và một đơn vị dùng `^ ^` (ví dụ `m/s^2^`), hai `topic:` khác nhau để ma trận có hai dòng, và **không** có mục `## CAN SOAT`.

- [ ] **Step 4: Chạy, kiểm đáp án và ma trận**

```bash
python tools/vi/de_thi.py projects/_de-thi/thu-nghiem-khtn-8 --phan dap-an
```

Kiểm bằng XML: `dap-an.docx` có `Ma trận đặc tả`, có `Thang điểm`, có `Không có mục nào cần soát`, và có đúng hai tên chủ đề đã đặt.

- [ ] **Step 5: Kiểm `--plan-only` không ghi file**

Xoá ba file docx trong `projects/_de-thi/thu-nghiem-khtn-8/`, rồi chạy:

```bash
python tools/vi/de_thi.py projects/_de-thi/thu-nghiem-khtn-8 --plan-only
```

Kỳ vọng: `ready` là `true`, `files` là danh sách rỗng, và trong thư mục không có file `.docx` nào.

- [ ] **Step 6: Kiểm ca thiếu thư viện**

```bash
python -c "import sys; sys.argv=['de_thi.py','projects/_de-thi/thu-nghiem-khtn-8']; sys.path.insert(0,'tools/vi'); import de_thi; de_thi.load_docx_build=lambda: (_ for _ in ()).throw(ImportError('No module named docx')); sys.exit(de_thi.main())"
```

Kỳ vọng: một dòng JSON có `"step": "docx"` và `fix` nêu `requirements-vi.txt`; mã thoát 1.

- [ ] **Step 7: Viết báo cáo**

Tạo `docs/vi/phat-trien/2026-09-12-de-thi-kiem-thu.md` ghi: lệnh đã chạy, dòng JSON thật (rút gọn đường dẫn về dạng tương đối, **không** để lộ tên thư mục riêng của chủ repo), số câu, tổng điểm, các cảnh báo, và kết quả kiểm XML. Ghi rõ một dòng: `Chưa mở Word để xem; phần đánh giá trình bày do chủ repo tự kiểm.`

- [ ] **Step 8: Chạy cả bộ test lần cuối**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests`
rồi `python skills/ppt-master/scripts/attribution_guard.py; echo $?`
Kỳ vọng: test OK, guard trả 0.

- [ ] **Step 9: Commit**

```bash
git add docs/vi/phat-trien/2026-09-12-de-thi-kiem-thu.md
git commit -m "docs(vi): record the real run of the exam exporter"
```

Kiểm `git status --porcelain` không có gì trong `projects/`.

---

## Tự soát của người viết plan

**Phủ spec:** §1 tiêu chí 1–2 → Task 6 (lượt hỏi) + Task 4 (lệnh) + Task 8 (chạy thật); tiêu chí 3 → Task 6 file `tieng-anh-khoa-hoc.md`; tiêu chí 4 → Task 6 điều cấm; tiêu chí 5 → Task 2 `review_notes` + Task 3 `_add_review`; tiêu chí 6 → Task 5; tiêu chí 7 → Global Constraints + test `UpstreamBoundaryTest` có sẵn; tiêu chí 8 → mọi task kết bằng bước chạy cả bộ test. §4.3 bốn phép kiểm đang khoá → Task 6 Step 1 và Task 7 Step 1. §6.1 ngữ pháp → Task 2. §6.2 thang điểm → Task 2 `total_points` + Task 3 `_add_scoring`. §6.3 lệnh và JSON → Task 4. §6.4 câu hỏi → Task 6 Step 3. §7 mọi ca biên → Task 2 và Task 4. §8 → Task 8. §9 phát hành → ngoài plan, làm sau khi review cuối xanh.

**Không có chỗ trống:** không có "TBD", mọi bước code đều có khối mã đầy đủ; các bước viết tài liệu nêu đủ nội dung bắt buộc và chuỗi mà test sẽ tìm.

**Nhất quán kiểu:** `Question`/`Exam` khai ở Task 2 và dùng đúng tên trường ở Task 3; `docx_build.FILENAMES` khai ở Task 3 và dùng ở Task 4; `load_docx_build` khai ở Task 4 và được test mock ở Task 4 Step 1 và Task 8 Step 6; `check_packages(name=, level=, fix=)` khai ở Task 5 và test ở cùng task.
