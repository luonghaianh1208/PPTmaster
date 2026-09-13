"""Test cho lớp soạn đề KHTN tiếng Anh của bản Việt."""

import sys
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT / "tools" / "vi"))

from de_thi_parts import parse  # noqa: E402
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

    def test_superscript_digit_in_question_number_raises_parse_error(self):
        broken = VALID_SOURCE.replace("### 1\n", "### ²\n", 1)
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_exam(broken)
        self.assertNotIsInstance(caught.exception, ValueError)
        expected_line = broken.splitlines().index("### ²") + 1
        self.assertEqual(caught.exception.line_no, expected_line)

    def test_superscript_digit_in_time_raises_parse_error(self):
        broken = VALID_SOURCE.replace("time: 45", "time: ²")
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_exam(broken)
        self.assertIn("time", caught.exception.message)

    def test_part_three_key_must_be_numeric(self):
        broken = VALID_SOURCE.replace("key: 36", "key: ba mươi sáu")
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_exam(broken)
        self.assertIn("PART III", caught.exception.message)
        expected_line = broken.splitlines().index("key: ba mươi sáu") + 1
        self.assertEqual(caught.exception.line_no, expected_line)

    def test_part_three_key_accepts_decimal_and_negative_numbers(self):
        source_with_decimal = VALID_SOURCE.replace("key: 36", "key: 12,5")
        exam1 = parse.parse_exam(source_with_decimal)
        self.assertEqual(exam1.part(3)[0].key, "12,5")
        source_with_negative = VALID_SOURCE.replace("key: 36", "key: -0.25")
        exam2 = parse.parse_exam(source_with_negative)
        self.assertEqual(exam2.part(3)[0].key, "-0.25")


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


if __name__ == "__main__":
    unittest.main()
