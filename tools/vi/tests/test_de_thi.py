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
