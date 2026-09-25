"""Test bộ đọc cmap của font Itim đóng gói: đủ 134 chữ tiếng Việt có dấu."""

import sys
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))

from video_ma_parts import phong  # noqa: E402

SEGOE_PRINT = Path("C:/Windows/Fonts/segoepr.ttf")


class ChuVietTest(unittest.TestCase):
    def test_chu_viet_has_134_distinct_characters(self):
        self.assertEqual(len(phong.CHU_VIET), 134)
        self.assertEqual(len(set(phong.CHU_VIET)), 134)


class BangMaTest(unittest.TestCase):
    def test_itim_covers_every_vietnamese_letter(self):
        ma = phong.bang_ma(phong.FONT)
        thieu = [c for c in phong.CHU_VIET if ord(c) not in ma]
        self.assertEqual(thieu, [])
        for c in "ĐƯƠ":
            self.assertIn(ord(c), ma)

    @unittest.skipUnless(SEGOE_PRINT.is_file(), "máy không có segoepr.ttf")
    def test_reader_really_distinguishes_a_font_missing_letters(self):
        ma = phong.bang_ma(SEGOE_PRINT)
        self.assertNotIn(ord("ề"), ma)


class FontCssTest(unittest.TestCase):
    def test_font_css_embeds_the_font_as_base64_data_url(self):
        css = phong.font_css()
        self.assertIn("@font-face", css)
        self.assertIn(f"font-family:'{phong.TEN}'", css)
        self.assertIn("url(data:font/ttf;base64,", css)
        self.assertNotIn("http://", css)
        self.assertNotIn("https://", css)


if __name__ == "__main__":
    unittest.main()
