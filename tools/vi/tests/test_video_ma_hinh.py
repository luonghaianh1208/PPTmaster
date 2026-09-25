"""Test hình vẽ nét (tabler-outline) và ảnh thật cho video giải thích.

Không cần Chromium, FFmpeg hay mạng, trừ test JPEG (tự bỏ qua khi máy không có ffmpeg).
"""

import json
import shutil
import struct
import subprocess
import sys
import tempfile
import unittest
import zlib
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))

from video_ma_parts import anh, hinh, kiem, parse  # noqa: E402

META = "tieu-de: T\nmon: Toán\nlop: 8\n"
HAS_FFMPEG = shutil.which("ffmpeg") is not None


def doc(canh: str, meta: str = META) -> str:
    return f"---\n{meta}---\n\n{canh}"


def line_of(text: str, needle: str) -> int:
    for number, line in enumerate(text.splitlines(), 1):
        if needle in line:
            return number
    raise AssertionError(needle)


def _png(path: Path, width: int = 1, height: int = 1) -> None:
    def chunk(tag: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    raw = b"".join(b"\x00" + b"\xff\x00\x00" * width for _ in range(height))
    idat = zlib.compress(raw)
    path.write_bytes(sig + chunk(b"IHDR", ihdr) + chunk(b"IDAT", idat) + chunk(b"IEND", b""))


def _jpeg(path: Path, width: int = 4, height: int = 3) -> None:
    subprocess.run(
        ["ffmpeg", "-y", "-hide_banner", "-loglevel", "error", "-f", "lavfi",
         "-i", f"color=c=red:s={width}x{height}", "-frames:v", "1", str(path)],
        check=True, timeout=60,
    )


class ChuanTenTest(unittest.TestCase):
    def test_strips_prefix_suffix_case_and_spaces(self):
        self.assertEqual(hinh.chuan_ten("  tabler-outline/Flask.svg  "), "flask")
        self.assertEqual(hinh.chuan_ten("Flask"), "flask")
        self.assertEqual(hinh.chuan_ten("flask.svg"), "flask")
        self.assertEqual(hinh.chuan_ten("tabler-outline/flask"), "flask")


class HinhDocTest(unittest.TestCase):
    def test_flask_has_three_path_elements(self):
        h = hinh.doc("flask")
        self.assertEqual(h["ten"], "flask")
        self.assertEqual(h["viewBox"], "0 0 24 24")
        self.assertEqual([p["the"] for p in h["phanTu"]], ["path", "path", "path"])

    def test_atom_has_no_frame_path(self):
        h = hinh.doc("atom")
        self.assertTrue(h["phanTu"])
        for p in h["phanTu"]:
            self.assertNotEqual(p["thuocTinh"].get("d"), "M0 0h24v24H0z")
            self.assertNotEqual(p["thuocTinh"].get("stroke"), "none")

    def test_accepts_prefixed_and_suffixed_names(self):
        self.assertEqual(hinh.doc("tabler-outline/Flask.svg")["ten"], "flask")

    def test_unknown_name_raises_with_suggestions(self):
        with self.assertRaises(hinh.HinhError) as caught:
            hinh.doc("binh-thi-nghiem")
        self.assertIn("Có thể bạn muốn", str(caught.exception))


class TachMinhHoaTest(unittest.TestCase):
    def test_splits_name_and_label(self):
        self.assertEqual(hinh.tach_minh_hoa("clock | Đồng hồ bấm giây"), ("clock", "Đồng hồ bấm giây"))

    def test_missing_bar_raises(self):
        with self.assertRaises(hinh.HinhError) as caught:
            hinh.tach_minh_hoa("clock")
        self.assertIn("tên | nhãn", str(caught.exception))

    def test_missing_label_raises(self):
        with self.assertRaises(hinh.HinhError) as caught:
            hinh.tach_minh_hoa("clock | ")
        self.assertIn("tên | nhãn", str(caught.exception))

    def test_label_over_thirty_characters_raises(self):
        with self.assertRaises(hinh.HinhError):
            hinh.tach_minh_hoa("clock | " + "a" * 31)

    def test_label_at_thirty_characters_is_fine(self):
        ten, nhan = hinh.tach_minh_hoa("clock | " + "a" * 30)
        self.assertEqual((ten, nhan), ("clock", "a" * 30))


class KiemHinhTest(unittest.TestCase):
    def test_prefixed_name_passes_kiem(self):
        text = doc("## Cảnh 1\nloai: tieu-de\nchu: A\nhinh: tabler-outline/Flask.svg\nloi: Xin chào.\n")
        self.assertEqual(kiem.kiem(parse.parse(text), Path(".")), [])

    def test_unknown_icon_is_a_canh_error_with_line_number(self):
        text = doc("## Cảnh 1\nloai: tieu-de\nchu: A\nhinh: binh-thi-nghiem\nloi: Xin chào.\n")
        with self.assertRaises(kiem.CanhError) as caught:
            kiem.kiem(parse.parse(text), Path("."))
        self.assertEqual(caught.exception.so, 1)
        self.assertIn(f"dòng {line_of(text, 'binh-thi-nghiem')}", str(caught.exception))

    def test_minh_hoa_picture_starts_at_its_own_line(self):
        text = doc("## Cảnh 1\nloai: minh-hoa\ntieu-de: A\nhinh: clock | Đồng hồ\nhinh: binh-thi-nghiem | Xấu\nloi: Xin chào.\n")
        with self.assertRaises(kiem.CanhError) as caught:
            kiem.kiem(parse.parse(text), Path("."))
        self.assertIn(f"dòng {line_of(text, 'binh-thi-nghiem')}", str(caught.exception))

    def test_minh_hoa_label_over_limit_is_a_canh_error(self):
        text = doc(f"## Cảnh 1\nloai: minh-hoa\ntieu-de: A\nhinh: clock | {'a' * 31}\nloi: Xin chào.\n")
        with self.assertRaises(kiem.CanhError):
            kiem.kiem(parse.parse(text), Path("."))

    def test_minh_hoa_missing_bar_is_a_canh_error_naming_the_shape(self):
        text = doc("## Cảnh 1\nloai: minh-hoa\ntieu-de: A\nhinh: clock\nloi: Xin chào.\n")
        with self.assertRaises(kiem.CanhError) as caught:
            kiem.kiem(parse.parse(text), Path("."))
        self.assertIn("tên | nhãn", str(caught.exception))


class AnhDocTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.thu_muc = Path(self.tmp.name)
        (self.thu_muc / "anh").mkdir()

    def test_reads_png_dimensions(self):
        _png(self.thu_muc / "anh" / "x.png", 3, 2)
        info = anh.doc(self.thu_muc, "x.png", "Ảnh tự chụp")
        self.assertEqual((info["rong"], info["cao"]), (3, 2))
        self.assertTrue(info["dataUrl"].startswith("data:image/png;base64,"))
        self.assertEqual(info["nguon"], "Ảnh tự chụp")

    def test_missing_file_raises(self):
        with self.assertRaises(anh.AnhError):
            anh.doc(self.thu_muc, "khong-co.png", "nguồn")

    def test_bad_extension_raises(self):
        (self.thu_muc / "anh" / "x.gif").write_bytes(b"GIF89a")
        with self.assertRaises(anh.AnhError):
            anh.doc(self.thu_muc, "x.gif", "nguồn")

    def test_oversized_file_raises(self):
        path = self.thu_muc / "anh" / "to.png"
        path.write_bytes(b"\x89PNG\r\n\x1a\n" + b"\x00" * anh.TOI_DA)
        with self.assertRaises(anh.AnhError):
            anh.doc(self.thu_muc, "to.png", "nguồn")

    def test_missing_source_raises(self):
        _png(self.thu_muc / "anh" / "x.png")
        with self.assertRaises(anh.AnhError):
            anh.doc(self.thu_muc, "x.png", None)

    def test_manifest_source_has_author_and_license(self):
        _png(self.thu_muc / "anh" / "x.png")
        (self.thu_muc / "anh" / "image_sources.json").write_text(
            json.dumps({"items": [{"filename": "x.png", "author": "Nguyễn Văn A",
                                    "license_name": "CC BY 4.0", "provider": "openverse"}]}),
            encoding="utf-8",
        )
        info = anh.doc(self.thu_muc, "x.png", None)
        self.assertIn("Nguyễn Văn A", info["nguon"])
        self.assertIn("CC BY 4.0", info["nguon"])

    def test_hand_written_source_wins_over_manifest(self):
        _png(self.thu_muc / "anh" / "x.png")
        (self.thu_muc / "anh" / "image_sources.json").write_text(
            json.dumps({"items": [{"filename": "x.png", "author": "Manifest", "license_name": "CC0", "provider": "p"}]}),
            encoding="utf-8",
        )
        info = anh.doc(self.thu_muc, "x.png", "Ảnh tự chụp")
        self.assertEqual(info["nguon"], "Ảnh tự chụp")

    def test_filename_with_vietnamese_accents_and_spaces(self):
        _png(self.thu_muc / "anh" / "ảnh con lắc.png")
        info = anh.doc(self.thu_muc, "ảnh con lắc.png", "nguồn")
        self.assertEqual((info["rong"], info["cao"]), (1, 1))

    @unittest.skipUnless(HAS_FFMPEG, "máy không có ffmpeg")
    def test_reads_jpeg_dimensions(self):
        _jpeg(self.thu_muc / "anh" / "x.jpg", 4, 3)
        info = anh.doc(self.thu_muc, "x.jpg", "nguồn")
        self.assertEqual((info["rong"], info["cao"]), (4, 3))


class KiemAnhSceneTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.thu_muc = Path(self.tmp.name)
        (self.thu_muc / "anh").mkdir()
        _png(self.thu_muc / "anh" / "x.png")

    def test_anh_scene_with_its_own_source_passes(self):
        text = doc("## Cảnh 1\nloai: anh\nanh: x.png\nchu-thich: Chú thích\nnguon: Ảnh tự chụp\nloi: Xin chào.\n")
        self.assertEqual(kiem.kiem(parse.parse(text), self.thu_muc), [])

    def test_anh_scene_missing_source_is_a_canh_error(self):
        text = doc("## Cảnh 1\nloai: anh\nanh: x.png\nchu-thich: Chú thích\nloi: Xin chào.\n")
        with self.assertRaises(kiem.CanhError):
            kiem.kiem(parse.parse(text), self.thu_muc)


if __name__ == "__main__":
    unittest.main()
