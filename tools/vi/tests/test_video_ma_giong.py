"""Test giọng đọc: file có sẵn, giọng máy (giả), sổ giọng, lỗi. Không bao giờ gọi giọng thật."""

import json
import sys
import tempfile
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))

from video_ma_parts import giong  # noqa: E402
from video_parts import media  # noqa: E402


class FakeTts:
    def __init__(self, marks=(0.0, 2.0), fail=False, partial=False):
        self.calls = []
        self.marks = list(marks)
        self.fail = fail
        self.partial = partial

    def __call__(self, text, voice, rate, out_path):
        self.calls.append((text, voice, rate))
        Path(out_path).write_bytes(b"ID3fake")
        if self.partial:
            raise media.MediaError("giong", "mất mạng giữa chừng", giong.FIX_GIONG)
        if self.fail:
            raise media.MediaError("giong", "mất mạng", giong.FIX_GIONG)
        return list(self.marks)


def fixed(giay):
    return lambda path: giay


class GiongTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name) / "giong"
        self.addCleanup(self.tmp.cleanup)

    def get(self, so=1, loi="Xin chào. Tạm biệt.", tts=None, do_dai=None, voice="nu", rate="vua"):
        return giong.lay_giong(so, loi, self.dir, voice, rate, tong_hop=tts or FakeTts(), do_dai=do_dai or fixed(4.0))

    def test_machine_voice_writes_mp3_and_ledger(self):
        tts = FakeTts()
        info = self.get(tts=tts)
        self.assertEqual(info.nguon, "may")
        self.assertEqual(info.giay, 4.0)
        self.assertEqual(info.moc_cau, [0.0, 2.0])
        self.assertFalse(info.uoc_luong)
        self.assertEqual(tts.calls, [("Xin chào. Tạm biệt.", "vi-VN-HoaiMyNeural", "+0%")])
        ledger = json.loads((self.dir / "canh-1.json").read_text(encoding="utf-8"))
        self.assertEqual(ledger["moc"], [0.0, 2.0])
        self.assertEqual(ledger["bam"], giong.bam("Xin chào. Tạm biệt.", "vi-VN-HoaiMyNeural", "+0%"))

    def test_voice_and_rate_choices(self):
        tts = FakeTts()
        self.get(tts=tts, voice="nam", rate="nhanh")
        self.assertEqual(tts.calls[0][1:], ("vi-VN-NamMinhNeural", "+15%"))

    def test_same_text_reuses_the_saved_voice(self):
        self.get()
        tts = FakeTts()
        info = self.get(tts=tts)
        self.assertEqual(tts.calls, [])
        self.assertEqual(info.moc_cau, [0.0, 2.0])

    def test_edited_text_regenerates_only_that_scene(self):
        self.get(so=1)
        self.get(so=2, loi="Cảnh hai.")
        tts = FakeTts()
        self.get(so=1, loi="Lời đã sửa.", tts=tts)
        self.assertEqual(len(tts.calls), 1)
        tts2 = FakeTts()
        self.get(so=2, loi="Cảnh hai.", tts=tts2)
        self.assertEqual(tts2.calls, [])

    def test_teacher_supplied_mp3_is_used_and_never_overwritten(self):
        self.dir.mkdir(parents=True)
        (self.dir / "canh-1.mp3").write_bytes(b"ID3thay-co")
        tts = FakeTts()
        info = self.get(tts=tts)
        self.assertEqual(info.nguon, "co-san")
        self.assertTrue(info.uoc_luong)
        self.assertEqual(info.moc_cau, [])
        self.assertEqual(tts.calls, [])
        info2 = self.get(loi="Lời khác hẳn.", tts=tts)
        self.assertEqual(info2.nguon, "co-san")
        self.assertEqual((self.dir / "canh-1.mp3").read_bytes(), b"ID3thay-co")
        self.assertFalse((self.dir / "canh-1.json").exists())

    def test_ledger_records_size_and_sha256_of_the_machine_mp3(self):
        self.get()
        ledger = json.loads((self.dir / "canh-1.json").read_text(encoding="utf-8"))
        self.assertEqual(ledger["kich_thuoc"], len(b"ID3fake"))
        import hashlib
        self.assertEqual(ledger["sha256"], hashlib.sha256(b"ID3fake").hexdigest())

    def test_teacher_mp3_over_a_stale_ledger_is_co_san_with_same_text(self):
        self.get()
        (self.dir / "canh-1.mp3").write_bytes(b"ID3thay-co-moi")
        tts = FakeTts()
        info = self.get(tts=tts)
        self.assertEqual(info.nguon, "co-san")
        self.assertTrue(info.uoc_luong)
        self.assertEqual(info.moc_cau, [])
        self.assertEqual(tts.calls, [])
        self.assertEqual((self.dir / "canh-1.mp3").read_bytes(), b"ID3thay-co-moi")

    def test_teacher_mp3_over_a_stale_ledger_is_never_overwritten_when_text_changes(self):
        self.get()
        (self.dir / "canh-1.mp3").write_bytes(b"ID3thay-co-moi")
        tts = FakeTts()
        info = self.get(loi="Lời đã sửa hẳn.", tts=tts)
        self.assertEqual(info.nguon, "co-san")
        self.assertEqual(tts.calls, [])
        self.assertEqual((self.dir / "canh-1.mp3").read_bytes(), b"ID3thay-co-moi")

    def test_ledger_without_size_and_hash_means_teacher_file(self):
        self.dir.mkdir(parents=True)
        (self.dir / "canh-1.mp3").write_bytes(b"ID3fake")
        ma = giong.bam("Xin chào. Tạm biệt.", "vi-VN-HoaiMyNeural", "+0%")
        (self.dir / "canh-1.json").write_text(json.dumps({"bam": ma, "moc": [0.0, 2.0]}), encoding="utf-8")
        tts = FakeTts()
        info = self.get(loi="Lời khác.", tts=tts)
        self.assertEqual(info.nguon, "co-san")
        self.assertEqual(tts.calls, [])

    def test_ledger_is_written_before_the_mp3_appears(self):
        seen = []
        real_replace = giong.os.replace

        def replace(src, dst):
            seen.append((self.dir / "canh-1.json").is_file())
            real_replace(src, dst)

        giong.os.replace = replace
        try:
            self.get()
        finally:
            giong.os.replace = real_replace
        self.assertEqual(seen, [True])

    def test_markup_is_stripped_before_synthesis_and_hashing(self):
        tts = FakeTts()
        self.get(loi="Học **H~2~SO~4~** và m/s^2^.", tts=tts)
        self.assertEqual(tts.calls[0][0], "Học H2SO4 và m/s2.")
        ledger = json.loads((self.dir / "canh-1.json").read_text(encoding="utf-8"))
        self.assertEqual(ledger["bam"], giong.bam("Học H2SO4 và m/s2.", "vi-VN-HoaiMyNeural", "+0%"))
        tts2 = FakeTts()
        self.get(loi="Học H2SO4 và m/s2.", tts=tts2)
        self.assertEqual(tts2.calls, [])

    def test_empty_supplied_file_is_a_giong_error_naming_the_file(self):
        self.dir.mkdir(parents=True)
        (self.dir / "canh-2.mp3").write_bytes(b"")
        with self.assertRaises(media.MediaError) as caught:
            self.get(so=2)
        self.assertEqual(caught.exception.step, "giong")
        self.assertIn("canh-2.mp3", str(caught.exception))

    def test_unreadable_supplied_file_is_a_giong_error(self):
        self.dir.mkdir(parents=True)
        (self.dir / "canh-1.mp3").write_bytes(b"khong phai mp3")

        def bad(path):
            raise media.MediaError("ffmpeg", f"ffprobe lỗi khi đọc {path.name} (mã 1).", media.FIX_FFMPEG)

        with self.assertRaises(media.MediaError) as caught:
            self.get(do_dai=bad)
        self.assertEqual(caught.exception.step, "giong")
        self.assertIn("canh-1.mp3", str(caught.exception))

    def test_missing_ffprobe_stays_an_ffmpeg_error(self):
        self.dir.mkdir(parents=True)
        (self.dir / "canh-1.mp3").write_bytes(b"ID3x")

        def missing(path):
            raise media.MediaError("ffmpeg", "Không chạy được ffprobe: [WinError 2]", media.FIX_FFMPEG)

        with self.assertRaises(media.MediaError) as caught:
            self.get(do_dai=missing)
        self.assertEqual(caught.exception.step, "ffmpeg")

    def test_failed_synthesis_leaves_no_half_written_mp3(self):
        with self.assertRaises(media.MediaError) as caught:
            self.get(tts=FakeTts(partial=True))
        self.assertEqual(caught.exception.step, "giong")
        self.assertFalse((self.dir / "canh-1.mp3").exists())
        self.assertFalse((self.dir / "canh-1.mp3.tmp").exists())
        info = self.get(tts=FakeTts())
        self.assertEqual(info.nguon, "may")

    def test_failed_replace_restores_the_old_ledger(self):
        import os
        self.get(so=1, loi="Lời cũ.")
        ledger_before = (self.dir / "canh-1.json").read_bytes()
        mp3_before = (self.dir / "canh-1.mp3").read_bytes()
        real_replace = os.replace

        def deny(src, dst):
            raise PermissionError("dang mo trong trinh phat")

        os.replace = deny
        try:
            with self.assertRaises(PermissionError):
                self.get(so=1, loi="Lời mới.", tts=FakeTts(marks=(0.0,)))
        finally:
            os.replace = real_replace
        self.assertEqual((self.dir / "canh-1.json").read_bytes(), ledger_before)
        self.assertEqual((self.dir / "canh-1.mp3").read_bytes(), mp3_before)
        self.assertFalse((self.dir / "canh-1.mp3.tmp").exists())
        tts = FakeTts(marks=(0.0,))
        info = self.get(so=1, loi="Lời mới.", tts=tts)
        self.assertEqual(info.nguon, "may")
        self.assertEqual(len(tts.calls), 1)

    def test_edge_tts_missing_message(self):
        import builtins
        real_import = builtins.__import__

        def no_edge(name, *args, **kwargs):
            if name == "edge_tts":
                raise ImportError("no edge_tts")
            return real_import(name, *args, **kwargs)

        builtins.__import__ = no_edge
        try:
            with self.assertRaises(media.MediaError) as caught:
                giong.tong_hop_edge("Xin chào.", "vi-VN-HoaiMyNeural", "+0%", self.dir / "x.mp3")
        finally:
            builtins.__import__ = real_import
        self.assertEqual(caught.exception.step, "giong")
        self.assertIn("edge-tts", str(caught.exception))


if __name__ == "__main__":
    unittest.main()
