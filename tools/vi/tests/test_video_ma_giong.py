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
    def __init__(self, marks=(0.0, 2.0), fail=False, partial=False, tu=None):
        self.calls = []
        self.marks = list(marks)
        self.fail = fail
        self.partial = partial
        self.tu = tu

    def __call__(self, text, voice, rate, out_path):
        self.calls.append((text, voice, rate))
        Path(out_path).write_bytes(b"ID3fake")
        if self.partial:
            raise media.MediaError("giong", "mất mạng giữa chừng", giong.FIX_GIONG)
        if self.fail:
            raise media.MediaError("giong", "mất mạng", giong.FIX_GIONG)
        if self.tu is not None:
            return {"cau": list(self.marks), "tu": list(self.tu)}
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

    def test_highlight_markup_is_stripped_before_synthesis_and_hashing(self):
        tts = FakeTts()
        self.get(loi="==Chu kì== là ((thời gian)) của __một dao động__ và {{1500}} vòng.", tts=tts)
        self.assertEqual(tts.calls[0][0], "Chu kì là thời gian của một dao động và 1500 vòng.")
        ledger = json.loads((self.dir / "canh-1.json").read_text(encoding="utf-8"))
        self.assertEqual(ledger["bam"], giong.bam("Chu kì là thời gian của một dao động và 1500 vòng.",
                                                    "vi-VN-HoaiMyNeural", "+0%"))

    def test_dict_tts_writes_word_marks_and_ledger_reuses_them(self):
        tu = [
            {"t": 0.0, "d": 0.2, "chu": "Xin"},
            {"t": 0.2, "d": 0.2, "chu": "chào."},
            {"t": 2.0, "d": 0.2, "chu": "Tạm"},
            {"t": 2.2, "d": 0.2, "chu": "biệt."},
        ]
        tts = FakeTts(marks=(0.0, 2.0), tu=tu)
        info = self.get(tts=tts)
        self.assertEqual(info.moc_cau, [0.0, 2.0])
        self.assertEqual(info.moc_tu, tu)
        self.assertFalse(info.uoc_luong_tu)
        ledger = json.loads((self.dir / "canh-1.json").read_text(encoding="utf-8"))
        self.assertEqual(ledger["tu"], tu)
        tts2 = FakeTts(marks=(0.0, 2.0), tu=tu)
        info2 = self.get(tts=tts2)
        self.assertEqual(tts2.calls, [])
        self.assertEqual(info2.moc_tu, tu)
        self.assertFalse(info2.uoc_luong_tu)

    def test_old_ledger_without_tu_is_estimated(self):
        self.get()
        so_giong = self.dir / "canh-1.json"
        ghi = json.loads(so_giong.read_text(encoding="utf-8"))
        ghi.pop("tu", None)
        so_giong.write_text(json.dumps(ghi, ensure_ascii=False), encoding="utf-8")
        tts = FakeTts()
        info = self.get(tts=tts)
        self.assertEqual(tts.calls, [])
        self.assertEqual(info.moc_tu, [])
        self.assertTrue(info.uoc_luong_tu)

    def test_teacher_supplied_mp3_has_estimated_word_marks(self):
        self.dir.mkdir(parents=True)
        (self.dir / "canh-1.mp3").write_bytes(b"ID3thay-co")
        info = self.get(tts=FakeTts())
        self.assertEqual(info.nguon, "co-san")
        self.assertEqual(info.moc_tu, [])
        self.assertTrue(info.uoc_luong_tu)

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


class AnswerVoiceTest(unittest.TestCase):
    """Giọng lời giải của cảnh câu hỏi: file riêng `canh-N-giai.mp3` (Review Focus 2)."""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name) / "giong"
        self.addCleanup(self.tmp.cleanup)

    def giai(self, tts, so=3, loi="Đáp án B. Vì chu kì tăng."):
        return giong.lay_giong(so, loi, self.dir, "nu", "vua", tong_hop=tts, do_dai=fixed(2.5), ten=f"canh-{so}-giai")

    def test_named_voice_writes_its_own_mp3_and_ledger(self):
        info = self.giai(FakeTts(marks=(0.0, 1.0)))
        self.assertEqual(info.mp3, self.dir / "canh-3-giai.mp3")
        self.assertEqual(info.nguon, "may")
        self.assertEqual(info.moc_cau, [0.0, 1.0])
        self.assertTrue((self.dir / "canh-3-giai.json").is_file())
        self.assertFalse((self.dir / "canh-3.mp3").exists())
        self.assertFalse((self.dir / "canh-3.json").exists())
        tts = FakeTts()
        self.giai(tts)
        self.assertEqual(tts.calls, [], "lần sau dùng lại giọng lời giải đã tạo")

    def test_question_and_answer_voices_do_not_share_a_ledger(self):
        giong.lay_giong(3, "Câu hỏi?", self.dir, "nu", "vua", tong_hop=FakeTts(), do_dai=fixed(4.0))
        tts = FakeTts()
        self.giai(tts)
        self.assertEqual(len(tts.calls), 1)
        tts_hoi = FakeTts()
        giong.lay_giong(3, "Câu hỏi?", self.dir, "nu", "vua", tong_hop=tts_hoi, do_dai=fixed(4.0))
        self.assertEqual(tts_hoi.calls, [])

    def test_teacher_answer_mp3_is_used_and_never_overwritten(self):
        self.dir.mkdir(parents=True)
        (self.dir / "canh-3-giai.mp3").write_bytes(b"ID3giai-thay-co")
        tts = FakeTts()
        info = self.giai(tts)
        self.assertEqual((info.nguon, info.uoc_luong), ("co-san", True))
        self.assertEqual(tts.calls, [])
        self.giai(tts, loi="Lời giải đã sửa hẳn.")
        self.assertEqual(tts.calls, [])
        self.assertEqual((self.dir / "canh-3-giai.mp3").read_bytes(), b"ID3giai-thay-co")
        self.assertFalse((self.dir / "canh-3-giai.json").exists())

    def test_failed_answer_voice_is_a_giong_error_naming_the_answer_file(self):
        self.dir.mkdir(parents=True)
        (self.dir / "canh-3.mp3").write_bytes(b"ID3cau-hoi-thay-co")  # thầy cô chỉ đặt sẵn file câu hỏi
        hoi = giong.lay_giong(3, "Câu hỏi?", self.dir, "nu", "vua", tong_hop=FakeTts(), do_dai=fixed(4.0))
        self.assertEqual(hoi.nguon, "co-san")
        with self.assertRaises(media.MediaError) as caught:
            self.giai(FakeTts(fail=True))
        self.assertEqual(caught.exception.step, "giong")
        self.assertIn("canh-3-giai.mp3", str(caught.exception))
        self.assertIn("canh-3-giai.mp3", caught.exception.fix)
        self.assertFalse((self.dir / "canh-3-giai.mp3").exists())
        self.assertFalse((self.dir / "canh-3-giai.mp3.tmp").exists())
        self.assertEqual((self.dir / "canh-3.mp3").read_bytes(), b"ID3cau-hoi-thay-co")

    def test_failed_question_voice_names_the_question_file(self):
        with self.assertRaises(media.MediaError) as caught:
            giong.lay_giong(4, "Câu hỏi?", self.dir, "nu", "vua", tong_hop=FakeTts(fail=True), do_dai=fixed(4.0))
        self.assertIn("canh-4.mp3", str(caught.exception))

    def test_empty_teacher_answer_file_names_it(self):
        self.dir.mkdir(parents=True)
        (self.dir / "canh-3-giai.mp3").write_bytes(b"")
        with self.assertRaises(media.MediaError) as caught:
            self.giai(FakeTts())
        self.assertIn("canh-3-giai.mp3", str(caught.exception))


def _tu(*items):
    """items: (giay, chu) -> danh sách sự kiện WordBoundary giả, thời lượng 0.09s mỗi từ."""
    return [{"t": giay, "d": 0.09, "chu": chu} for giay, chu in items]


class SentenceMarkFromWordsTest(unittest.TestCase):
    def test_clean_case_is_unchanged(self):
        tu = _tu((0.125, "Chu"), (0.275, "kì."), (2.0, "Tạm"), (2.2, "biệt."))
        marks = giong._moc_cau_theo_tu("Chu kì. Tạm biệt.", tu)
        self.assertEqual(marks, [0.125, 2.0])

    def test_a_number_spoken_as_several_events_does_not_shift_the_next_sentence(self):
        # "1500" tổng hợp thành 4 sự kiện chữ số riêng lẻ; đếm số từ kịch bản (1 từ)
        # sẽ lệch mốc câu 2; phải so khớp theo chữ để câu 2 vẫn đúng tại 0.7.
        tu = _tu(
            (0.0, "Có"), (0.1, "1"), (0.2, "5"), (0.3, "0"), (0.4, "0"),
            (0.5, "vòng"), (0.6, "quay."),
            (0.7, "Kết"), (0.8, "thúc"), (0.9, "thí"), (1.0, "nghiệm."),
        )
        marks = giong._moc_cau_theo_tu("Có 1500 vòng quay. Kết thúc thí nghiệm.", tu)
        self.assertEqual(len(marks), 2)
        self.assertEqual(marks[0], 0.0)
        self.assertAlmostEqual(marks[1], 0.7, places=2)

    def test_a_merged_token_matches_across_its_split_events(self):
        # Kịch bản viết liền "H2O"; dịch vụ tách thành ba sự kiện "H", "2", "O".
        tu = _tu((0.0, "H"), (0.1, "2"), (0.2, "O"), (0.3, "là"), (0.4, "nước."), (0.5, "Ok."))
        marks = giong._moc_cau_theo_tu("H2O là nước. Ok.", tu)
        self.assertEqual(marks, [0.0, 0.5])

    def test_a_dropped_word_mid_sentence_does_not_break_the_sentence_marks(self):
        # Dịch vụ không phát ra sự kiện cho "các" (rớt từ), các từ khác vẫn khớp đúng.
        tu = _tu(
            (0.0, "Xin"), (0.1, "chào"), (0.3, "em"), (0.4, "hôm"), (0.5, "nay."),
            (0.6, "Tạm"), (0.7, "biệt."),
        )
        marks = giong._moc_cau_theo_tu("Xin chào các em hôm nay. Tạm biệt.", tu)
        self.assertEqual(marks, [0.0, 0.6])

    def test_unmatchable_first_word_falls_back_to_empty_instead_of_a_wrong_mark(self):
        # Không đủ sự kiện để tìm ra mốc câu 2 một cách chắc chắn: trả về [] để
        # lich.dung_lich tự ước lượng lại, không bao giờ đoán mốc sai một cách âm thầm.
        tu = _tu((0.0, "Một"))
        marks = giong._moc_cau_theo_tu("Một. Hai.", tu)
        self.assertEqual(marks, [])

    def test_no_events_at_all_falls_back_to_empty(self):
        self.assertEqual(giong._moc_cau_theo_tu("Một. Hai.", []), [])


if __name__ == "__main__":
    unittest.main()
