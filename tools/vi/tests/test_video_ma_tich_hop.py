"""Tích hợp: dựng video 2 cảnh thật với tiếng giả. Tự bỏ qua nếu máy thiếu Chromium/playwright hoặc FFmpeg/ffprobe."""

import contextlib
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))

import video_ma  # noqa: E402
from video_ma_parts import lich  # noqa: E402

CO = video_ma.co_chromium() and video_ma.co_ffmpeg()
NEED = "máy thiếu Chromium/playwright hoặc FFmpeg/ffprobe"

VIDEO_MD = """---
tieu-de: Con lắc đơn
mon: Vật lí
lop: 11
phu-de: {phu_de}
---

## Cảnh 1
loai: y-tung-y
tieu-de: Chu kì phụ thuộc vào gì
y: Chiều dài dây l
y: Gia tốc trọng trường g
loi: Thứ nhất, chu kì phụ thuộc chiều dài dây. Thứ hai, chu kì phụ thuộc gia tốc trọng trường.

## Cảnh 2
loai: thi-nghiem
mau: li-con-lac-don
tham-so: 0 chieu-dai 0.4
tham-so: 3 chieu-dai 1.6
do: chu-ki
loi: Hãy quan sát chu kì.
"""


def tao_tieng(path: Path, giay: float) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    subprocess.run(["ffmpeg", "-y", "-hide_banner", "-loglevel", "error", "-f", "lavfi", "-i", f"sine=frequency=440:duration={giay}",
                    "-q:a", "9", str(path)], check=True, timeout=60)


def thong_so(video: Path) -> dict:
    out = subprocess.run(["ffprobe", "-v", "error", "-show_streams", "-show_format", "-of", "json", str(video)],
                         capture_output=True, text=True, encoding="utf-8", check=True, timeout=60).stdout
    return json.loads(out)


@unittest.skipUnless(CO, NEED)
class EndToEndTest(unittest.TestCase):
    def dung(self, ten: str, phu_de: str):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        thu_muc = Path(tmp.name) / ten
        thu_muc.mkdir()
        (thu_muc / "video.md").write_text(VIDEO_MD.format(phu_de=phu_de), encoding="utf-8")
        tao_tieng(thu_muc / "giong" / "canh-1.mp3", 3.0)
        tao_tieng(thu_muc / "giong" / "canh-2.mp3", 4.0)
        out = io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(io.StringIO()):
            code = video_ma.main([str(thu_muc)])
        lines = [l for l in out.getvalue().splitlines() if l.strip()]
        self.assertEqual(len(lines), 1, lines)
        return thu_muc, code, json.loads(lines[0])

    def test_builds_a_playable_video_in_a_hard_folder_name(self):
        thu_muc, code, data = self.dung("Bài 5 – Sulfur dioxide (thử) 'a' 50%", "hinh")
        self.assertEqual(code, 0, data)
        self.assertTrue(data["ready"], data)
        self.assertEqual(data["giong"], "co-san")
        self.assertTrue(any("ước lượng" in w for w in data["warnings"]))
        video = thu_muc / "video.mp4"
        self.assertTrue(video.is_file())
        info = thong_so(video)
        kinds = {s["codec_type"]: s for s in info["streams"]}
        self.assertEqual((kinds["video"]["width"], kinds["video"]["height"]), (1280, 720))
        self.assertEqual(kinds["video"]["r_frame_rate"], "30/1")
        self.assertIn("audio", kinds)
        expect = lich.thoi_luong_canh(3.0) + lich.thoi_luong_canh(4.0)
        self.assertAlmostEqual(float(info["format"]["duration"]), expect, delta=0.25)
        self.assertAlmostEqual(data["thoi_luong_giay"], expect, delta=0.01)
        self.assertFalse((thu_muc / ".khung").exists())
        self.assertEqual((thu_muc / "giong" / "canh-1.mp3").is_file(), True)

    def test_subtitle_file_mode_writes_srt_next_to_the_video(self):
        thu_muc, code, data = self.dung("phu de rieng", "file")
        self.assertEqual(code, 0, data)
        self.assertIn("phu-de.srt", data["files"])
        text = (thu_muc / "phu-de.srt").read_text(encoding="utf-8")
        self.assertIn("Thứ nhất, chu kì phụ thuộc chiều dài dây.", text)

    def test_rerun_after_editing_one_scene_reuses_supplied_voice(self):
        thu_muc, code, _ = self.dung("chay-lai", "khong")
        self.assertEqual(code, 0)
        before = (thu_muc / "giong" / "canh-1.mp3").read_bytes()
        md = thu_muc / "video.md"
        md.write_text(md.read_text(encoding="utf-8").replace("Hãy quan sát chu kì.", "Hãy quan sát kĩ chu kì."), encoding="utf-8")
        out = io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(io.StringIO()):
            self.assertEqual(video_ma.main([str(thu_muc)]), 0)
        self.assertEqual((thu_muc / "giong" / "canh-1.mp3").read_bytes(), before)


if __name__ == "__main__":
    unittest.main()
