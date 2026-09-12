# Video bài giảng có lời giảng (`v6.3.2-vi.4`): Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Thầy cô nhắn một câu là bài giảng trong PPT Master thành video MP4 có lời giảng tiếng Việt và phụ đề, chạy được cả khi máy không có PowerPoint.

**Architecture:** Lớp Việt chỉ điều phối và ghép. Phần sinh nội dung (slide, ghi chú, tiếng nói, PPTX) vẫn do upstream làm. `tools/vi/video.py` là cổng vào duy nhất: chọn cách dựng (PowerPoint hoặc FFmpeg + ảnh slide), gọi script upstream qua tiến trình con, in một dòng JSON ra stdout. Phần tính toán thuần (phụ đề, thời lượng, câu lệnh FFmpeg, chọn cách dựng) tách ra `tools/vi/video_parts/` để test được mà không cần FFmpeg, không cần mạng.

**Tech Stack:** Python ≥ 3.10 (chỉ thư viện chuẩn), FFmpeg/ffprobe, Playwright + Chromium (chụp ảnh slide), Windows PowerShell 5.1 (`pptmaster.ps1`), `unittest`, git, `gh`.

**Spec:** [2026-09-12-video-bai-giang-design.md](2026-09-12-video-bai-giang-design.md)

## Global Constraints

- Nhánh làm việc: `feat/vi-video`, đã tách từ `main` @ `615c87b2` (bản Việt `v6.3.2-vi.3`); spec đã commit ở `b160853a`.
- Không sửa bất kỳ file nào dưới `skills/`: `git diff --stat v6.3.2 HEAD -- skills/` phải rỗng; `attribution_guard.py` exit 0. Không sửa `CAI-DAT.bat`, `KIEM-TRA.bat`, `CAP-NHAT.bat`, `AGENTS.md`, `CLAUDE.md`.
- File được tạo/sửa: `tools/vi/video.py`, `tools/vi/video_parts/*`, `tools/vi/pptmaster.ps1`, `tools/vi/tests/test_video.py`, `tools/vi/tests/test_installer.py`, `tools/vi/tests/test_vi_layer.py`, `docs/vi/lam-video.md`, `docs/vi/tro-ly/video-bai-giang.md`, `docs/vi/tro-ly/quy-trinh-hoi.md`, `AGENTS.vi.md`, `docs/vi/bat-dau-nhanh.md`, `docs/vi/xu-ly-loi.md`, `docs/vi/phat-trien/2026-09-12-video-bai-giang-design.md`, `CHANGELOG-VI.md`, `README.md`.
- Python chỉ dùng thư viện chuẩn. Mọi lệnh ngoài (`ffmpeg`, `ffprobe`, script upstream) gọi qua `subprocess` với danh sách tham số, không dùng `shell=True`.
- `tools/vi/video.py` và `-Action tool`: stdout đúng một dòng JSON; mọi dòng tiến trình ra stderr. Mã thoát 0 khi `ready`, ngược lại 1.
- `tools/vi/pptmaster.ps1` giữ UTF-8 BOM.
- File Markdown viết hoàn toàn bằng tiếng Việt; chuỗi không phải tiếng Việt chỉ ở lệnh, đường dẫn, key JSON, tên giọng đọc và key canvas.
- Trong `docs/vi/tro-ly/video-bai-giang.md`: đúng 8 mục cấp 2 theo thứ tự chuẩn; 1–7 câu đánh số, mỗi câu có "Gợi ý:"; mục "Tạo nhanh" 2–3 mục đánh số; trong `## Khổ slide` chỉ đặt key canvas trong backtick; không có link Markdown.
- Test: `unittest` chuẩn trong `tools/vi/tests/`; không tải, không cài, không mở cửa sổ, không cần mạng. `$PY` = `C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe` (định nghĩa lại trong mỗi lần gọi PowerShell). Chạy: `& $PY -m unittest discover -s tools/vi/tests -v`.
- Commit theo Conventional Commits: tiêu đề, một dòng trống, rồi một dòng `Co-Authored-By: Claude … <noreply@anthropic.com>` theo hướng dẫn ghi công của phiên. Kiểm bằng `git log -1 --format=%B`.
- Không push trước Task 8; không force push; mọi lệnh `gh` đều có `--repo luonghaianh1208/PPTmaster`.

## File Structure

| File | Trách nhiệm |
|---|---|
| `tools/vi/video_parts/srt.py` | Đọc, dịch mốc thời gian, ghép và ghi phụ đề SRT |
| `tools/vi/video_parts/media.py` | Đo thời lượng bằng `ffprobe`, dựng file concat, dựng câu lệnh FFmpeg, thoát ký tự đường dẫn phụ đề |
| `tools/vi/video_parts/selection.py` | Chọn cách dựng và liệt kê các bước (dùng cho `--plan-only`) |
| `tools/vi/video.py` | Cổng vào: đọc trạng thái dự án, gọi upstream, ghép video, in JSON |
| `tools/vi/pptmaster.ps1` | Thêm `-Action tool -Name chromium` |
| `docs/vi/tro-ly/video-bai-giang.md` | Bộ câu hỏi cho loại việc thứ 6 |
| `docs/vi/lam-video.md` | Hướng dẫn cho thầy cô |
| `AGENTS.vi.md` | Mục 3 (câu lệnh) và mục 11 (quy trình làm video) |
| `tools/vi/tests/test_video.py` | Test phần tính toán thuần và `--plan-only` |

Thứ tự task: 1 phụ đề → 2 câu lệnh FFmpeg và đo thời lượng → 3 chọn cách dựng và cổng vào → 4 cài Chromium → 5 hướng dẫn cho AI → 6 tài liệu cho thầy cô → 7 chạy thật trên máy này → 8 nghiệm thu (điểm dừng) → 9 phát hành.

---

### Task 1: Phụ đề SRT

**Files:**
- Create: `tools/vi/video_parts/__init__.py`, `tools/vi/video_parts/srt.py`
- Create: `tools/vi/tests/test_video.py`

**Interfaces:**
- Consumes: không có (task đầu tiên).
- Produces:
  - `format_timestamp(seconds: float) -> str` → `"00:01:02,500"`.
  - `parse_timestamp(text: str) -> float`.
  - `Cue` = `dataclass(index: int, start: float, end: float, text: str)`.
  - `parse_srt(text: str) -> list[Cue]`.
  - `render_srt(cues: Sequence[Cue]) -> str` (kết thúc bằng một dòng trống, xuống dòng `\n`).
  - `merge_srt(sources: Sequence[tuple[str, float]]) -> str` — mỗi phần tử là (nội dung SRT của một slide, mốc bắt đầu của slide đó tính bằng giây); trả về nội dung SRT đã dịch mốc và đánh số lại từ 1.

- [ ] **Step 1: Viết test lỗi**

Tạo `tools/vi/tests/test_video.py`:

```python
"""Test cho lớp làm video của bản Việt (không chạy FFmpeg, không cần mạng)."""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT / "tools" / "vi"))

from video_parts import srt  # noqa: E402


SAMPLE_SRT = (
    "1\n"
    "00:00:00,000 --> 00:00:02,500\n"
    "Chào các em.\n"
    "\n"
    "2\n"
    "00:00:02,500 --> 00:00:05,000\n"
    "Hôm nay ta học phân số.\n"
    "\n"
)


class TimestampTest(unittest.TestCase):
    def test_format_timestamp_uses_srt_form(self):
        self.assertEqual(srt.format_timestamp(0), "00:00:00,000")
        self.assertEqual(srt.format_timestamp(62.5), "00:01:02,500")
        self.assertEqual(srt.format_timestamp(3723.004), "01:02:03,004")

    def test_parse_timestamp_round_trips(self):
        self.assertAlmostEqual(srt.parse_timestamp("01:02:03,004"), 3723.004, places=3)
        self.assertEqual(srt.format_timestamp(srt.parse_timestamp("00:01:02,500")), "00:01:02,500")


class ParseRenderTest(unittest.TestCase):
    def test_parse_srt_reads_cues(self):
        cues = srt.parse_srt(SAMPLE_SRT)
        self.assertEqual(len(cues), 2)
        self.assertAlmostEqual(cues[1].start, 2.5, places=3)
        self.assertAlmostEqual(cues[1].end, 5.0, places=3)
        self.assertEqual(cues[1].text, "Hôm nay ta học phân số.")

    def test_parse_srt_accepts_crlf_and_bom(self):
        cues = srt.parse_srt("\ufeff" + SAMPLE_SRT.replace("\n", "\r\n"))
        self.assertEqual(len(cues), 2)
        self.assertEqual(cues[0].text, "Chào các em.")

    def test_parse_srt_keeps_multiline_text(self):
        text = "1\n00:00:00,000 --> 00:00:01,000\ndòng một\ndòng hai\n\n"
        self.assertEqual(srt.parse_srt(text)[0].text, "dòng một\ndòng hai")

    def test_render_srt_round_trips(self):
        self.assertEqual(srt.render_srt(srt.parse_srt(SAMPLE_SRT)), SAMPLE_SRT)


class MergeTest(unittest.TestCase):
    def test_merge_shifts_and_renumbers(self):
        merged = srt.merge_srt([(SAMPLE_SRT, 0.0), (SAMPLE_SRT, 5.0)])
        cues = srt.parse_srt(merged)
        self.assertEqual([cue.index for cue in cues], [1, 2, 3, 4])
        self.assertAlmostEqual(cues[2].start, 5.0, places=3)
        self.assertAlmostEqual(cues[3].end, 10.0, places=3)

    def test_merge_skips_empty_sources(self):
        merged = srt.merge_srt([("", 0.0), (SAMPLE_SRT, 3.0)])
        cues = srt.parse_srt(merged)
        self.assertEqual(len(cues), 2)
        self.assertAlmostEqual(cues[0].start, 3.0, places=3)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Chạy test, xác nhận lỗi**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -p test_video.py -v
```
Expected: ERROR `ModuleNotFoundError: No module named 'video_parts'`.

- [ ] **Step 3: Viết code**

Tạo `tools/vi/video_parts/__init__.py` với đúng một dòng:

```python
"""Phần tính toán thuần của lớp làm video bản Việt."""
```

Tạo `tools/vi/video_parts/srt.py`:

```python
#!/usr/bin/env python3
"""Đọc, dịch mốc thời gian và ghép phụ đề SRT.

Chỉ dùng thư viện chuẩn Python.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Sequence

TIMESTAMP_RE = re.compile(r"(\d{2}):(\d{2}):(\d{2}),(\d{3})")
ARROW_RE = re.compile(r"(\d{2}:\d{2}:\d{2},\d{3})\s*-->\s*(\d{2}:\d{2}:\d{2},\d{3})")


@dataclass
class Cue:
    index: int
    start: float
    end: float
    text: str


def format_timestamp(seconds: float) -> str:
    if seconds < 0:
        seconds = 0.0
    total_ms = int(round(seconds * 1000))
    hours, rest = divmod(total_ms, 3_600_000)
    minutes, rest = divmod(rest, 60_000)
    secs, millis = divmod(rest, 1000)
    return f"{hours:02d}:{minutes:02d}:{secs:02d},{millis:03d}"


def parse_timestamp(text: str) -> float:
    match = TIMESTAMP_RE.fullmatch(text.strip())
    if match is None:
        raise ValueError(f"Mốc thời gian SRT không hợp lệ: {text!r}")
    hours, minutes, seconds, millis = (int(part) for part in match.groups())
    return hours * 3600 + minutes * 60 + seconds + millis / 1000


def parse_srt(text: str) -> list[Cue]:
    cleaned = text.lstrip("\ufeff").replace("\r\n", "\n").replace("\r", "\n")
    cues: list[Cue] = []
    for block in cleaned.split("\n\n"):
        lines = [line for line in block.split("\n") if line.strip()]
        if not lines:
            continue
        arrow_at = next((i for i, line in enumerate(lines) if ARROW_RE.search(line)), None)
        if arrow_at is None:
            continue
        match = ARROW_RE.search(lines[arrow_at])
        index = len(cues) + 1
        if arrow_at > 0:
            try:
                index = int(lines[arrow_at - 1].strip())
            except ValueError:
                index = len(cues) + 1
        cues.append(
            Cue(
                index=index,
                start=parse_timestamp(match.group(1)),
                end=parse_timestamp(match.group(2)),
                text="\n".join(lines[arrow_at + 1:]).strip(),
            )
        )
    return cues


def render_srt(cues: Sequence[Cue]) -> str:
    blocks = []
    for number, cue in enumerate(cues, start=1):
        blocks.append(
            f"{number}\n"
            f"{format_timestamp(cue.start)} --> {format_timestamp(cue.end)}\n"
            f"{cue.text}\n"
        )
    return "\n".join(blocks) + ("\n" if blocks else "")


def merge_srt(sources: Sequence[tuple[str, float]]) -> str:
    merged: list[Cue] = []
    for text, offset in sources:
        for cue in parse_srt(text):
            merged.append(Cue(index=0, start=cue.start + offset, end=cue.end + offset, text=cue.text))
    return render_srt(merged)
```

- [ ] **Step 4: Chạy test, xác nhận đạt**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -v
```
Expected: toàn bộ test OK (137 test cũ + 8 test mới).

- [ ] **Step 5: Commit**

```powershell
git add tools/vi/video_parts tools/vi/tests/test_video.py
git commit -m "feat(vi): merge per-slide subtitles into one SRT" -m "Co-Authored-By: Claude <noreply@anthropic.com>"
```
(Thay dòng `Co-Authored-By` bằng dòng ghi công đúng của phiên.)

---

### Task 2: Đo thời lượng và dựng câu lệnh FFmpeg

**Files:**
- Create: `tools/vi/video_parts/media.py`
- Modify: `tools/vi/tests/test_video.py` (thêm class mới trước `if __name__ == "__main__":`)

**Interfaces:**
- Consumes: không có.
- Produces:
  - `probe_duration(path: Path, run=subprocess.run) -> float` — gọi `ffprobe -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 <path>`; lỗi hoặc số không hợp lệ → `MediaError`.
  - `MediaError(Exception)` với thuộc tính `step: str`, `message: str`, `fix: str`.
  - `build_concat_text(entries: Sequence[tuple[Path, float]]) -> str` — file concat của FFmpeg cho ảnh; lặp lại file cuối theo đúng quy ước của FFmpeg.
  - `build_audio_concat_text(paths: Sequence[Path]) -> str`.
  - `escape_subtitles_filter(path: Path) -> str` — chuỗi cho `-vf subtitles=…` trên Windows.
  - `build_render_command(images_concat, audio_concat, out_path, fps, height, burn_srt=None) -> list[str]`.

- [ ] **Step 1: Viết test lỗi**

Thêm vào `tools/vi/tests/test_video.py`, ngay trước `if __name__ == "__main__":`:

```python
from video_parts import media  # noqa: E402


def fake_run(stdout="", returncode=0):
    def run(cmd, **kwargs):
        return subprocess.CompletedProcess(cmd, returncode, stdout, "")
    return run


class ProbeDurationTest(unittest.TestCase):
    def test_reads_duration_from_ffprobe(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "a.mp3"
            path.write_bytes(b"")
            self.assertAlmostEqual(media.probe_duration(path, run=fake_run("12.340000\n")), 12.34, places=3)

    def test_raises_media_error_when_ffprobe_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "a.mp3"
            path.write_bytes(b"")
            with self.assertRaises(media.MediaError) as ctx:
                media.probe_duration(path, run=fake_run("", returncode=1))
            self.assertEqual(ctx.exception.step, "ffmpeg")

    def test_raises_media_error_on_unparsable_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "a.mp3"
            path.write_bytes(b"")
            with self.assertRaises(media.MediaError):
                media.probe_duration(path, run=fake_run("N/A\n"))


class ConcatTest(unittest.TestCase):
    def test_image_concat_repeats_last_file(self):
        entries = [(Path("C:/p/01.png"), 2.0), (Path("C:/p/02.png"), 3.5)]
        lines = media.build_concat_text(entries).splitlines()
        self.assertEqual(lines[0], "file 'C:/p/01.png'")
        self.assertEqual(lines[1], "duration 2.000")
        self.assertEqual(lines[2], "file 'C:/p/02.png'")
        self.assertEqual(lines[3], "duration 3.500")
        self.assertEqual(lines[4], "file 'C:/p/02.png'")

    def test_audio_concat_lists_every_file(self):
        text = media.build_audio_concat_text([Path("C:/p/01.mp3"), Path("C:/p/02.mp3")])
        self.assertEqual(text.splitlines(), ["file 'C:/p/01.mp3'", "file 'C:/p/02.mp3'"])

    def test_paths_with_quote_are_escaped(self):
        text = media.build_audio_concat_text([Path("C:/p/it's.mp3")])
        self.assertIn("'\\''", text)


class RenderCommandTest(unittest.TestCase):
    def test_command_uses_concat_inputs_and_h264(self):
        cmd = media.build_render_command(
            Path("C:/p/images.txt"), Path("C:/p/audio.txt"), Path("C:/p/out.mp4"), fps=30, height=1080,
        )
        self.assertEqual(cmd[0], "ffmpeg")
        self.assertIn("-c:v", cmd)
        self.assertIn("libx264", cmd)
        self.assertIn("-c:a", cmd)
        self.assertIn("aac", cmd)
        self.assertIn("-shortest", cmd)
        self.assertEqual(cmd[-1], str(Path("C:/p/out.mp4")))
        self.assertIn("scale=-2:1080", " ".join(cmd))

    def test_burned_subtitles_add_filter(self):
        cmd = media.build_render_command(
            Path("C:/p/images.txt"), Path("C:/p/audio.txt"), Path("C:/p/out.mp4"),
            fps=30, height=720, burn_srt=Path("C:/p/phu de.srt"),
        )
        joined = " ".join(cmd)
        self.assertIn("subtitles=", joined)
        self.assertIn("scale=-2:720", joined)

    def test_escape_subtitles_filter_escapes_drive_and_backslash(self):
        escaped = media.escape_subtitles_filter(Path(r"C:\du an\phu de.srt"))
        self.assertNotIn("\\d", escaped.replace("\\\\", ""))
        self.assertIn(r"C\:", escaped)
```

- [ ] **Step 2: Chạy test, xác nhận lỗi**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -p test_video.py -v
```
Expected: ERROR `ImportError: cannot import name 'media'`.

- [ ] **Step 3: Viết code**

Tạo `tools/vi/video_parts/media.py`:

```python
#!/usr/bin/env python3
"""Đo thời lượng và dựng câu lệnh FFmpeg cho video bài giảng.

Chỉ dùng thư viện chuẩn Python. Không tự chạy FFmpeg ở mức module này.
"""

from __future__ import annotations

import subprocess
from pathlib import Path
from typing import Callable, Optional, Sequence

FIX_FFMPEG = 'Cài FFmpeg bằng: powershell -NoProfile -ExecutionPolicy Bypass -File tools\\vi\\pptmaster.ps1 -Action tool -Name ffmpeg'


class MediaError(Exception):
    def __init__(self, step: str, message: str, fix: str) -> None:
        super().__init__(message)
        self.step = step
        self.message = message
        self.fix = fix


def probe_duration(path: Path, run: Callable = subprocess.run) -> float:
    cmd = [
        "ffprobe", "-v", "error",
        "-show_entries", "format=duration",
        "-of", "default=noprint_wrappers=1:nokey=1",
        str(path),
    ]
    try:
        proc = run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=120)
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise MediaError("ffmpeg", f"Không chạy được ffprobe: {exc}", FIX_FFMPEG) from exc
    if proc.returncode != 0:
        raise MediaError("ffmpeg", f"ffprobe lỗi khi đọc {path.name} (mã {proc.returncode}).", FIX_FFMPEG)
    try:
        return float((proc.stdout or "").strip())
    except ValueError as exc:
        raise MediaError("ffmpeg", f"Không đọc được thời lượng của {path.name}.", FIX_FFMPEG) from exc


def _concat_path(path: Path) -> str:
    return str(path).replace("\\", "/").replace("'", "'\\''")


def build_concat_text(entries: Sequence[tuple[Path, float]]) -> str:
    lines: list[str] = []
    for path, duration in entries:
        lines.append(f"file '{_concat_path(path)}'")
        lines.append(f"duration {duration:.3f}")
    if entries:
        lines.append(f"file '{_concat_path(entries[-1][0])}'")
    return "\n".join(lines) + ("\n" if lines else "")


def build_audio_concat_text(paths: Sequence[Path]) -> str:
    lines = [f"file '{_concat_path(path)}'" for path in paths]
    return "\n".join(lines) + ("\n" if lines else "")


def escape_subtitles_filter(path: Path) -> str:
    text = str(path).replace("\\", "\\\\").replace(":", "\\:").replace("'", "\\'")
    return text


def build_render_command(
    images_concat: Path,
    audio_concat: Path,
    out_path: Path,
    fps: int,
    height: int,
    burn_srt: Optional[Path] = None,
) -> list[str]:
    video_filter = f"scale=-2:{height},format=yuv420p"
    if burn_srt is not None:
        video_filter = f"subtitles='{escape_subtitles_filter(burn_srt)}',{video_filter}"
    return [
        "ffmpeg", "-y", "-hide_banner", "-loglevel", "error",
        "-f", "concat", "-safe", "0", "-i", str(images_concat),
        "-f", "concat", "-safe", "0", "-i", str(audio_concat),
        "-vf", video_filter,
        "-r", str(fps),
        "-c:v", "libx264", "-preset", "medium", "-crf", "20",
        "-c:a", "aac", "-b:a", "160k",
        "-movflags", "+faststart",
        "-shortest",
        str(out_path),
    ]
```

- [ ] **Step 4: Chạy test, xác nhận đạt**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -v
```
Expected: toàn bộ test OK.

- [ ] **Step 5: Commit**

```powershell
git add tools/vi/video_parts/media.py tools/vi/tests/test_video.py
git commit -m "feat(vi): build ffmpeg render command for lecture video" -m "Co-Authored-By: Claude <noreply@anthropic.com>"
```

---

### Task 3: Chọn cách dựng và cổng vào `video.py`

**Files:**
- Create: `tools/vi/video_parts/selection.py`, `tools/vi/video.py`
- Modify: `tools/vi/tests/test_video.py`

**Interfaces:**
- Consumes: `srt.merge_srt`, `media.probe_duration`, `media.build_concat_text`, `media.build_audio_concat_text`, `media.build_render_command`, `media.MediaError`.
- Produces:
  - `ProjectState` = `dataclass(slides: list[str], audio: list[str], narrated_pptx: str | None, has_powerpoint: bool, has_chromium: bool, previews: list[str])`.
  - `select_backend(state: ProjectState, requested: str) -> tuple[str, list[str]]` — trả về (tên cách dựng, cảnh báo); không chọn được → `SelectionError(step, message, fix)`.
  - `plan_steps(state: ProjectState, backend: str, subtitle_mode: str) -> list[dict]` — mỗi bước là `{"step", "action", "method"}`.
  - CLI: `python tools/vi/video.py <đường_dẫn_dự_án> [--cach auto|powerpoint|ffmpeg] [--phu-de file|hinh|khong] [--do-phan-giai 1080|720] [--plan-only]`, in một dòng JSON như spec §6.1.

- [ ] **Step 1: Viết test lỗi**

Thêm vào `tools/vi/tests/test_video.py`, trước `if __name__ == "__main__":`:

```python
from video_parts import selection  # noqa: E402

VIDEO_CLI = REPO_ROOT / "tools" / "vi" / "video.py"


def state(**kwargs):
    base = dict(
        slides=["01_mo_dau", "02_noi_dung"],
        audio=["01_mo_dau", "02_noi_dung"],
        narrated_pptx="exports/bai_narrated.pptx",
        has_powerpoint=True,
        has_chromium=True,
        previews=[],
    )
    base.update(kwargs)
    return selection.ProjectState(**base)


class SelectBackendTest(unittest.TestCase):
    def test_auto_prefers_powerpoint_when_available(self):
        backend, warnings = selection.select_backend(state(), "auto")
        self.assertEqual(backend, "powerpoint")
        self.assertEqual(warnings, [])

    def test_auto_falls_back_to_ffmpeg_without_powerpoint(self):
        backend, warnings = selection.select_backend(state(has_powerpoint=False), "auto")
        self.assertEqual(backend, "ffmpeg")
        self.assertTrue(any("PowerPoint" in warning for warning in warnings))

    def test_auto_falls_back_to_ffmpeg_without_narrated_pptx(self):
        backend, _ = selection.select_backend(state(narrated_pptx=None), "auto")
        self.assertEqual(backend, "ffmpeg")

    def test_explicit_powerpoint_without_powerpoint_is_an_error(self):
        with self.assertRaises(selection.SelectionError) as ctx:
            selection.select_backend(state(has_powerpoint=False), "powerpoint")
        self.assertEqual(ctx.exception.step, "powerpoint")

    def test_missing_audio_is_an_error_for_every_backend(self):
        for requested in ("auto", "powerpoint", "ffmpeg"):
            with self.subTest(requested=requested):
                with self.assertRaises(selection.SelectionError) as ctx:
                    selection.select_backend(state(audio=[]), requested)
                self.assertEqual(ctx.exception.step, "audio")

    def test_audio_missing_for_one_slide_is_an_error(self):
        with self.assertRaises(selection.SelectionError) as ctx:
            selection.select_backend(state(audio=["01_mo_dau"]), "ffmpeg")
        self.assertEqual(ctx.exception.step, "audio")
        self.assertIn("02_noi_dung", ctx.exception.message)


class PlanStepsTest(unittest.TestCase):
    def test_ffmpeg_plan_lists_capture_and_render(self):
        steps = [(s["step"], s["action"], s["method"]) for s in selection.plan_steps(state(), "ffmpeg", "file")]
        self.assertEqual(steps, [
            ("preview", "capture", "visual_review.py"),
            ("subtitle", "merge", "srt"),
            ("render", "run", "ffmpeg"),
        ])

    def test_ffmpeg_plan_adds_chromium_install_when_missing(self):
        steps = selection.plan_steps(state(has_chromium=False), "ffmpeg", "file")
        self.assertEqual(steps[0]["step"], "chromium")
        self.assertEqual(steps[0]["method"], "pip+playwright")

    def test_ffmpeg_plan_skips_capture_when_previews_exist(self):
        steps = [s["step"] for s in selection.plan_steps(state(previews=["01_mo_dau", "02_noi_dung"]), "ffmpeg", "file")]
        self.assertNotIn("preview", steps)

    def test_powerpoint_plan_uses_upstream_exporter(self):
        steps = [(s["step"], s["method"]) for s in selection.plan_steps(state(), "powerpoint", "hinh")]
        self.assertIn(("render", "powerpoint_video.py"), steps)
        self.assertIn(("subtitle", "srt"), steps)
        self.assertIn(("burn", "ffmpeg"), steps)

    def test_no_subtitle_mode_drops_subtitle_steps(self):
        steps = [s["step"] for s in selection.plan_steps(state(), "ffmpeg", "khong")]
        self.assertNotIn("subtitle", steps)
        self.assertNotIn("burn", steps)


class VideoCliPlanTest(unittest.TestCase):
    def build_project(self, root, with_audio=True):
        (root / "svg_output").mkdir(parents=True)
        (root / "audio").mkdir(parents=True)
        (root / "exports").mkdir(parents=True)
        for stem in ("01_mo_dau", "02_noi_dung"):
            (root / "svg_output" / f"{stem}.svg").write_text("<svg/>", encoding="utf-8")
            if with_audio:
                (root / "audio" / f"{stem}.mp3").write_bytes(b"")
                (root / "audio" / f"{stem}.srt").write_text(SAMPLE_SRT, encoding="utf-8")
        return root

    def run_cli(self, *args):
        proc = subprocess.run(
            [sys.executable, str(VIDEO_CLI), *args],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=120,
            env={**os.environ, "PYTHONIOENCODING": "utf-8"},
        )
        stdout = proc.stdout.strip()
        self.assertEqual(len(stdout.splitlines()), 1, f"stdout phải là một dòng JSON:\n{proc.stdout}\n{proc.stderr}")
        return proc.returncode, json.loads(stdout)

    def test_plan_only_reports_backend_and_steps(self):
        with tempfile.TemporaryDirectory() as tmp:
            project = self.build_project(Path(tmp) / "du_an")
            code, data = self.run_cli(str(project), "--cach", "ffmpeg", "--plan-only")
            self.assertEqual(code, 0)
            self.assertEqual(data["backend"], "ffmpeg")
            self.assertIn("render", [step["step"] for step in data["steps"]])

    def test_missing_audio_reports_error_json(self):
        with tempfile.TemporaryDirectory() as tmp:
            project = self.build_project(Path(tmp) / "du_an", with_audio=False)
            code, data = self.run_cli(str(project), "--plan-only")
            self.assertEqual(code, 1)
            self.assertEqual(data["error"]["step"], "audio")
            self.assertIn("notes_to_audio.py", data["error"]["fix"])

    def test_unknown_project_reports_error_json(self):
        with tempfile.TemporaryDirectory() as tmp:
            code, data = self.run_cli(str(Path(tmp) / "khong_co"), "--plan-only")
            self.assertEqual(code, 1)
            self.assertEqual(data["error"]["step"], "project")
```

- [ ] **Step 2: Chạy test, xác nhận lỗi**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -p test_video.py -v
```
Expected: ERROR `ImportError: cannot import name 'selection'`.

- [ ] **Step 3: Viết `selection.py`**

```python
#!/usr/bin/env python3
"""Chọn cách dựng video và liệt kê các bước sẽ chạy."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

FIX_AUDIO = (
    "Chạy bước thuyết minh của dự án gốc trước: "
    "python skills/ppt-master/scripts/notes_to_audio.py <đường_dẫn_dự_án> --voice vi-VN-HoaiMyNeural"
)
FIX_POWERPOINT = "Máy không có PowerPoint. Chạy lại với --cach ffmpeg."
FIX_NARRATED = (
    "Chưa có bản PPTX đã gắn tiếng. Chạy lại bước xuất PPTX có thuyết minh của dự án gốc, "
    "hoặc chạy lại với --cach ffmpeg."
)


class SelectionError(Exception):
    def __init__(self, step: str, message: str, fix: str) -> None:
        super().__init__(message)
        self.step = step
        self.message = message
        self.fix = fix


@dataclass
class ProjectState:
    slides: list[str]
    audio: list[str]
    narrated_pptx: Optional[str] = None
    has_powerpoint: bool = False
    has_chromium: bool = False
    previews: list[str] = field(default_factory=list)


def _check_audio(state: ProjectState) -> None:
    if not state.slides:
        raise SelectionError("project", "Dự án chưa có slide nào trong svg_output/.", "Tạo slide trước khi làm video.")
    if not state.audio:
        raise SelectionError("audio", "Dự án chưa có file tiếng thuyết minh nào trong audio/.", FIX_AUDIO)
    missing = [stem for stem in state.slides if stem not in set(state.audio)]
    if missing:
        raise SelectionError("audio", "Thiếu tiếng thuyết minh cho: " + ", ".join(missing), FIX_AUDIO)


def select_backend(state: ProjectState, requested: str) -> tuple[str, list[str]]:
    _check_audio(state)
    warnings: list[str] = []
    if requested == "powerpoint":
        if not state.has_powerpoint:
            raise SelectionError("powerpoint", "Máy không có PowerPoint để xuất video.", FIX_POWERPOINT)
        if not state.narrated_pptx:
            raise SelectionError("narrated_pptx", "Chưa có bản PPTX đã gắn tiếng thuyết minh.", FIX_NARRATED)
        return "powerpoint", warnings
    if requested == "ffmpeg":
        return "ffmpeg", warnings
    if state.has_powerpoint and state.narrated_pptx:
        return "powerpoint", warnings
    if not state.has_powerpoint:
        warnings.append("Máy không có PowerPoint nên em dựng video bằng FFmpeg.")
    else:
        warnings.append("Chưa có bản PPTX gắn tiếng nên em dựng video bằng FFmpeg.")
    return "ffmpeg", warnings


def plan_steps(state: ProjectState, backend: str, subtitle_mode: str) -> list[dict]:
    steps: list[dict] = []
    if backend == "ffmpeg":
        if not state.has_chromium:
            steps.append({"step": "chromium", "action": "install", "method": "pip+playwright"})
        if sorted(state.previews) != sorted(state.slides):
            steps.append({"step": "preview", "action": "capture", "method": "visual_review.py"})
        if subtitle_mode != "khong":
            steps.append({"step": "subtitle", "action": "merge", "method": "srt"})
        steps.append({"step": "render", "action": "run", "method": "ffmpeg"})
        return steps
    if subtitle_mode != "khong":
        steps.append({"step": "subtitle", "action": "merge", "method": "srt"})
    steps.append({"step": "render", "action": "run", "method": "powerpoint_video.py"})
    if subtitle_mode == "hinh":
        steps.append({"step": "burn", "action": "run", "method": "ffmpeg"})
    return steps
```

- [ ] **Step 4: Viết `video.py`**

```python
#!/usr/bin/env python3
"""Làm video bài giảng có lời giảng từ một dự án PPT Master.

Cách dùng:
    python tools/vi/video.py <đường_dẫn_dự_án> [--cach auto|powerpoint|ffmpeg]
        [--phu-de file|hinh|khong] [--do-phan-giai 1080|720] [--plan-only]

stdout: đúng một dòng JSON. Tiến trình và log của lệnh ngoài đi ra stderr.
Mã thoát: 0 khi dựng xong, 1 khi lỗi.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from video_parts import media, selection, srt  # noqa: E402

REPO_ROOT = Path(__file__).resolve().parents[2]
SCRIPTS = REPO_ROOT / "skills" / "ppt-master" / "scripts"
FIX_CHROMIUM = (
    "Cài Chromium bằng: powershell -NoProfile -ExecutionPolicy Bypass -File tools\\vi\\pptmaster.ps1 "
    "-Action tool -Name chromium"
)


def log(text: str) -> None:
    print(text, file=sys.stderr, flush=True)


def emit(payload: dict) -> None:
    sys.stdout.write(json.dumps(payload, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def python_exe() -> str:
    venv = REPO_ROOT / "venv" / "Scripts" / "python.exe"
    return str(venv) if venv.is_file() else sys.executable


def has_powerpoint() -> bool:
    proc = subprocess.run(
        [python_exe(), str(SCRIPTS / "powerpoint_video.py"), "--check"],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180,
    )
    return proc.returncode == 0


def has_chromium() -> bool:
    root = os.environ.get("LOCALAPPDATA")
    if not root:
        return False
    browsers = Path(root) / "ms-playwright"
    return browsers.is_dir() and any(browsers.glob("chromium-*"))


def read_state(project: Path, check_powerpoint: bool) -> selection.ProjectState:
    slides = sorted(path.stem for path in (project / "svg_output").glob("*.svg"))
    audio = sorted(path.stem for path in (project / "audio").glob("*.mp3"))
    previews = sorted(path.stem for path in (project / ".preview").glob("*.png"))
    narrated = sorted((project / "exports").glob("*_narrated.pptx"), key=lambda p: p.stat().st_mtime)
    return selection.ProjectState(
        slides=slides,
        audio=audio,
        narrated_pptx=str(narrated[-1]) if narrated else None,
        has_powerpoint=has_powerpoint() if check_powerpoint else False,
        has_chromium=has_chromium(),
        previews=previews,
    )


def capture_previews(project: Path) -> None:
    server = str(SCRIPTS / "svg_editor" / "server.py")
    already_running = any((project / "live_preview").glob("*.lock"))
    if already_running:
        log("Máy chủ xem trước đang chạy sẵn, dùng lại.")
    else:
        log("Bật máy chủ xem trước để chụp ảnh slide...")
        started = subprocess.run(
            [python_exe(), server, str(project), "--daemon", "--live", "--no-browser"],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=300,
        )
        if started.returncode != 0:
            raise media.MediaError("render", "Không bật được máy chủ xem trước.", "Xem docs/vi/xu-ly-loi.md, mục Dựng video thất bại.")
    try:
        log("Chụp ảnh từng slide...")
        proc = subprocess.run(
            [python_exe(), str(SCRIPTS / "visual_review.py"), str(project)],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=1800,
        )
        for line in (proc.stderr or "").splitlines():
            log(line)
        if proc.returncode == 3:
            raise media.MediaError("chromium", "Chưa cài Chromium để chụp ảnh slide.", FIX_CHROMIUM)
        if proc.returncode != 0:
            raise media.MediaError("render", f"Chụp ảnh slide thất bại (mã {proc.returncode}).", "Xem docs/vi/xu-ly-loi.md, mục Dựng video thất bại.")
    finally:
        if not already_running:
            subprocess.run(
                [python_exe(), server, str(project), "--shutdown"],
                capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=120,
            )
            log("Đã tắt máy chủ xem trước.")


def build_subtitle(project: Path, stems: list[str], durations: list[float], out_path: Path) -> Path:
    sources = []
    offset = 0.0
    for stem, duration in zip(stems, durations):
        srt_path = project / "audio" / f"{stem}.srt"
        text = srt_path.read_text(encoding="utf-8") if srt_path.is_file() else ""
        sources.append((text, offset))
        offset += duration
    out_path.write_text(srt.merge_srt(sources), encoding="utf-8")
    return out_path


def render_ffmpeg(project: Path, stems: list[str], durations: list[float], out_path: Path,
                  height: int, burn_srt: Path | None) -> None:
    entries = [(project / ".preview" / f"{stem}.png", duration) for stem, duration in zip(stems, durations)]
    audio_files = [project / "audio" / f"{stem}.mp3" for stem in stems]
    with tempfile.TemporaryDirectory(prefix="pptmaster-vi-video-") as tmp:
        images_concat = Path(tmp) / "images.txt"
        audio_concat = Path(tmp) / "audio.txt"
        images_concat.write_text(media.build_concat_text(entries), encoding="utf-8")
        audio_concat.write_text(media.build_audio_concat_text(audio_files), encoding="utf-8")
        cmd = media.build_render_command(images_concat, audio_concat, out_path, fps=30, height=height, burn_srt=burn_srt)
        log("Dựng video bằng FFmpeg (có thể mất vài phút)...")
        proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=7200)
        for line in (proc.stderr or "").splitlines():
            log(line)
        if proc.returncode != 0:
            raise media.MediaError("render", f"FFmpeg lỗi (mã {proc.returncode}).", "Xem docs/vi/xu-ly-loi.md, mục Dựng video thất bại.")


def render_powerpoint(pptx: Path, out_path: Path, height: int) -> None:
    log("Xuất video bằng PowerPoint (cửa sổ PowerPoint sẽ hiện lên)...")
    cmd = [python_exe(), str(SCRIPTS / "powerpoint_video.py"), str(pptx), "-o", str(out_path), "--resolution", str(height)]
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=7200)
    for line in (proc.stderr or "").splitlines():
        log(line)
    if proc.returncode != 0:
        raise media.MediaError("powerpoint", f"PowerPoint không xuất được video (mã {proc.returncode}).", "Chạy lại với --cach ffmpeg.")


def burn_subtitles(video: Path, subtitle: Path, height: int) -> Path:
    burned = video.with_name(video.stem + "_phude.mp4")
    cmd = [
        "ffmpeg", "-y", "-hide_banner", "-loglevel", "error", "-i", str(video),
        "-vf", f"subtitles='{media.escape_subtitles_filter(subtitle)}',scale=-2:{height},format=yuv420p",
        "-c:v", "libx264", "-preset", "medium", "-crf", "20", "-c:a", "copy", str(burned),
    ]
    log("In phụ đề lên video...")
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=7200)
    for line in (proc.stderr or "").splitlines():
        log(line)
    if proc.returncode != 0:
        raise media.MediaError("render", f"Không in được phụ đề (mã {proc.returncode}).", "Dùng --phu-de file để lấy phụ đề rời.")
    return burned


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Làm video bài giảng có lời giảng")
    parser.add_argument("project_path", type=Path)
    parser.add_argument("--cach", choices=["auto", "powerpoint", "ffmpeg"], default="auto")
    parser.add_argument("--phu-de", dest="phu_de", choices=["file", "hinh", "khong"], default="file")
    parser.add_argument("--do-phan-giai", dest="height", type=int, choices=[720, 1080], default=1080)
    parser.add_argument("--plan-only", action="store_true")
    args = parser.parse_args(argv)

    project = args.project_path
    payload = {
        "ready": False, "backend": None, "video": None, "subtitle": None,
        "duration_seconds": None, "size_mb": None, "slides": 0,
        "installed": [], "warnings": [], "error": None,
    }
    if not (project / "svg_output").is_dir():
        payload["error"] = {
            "step": "project",
            "message": f"Không thấy thư mục dự án hợp lệ: {project}",
            "fix": "Kiểm tra lại đường dẫn dự án trong projects/.",
        }
        emit(payload)
        return 1

    try:
        state = read_state(project, check_powerpoint=args.cach in ("auto", "powerpoint"))
        backend, warnings = selection.select_backend(state, args.cach)
        payload["backend"] = backend
        payload["warnings"] = warnings
        payload["slides"] = len(state.slides)
        free_gb = shutil.disk_usage(project.resolve().anchor).free / (1024 ** 3)
        if free_gb < 2:
            payload["warnings"].append(f"Ổ đĩa chỉ còn {free_gb:.1f} GB; video có thể nặng 100-300 MB.")
        if len(str(project.resolve())) > 120:
            payload["warnings"].append("Đường dẫn dự án dài; nên chuyển bộ công cụ sang D:\\PPTmaster nếu dựng video lỗi.")
        steps = selection.plan_steps(state, backend, args.phu_de)
        if args.plan_only:
            payload["steps"] = steps
            payload["ready"] = True
            emit(payload)
            return 0
        if backend == "ffmpeg" and not state.has_chromium:
            raise media.MediaError("chromium", "Chưa cài Chromium để chụp ảnh slide.", FIX_CHROMIUM)
        if shutil.which("ffprobe") is None or shutil.which("ffmpeg") is None:
            raise media.MediaError("ffmpeg", "Máy chưa có FFmpeg.", media.FIX_FFMPEG)

        stems = state.slides
        durations = [media.probe_duration(project / "audio" / f"{stem}.mp3") for stem in stems]
        exports = project / "exports"
        exports.mkdir(parents=True, exist_ok=True)
        stamp = __import__("time").strftime("%Y%m%d_%H%M%S")
        video_path = exports / f"{project.name}_video_{stamp}.mp4"

        subtitle_path = None
        if args.phu_de != "khong":
            subtitle_path = build_subtitle(project, stems, durations, video_path.with_suffix(".srt"))

        if backend == "ffmpeg":
            if sorted(state.previews) != sorted(stems):
                capture_previews(project)
            burn = subtitle_path if args.phu_de == "hinh" else None
            render_ffmpeg(project, stems, durations, video_path, args.height, burn)
        else:
            render_powerpoint(Path(state.narrated_pptx), video_path, args.height)
            if args.phu_de == "hinh" and subtitle_path is not None:
                video_path = burn_subtitles(video_path, subtitle_path, args.height)

        total_audio = sum(durations)
        actual = media.probe_duration(video_path)
        if args.phu_de != "khong" and total_audio > 0 and abs(actual - total_audio) / total_audio > 0.02:
            payload["warnings"].append(
                "Thời lượng video lệch hơn 2% so với tổng thời lượng tiếng, phụ đề có thể lệch ở cuối bài."
            )
        payload["video"] = str(video_path)
        payload["subtitle"] = str(subtitle_path) if subtitle_path else None
        payload["duration_seconds"] = round(actual, 1)
        payload["size_mb"] = round(video_path.stat().st_size / (1024 * 1024), 1)
        payload["ready"] = True
        emit(payload)
        return 0
    except (selection.SelectionError, media.MediaError) as exc:
        payload["error"] = {"step": exc.step, "message": exc.message, "fix": exc.fix}
        emit(payload)
        return 1


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 5: Chạy test, xác nhận đạt**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -v
```
Expected: toàn bộ test OK. Test `--plan-only` chạy `video.py` bằng `sys.executable`, không gọi FFmpeg và không gọi PowerPoint (vì `--cach ffmpeg` hoặc lỗi xảy ra trước).

- [ ] **Step 6: Commit**

```powershell
git add tools/vi/video.py tools/vi/video_parts/selection.py tools/vi/tests/test_video.py
git commit -m "feat(vi): add lecture video entry point" -m "Co-Authored-By: Claude <noreply@anthropic.com>"
```

---

### Task 4: Cài Chromium theo nhu cầu

**Files:**
- Modify: `tools/vi/pptmaster.ps1` (`$OptionalTools` dòng 46-49, `Find-ToolDir` dòng 418, `Invoke-Tool` dòng 441)
- Modify: `tools/vi/tests/test_installer.py`

**Interfaces:**
- Consumes: `video.py` kiểm Chromium bằng `%LOCALAPPDATA%\ms-playwright\chromium-*`.
- Produces: `powershell -File tools\vi\pptmaster.ps1 -Action tool -Name chromium` → JSON `{tool, found, installed, dir, error}`; `-PlanOnly` → `{tool, found, dir, steps}` với bước `{"step": "chromium", "action": "install", "method": "pip+playwright"}`.

- [ ] **Step 1: Viết test lỗi**

Thêm vào `tools/vi/tests/test_installer.py`, trong `class InstallerPlanTest`, sau test `test_tool_found_in_user_install_folder`:

```python
    def test_tool_chromium_missing_plans_pip_playwright(self):
        plan = self.run_plan("-Action", "tool", "-Name", "chromium")
        self.assertEqual(plan["tool"], "chromium")
        self.assertFalse(plan["found"])
        self.assertEqual(self.steps(plan), [("chromium", "install", "pip+playwright")])

    def test_tool_chromium_found_when_browser_folder_exists(self):
        browser = self.localappdata / "ms-playwright" / "chromium-1234" / "chrome-win"
        browser.mkdir(parents=True)
        (browser / "headless_shell.exe").write_bytes(b"")
        plan = self.run_plan("-Action", "tool", "-Name", "chromium")
        self.assertTrue(plan["found"])
        self.assertIn("chromium-1234", plan["dir"])
        self.assertEqual(plan["steps"], [])
```

- [ ] **Step 2: Chạy test, xác nhận lỗi**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -p test_installer.py -v
```
Expected: FAIL cả hai test mới — `-Name chromium` bị coi là tên công cụ không hợp lệ nên mã thoát 1 và JSON có `error.step = "tool"`.

- [ ] **Step 3: Sửa `pptmaster.ps1`**

Thay khối `$OptionalTools` (dòng 46-49) bằng:

```powershell
$OptionalTools = @{
    ffmpeg = @{ Id = 'Gyan.FFmpeg'; Exe = 'ffmpeg.exe'; Manual = 'https://ffmpeg.org/download.html' }
    pandoc = @{ Id = 'JohnMacFarlane.Pandoc'; Exe = 'pandoc.exe'; Manual = 'https://pandoc.org/installing.html' }
    chromium = @{ Id = $null; Exe = $null; Manual = 'https://playwright.dev/python/docs/browsers' }
}
```

Trong `Find-ToolDir`, ngay sau dòng `function Find-ToolDir([string]$ToolName) {`, thêm nhánh riêng cho Chromium trước khi lấy `$exe`:

```powershell
    if ($ToolName -eq 'chromium') {
        if (-not $env:LOCALAPPDATA) { return $null }
        $browsers = Join-Path $env:LOCALAPPDATA 'ms-playwright'
        if (-not (Test-Path $browsers)) { return $null }
        $folder = @(Get-ChildItem -Path $browsers -Directory -Filter 'chromium-*' -ErrorAction SilentlyContinue | Sort-Object Name)
        if ($folder.Count -eq 0) { return $null }
        return $folder[-1].FullName
    }
```

Trong `Invoke-Tool`, sửa dòng `fix` của lỗi tên công cụ thành `'Chạy lại với -Name ffmpeg, -Name pandoc hoặc -Name chromium.'`, sửa nhánh `-PlanOnly` để chọn phương pháp theo công cụ:

```powershell
        if (-not $dir) {
            $method = if ($Name -eq 'chromium') { 'pip+playwright' } elseif (Test-Winget) { 'winget' } else { 'manual' }
            $steps += [pscustomobject]@{ step = $Name; action = 'install'; method = $method }
        }
```

và thêm nhánh cài thật cho Chromium ngay sau khối `if ($dir) { … return 0 }`:

```powershell
    if ($Name -eq 'chromium') {
        if (-not (Test-Path $VenvPython)) {
            Write-Json ([pscustomobject]@{ tool = $Name; found = $false; installed = $false; dir = $null; error = [pscustomobject]@{ step = 'tool'; message = 'Chưa có môi trường Python riêng (venv) để cài Chromium.'; fix = 'Chạy lệnh cài đặt trước: -Action setup -Auto' } })
            return 1
        }
        Write-Log 'Cài Playwright và tải Chromium (khoảng 150-300 MB, có thể mất vài phút)...'
        Invoke-Logged { & $VenvPython -m pip install playwright }
        Invoke-Logged { & $VenvPython -m playwright install chromium }
        $dir = Find-ToolDir $Name
        if (-not $dir) {
            Write-Json ([pscustomobject]@{ tool = $Name; found = $false; installed = $false; dir = $null; error = [pscustomobject]@{ step = 'tool'; message = 'Không tải được Chromium.'; fix = "Kiểm tra mạng rồi chạy lại; hướng dẫn thủ công: $($tool.Manual)" } })
            return 1
        }
        Write-Json ([pscustomobject]@{ tool = $Name; found = $true; installed = $true; dir = $dir; error = $null })
        return 0
    }
```

Sau khi sửa, kiểm BOM:

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -c "import pathlib; p = pathlib.Path('tools/vi/pptmaster.ps1'); b = p.read_bytes(); p.write_bytes(b if b.startswith(b'\xef\xbb\xbf') else b'\xef\xbb\xbf' + b); print(p.read_bytes()[:3])"
```
Expected: in `b'\xef\xbb\xbf'`.

- [ ] **Step 4: Chạy test, xác nhận đạt**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -v
```
Expected: toàn bộ test OK, gồm hai test Chromium mới và `test_powershell_scripts_have_utf8_bom`.

- [ ] **Step 5: Commit**

```powershell
git add tools/vi/pptmaster.ps1 tools/vi/tests/test_installer.py
git commit -m "feat(vi): install chromium on demand for slide capture" -m "Co-Authored-By: Claude <noreply@anthropic.com>"
```

---

### Task 5: Hướng dẫn cho AI (loại việc thứ 6)

**Files:**
- Create: `docs/vi/tro-ly/video-bai-giang.md`
- Modify: `docs/vi/tro-ly/quy-trinh-hoi.md` (bảng trong `## Khi nào áp dụng`), `AGENTS.vi.md` (mục 3 và mục 11 mới)
- Modify: `tools/vi/tests/test_vi_layer.py`

**Interfaces:**
- Consumes: CLI của `tools/vi/video.py` (Task 3) và `-Action tool -Name chromium` (Task 4).
- Produces: hằng `VIDEO_GUIDE = "video-bai-giang.md"` trong `GUIDE_FILES`; hằng `AGENTS_VI_VIDEO_HEADING = "## 11. Làm video bài giảng"`; class `VideoGuideTest`.

- [ ] **Step 1: Viết test lỗi**

Trong `tools/vi/tests/test_vi_layer.py`:
1. Thêm `"video-bai-giang.md",` vào cuối tuple `GUIDE_FILES` (dòng 176).
2. Sửa `test_agents_vi_has_assistant_section_last` thành:

```python
    def test_agents_vi_keeps_assistant_then_video_sections_last(self):
        headings = h2_headings(read("AGENTS.vi.md"))
        self.assertEqual(headings[-2], AGENTS_VI_ASSISTANT_HEADING)
        self.assertEqual(headings[-1], AGENTS_VI_VIDEO_HEADING)
```

3. Thêm trước `if __name__ == "__main__":`:

```python
AGENTS_VI_VIDEO_HEADING = "## 11. Làm video bài giảng"
VIDEO_COMMAND = r"venv\Scripts\python.exe tools\vi\video.py"


class VideoGuideTest(unittest.TestCase):
    def test_common_rules_table_lists_video_task(self):
        body = section(read("docs/vi/tro-ly/quy-trinh-hoi.md"), "## Khi nào áp dụng")
        self.assertIn("video-bai-giang.md", body)
        for keyword in ("làm video", "xuất video", "lồng tiếng"):
            self.assertIn(keyword, body)

    def test_agents_vi_video_section_explains_order_and_command(self):
        text = read("AGENTS.vi.md")
        self.assertIn(AGENTS_VI_VIDEO_HEADING, h2_headings(text))
        body = section(text, AGENTS_VI_VIDEO_HEADING)
        self.assertIn(VIDEO_COMMAND, body)
        for phrase in ("notes_to_audio.py", "chromium", "cửa sổ PowerPoint", "(docs/vi/tro-ly/video-bai-giang.md)"):
            self.assertIn(phrase, body)

    def test_agents_vi_triggers_include_video_phrases(self):
        body = section(read("AGENTS.vi.md"), "## 3. Câu lệnh tiếng Việt kích hoạt skill `ppt-master`")
        for phrase in ("làm video bài giảng", "xuất video"):
            self.assertIn(phrase, body)

    def test_video_guide_quick_section_covers_voice_and_subtitles(self):
        body = section(read("docs/vi/tro-ly/video-bai-giang.md"), "## Tạo nhanh")
        self.assertIn("giọng", body)
        self.assertIn("phụ đề", body)
```

- [ ] **Step 2: Chạy test, xác nhận lỗi**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -p test_vi_layer.py -v
```
Expected: FAIL ở các test khuôn hướng dẫn (thiếu file `video-bai-giang.md`), ở `VideoGuideTest`, và ở test thứ tự mục của `AGENTS.vi.md`.

- [ ] **Step 3: Viết `docs/vi/tro-ly/video-bai-giang.md`**

```markdown
# Loại việc: Video bài giảng

File dành cho AI. Luôn đọc `docs/vi/tro-ly/quy-trinh-hoi.md` trước file này.

## Khi nào dùng

Thầy cô muốn biến một bài giảng đã có trong bộ công cụ thành video có lời giảng, để đăng LMS, YouTube hoặc gửi học sinh tự học.

Ví dụ câu lệnh:
- "Làm video bài giảng này"
- "Lồng tiếng rồi xuất video bài Phân số"

## Câu hỏi bắt buộc

1. Dùng giọng đọc nào: giọng nữ hay giọng nam?
   Gợi ý: giọng nữ `vi-VN-HoaiMyNeural`.
2. Tốc độ đọc thế nào: chậm cho lớp nhỏ, vừa, hay nhanh?
   Gợi ý: vừa, giữ tốc độ mặc định.
3. Phụ đề để thành file riêng hay in thẳng lên hình?
   Gợi ý: file riêng, vì YouTube nhận file phụ đề và học sinh bật tắt được.
4. Độ phân giải video: 1080 cho máy chiếu và YouTube, hay 720 cho file nhẹ?
   Gợi ý: 1080.

## Câu hỏi tuỳ chọn

- Thầy cô muốn giữ hiệu ứng chuyển cảnh của slide không? Chỉ hỏi khi máy có PowerPoint và bài giảng có hiệu ứng.

## Tạo nhanh

1. Giọng đọc (câu hỏi bắt buộc 1).
2. Phụ đề (câu hỏi bắt buộc 3).

## Cấu trúc gợi ý

- Giữ nguyên thứ tự slide của bài giảng; mỗi slide một đoạn lời giảng trong `notes/`.
- Chưa có ghi chú thì viết ghi chú trước, mỗi slide 3–6 câu, đọc khoảng 30–60 giây.

## Phong cách gợi ý

- Lời giảng viết như nói: câu ngắn, xưng hô thống nhất, đọc số và công thức thành lời.
- Tên riêng tiếng Anh nên viết lại theo cách đọc để máy đọc đúng.

## Khổ slide

- Mặc định `ppt169`. Video xuất ra 1080 hoặc 720 dòng, 30 hình mỗi giây.

## Ghi vào brief

- Loại việc: Video bài giảng.
- Thầy cô yêu cầu: giọng đọc, tốc độ, phụ đề, độ phân giải, cách dựng nếu thầy cô có ý riêng.
- AI đề xuất (thầy cô đã đồng ý): các gợi ý thầy cô chấp nhận.
- Viết theo mẫu `docs/vi/tro-ly/mau-brief.md`.
```

- [ ] **Step 4: Sửa `quy-trinh-hoi.md` và `AGENTS.vi.md`**

Trong `docs/vi/tro-ly/quy-trinh-hoi.md`, thêm một dòng vào cuối bảng ở mục `## Khi nào áp dụng`:

```markdown
| Video bài giảng | "làm video", "xuất video", "lồng tiếng", "video bài giảng" | [video-bai-giang.md](video-bai-giang.md) |
```

Trong `AGENTS.vi.md` mục 3, thêm vào cuối danh sách câu lệnh: `"làm video bài giảng"`, `"lồng tiếng"`, `"xuất video"`.

Thêm mục 11 vào cuối `AGENTS.vi.md`:

```markdown
## 11. Làm video bài giảng

Khi người dùng yêu cầu làm video từ một bài giảng đã có, đọc [docs/vi/tro-ly/video-bai-giang.md](docs/vi/tro-ly/video-bai-giang.md), hỏi một lượt theo file đó, rồi làm đúng thứ tự sau:

1. Chưa có `notes/*.md`: viết ghi chú lời giảng cho từng slide theo quy trình của upstream.
2. Chưa có `audio/*.mp3`: chạy `skills/ppt-master/scripts/notes_to_audio.py <đường_dẫn_dự_án> --voice vi-VN-HoaiMyNeural` (hoặc `vi-VN-NamMinhNeural`).
3. Dựng video: `venv\Scripts\python.exe tools\vi\video.py <đường_dẫn_dự_án> --cach auto --phu-de file --do-phan-giai 1080`.
4. Đọc dòng JSON ở stdout. `ready` là `true` thì báo thầy cô đường dẫn video, thời lượng, dung lượng và nơi để phụ đề; `error` khác `null` thì làm theo `error.fix`, tối đa một lần, rồi báo thầy cô.

- `error.step` là `chromium`: hỏi thầy cô trước rồi chạy `powershell -NoProfile -ExecutionPolicy Bypass -File tools\vi\pptmaster.ps1 -Action tool -Name chromium`, vì bước này tải khoảng 150–300 MB.
- `error.step` là `audio`: quay lại bước 2.
- Cách dựng `powerpoint` sẽ mở **cửa sổ PowerPoint** và chiếm máy vài phút; báo trước cho thầy cô một dòng.
- Không tự cài phần mềm nào khác, không tự chạy FFmpeg theo cách riêng.
```

- [ ] **Step 5: Chạy test, xác nhận đạt**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -v
```
Expected: toàn bộ test OK, gồm các test khuôn hướng dẫn chạy trên 6 file loại việc.

- [ ] **Step 6: Commit**

```powershell
git add docs/vi/tro-ly/video-bai-giang.md docs/vi/tro-ly/quy-trinh-hoi.md AGENTS.vi.md tools/vi/tests/test_vi_layer.py
git commit -m "docs(vi): add lecture video task type for the AI" -m "Co-Authored-By: Claude <noreply@anthropic.com>"
```

---

### Task 6: Tài liệu cho thầy cô

**Files:**
- Create: `docs/vi/lam-video.md`
- Modify: `docs/vi/bat-dau-nhanh.md` (thêm mục `## Làm video bài giảng` ngay trước `## Lấy file kết quả`), `docs/vi/xu-ly-loi.md` (thêm mục `## Dựng video thất bại` ngay trước `## Đường dẫn quá dài`)
- Modify: `tools/vi/tests/test_vi_layer.py`

**Interfaces:**
- Consumes: hành vi của `video.py` (Task 3) và thông báo lỗi trong `media.py`.
- Produces: class `VideoUserDocsTest`; `REQUIRED_DOCS` thêm `"lam-video.md"`.

- [ ] **Step 1: Viết test lỗi**

Thêm `"lam-video.md",` vào `REQUIRED_DOCS` (dòng 135) và thêm class sau vào `tools/vi/tests/test_vi_layer.py`:

```python
class VideoUserDocsTest(unittest.TestCase):
    def test_video_doc_explains_time_size_and_subtitles(self):
        text = read("docs/vi/lam-video.md")
        for phrase in ("phụ đề", "YouTube", "PowerPoint", "FFmpeg", "MB"):
            self.assertIn(phrase, text)

    def test_quick_start_mentions_video(self):
        text = read("docs/vi/bat-dau-nhanh.md")
        headings = h2_headings(text)
        self.assertIn("## Làm video bài giảng", headings)
        self.assertLess(headings.index("## Làm video bài giảng"), headings.index("## Lấy file kết quả"))
        self.assertIn("(lam-video.md)", section(text, "## Làm video bài giảng"))

    def test_troubleshooting_has_video_section(self):
        text = read("docs/vi/xu-ly-loi.md")
        headings = h2_headings(text)
        self.assertIn("## Dựng video thất bại", headings)
        body = section(text, "## Dựng video thất bại")
        for phrase in ("Chromium", "FFmpeg", "PowerPoint", "venv\\Scripts\\python.exe tools\\vi\\video.py"):
            self.assertIn(phrase, body)
```

- [ ] **Step 2: Chạy test, xác nhận lỗi**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -p test_vi_layer.py -v
```
Expected: FAIL ở `test_required_vietnamese_docs_exist` và cả ba test mới.

- [ ] **Step 3: Viết tài liệu**

Tạo `docs/vi/lam-video.md`:

```markdown
# Làm video bài giảng

Bộ công cụ biến một bài giảng đã làm thành video có lời giảng tiếng Việt và phụ đề.

## Cần gì trước

- Một bài giảng đã tạo trong `projects\`.
- Lời giảng cho từng slide. Chưa có thì cứ nhắn AI viết giúp, rồi thầy cô sửa lại cho đúng ý.

## Cách làm

Nhắn cho AI: `Làm video bài giảng này`. AI hỏi bốn câu ngắn (giọng đọc, tốc độ, phụ đề, độ phân giải) rồi làm.

Máy có PowerPoint thì AI dùng PowerPoint để giữ hiệu ứng chuyển cảnh; lúc đó **cửa sổ PowerPoint sẽ hiện lên và chiếm máy vài phút**. Máy không có PowerPoint thì AI ghép bằng FFmpeg, lần đầu phải tải Chromium khoảng 150–300 MB để chụp ảnh từng slide.

## Thời gian và dung lượng

- Tạo tiếng đọc: khoảng một phút cho mỗi 10 slide.
- Dựng video: vài phút, tuỳ độ dài bài và tốc độ máy.
- Video 10 phút nặng khoảng 100–300 MB. Ổ đĩa nên còn trống ít nhất 2 GB.

## Phụ đề

- **File riêng** (nên dùng): được file `.srt` cạnh video. YouTube nhận file này, học sinh bật tắt được.
- **In lên hình**: chữ nằm sẵn trong video, gửi Zalo hay chiếu ngoại tuyến đều thấy, nhưng không tắt được.

## Trước khi giao cho học sinh

Nghe lại vài đoạn. Máy đọc tên riêng nước ngoài và công thức đôi khi chưa đúng; sửa lại lời giảng trong ghi chú rồi nhắn AI làm lại là được.

Gặp lỗi, xem [Xử lý lỗi](xu-ly-loi.md).
```

Trong `docs/vi/bat-dau-nhanh.md`, chèn ngay trước `## Lấy file kết quả`:

```markdown
## Làm video bài giảng

Có bài giảng rồi, nhắn `Làm video bài giảng này` là AI đọc lời giảng, ghép thành video kèm phụ đề. Chi tiết trong [Làm video bài giảng](lam-video.md).
```

Trong `docs/vi/xu-ly-loi.md`, chèn ngay trước `## Đường dẫn quá dài`:

```markdown
## Dựng video thất bại

Xem dòng kết quả AI đọc được, phần `error`:

- `chromium`: máy chưa có Chromium để chụp ảnh slide. Cho AI chạy `powershell -NoProfile -ExecutionPolicy Bypass -File tools\vi\pptmaster.ps1 -Action tool -Name chromium` (tải 150–300 MB).
- `ffmpeg`: máy chưa có FFmpeg. Cho AI chạy lệnh trên với `-Name ffmpeg`.
- `audio`: bài giảng chưa có tiếng đọc. Nhờ AI tạo lời giảng và tiếng đọc trước.
- `powerpoint`: PowerPoint không xuất được video. Thử lại bằng cách ghép ảnh: `venv\Scripts\python.exe tools\vi\video.py <đường_dẫn_dự_án> --cach ffmpeg`.
- `render`: thường do hết dung lượng ổ đĩa hoặc đường dẫn quá dài. Dọn ổ đĩa, hoặc chuyển bộ công cụ sang `D:\PPTmaster` rồi làm lại.
```

- [ ] **Step 4: Chạy test, xác nhận đạt**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -v
```
Expected: toàn bộ test OK, gồm `test_relative_markdown_links_resolve`.

- [ ] **Step 5: Commit**

```powershell
git add docs/vi/lam-video.md docs/vi/bat-dau-nhanh.md docs/vi/xu-ly-loi.md tools/vi/tests/test_vi_layer.py
git commit -m "docs(vi): explain lecture video to teachers" -m "Co-Authored-By: Claude <noreply@anthropic.com>"
```

---

### Task 7: Chạy thật trên máy này

**Files:** không sửa file trong repo, trừ khi Step 5 cần sửa lỗi.

**Interfaces:**
- Consumes: toàn bộ Task 1–6; dự án `projects/thpt_gioi_thieu_ppt169_20260911` (3 slide, đã có PPTX, chưa có ghi chú và tiếng).
- Produces: bằng chứng cho tiêu chí thành công #1, #2, #3 của spec.

**Cảnh báo trước khi chạy:** bước 4 gọi PowerPoint thật nên **màn hình sẽ hiện cửa sổ PowerPoint**; bước 2 tải khoảng 150–300 MB.

- [ ] **Step 1: Tạo ghi chú và tiếng thuyết minh**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
$env:PYTHONIOENCODING = 'utf-8'
$project = 'projects\thpt_gioi_thieu_ppt169_20260911'
New-Item -ItemType Directory -Force (Join-Path $project 'notes') | Out-Null
```

Viết ba file ghi chú bằng công cụ Write (không dùng `Set-Content` vì có chữ tiếng Việt): `notes\01_trang_bia.md`, `notes\02_gia_tri_cot_loi.md`, `notes\03_thanh_tuu_va_cam_ket.md`. Mỗi file 3–4 câu lời giảng tiếng Việt bám nội dung slide tương ứng, ví dụ cho `notes\01_trang_bia.md`:

```markdown
Chào quý thầy cô và các em. Hôm nay chúng ta cùng tìm hiểu về ngôi trường của chúng mình.

Phần trình bày gồm ba nội dung chính: giới thiệu chung, giá trị cốt lõi, và những thành tựu nổi bật.
```

Sau đó:

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
$env:PYTHONIOENCODING = 'utf-8'
& $PY skills/ppt-master/scripts/notes_to_audio.py projects\thpt_gioi_thieu_ppt169_20260911 --voice vi-VN-HoaiMyNeural
Get-ChildItem projects\thpt_gioi_thieu_ppt169_20260911\audio | Select-Object Name, Length
```
Expected: ba file `.mp3` và ba file `.srt`. Lỗi mạng thì ghi vào báo cáo và dừng task.

- [ ] **Step 2: Cài Chromium**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\vi\pptmaster.ps1 -Action tool -Name chromium
```
Expected: một dòng JSON có `"found": true`, `"installed": true`, `dir` trỏ tới `ms-playwright\chromium-…`. Chạy nền, không mở cửa sổ.

- [ ] **Step 3: Dựng video bằng FFmpeg**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
$env:PYTHONIOENCODING = 'utf-8'
& $PY -c "import json, subprocess, sys; p = subprocess.run([sys.argv[1], 'tools/vi/video.py', 'projects/thpt_gioi_thieu_ppt169_20260911', '--cach', 'ffmpeg', '--phu-de', 'file'], capture_output=True, timeout=3600); out = p.stdout.decode('utf-8'); print('exit', p.returncode, 'stdout_lines', len(out.strip().splitlines())); print(json.loads(out))" $PY
```
Expected: `exit 0`, đúng một dòng JSON, `ready` là `true`, có `video`, `subtitle`, `duration_seconds` xấp xỉ tổng thời lượng ba file tiếng, `size_mb` lớn hơn 0. Kiểm thêm: mở file `.srt` xem mốc thời gian tăng dần; kiểm `projects\…\live_preview` không còn tiến trình máy chủ (`Get-Process python` không có tiến trình lạ).

- [ ] **Step 4: Dựng video bằng PowerPoint (sẽ hiện cửa sổ PowerPoint)**

Chỉ chạy khi chủ repo đồng ý ngay lúc đó. Trước hết cần bản PPTX đã gắn tiếng:

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
$env:PYTHONIOENCODING = 'utf-8'
& $PY skills/ppt-master/scripts/svg_to_pptx.py projects\thpt_gioi_thieu_ppt169_20260911 --recorded-narration audio
Get-ChildItem projects\thpt_gioi_thieu_ppt169_20260911\exports -Filter *_narrated.pptx | Select-Object Name
& $PY tools\vi\video.py projects\thpt_gioi_thieu_ppt169_20260911 --cach powerpoint --phu-de file
```
Expected: JSON `ready` là `true`, `backend` là `powerpoint`. Nếu `svg_to_pptx.py` báo thiếu tham số, đọc `skills/ppt-master/scripts/docs/narration.md` mục "Narrated `svg_to_pptx.py` export" và dùng đúng bộ cờ ở đó; ghi lệnh đã dùng vào báo cáo.

- [ ] **Step 5: Nếu có lỗi**

Sửa tối thiểu trong `tools/vi/`, chạy lại toàn bộ test, commit `fix(vi): …`, rồi chạy lại bước hỏng. Tối đa 2 vòng. Không nới test.

- [ ] **Step 6: Dọn và ghi báo cáo**

Xoá video và ảnh thử trong `projects\…\exports` và `.preview` nếu chủ repo không cần giữ; giữ nguyên `notes` và `audio`. Ghi vào báo cáo: mọi lệnh, mọi dòng JSON, thời gian dựng, dung lượng file.

---

### Task 8: Nghiệm thu của chủ repo (điểm dừng)

**Files:** không có.

**Interfaces:**
- Consumes: nhánh `feat/vi-video` sau Task 7.
- Produces: xác nhận "đạt" cho tiêu chí #1, #2, #3 của spec.

- [ ] **Step 1: Kiểm tra toàn nhánh**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -v
& $PY skills/ppt-master/scripts/attribution_guard.py; "guard=$LASTEXITCODE"
git diff --stat v6.3.2 HEAD -- skills/; "(skills diff end)"
git status --porcelain; "(status end)"
```
Expected: test OK; `guard=0`; diff `skills/` rỗng; status rỗng.

- [ ] **Step 2: Gửi chủ repo danh sách nghiệm thu và chờ trả lời**

Nhờ chủ repo (sẽ mở Antigravity, PowerPoint và trình phát video — báo trước):
1. Mở một bài giảng thật đã có trong `projects\`, nhắn "Làm video bài giảng này".
2. Kiểm AI hỏi đúng bốn câu, rồi tự tạo lời giảng và tiếng đọc nếu chưa có.
3. Nhận video, mở bằng trình phát: tiếng khớp slide, chữ tiếng Việt trong video đúng dấu, phụ đề đúng chỗ.
4. Thử một lần với `--phu-de hinh` để xem chữ in lên hình có dễ đọc không.
5. Nếu máy có PowerPoint, kiểm cả hai cách dựng và cho biết cách nào thầy cô thích hơn.

Không sang Task 9 khi chưa có xác nhận "đạt".

---

### Task 9: Phát hành `v6.3.2-vi.4`

**Files:**
- Modify: `CHANGELOG-VI.md`, `README.md` (dòng `Phiên bản:`)

**Interfaces:**
- Consumes: xác nhận "đạt" (Task 8); `gh` đã đăng nhập `luonghaianh1208`.
- Produces: `origin/main` chứa tính năng, tag `v6.3.2-vi.4`, GitHub Release.

- [ ] **Step 1: Cập nhật CHANGELOG và dòng phiên bản**

Lấy ngày bằng `Get-Date -Format yyyy-MM-dd`, thêm vào `CHANGELOG-VI.md` ngay dưới dòng `# Nhật ký thay đổi — Bản Việt`:

```markdown
## 6.3.2-vi.4 — NGAY_PHAT_HANH

Làm video bài giảng: bài giảng đã có thành video MP4 có lời giảng tiếng Việt và phụ đề.

### Thêm
- `tools/vi/video.py`: một lệnh để dựng video; tự chọn PowerPoint (giữ hiệu ứng chuyển cảnh) hoặc FFmpeg (chạy được cả khi máy không có PowerPoint).
- Phụ đề `.srt` dựng từ phụ đề từng slide; chọn để file rời hoặc in lên hình.
- `-Action tool -Name chromium` cài Chromium khi cần chụp ảnh slide.
- Loại việc thứ 6 cho AI: `docs/vi/tro-ly/video-bai-giang.md`; `AGENTS.vi.md` mục 11.
- Tài liệu `docs/vi/lam-video.md` và mục "Dựng video thất bại" trong Xử lý lỗi.

### Không thay đổi
- Lõi PPT Master v6.3.2 của Hugo He giữ nguyên; phần thuyết minh vẫn dùng `edge-tts` của dự án gốc.

### Rủi ro
- Cách dựng bằng PowerPoint mở cửa sổ PowerPoint và chiếm máy vài phút.
- Chromium tải khoảng 150–300 MB ở lần dùng đầu trên máy không có PowerPoint.
- Video 10 phút nặng khoảng 100–300 MB; nên còn trống ít nhất 2 GB.
- Máy đọc có thể đọc sai tên riêng nước ngoài và công thức; nên nghe lại trước khi giao cho học sinh.
```

Sửa README: `Phiên bản: **6.3.2-vi.3**` thành `Phiên bản: **6.3.2-vi.4**`.

- [ ] **Step 2: Chạy test và commit**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
& $PY -m unittest discover -s tools/vi/tests -v
git add CHANGELOG-VI.md README.md
git commit -m "docs(vi): release 6.3.2-vi.4" -m "Co-Authored-By: Claude <noreply@anthropic.com>"
```

- [ ] **Step 3: Đưa vào `main`, gắn tag, push**

```powershell
git switch main
git pull --ff-only origin main
git merge --ff-only feat/vi-video
git tag -a v6.3.2-vi.4 -m "PPT Master ban Viet 6.3.2-vi.4"
git push origin main
git push origin v6.3.2-vi.4
git ls-remote origin refs/heads/main refs/tags/v6.3.2-vi.4
```
Expected: fast-forward thành công; không dùng `--force`; `ls-remote` in đủ hai ref.

- [ ] **Step 4: Tạo GitHub Release**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
$notes = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\release-notes-vi4.md'
& $PY -c "import re, pathlib, sys; t = pathlib.Path('CHANGELOG-VI.md').read_text(encoding='utf-8'); m = re.search(r'^## 6\.3\.2-vi\.4.*?(?=^## |\Z)', t, re.S | re.M); pathlib.Path(sys.argv[1]).write_text(m.group(0), encoding='utf-8')" $notes
gh release create v6.3.2-vi.4 --repo luonghaianh1208/PPTmaster --title "PPT Master bản Việt 6.3.2-vi.4" --notes-file $notes
gh release view v6.3.2-vi.4 --repo luonghaianh1208/PPTmaster --json url,tagName,isDraft
gh release view v6.3.2-vi.4 --repo hugohe3/ppt-master --json url
```
Expected: in URL release, `isDraft` là `false`; lệnh cuối báo `release not found`.

- [ ] **Step 5: Xác minh sau phát hành**

```powershell
$PY = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\venv312\Scripts\python.exe'
$env:PYTHONIOENCODING = 'utf-8'
$fresh = 'C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\freshclone-vi4'
git clone -q --depth 1 https://github.com/luonghaianh1208/PPTmaster.git $fresh
Push-Location $fresh
& $PY skills/ppt-master/scripts/attribution_guard.py; "guard=$LASTEXITCODE"
& $PY -m unittest discover -s tools/vi/tests
Test-Path docs/vi/lam-video.md
Pop-Location
```
Expected: `guard=0`; test OK; `True`. Sau đó xoá `$fresh` và `release-notes-vi4.md`. Xoá nhánh phải hỏi chủ repo.
