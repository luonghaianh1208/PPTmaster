# Video giải thích dựng bằng mã (kiểu viết tay) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Từ một file `video.md` (kịch bản chữ), dựng `video.mp4` 1280×720 kiểu viết tay có giọng đọc tiếng Việt và phụ đề in lên hình, gồm 8 loại cảnh và cảnh quay thí nghiệm ảo chạy đúng 8 mô hình đã kiểm; nối vào AI thành loại việc thứ 10 của lớp Việt (bản v6.3.2-vi.9).

**Architecture:** `tools/vi/video_ma.py` đọc → kiểm → tạo giọng → lập lịch → dựng trang HTML mỗi cảnh (mỗi cảnh là danh sách "mục" có mốc thời gian cố định, hàm `datThoiDiem(t)` xác định) → Chromium chụp khung 15 khung/giây → FFmpeg ghép hình + tiếng + phụ đề. Phần vẽ là JavaScript thuần đặt ở `video_ma_parts/runtime/`, logic thuần (mốc thời gian, cắt chữ, nội suy tham số) kiểm bằng Node; phần chụp kiểm bằng Chromium thật. Không sửa `video.py`, `video_parts/`, `thi_nghiem_parts/`, `skills/`.

**Tech Stack:** Python 3.10+ chỉ thư viện chuẩn + `edge-tts` (đã có) + `playwright` (cài theo yêu cầu, import lười); JavaScript thuần (Chromium, Node cho test); FFmpeg/ffprobe; test bằng `unittest`.

**Spec:** `docs/vi/phat-trien/2026-09-25-video-ma-viet-tay-design.md` (Task 1 sửa ba chỗ của spec theo các quyết định dưới đây; đọc cả hai).

## Global Constraints

- Chỉ chạy trên Windows; font `Segoe Print`, dự phòng `Ink Free`, `Comic Sans MS`; không nhúng font.
- Python chỉ dùng thư viện chuẩn, `edge-tts` và `playwright` (import lười, trong hàm); phần vẽ là JavaScript thuần, không thư viện ngoài, không địa chỉ web trong trang dựng.
- Không sửa `tools/vi/video.py`, `tools/vi/video_parts/`, `tools/vi/thi_nghiem_parts/`, `skills/`, `LICENSE`, `SPONSORS*.md`, `requirements.txt` gốc, `attribution_guard.py`, frontmatter của `SKILL.md`. Chỉ đọc file mô hình và `khung.js` từ `thi_nghiem_parts`.
- Không chạy `project_manager.py init`, không tạo SVG của upstream; đầu ra ở `projects/_video/<tên>/`; không commit gì trong `projects/`; không tạo hay commit `.env`.
- Giọng máy edge-tts **không bao giờ** được gọi trong test; test dùng hàm giả hoặc tiếng giả tạo bằng FFmpeg.
- Chuỗi hiển thị và thông báo bằng tiếng Việt; tên file, khoá và mã bằng tiếng Anh/không dấu như spec; mỗi file Markdown một ngôn ngữ (tiếng Việt).
- Mọi câu lệnh in đúng **một dòng JSON** ra stdout, mã thoát 0 hoặc 1; log đi ra stderr.
- Toàn bộ test hiện có của lớp Việt vẫn xanh (`venv\Scripts\python.exe -m unittest discover -s tools/vi/tests -t .`, chưa kể test mới có Chromium).
- Trailer commit: `Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>`. Không push, không phát hành khi chưa hỏi chủ repo.
- Lệnh Python: dùng `venv\Scripts\python.exe` ở gốc repo. Test có Chromium cần Python có `playwright`; trên máy chủ repo dùng `C:/Users/ADMIN/vmt/v/Scripts/python.exe` (đường dẫn ngắn, đã cài `playwright`; cài trong đường dẫn dài sẽ lỗi long-path).

## Quyết định của kế hoạch (bổ sung spec, Task 1 ghi vào spec)

- **Dẫn đầu 0,7 giây:** mỗi cảnh có 0,7 s trước khi tiếng bắt đầu để tiêu đề kịp viết. Thời lượng cảnh = 0,7 + giọng + 0,6, tối thiểu 2,5 s, rồi làm tròn lên bội của 1/15 s (nên tối thiểu thực tế là 2,533 s). Mốc câu và mốc hiện ý tính theo thời gian cảnh (đã cộng 0,7). Tiếng được đệm 700 ms im lặng ở đầu.
- **Sổ giọng:** `giong/canh-N.json` (băm của lời + giọng + tốc độ, và mốc câu) thay cho `canh-N.txt` của spec. Có `canh-N.mp3` mà **không có** `canh-N.json` là file thầy cô đặt sẵn, không bao giờ bị ghi đè.
- **`fps_chụp` = 15, định dạng PNG** (đo ở Task 1: PNG 45,3 ms/khung, JPEG 50,9 ms/khung; video 5 phút ở 15 khung/giây mất khoảng 3,4 phút).
- Cảnh thí nghiệm: tối đa 3 mã tham số khác nhau trong `tham-so:` và tối đa 3 đại lượng trong `do:` (mặc định hai đại lượng đầu của mô hình); `giây` của `tham-so:` tính từ đầu cảnh.
- Giới hạn chữ bổ sung: tiêu đề cảnh ≤ 90 ký tự; `trai`, `phai` ≤ 24; `truc-ngang`, `truc-doc` ≤ 40.
- Kiểm chữ tràn khung dùng `scrollHeight/scrollWidth` của ô chữ ở trạng thái cuối cảnh.

## Review Focus

Năm điều spec ngụ ý mà không có test nào tự nhiên gọi tới; mỗi điều được ghim bằng test trong task sở hữu mã:

1. **Lời một câu nhưng cảnh 4–6 ý, hoặc lời không có dấu chấm cuối** → các ý chia đều trên thời lượng giọng, không dồn một lúc (Task 3, `test_loi_mot_cau_nhieu_y`).
2. **Chữ thầy cô có `<`, `&`, `"`, `</script>`** (ví dụ "a < b", "P&V", "H₂SO₄ </script>") → hiện nguyên chữ, trang không vỡ, không chạy mã (Task 5 `test_canh.js` kiểm `catDanhDau`; Task 6 `test_trang_khong_chua_ma_chay_duoc`).
3. **File giọng có sẵn rỗng hoặc hỏng, hoặc tổng hợp giọng dở dang** → lỗi `giong` nêu đúng file; không để lại mp3 dở dang bị lần sau coi là "thầy cô đặt sẵn" (Task 4).
4. **Thư mục dự án có dấu cách, dấu tiếng Việt, dấu nháy đơn** (`Bài 5 – Sulfur dioxide (thử)`) → ghép hình, tiếng và phụ đề chạy được (Task 7 danh sách nối tiếng; Task 9 dựng thật trong thư mục như vậy).
5. **Chạy lại sau khi sửa lời một cảnh, hoặc sau lỗi giữa chừng** → chỉ cảnh đổi bị tạo giọng lại, file thầy cô đặt sẵn giữ nguyên, thư mục khung tạm không để khung cũ lọt vào video mới (Task 4 và Task 8 `test_thu_muc_khung_tam_duoc_don`).

---

## Cấu trúc file

```
tools/vi/video_ma.py                        CLI (Task 8)
tools/vi/video_ma_parts/__init__.py         (Task 2)
tools/vi/video_ma_parts/parse.py            (Task 2)  đọc video.md, lỗi nêu đúng dòng
tools/vi/video_ma_parts/kiem.py             (Task 2)  giới hạn chữ, mô hình, tham số
tools/vi/video_ma_parts/lich.py             (Task 3)  câu, mốc, thời lượng, dữ liệu cảnh
tools/vi/video_ma_parts/giong.py            (Task 4)  edge-tts / mp3 có sẵn
tools/vi/video_ma_parts/runtime/khung-video.js   (Task 5)
tools/vi/video_ma_parts/runtime/viet-tay.css     (Task 5)
tools/vi/video_ma_parts/runtime/canh/<loại>.js   (Task 5) 8 file
tools/vi/video_ma_parts/trang.py            (Task 6)
tools/vi/video_ma_parts/chup.py             (Task 6)
tools/vi/video_ma_parts/ghep.py             (Task 7)
tools/vi/fixtures/video-mau/video.md        (Task 2)  kịch bản mẫu 8 cảnh (con lắc đơn)
tools/vi/tests/test_video_ma_parse.py       (Task 2)
tools/vi/tests/test_video_ma_lich.py        (Task 3)
tools/vi/tests/test_video_ma_giong.py       (Task 4)
tools/vi/tests/js/test_canh.js              (Task 5)
tools/vi/tests/test_video_ma_canh.py        (Task 5)  chạy test_canh.js
tools/vi/tests/test_video_ma_chup.py        (Task 6)  cần Chromium
tools/vi/tests/test_video_ma_ghep.py        (Task 7)
tools/vi/tests/test_video_ma_cong_cu.py     (Task 8)
tools/vi/tests/test_video_ma_tich_hop.py    (Task 9)  cần Chromium + FFmpeg
docs/vi/phat-trien/2026-09-25-video-ma-kiem-thu.md   (Task 1, hoàn thiện Task 11)
docs/vi/tro-ly/video-giai-thich.md, canh-video.md; docs/vi/video-giai-thich.md  (Task 10)
```

Sửa ở Task 10: `AGENTS.vi.md`, `.agents/rules/ppt-master-vi.md`, `docs/vi/tro-ly/quy-trinh-hoi.md`, `docs/vi/tro-ly/mau-brief.md`, `docs/vi/xu-ly-loi.md`, `docs/vi/cau-lenh-mau.md`, `docs/vi/bat-dau-nhanh.md`, `README.md`, `CHANGELOG-VI.md`, `tools/vi/tests/test_vi_layer.py`.

---

### Task 1: Biên bản đo và chỉnh spec

**Files:**
- Create: `docs/vi/phat-trien/2026-09-25-video-ma-kiem-thu.md`
- Modify: `docs/vi/phat-trien/2026-09-25-video-ma-viet-tay-design.md` (mục 8, mục 9, mục 6/7)

**Interfaces:**
- Produces: hằng số dùng ở các task sau: `FPS = 15`, ảnh PNG, `DAN_DAU = 0.7`, `DUOI = 0.6`, `TOI_THIEU = 2.5`.

Hai phép đo của spec mục 12 đã làm xong trước khi viết kế hoạch (máy chủ repo, Windows 11, Chromium headless shell, viewport 1280×720):

1. Font: `Segoe Print`, `Ink Free` và `Comic Sans MS` đều có trong hệ thống (`document.fonts.check` đúng) và đều hiện đủ dấu chồng ("Nhờ ướt nhẫm quyết định Đường thẳng khớp", "ẫ ỡ ự ệ ầ ẩ ỗ ử", "ƯỚC LƯỢNG – ĐỒ THỊ – PHƯƠNG TRÌNH"). Segoe Print đọc rõ hơn; Ink Free mảnh, dùng dự phòng.
2. Tốc độ: 100 khung liên tiếp có nét SVG đang vẽ và chữ đang viết: PNG 45,3 ms/khung, JPEG 50,9 ms/khung. Video 5 phút: 12 khung/giây → 2,7 phút; 15 → 3,4 phút; 30 → 6,8 phút. Không cần chụp song song.

- [ ] **Step 1: Tạo biên bản kiểm thử**

Tạo `docs/vi/phat-trien/2026-09-25-video-ma-kiem-thu.md`:

```markdown
# Biên bản kiểm thử: video giải thích dựng bằng mã (v6.3.2-vi.9)

Ngày 2026-09-25. Máy chủ repo: Windows 11, Chromium headless shell của Playwright, FFmpeg bản Gyan.

## 1. Phép đo trước khi làm (spec mục 12)

### 1.1 Font
Segoe Print, Ink Free và Comic Sans MS đều hiện đủ dấu chồng tiếng Việt. Chọn Segoe Print làm font chính, Ink Free và Comic Sans MS dự phòng.

### 1.2 Tốc độ chụp khung
| Định dạng | ms/khung | 5 phút ở 12 khung/giây | ở 15 | ở 30 |
|---|---|---|---|---|
| PNG | 45,3 | 2,7 phút | 3,4 phút | 6,8 phút |
| JPEG chất lượng 92 | 50,9 | 3,1 phút | 3,8 phút | 7,6 phút |

Quyết định: chụp PNG, 15 khung/giây, ghép ra 30 khung/giây (nhân khung). Ngưỡng 20 phút của spec không bị chạm, không cần chụp song song.

## 2. Kết quả kiểm thử

(Điền ở Task 11.)
```

- [ ] **Step 2: Sửa spec cho khớp các quyết định**

Trong `docs/vi/phat-trien/2026-09-25-video-ma-viet-tay-design.md`:

Mục 8, dòng 1: đổi `công cụ ghi mã băm của lời trong `giong/canh-N.txt`` thành `công cụ ghi sổ `giong/canh-N.json` gồm mã băm của lời, giọng, tốc độ và mốc câu; có `canh-N.mp3` mà không có `canh-N.json` là file thầy cô đặt sẵn, không bao giờ bị ghi đè`.

Mục 8, dòng 2: thay bằng `Thời lượng cảnh = 0,7 giây dẫn đầu (tiêu đề viết trước khi tiếng bắt đầu) + thời lượng file giọng (`probe_duration`) + 0,6 giây, tối thiểu 2,5 giây, làm tròn lên bội của 1/15 giây. Tiếng được đệm 0,7 giây im lặng ở đầu; mốc câu và mốc hiện ý tính theo thời gian cảnh.`

Mục 9, dòng cuối: thay `fps_chụp chốt sau khi đo (mục 12); khung viết tay mặc định 12–15.` bằng `fps_chụp = 15, ảnh PNG (đã đo ở biên bản kiểm thử: 45,3 ms/khung).`

Mục 7: thêm dòng cuối `Tối đa 3 mã tham số khác nhau trong `tham-so:` và tối đa 3 đại lượng trong `do:` (mặc định hai đại lượng đầu của mô hình); số giây của `tham-so:` tính từ đầu cảnh.`

Mục 6, sau bảng: thêm `Giới hạn bổ sung: tiêu đề cảnh ≤ 90 ký tự, `trai` và `phai` ≤ 24, `truc-ngang` và `truc-doc` ≤ 40.`

- [ ] **Step 3: Commit**

```bash
git add docs/vi/phat-trien/2026-09-25-video-ma-kiem-thu.md docs/vi/phat-trien/2026-09-25-video-ma-viet-tay-design.md
git commit -m "docs(vi): record font and capture-speed measurements for code-built videos"
```

---

### Task 2: Bộ đọc `video.md` và bộ kiểm giới hạn

**Files:**
- Create: `tools/vi/video_ma_parts/__init__.py` (rỗng), `tools/vi/video_ma_parts/parse.py`, `tools/vi/video_ma_parts/kiem.py`
- Create: `tools/vi/fixtures/video-mau/video.md`
- Test: `tools/vi/tests/test_video_ma_parse.py`

**Interfaces:**
- Produces (`parse.py`): `ParseError(line_no: int, message: str)` (str là `"Dòng N: ..."`); `Scene(so, dong, loai, loi, truong: dict[str, list[str]], dong_truong: dict[str, list[int]])`; `Video(meta: dict[str, str], canh: list[Scene])` (meta đã điền mặc định `phong-cach=viet-tay`, `giong=nu`, `toc-do=vua`, `phu-de=hinh`); `parse(text: str) -> Video`; hằng `SCENE_SPEC`, `SCENE_TYPES`.
- Produces (`kiem.py`): `CanhError(so: int, message: str)` (str là `"Cảnh N: ..."`); `kiem(video, thu_muc: Path) -> list[str]` (warnings); `tham_so_theo_thoi_gian(scene) -> dict[str, list[tuple[float, float]]]` (mỗi mã: danh sách `(giây, giá trị)` sắp theo giây, ổn định); `ma_do(scene, model) -> list[str]`; `hien_thi(chu) -> int` (đếm ký tự sau khi bỏ `**`, `~`, `^`); `LIMITS`.

- [ ] **Step 1: Viết kịch bản mẫu**

Tạo `tools/vi/fixtures/video-mau/video.md` (một video Vật lí, đủ 8 loại cảnh, dùng cho test bộ đọc, xem trước và kiểm bằng mắt):

```markdown
---
tieu-de: Con lắc đơn
mon: Vật lí
lop: 11
---

## Cảnh 1
loai: tieu-de
chu: Con lắc đơn và chu kì dao động
phu: Vật lí 11
loi: Hôm nay chúng ta tìm hiểu chu kì của con lắc đơn phụ thuộc vào yếu tố nào.

## Cảnh 2
loai: khai-niem
thuat-ngu: Chu kì T
dinh-nghia: Khoảng thời gian ngắn nhất để con lắc thực hiện một dao động toàn phần, đo bằng giây.
loi: Chu kì là khoảng thời gian ngắn nhất để con lắc thực hiện một dao động toàn phần. Đơn vị của chu kì là giây.

## Cảnh 3
loai: cong-thuc
bieu-thuc: T = 2π√(l/g)
giai-thich: T là chu kì, đơn vị giây
giai-thich: l là chiều dài dây, đơn vị mét
giai-thich: g là gia tốc trọng trường, khoảng 9,8 m/s^2^
loi: Chu kì bằng hai pi nhân căn của l chia g. Trong đó T là chu kì. Chữ l là chiều dài dây. Chữ g là gia tốc trọng trường.

## Cảnh 4
loai: y-tung-y
tieu-de: Chu kì phụ thuộc vào gì
y: Chiều dài dây l: dây dài hơn thì chu kì lớn hơn
y: Gia tốc trọng trường g: g lớn thì chu kì nhỏ
y: Không phụ thuộc khối lượng quả nặng
loi: Thứ nhất, chu kì phụ thuộc chiều dài dây. Thứ hai, chu kì phụ thuộc gia tốc trọng trường. Thứ ba, chu kì không phụ thuộc khối lượng quả nặng.

## Cảnh 5
loai: quy-trinh
tieu-de: Cách đo chu kì
buoc: Treo quả nặng vào dây
buoc: Kéo lệch một góc nhỏ rồi thả
buoc: Đo thời gian 10 dao động
loi: Bước một, treo quả nặng vào dây. Bước hai, kéo lệch một góc nhỏ rồi thả. Bước ba, đo thời gian mười dao động rồi chia cho mười.

## Cảnh 6
loai: so-sanh
tieu-de: Yếu tố nào ảnh hưởng
trai: Có ảnh hưởng
phai: Không ảnh hưởng
y-trai: Chiều dài dây
y-trai: Gia tốc trọng trường
y-phai: Khối lượng quả nặng
loi: Chiều dài dây và gia tốc trọng trường có ảnh hưởng đến chu kì. Khối lượng quả nặng thì không.

## Cảnh 7
loai: do-thi
tieu-de: Chu kì theo chiều dài dây
truc-ngang: Chiều dài l (m)
truc-doc: Chu kì T (s)
diem: 0.25, 1.0
diem: 0.5, 1.42
diem: 1.0, 2.01
diem: 2.0, 2.84
loi: Khi chiều dài tăng từ một phần tư mét đến hai mét, chu kì tăng từ một giây đến gần ba giây. Đường tăng chậm dần.

## Cảnh 8
loai: thi-nghiem
mau: li-con-lac-don
tham-so: 0 chieu-dai 0.4
tham-so: 6 chieu-dai 1.6
do: chu-ki
loi: Hãy quan sát. Khi ta tăng chiều dài dây, chu kì dao động tăng theo.
```

- [ ] **Step 2: Viết test thất bại**

Tạo `tools/vi/tests/test_video_ma_parse.py`:

```python
"""Test bộ đọc và bộ kiểm giới hạn của video.md (không cần Chromium, FFmpeg, mạng)."""

import sys
import tempfile
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))

from video_ma_parts import kiem, parse  # noqa: E402

MAU = TOOLS_VI / "fixtures" / "video-mau" / "video.md"
META = "tieu-de: T\nmon: Toán\nlop: 8\n"


def doc(canh: str, meta: str = META) -> str:
    return f"---\n{meta}---\n\n{canh}"


def line_of(text: str, needle: str) -> int:
    for number, line in enumerate(text.splitlines(), 1):
        if needle in line:
            return number
    raise AssertionError(needle)


class FixtureTest(unittest.TestCase):
    def test_sample_parses_all_eight_scene_types_in_order(self):
        video = parse.parse(MAU.read_text(encoding="utf-8"))
        self.assertEqual([c.loai for c in video.canh], list(parse.SCENE_TYPES))
        self.assertEqual([c.so for c in video.canh], list(range(1, 9)))
        self.assertEqual(video.meta["giong"], "nu")
        self.assertEqual(video.meta["toc-do"], "vua")
        self.assertEqual(video.meta["phu-de"], "hinh")
        self.assertEqual(video.meta["phong-cach"], "viet-tay")

    def test_sample_passes_the_limits_without_warnings(self):
        video = parse.parse(MAU.read_text(encoding="utf-8"))
        self.assertEqual(kiem.kiem(video, Path(".")), [])

    def test_repeated_fields_keep_order_and_line_numbers(self):
        text = MAU.read_text(encoding="utf-8")
        video = parse.parse(text)
        y = video.canh[3]
        self.assertEqual(len(y.truong["y"]), 3)
        self.assertEqual(y.dong_truong["y"][0], line_of(text, "Chiều dài dây l: dây dài"))


class ParseErrorTest(unittest.TestCase):
    def assert_error(self, text: str, line: int, fragment: str):
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse(text)
        self.assertEqual(caught.exception.line_no, line, str(caught.exception))
        self.assertIn(fragment, str(caught.exception))

    def test_missing_front_matter(self):
        self.assert_error("## Cảnh 1\nloai: tieu-de\n", 1, "---")

    def test_missing_required_meta_points_at_the_closing_line(self):
        self.assert_error("---\ntieu-de: T\nmon: Toán\n---\n", 4, "lop")

    def test_unknown_meta_key(self):
        text = doc("## Cảnh 1\n", META + "mau-nen: do\n")
        self.assert_error(text, 5, "mau-nen")

    def test_bad_meta_value(self):
        text = doc("## Cảnh 1\n", META + "giong: tre-em\n")
        self.assert_error(text, 5, "giong")

    def test_no_scene(self):
        self.assert_error(doc(""), 6, "Cảnh 1")

    def test_scene_numbers_must_be_sequential(self):
        text = doc("## Cảnh 2\nloai: tieu-de\nchu: A\nloi: Xin chào.\n")
        self.assert_error(text, 7, "Cảnh 1")

    def test_missing_loi(self):
        text = doc("## Cảnh 1\nloai: tieu-de\nchu: A\n")
        self.assert_error(text, 7, "loi")

    def test_unknown_scene_type(self):
        text = doc("## Cảnh 1\nloai: hoat-hinh\nloi: Xin chào.\n")
        self.assert_error(text, 8, "loai")

    def test_unknown_field_for_the_type(self):
        text = doc("## Cảnh 1\nloai: tieu-de\nchu: A\nbuoc: B\nloi: Xin chào.\n")
        self.assert_error(text, 10, "buoc")

    def test_duplicate_single_field(self):
        text = doc("## Cảnh 1\nloai: tieu-de\nchu: A\nchu: B\nloi: Xin chào.\n")
        self.assert_error(text, 10, "chu")

    def test_too_many_repeated_items(self):
        ys = "".join(f"y: ý {i}\n" for i in range(7))
        text = doc(f"## Cảnh 1\nloai: y-tung-y\ntieu-de: A\n{ys}loi: Xin chào.\n")
        self.assert_error(text, line_of(text, "ý 6"), "tối đa 6")

    def test_too_few_repeated_items(self):
        text = doc("## Cảnh 1\nloai: quy-trinh\ntieu-de: A\nbuoc: B\nloi: Xin chào.\n")
        self.assert_error(text, 7, "buoc")

    def test_point_needs_two_numbers_with_dot_decimals(self):
        text = doc("## Cảnh 1\nloai: do-thi\ntieu-de: A\ntruc-ngang: x\ntruc-doc: y\ndiem: 1;2\ndiem: 2, 3\nloi: Xin chào.\n")
        self.assert_error(text, line_of(text, "1;2"), "x, y")

    def test_parameter_line_needs_three_parts(self):
        text = doc("## Cảnh 1\nloai: thi-nghiem\nmau: li-con-lac-don\ntham-so: 0 chieu-dai\nloi: Xin chào.\n")
        self.assert_error(text, line_of(text, "tham-so"), "<giây> <mã> <giá trị>")

    def test_web_addresses_are_refused(self):
        text = doc("## Cảnh 1\nloai: tieu-de\nchu: xem https://a.vn\nloi: Xin chào.\n")
        self.assert_error(text, line_of(text, "https"), "địa chỉ web")

    def test_empty_value(self):
        text = doc("## Cảnh 1\nloai: tieu-de\nchu:\nloi: Xin chào.\n")
        self.assert_error(text, line_of(text, "chu:"), "trống")

    def test_line_that_is_not_a_field(self):
        text = doc("## Cảnh 1\nloai: tieu-de\nĐây là một dòng lạ\nloi: Xin chào.\n")
        self.assert_error(text, line_of(text, "dòng lạ"), "khoá: giá trị")


class LimitTest(unittest.TestCase):
    def video(self, canh: str):
        return parse.parse(doc(canh))

    def test_title_over_limit_names_scene_and_line(self):
        text = doc(f"## Cảnh 1\nloai: tieu-de\nchu: {'a' * 91}\nloi: Xin chào.\n")
        with self.assertRaises(kiem.CanhError) as caught:
            kiem.kiem(parse.parse(text), Path("."))
        self.assertEqual(caught.exception.so, 1)
        self.assertIn("91", str(caught.exception))
        self.assertIn(f"dòng {line_of(text, 'aaa')}", str(caught.exception))

    def test_markup_characters_are_not_counted(self):
        video = self.video(f"## Cảnh 1\nloai: tieu-de\nchu: **{'a' * 90}**\nloi: Xin chào.\n")
        self.assertEqual(kiem.kiem(video, Path(".")), [])
        self.assertEqual(kiem.hien_thi("H~2~SO~4~ m/s^2^"), len("H2SO4 m/s2"))

    def test_long_narration_is_a_warning_not_an_error(self):
        video = self.video(f"## Cảnh 1\nloai: tieu-de\nchu: A\nloi: {'Câu này dài. ' * 60}\n")
        warnings = kiem.kiem(video, Path("."))
        self.assertEqual(len(warnings), 1)
        self.assertIn("Cảnh 1", warnings[0])
        self.assertIn("700", warnings[0])


class ExperimentSceneTest(unittest.TestCase):
    def check(self, extra: str):
        text = doc(f"## Cảnh 1\nloai: thi-nghiem\nmau: li-con-lac-don\n{extra}loi: Xin chào.\n")
        return text, lambda: kiem.kiem(parse.parse(text), Path("."))

    def test_new_model_is_not_allowed_in_video(self):
        text = doc("## Cảnh 1\nloai: thi-nghiem\nmau: moi\nloi: Xin chào.\n")
        with self.assertRaises(kiem.CanhError) as caught:
            kiem.kiem(parse.parse(text), Path("."))
        self.assertIn("thư viện", str(caught.exception))

    def test_unknown_model_lists_the_library(self):
        text = doc("## Cảnh 1\nloai: thi-nghiem\nmau: li-khong-co\nloi: Xin chào.\n")
        with self.assertRaises(kiem.CanhError) as caught:
            kiem.kiem(parse.parse(text), Path("."))
        self.assertIn("li-con-lac-don", str(caught.exception))

    def test_unknown_parameter(self):
        _, run = self.check("tham-so: 0 toc-do 3\n")
        with self.assertRaises(kiem.CanhError) as caught:
            run()
        self.assertIn("toc-do", str(caught.exception))

    def test_value_outside_the_allowed_range_names_the_line(self):
        text, run = self.check("tham-so: 0 chieu-dai 5\n")
        with self.assertRaises(parse.ParseError) as caught:
            run()
        self.assertEqual(caught.exception.line_no, line_of(text, "tham-so"))
        self.assertIn("0.2", str(caught.exception))

    def test_more_than_three_parameters(self):
        _, run = self.check("tham-so: 0 chieu-dai 1\ntham-so: 0 g 9.8\ntham-so: 0 goc-lech 8\ntham-so: 0 khoi-luong 0.2\n")
        with self.assertRaises(kiem.CanhError):
            run()

    def test_unknown_measured_quantity(self):
        _, run = self.check("do: van-toc\n")
        with self.assertRaises(kiem.CanhError) as caught:
            run()
        self.assertIn("chu-ki", str(caught.exception))

    def test_schedule_is_sorted_and_default_measures_are_the_first_two(self):
        from thi_nghiem_parts import thu_vien

        text = doc("## Cảnh 1\nloai: thi-nghiem\nmau: li-con-lac-don\ntham-so: 6 chieu-dai 1.6\ntham-so: 0 chieu-dai 0.4\nloi: Xin chào.\n")
        scene = parse.parse(text).canh[0]
        self.assertEqual(kiem.tham_so_theo_thoi_gian(scene), {"chieu-dai": [(0.0, 0.4), (6.0, 1.6)]})
        model = thu_vien.load("li-con-lac-don", Path("."))
        self.assertEqual(kiem.ma_do(scene, model), ["chu-ki", "thoi-gian-10-dao-dong"])


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 3: Chạy test để thấy lỗi**

Run: `venv\Scripts\python.exe -m unittest tools/vi/tests/test_video_ma_parse.py`
Expected: FAIL — `ModuleNotFoundError: video_ma_parts`.

- [ ] **Step 4: Viết `parse.py`**

Tạo `tools/vi/video_ma_parts/__init__.py` rỗng, rồi `tools/vi/video_ma_parts/parse.py`:

```python
"""Đọc video.md thành dữ liệu; lỗi luôn kèm số dòng. Ngữ pháp ở docs/vi/tro-ly/video-giai-thich.md. Chỉ dùng thư viện chuẩn."""

from __future__ import annotations

import re
from dataclasses import dataclass

META_REQUIRED = ("tieu-de", "mon", "lop")
META_CHOICES = {
    "phong-cach": ("viet-tay",),
    "giong": ("nu", "nam"),
    "toc-do": ("cham", "vua", "nhanh"),
    "phu-de": ("hinh", "file", "khong"),
}
META_DEFAULTS = {"phong-cach": "viet-tay", "giong": "nu", "toc-do": "vua", "phu-de": "hinh"}

# loại cảnh -> (trường đơn bắt buộc, trường đơn tuỳ chọn, trường lặp {khoá: (tối thiểu, tối đa)})
SCENE_SPEC = {
    "tieu-de": (("chu",), ("phu",), {}),
    "khai-niem": (("thuat-ngu", "dinh-nghia"), (), {}),
    "cong-thuc": (("bieu-thuc",), (), {"giai-thich": (0, 4)}),
    "y-tung-y": (("tieu-de",), (), {"y": (1, 6)}),
    "quy-trinh": (("tieu-de",), (), {"buoc": (2, 5)}),
    "so-sanh": (("tieu-de", "trai", "phai"), (), {"y-trai": (1, 4), "y-phai": (1, 4)}),
    "do-thi": (("tieu-de", "truc-ngang", "truc-doc"), (), {"diem": (2, 12)}),
    "thi-nghiem": (("mau",), ("do",), {"tham-so": (0, 99)}),
}
SCENE_TYPES = tuple(SCENE_SPEC)

_KEY_RE = re.compile(r"^([a-z][a-z0-9-]*):\s*(.*)$")
_SCENE_RE = re.compile(r"^##\s+Cảnh\s+(\d+)\s*$")
_URL_RE = re.compile(r"https?://|www\.", re.IGNORECASE)
_POINT_RE = re.compile(r"^-?\d+(?:\.\d+)?\s*,\s*-?\d+(?:\.\d+)?$")
_PARAM_RE = re.compile(r"^\d+(?:\.\d+)?\s+[a-z0-9-]+\s+-?\d+(?:\.\d+)?$")


class ParseError(Exception):
    """Lỗi trong video.md; luôn kèm số dòng để AI sửa được đúng chỗ."""

    def __init__(self, line_no: int, message: str) -> None:
        super().__init__(f"Dòng {line_no}: {message}")
        self.line_no = line_no
        self.message = message


@dataclass
class Scene:
    so: int
    dong: int
    loai: str
    loi: str
    truong: dict
    dong_truong: dict


@dataclass
class Video:
    meta: dict
    canh: list


def _check_value(no: int, key: str, value: str) -> None:
    if value == "":
        raise ParseError(no, f"`{key}` đang trống.")
    if _URL_RE.search(value):
        raise ParseError(no, f"`{key}` chứa địa chỉ web. Không chèn địa chỉ web vào video.")


def _read_meta(lines: list, start: int) -> tuple:
    meta: dict = {}
    i = start
    while i < len(lines):
        raw = lines[i].strip()
        if raw == "---":
            break
        if raw:
            match = _KEY_RE.match(raw)
            if match is None:
                raise ParseError(i + 1, "Dòng thông tin phải có dạng `khoá: giá trị`.")
            key, value = match.group(1), match.group(2).strip()
            if key not in META_REQUIRED and key not in META_CHOICES:
                raise ParseError(i + 1, f"Khoá `{key}` không có trong khối thông tin.")
            if key in meta:
                raise ParseError(i + 1, f"Khoá `{key}` bị lặp.")
            _check_value(i + 1, key, value)
            if key in META_CHOICES and value not in META_CHOICES[key]:
                raise ParseError(i + 1, f"`{key}` phải là một trong: {', '.join(META_CHOICES[key])}.")
            meta[key] = value
        i += 1
    else:
        raise ParseError(start, "Khối thông tin chưa đóng bằng dòng `---`.")
    for key in META_REQUIRED:
        if key not in meta:
            raise ParseError(i + 1, f"Khối thông tin thiếu `{key}`.")
    for key, default in META_DEFAULTS.items():
        meta.setdefault(key, default)
    return meta, i + 1


def _finish(so: int, dong0: int, fields: list) -> Scene:
    truong: dict = {}
    dong_truong: dict = {}
    for key, value, no in fields:
        truong.setdefault(key, []).append(value)
        dong_truong.setdefault(key, []).append(no)
    for key in ("loai", "loi"):
        if key not in truong:
            raise ParseError(dong0, f"Cảnh {so} thiếu `{key}`.")
        if len(truong[key]) > 1:
            raise ParseError(dong_truong[key][1], f"`{key}` bị lặp trong Cảnh {so}.")
    loai = truong["loai"][0]
    if loai not in SCENE_SPEC:
        raise ParseError(dong_truong["loai"][0], f"`loai` phải là một trong: {', '.join(SCENE_TYPES)}.")
    required, optional, repeated = SCENE_SPEC[loai]
    allowed = {"loai", "loi", *required, *optional, *repeated}
    for key, value, no in fields:
        if key not in allowed:
            raise ParseError(no, f"Cảnh loại `{loai}` không có trường `{key}`.")
    for key in (*required, *optional):
        if key in truong and len(truong[key]) > 1:
            raise ParseError(dong_truong[key][1], f"`{key}` bị lặp trong Cảnh {so}.")
    for key in required:
        if key not in truong:
            raise ParseError(dong0, f"Cảnh {so} (loại `{loai}`) thiếu `{key}`.")
    for key, (low, high) in repeated.items():
        count = len(truong.get(key, []))
        if count < low:
            raise ParseError(dong0, f"Cảnh {so} (loại `{loai}`) cần ít nhất {low} dòng `{key}`.")
        if count > high:
            raise ParseError(dong_truong[key][high], f"Cảnh loại `{loai}` chỉ có tối đa {high} dòng `{key}`.")
    for value, no in zip(truong.get("diem", []), dong_truong.get("diem", [])):
        if _POINT_RE.match(value) is None:
            raise ParseError(no, "Điểm đồ thị phải có dạng `x, y` (hai số, dấu thập phân là dấu chấm).")
    for value, no in zip(truong.get("tham-so", []), dong_truong.get("tham-so", [])):
        if _PARAM_RE.match(value) is None:
            raise ParseError(no, "Dòng `tham-so` phải có dạng `<giây> <mã> <giá trị>`, ví dụ `0 chieu-dai 0.4`.")
    loi = truong.pop("loi")[0]
    truong.pop("loai")
    dong_truong.pop("loai")
    dong_truong.pop("loi")
    return Scene(so=so, dong=dong0, loai=loai, loi=loi, truong=truong, dong_truong=dong_truong)


def parse(text: str) -> Video:
    lines = text.lstrip("\ufeff").splitlines()
    i = 0
    while i < len(lines) and not lines[i].strip():
        i += 1
    if i >= len(lines) or lines[i].strip() != "---":
        raise ParseError(i + 1, "video.md phải mở đầu bằng khối thông tin giữa hai dòng `---`.")
    meta, i = _read_meta(lines, i + 1)
    scenes: list = []
    current = None
    for index in range(i, len(lines)):
        no = index + 1
        raw = lines[index].strip()
        if not raw:
            continue
        heading = _SCENE_RE.match(raw)
        if heading:
            if current is not None:
                scenes.append(_finish(*current))
            expected = len(scenes) + 1
            if int(heading.group(1)) != expected:
                raise ParseError(no, f"Cảnh phải đánh số liên tiếp; ở đây phải là `## Cảnh {expected}`.")
            current = (expected, no, [])
            continue
        if current is None:
            raise ParseError(no, "Sau khối thông tin phải là `## Cảnh 1`.")
        match = _KEY_RE.match(raw)
        if match is None:
            raise ParseError(no, "Dòng trong cảnh phải có dạng `khoá: giá trị`.")
        key, value = match.group(1), match.group(2).strip()
        _check_value(no, key, value)
        current[2].append((key, value, no))
    if current is not None:
        scenes.append(_finish(*current))
    if not scenes:
        raise ParseError(i + 1, "video.md chưa có cảnh nào; bắt đầu bằng `## Cảnh 1`.")
    return Video(meta=meta, canh=scenes)
```

- [ ] **Step 5: Viết `kiem.py`**

Tạo `tools/vi/video_ma_parts/kiem.py`:

```python
"""Kiểm giới hạn chữ và nội dung cảnh của video.md. Chỉ dùng thư viện chuẩn."""

from __future__ import annotations

import re
from pathlib import Path

from thi_nghiem_parts import thu_vien

from .parse import ParseError, Scene, Video

LIMITS = {
    ("tieu-de", "chu"): 90, ("tieu-de", "phu"): 90,
    ("khai-niem", "thuat-ngu"): 60, ("khai-niem", "dinh-nghia"): 220,
    ("cong-thuc", "bieu-thuc"): 90, ("cong-thuc", "giai-thich"): 80,
    ("y-tung-y", "tieu-de"): 90, ("y-tung-y", "y"): 80,
    ("quy-trinh", "tieu-de"): 90, ("quy-trinh", "buoc"): 50,
    ("so-sanh", "tieu-de"): 90, ("so-sanh", "trai"): 24, ("so-sanh", "phai"): 24,
    ("so-sanh", "y-trai"): 60, ("so-sanh", "y-phai"): 60,
    ("do-thi", "tieu-de"): 90, ("do-thi", "truc-ngang"): 40, ("do-thi", "truc-doc"): 40,
}
LOI_DAI = 700
MAX_THAM_SO = 3
MAX_DO = 3
_MARKUP_RE = re.compile(r"\*\*|~|\^")


class CanhError(Exception):
    """Nội dung cảnh sai; luôn nêu số cảnh."""

    def __init__(self, so: int, message: str) -> None:
        super().__init__(f"Cảnh {so}: {message}")
        self.so = so
        self.message = message


def hien_thi(chu: str) -> int:
    return len(_MARKUP_RE.sub("", chu))


def tham_so_theo_thoi_gian(scene: Scene) -> dict:
    out: dict = {}
    for value in scene.truong.get("tham-so", []):
        giay, ma, gia_tri = value.split()
        out.setdefault(ma, []).append((float(giay), float(gia_tri)))
    for ma in out:
        out[ma].sort(key=lambda mot: mot[0])
    return out


def ma_do(scene: Scene, model) -> list:
    if "do" in scene.truong:
        return [ma.strip() for ma in scene.truong["do"][0].split(",") if ma.strip()]
    return [dl["ma"] for dl in model.khai_bao["daiLuongDo"][:2]]


def _kiem_thi_nghiem(scene: Scene, thu_muc: Path) -> None:
    mau = scene.truong["mau"][0]
    if mau == thu_vien.NEW_MODEL:
        raise CanhError(scene.so, "Video chỉ dùng mẫu có sẵn trong thư viện thí nghiệm, không dùng `moi`. Mẫu có: "
                        + ", ".join(thu_vien.list_models()) + ".")
    try:
        model = thu_vien.load(mau, thu_muc)
    except thu_vien.ModelError as exc:
        raise CanhError(scene.so, str(exc)) from exc
    lich = tham_so_theo_thoi_gian(scene)
    dong_theo_ma: dict = {}
    for value, no in zip(scene.truong.get("tham-so", []), scene.dong_truong.get("tham-so", [])):
        dong_theo_ma.setdefault(value.split()[1], []).append((no, float(value.split()[2])))
    for ma, cac_dong in dong_theo_ma.items():
        ts = model.tham_so(ma)
        if ts is None:
            co = ", ".join(t["ma"] for t in model.khai_bao["thamSo"])
            raise CanhError(scene.so, f"tham số `{ma}` không có trong mẫu `{mau}` (dòng {cac_dong[0][0]}). Có: {co}.")
        if ts.get("kieu") != "so":
            raise CanhError(scene.so, f"tham số `{ma}` không phải số nên chưa dùng được trong video.")
        for no, gia_tri in cac_dong:
            if not ts["min"] <= gia_tri <= ts["max"]:
                raise ParseError(no, f"`{ma}` = {gia_tri:g} ngoài khoảng cho phép {ts['min']:g}–{ts['max']:g}.")
    if len(lich) > MAX_THAM_SO:
        raise CanhError(scene.so, f"chỉ đổi tối đa {MAX_THAM_SO} tham số khác nhau trong một cảnh (đang có {len(lich)}).")
    codes = ma_do(scene, model)
    if len(codes) > MAX_DO:
        raise CanhError(scene.so, f"`do` chỉ liệt kê tối đa {MAX_DO} đại lượng (đang có {len(codes)}).")
    for ma in codes:
        if model.dai_luong(ma) is None:
            co = ", ".join(dl["ma"] for dl in model.khai_bao["daiLuongDo"])
            raise CanhError(scene.so, f"đại lượng đo `{ma}` không có trong mẫu `{mau}`. Có: {co}.")


def kiem(video: Video, thu_muc: Path) -> list:
    warnings: list = []
    for scene in video.canh:
        for key, values in scene.truong.items():
            gioi_han = LIMITS.get((scene.loai, key))
            if gioi_han is None:
                continue
            for value, no in zip(values, scene.dong_truong[key]):
                so_ky_tu = hien_thi(value)
                if so_ky_tu > gioi_han:
                    raise CanhError(scene.so, f"`{key}` dài {so_ky_tu} ký tự, tối đa {gioi_han} (dòng {no}). Rút gọn nội dung.")
        if len(scene.loi) > LOI_DAI:
            warnings.append(f"Cảnh {scene.so}: lời dài {len(scene.loi)} ký tự (quá {LOI_DAI}); nên tách thành hai cảnh.")
        if scene.loai == "thi-nghiem":
            _kiem_thi_nghiem(scene, thu_muc)
    return warnings
```

- [ ] **Step 6: Chạy test đến khi xanh**

Run: `venv\Scripts\python.exe -m unittest tools/vi/tests/test_video_ma_parse.py -v`
Expected: PASS (mọi test). Nếu một test lệch số dòng vì tính sai, sửa **test** chỉ khi bộ đọc trỏ đúng dòng lỗi thật; không nới lỏng kiểm tra.

- [ ] **Step 7: Commit**

```bash
git add tools/vi/video_ma_parts/__init__.py tools/vi/video_ma_parts/parse.py tools/vi/video_ma_parts/kiem.py tools/vi/fixtures/video-mau/video.md tools/vi/tests/test_video_ma_parse.py
git commit -m "feat(vi): read and validate video.md scripts for code-built videos"
```

---

### Task 3: Lịch thời gian

**Files:**
- Create: `tools/vi/video_ma_parts/lich.py`
- Test: `tools/vi/tests/test_video_ma_lich.py`

**Interfaces:**
- Consumes: `parse.Scene`; `kiem.CanhError`, `kiem.tham_so_theo_thoi_gian`, `kiem.ma_do`.
- Produces: hằng `FPS=15`, `DAN_DAU=0.7`, `DUOI=0.6`, `TOI_THIEU=2.5`, `CANH_DAI=40.0`, `VIDEO_DAI=480.0`; `GiongInfo(mp3: Path|None, giay: float, moc_cau: list[float], uoc_luong: bool, nguon: str)` (`moc_cau` **tính từ lúc tiếng bắt đầu**, rỗng nếu không có; `nguon` là `"may"` hoặc `"co-san"`); `CanhLich(so, bat_dau, thoi_luong, so_khung, giay_giong, cau, moc_cau, moc_cau_giong, uoc_luong)` (`moc_cau` tính từ đầu cảnh = `moc_cau_giong` + 0,7); `tach_cau(loi) -> list[str]`; `moc_uoc_luong(cau, giay) -> list[float]`; `thoi_luong_canh(giay_giong, fps=FPS) -> float`; `moc_hien(so_muc, moc_cau_giong, giay) -> list[float]` (tính từ lúc tiếng bắt đầu); `so_muc(scene) -> int`; `dung_lich(cac_canh, cac_giong, fps=FPS, kiem_moc=True) -> tuple[list[CanhLich], list[str]]`; `du_lieu_canh(scene, cl, model=None) -> dict` (dữ liệu JSON cho trang, khoá: `so`, `loai`, `thoiLuong`, `danDau`, `truong`, `moc`, và `diem` cho đồ thị, `khaiBao`/`thamSo`/`do` cho thí nghiệm).

- [ ] **Step 1: Viết test thất bại**

Tạo `tools/vi/tests/test_video_ma_lich.py`:

```python
"""Test lịch thời gian của video giải thích: câu, mốc hiện ý, thời lượng cảnh, dữ liệu cảnh."""

import sys
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))

from thi_nghiem_parts import thu_vien  # noqa: E402
from video_ma_parts import kiem, lich, parse  # noqa: E402

META = "tieu-de: T\nmon: Toán\nlop: 8\n"


def canh_dau(noi_dung: str, loi: str):
    text = f"---\n{META}---\n\n## Cảnh 1\n{noi_dung}loi: {loi}\n"
    return parse.parse(text).canh[0]


def giong(giay: float, moc=None, uoc=False):
    return lich.GiongInfo(mp3=None, giay=giay, moc_cau=moc or [], uoc_luong=uoc, nguon="may")


class SentenceTest(unittest.TestCase):
    def test_split_on_terminators_and_keep_ellipsis(self):
        self.assertEqual(lich.tach_cau("Một. Hai! Ba? Bốn… Năm."), ["Một.", "Hai!", "Ba?", "Bốn…", "Năm."])

    def test_no_terminator_is_one_sentence(self):
        self.assertEqual(lich.tach_cau("Không có dấu chấm cuối"), ["Không có dấu chấm cuối"])

    def test_estimated_marks_follow_character_ratio(self):
        marks = lich.moc_uoc_luong(["aaaa", "bb", "cccccc"], 12.0)
        self.assertEqual(marks, [0.0, 4.0, 6.0])


class DurationTest(unittest.TestCase):
    def test_minimum_and_frame_multiple(self):
        toi_thieu = lich.thoi_luong_canh(0.1)
        self.assertGreaterEqual(toi_thieu, 2.5)
        self.assertLess(toi_thieu, 2.5 + 1 / lich.FPS)
        long = lich.thoi_luong_canh(4.13)
        self.assertGreaterEqual(long, 0.7 + 4.13 + 0.6)
        self.assertAlmostEqual(long * lich.FPS, round(long * lich.FPS), places=6)
        self.assertLess(long, 0.7 + 4.13 + 0.6 + 1 / lich.FPS)


class RevealTimesTest(unittest.TestCase):
    def test_item_k_appears_at_sentence_k(self):
        self.assertEqual(lich.moc_hien(3, [0.0, 2.0, 5.0, 7.0], 9.0), [0.0, 2.0, 5.0])

    def test_fewer_sentences_than_items_spreads_evenly(self):
        self.assertEqual(lich.moc_hien(4, [0.0], 8.0), [0.0, 2.0, 4.0, 6.0])

    def test_loi_mot_cau_nhieu_y(self):
        scene = canh_dau("loai: y-tung-y\ntieu-de: A\ny: a\ny: b\ny: c\ny: d\ny: e\ny: g\n",
                         "Một câu duy nhất không có dấu chấm cuối")
        plan, _ = lich.dung_lich([scene], [giong(12.0)])
        du = lich.du_lieu_canh(scene, plan[0])
        self.assertEqual(len(du["moc"]), 6)
        gaps = [b - a for a, b in zip(du["moc"], du["moc"][1:])]
        self.assertTrue(all(abs(g - 2.0) < 1e-6 for g in gaps), gaps)
        self.assertAlmostEqual(du["moc"][0], lich.DAN_DAU)

    def test_items_per_scene_type(self):
        text = f"---\n{META}---\n\n## Cảnh 1\nloai: so-sanh\ntieu-de: A\ntrai: T\nphai: P\ny-trai: a\ny-trai: b\ny-phai: c\nloi: Ok.\n"
        self.assertEqual(lich.so_muc(parse.parse(text).canh[0]), 3)


class PlanTest(unittest.TestCase):
    def scenes(self):
        text = (f"---\n{META}---\n\n## Cảnh 1\nloai: tieu-de\nchu: A\nloi: Xin chào các em. Hôm nay học bài mới.\n\n"
                "## Cảnh 2\nloai: tieu-de\nchu: B\nloi: Tiếp theo.\n")
        return parse.parse(text).canh

    def test_offsets_accumulate_and_frames_match(self):
        plan, warnings = lich.dung_lich(self.scenes(), [giong(4.0, [0.0, 1.9]), giong(0.5, [0.0])])
        self.assertEqual(warnings, [])
        self.assertEqual(plan[0].bat_dau, 0.0)
        self.assertAlmostEqual(plan[1].bat_dau, plan[0].thoi_luong)
        self.assertEqual(plan[1].thoi_luong, lich.thoi_luong_canh(0.5))
        self.assertGreaterEqual(plan[1].thoi_luong, 2.5)
        for cl in plan:
            self.assertEqual(cl.so_khung, round(cl.thoi_luong * lich.FPS))
        self.assertEqual(plan[0].cau, ["Xin chào các em.", "Hôm nay học bài mới."])
        self.assertEqual(plan[0].moc_cau_giong, [0.0, 1.9])
        self.assertAlmostEqual(plan[0].moc_cau[1], 1.9 + lich.DAN_DAU)

    def test_missing_marks_are_estimated_with_a_warning(self):
        plan, warnings = lich.dung_lich(self.scenes(), [giong(4.0, [], True), giong(1.0, [0.0])])
        self.assertTrue(plan[0].uoc_luong)
        self.assertEqual(len(warnings), 1)
        self.assertIn("Cảnh 1", warnings[0])
        self.assertIn("ước lượng", warnings[0])

    def test_mark_count_mismatch_falls_back_to_estimate(self):
        plan, warnings = lich.dung_lich(self.scenes(), [giong(4.0, [0.0]), giong(1.0, [0.0])])
        self.assertTrue(plan[0].uoc_luong)
        self.assertEqual(len(warnings), 1)

    def test_long_scene_and_long_video_warn(self):
        plan, warnings = lich.dung_lich(self.scenes(), [giong(45.0, [0.0, 20.0]), giong(1.0, [0.0])])
        self.assertTrue(any("40 giây" in w for w in warnings))
        scenes = self.scenes() * 1
        _, warnings = lich.dung_lich(scenes, [giong(300.0, [0.0, 100.0]), giong(300.0, [0.0])])
        self.assertTrue(any("8 phút" in w for w in warnings))


class ExperimentDataTest(unittest.TestCase):
    def scene(self, extra: str):
        return canh_dau(f"loai: thi-nghiem\nmau: li-con-lac-don\n{extra}", "Hãy quan sát.")

    def test_mark_beyond_scene_duration_is_a_canh_error(self):
        scene = self.scene("tham-so: 30 chieu-dai 1.6\n")
        with self.assertRaises(kiem.CanhError) as caught:
            lich.dung_lich([scene], [giong(4.0, [0.0])])
        self.assertIn("30", str(caught.exception))
        self.assertIn("dòng", str(caught.exception))

    def test_preview_skips_the_mark_check(self):
        scene = self.scene("tham-so: 30 chieu-dai 1.6\n")
        plan, _ = lich.dung_lich([scene], [giong(4.0, [0.0])], kiem_moc=False)
        self.assertEqual(len(plan), 1)

    def test_scene_data_carries_model_schedule_and_measures(self):
        scene = self.scene("tham-so: 6 chieu-dai 1.6\ntham-so: 0 chieu-dai 0.4\ndo: chu-ki\n")
        plan, _ = lich.dung_lich([scene], [giong(8.0, [0.0])])
        model = thu_vien.load("li-con-lac-don", Path("."))
        du = lich.du_lieu_canh(scene, plan[0], model)
        self.assertEqual(du["loai"], "thi-nghiem")
        self.assertEqual(du["thamSo"], {"chieu-dai": [[0.0, 0.4], [6.0, 1.6]]})
        self.assertEqual(du["do"], ["chu-ki"])
        self.assertEqual(du["khaiBao"]["ma"], "li-con-lac-don")
        self.assertEqual(du["thoiLuong"], plan[0].thoi_luong)
        self.assertEqual(du["danDau"], lich.DAN_DAU)

    def test_graph_scene_data_has_numeric_points(self):
        scene = canh_dau("loai: do-thi\ntieu-de: A\ntruc-ngang: x\ntruc-doc: y\ndiem: 0.5, 1\ndiem: 2, 3.5\n", "Ok.")
        plan, _ = lich.dung_lich([scene], [giong(3.0, [0.0])])
        du = lich.du_lieu_canh(scene, plan[0])
        self.assertEqual(du["diem"], [[0.5, 1.0], [2.0, 3.5]])
        self.assertEqual(len(du["moc"]), 2)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Chạy test để thấy lỗi**

Run: `venv\Scripts\python.exe -m unittest tools/vi/tests/test_video_ma_lich.py`
Expected: FAIL — `ImportError: cannot import name 'lich'`.

- [ ] **Step 3: Viết `lich.py`**

```python
"""Lịch thời gian: tách câu, mốc câu, mốc hiện ý, thời lượng cảnh, dữ liệu JSON cho trang. Chỉ dùng thư viện chuẩn."""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from pathlib import Path

from .kiem import CanhError, ma_do, tham_so_theo_thoi_gian
from .parse import Scene

FPS = 15
DAN_DAU = 0.7
DUOI = 0.6
TOI_THIEU = 2.5
CANH_DAI = 40.0
VIDEO_DAI = 480.0
_CAU_RE = re.compile(r"(?<=[.!?…])\s+")


@dataclass
class GiongInfo:
    mp3: Path | None
    giay: float
    moc_cau: list
    uoc_luong: bool
    nguon: str


@dataclass
class CanhLich:
    so: int
    bat_dau: float
    thoi_luong: float
    so_khung: int
    giay_giong: float
    cau: list
    moc_cau: list
    moc_cau_giong: list
    uoc_luong: bool


def tach_cau(loi: str) -> list:
    return [cau.strip() for cau in _CAU_RE.split(loi.strip()) if cau.strip()]


def moc_uoc_luong(cau: list, giay: float) -> list:
    tong = sum(len(c) for c in cau) or 1
    da_qua = 0
    moc = []
    for c in cau:
        moc.append(round(da_qua / tong * giay, 3))
        da_qua += len(c)
    return moc


def thoi_luong_canh(giay_giong: float, fps: int = FPS) -> float:
    tho = max(TOI_THIEU, DAN_DAU + giay_giong + DUOI)
    return math.ceil(tho * fps - 1e-9) / fps


def moc_hien(so_muc: int, moc_cau_giong: list, giay: float) -> list:
    if so_muc <= 0:
        return []
    if len(moc_cau_giong) >= so_muc:
        return list(moc_cau_giong[:so_muc])
    return [round(k * giay / so_muc, 3) for k in range(so_muc)]


def so_muc(scene: Scene) -> int:
    t = scene.truong
    return {
        "y-tung-y": len(t.get("y", [])),
        "quy-trinh": len(t.get("buoc", [])),
        "so-sanh": len(t.get("y-trai", [])) + len(t.get("y-phai", [])),
        "do-thi": len(t.get("diem", [])),
        "cong-thuc": len(t.get("giai-thich", [])),
    }.get(scene.loai, 0)


def dung_lich(cac_canh: list, cac_giong: list, fps: int = FPS, kiem_moc: bool = True) -> tuple:
    plan: list = []
    warnings: list = []
    bat_dau = 0.0
    for scene, giong in zip(cac_canh, cac_giong):
        cau = tach_cau(scene.loi)
        uoc = giong.uoc_luong or len(giong.moc_cau) != len(cau)
        moc_giong = moc_uoc_luong(cau, giong.giay) if uoc else list(giong.moc_cau)
        if uoc:
            warnings.append(f"Cảnh {scene.so}: mốc câu ước lượng theo số ký tự; hình có thể lệch tiếng vài trăm mili giây.")
        thoi_luong = thoi_luong_canh(giong.giay, fps)
        if thoi_luong > CANH_DAI:
            warnings.append(f"Cảnh {scene.so}: dài {thoi_luong:.0f} giây (quá 40 giây); nên tách thành hai cảnh.")
        if kiem_moc and scene.loai == "thi-nghiem":
            for value, no in zip(scene.truong.get("tham-so", []), scene.dong_truong.get("tham-so", [])):
                giay = float(value.split()[0])
                if giay > thoi_luong:
                    raise CanhError(scene.so, f"mốc `tham-so` {giay:g} giây (dòng {no}) vượt thời lượng cảnh {thoi_luong:.1f} giây.")
        plan.append(CanhLich(
            so=scene.so, bat_dau=round(bat_dau, 6), thoi_luong=thoi_luong, so_khung=round(thoi_luong * fps),
            giay_giong=giong.giay, cau=cau, moc_cau=[round(DAN_DAU + m, 3) for m in moc_giong],
            moc_cau_giong=moc_giong, uoc_luong=uoc,
        ))
        bat_dau += thoi_luong
    if bat_dau > VIDEO_DAI:
        warnings.append(f"Video dài {bat_dau / 60:.1f} phút (quá 8 phút); nên tách thành nhiều video.")
    return plan, warnings


def du_lieu_canh(scene: Scene, cl: CanhLich, model=None) -> dict:
    du = {
        "so": scene.so,
        "loai": scene.loai,
        "thoiLuong": cl.thoi_luong,
        "danDau": DAN_DAU,
        "truong": scene.truong,
        "moc": [round(DAN_DAU + m, 3) for m in moc_hien(so_muc(scene), cl.moc_cau_giong, cl.giay_giong)],
    }
    if scene.loai == "do-thi":
        du["diem"] = [[float(p) for p in v.split(",")] for v in scene.truong["diem"]]
    if scene.loai == "thi-nghiem":
        du["khaiBao"] = model.khai_bao
        du["thamSo"] = {ma: [[g, v] for g, v in ds] for ma, ds in tham_so_theo_thoi_gian(scene).items()}
        du["do"] = ma_do(scene, model)
    return du
```

- [ ] **Step 4: Chạy test đến khi xanh**

Run: `venv\Scripts\python.exe -m unittest tools/vi/tests/test_video_ma_lich.py -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add tools/vi/video_ma_parts/lich.py tools/vi/tests/test_video_ma_lich.py
git commit -m "feat(vi): schedule sentence marks, reveal times and scene durations"
```

---

### Task 4: Giọng đọc

**Files:**
- Create: `tools/vi/video_ma_parts/giong.py`
- Test: `tools/vi/tests/test_video_ma_giong.py`

**Interfaces:**
- Consumes: `lich.GiongInfo`; `video_parts.media.MediaError`, `media.probe_duration`.
- Produces: `VOICES`, `RATES`; `bam(loi, voice, rate) -> str`; `tong_hop_edge(text, voice, rate, out_path) -> list[float]` (mốc câu, giây, tính từ đầu tiếng); `lay_giong(so, loi, thu_muc, giong, toc_do, tong_hop=tong_hop_edge, do_dai=media.probe_duration) -> GiongInfo`. Lỗi: `media.MediaError` với `step == "giong"` (hoặc `"ffmpeg"` nếu không chạy được `ffprobe`).

- [ ] **Step 1: Viết test thất bại**

Tạo `tools/vi/tests/test_video_ma_giong.py`:

```python
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
```

- [ ] **Step 2: Chạy test để thấy lỗi**

Run: `venv\Scripts\python.exe -m unittest tools/vi/tests/test_video_ma_giong.py`
Expected: FAIL — `ImportError: cannot import name 'giong'`.

- [ ] **Step 3: Viết `giong.py`**

```python
"""Giọng đọc từng cảnh: file thầy cô đặt sẵn, giọng máy edge-tts, hoặc bản tạo lần trước còn dùng được."""

from __future__ import annotations

import asyncio
import hashlib
import json
import os
from pathlib import Path
from typing import Callable

from video_parts import media

from .lich import GiongInfo

VOICES = {"nu": "vi-VN-HoaiMyNeural", "nam": "vi-VN-NamMinhNeural"}
RATES = {"cham": "-10%", "vua": "+0%", "nhanh": "+15%"}
FIX_GIONG = "Có mạng rồi chạy lại, hoặc đặt sẵn file giọng giong/canh-<số>.mp3 cho từng cảnh."
FIX_EDGE = "Cài edge-tts bằng: python -m pip install -r requirements.txt (ở thư mục gốc repo)."
FIX_FILE = "Xoá hoặc thay file giọng đó rồi chạy lại."


def bam(loi: str, voice: str, rate: str) -> str:
    return hashlib.sha256(f"{voice}|{rate}|{loi}".encode("utf-8")).hexdigest()[:16]


async def _tong_hop(text: str, voice: str, rate: str, out_path: Path) -> list:
    import edge_tts

    moc: list = []
    with open(out_path, "wb") as f:
        communicate = edge_tts.Communicate(text, voice, rate=rate, boundary="SentenceBoundary")
        async for chunk in communicate.stream():
            if chunk["type"] == "audio":
                f.write(chunk["data"])
            elif chunk["type"] == "SentenceBoundary":
                moc.append(round(chunk["offset"] / 1e7, 3))
    return moc


def tong_hop_edge(text: str, voice: str, rate: str, out_path: Path) -> list:
    try:
        import edge_tts  # noqa: F401
    except ImportError as exc:
        raise media.MediaError("giong", "Chưa cài edge-tts.", FIX_EDGE) from exc
    try:
        return asyncio.run(_tong_hop(text, voice, rate, out_path))
    except Exception as exc:
        raise media.MediaError("giong", f"Không tạo được giọng đọc (thường do mất mạng): {exc}", FIX_GIONG) from exc


def _giay(mp3: Path, do_dai: Callable) -> float:
    if not mp3.is_file() or mp3.stat().st_size == 0:
        raise media.MediaError("giong", f"{mp3.name} rỗng hoặc không có.", FIX_FILE)
    try:
        return float(do_dai(mp3))
    except media.MediaError as exc:
        if str(exc).startswith("ffprobe lỗi"):
            raise media.MediaError("giong", f"{mp3.name} không đọc được (file hỏng?).", FIX_FILE) from exc
        raise


def _so_giong(path: Path) -> dict:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def lay_giong(so: int, loi: str, thu_muc: Path, giong: str, toc_do: str,
              tong_hop: Callable = tong_hop_edge, do_dai: Callable = media.probe_duration) -> GiongInfo:
    mp3 = thu_muc / f"canh-{so}.mp3"
    so_giong = thu_muc / f"canh-{so}.json"
    voice, rate = VOICES[giong], RATES[toc_do]
    ma_bam = bam(loi, voice, rate)
    if mp3.is_file():
        ghi = _so_giong(so_giong)
        if not ghi:
            return GiongInfo(mp3=mp3, giay=_giay(mp3, do_dai), moc_cau=[], uoc_luong=True, nguon="co-san")
        if ghi.get("bam") == ma_bam:
            moc = [float(m) for m in ghi.get("moc", [])]
            return GiongInfo(mp3=mp3, giay=_giay(mp3, do_dai), moc_cau=moc, uoc_luong=not moc, nguon="may")
    thu_muc.mkdir(parents=True, exist_ok=True)
    tam = mp3.with_name(mp3.name + ".tmp")
    try:
        moc = tong_hop(loi, voice, rate, tam)
        os.replace(tam, mp3)
    except BaseException:
        tam.unlink(missing_ok=True)
        raise
    so_giong.write_text(json.dumps({"bam": ma_bam, "moc": moc}, ensure_ascii=False), encoding="utf-8")
    return GiongInfo(mp3=mp3, giay=_giay(mp3, do_dai), moc_cau=list(moc), uoc_luong=not moc, nguon="may")
```

- [ ] **Step 4: Chạy test đến khi xanh**

Run: `venv\Scripts\python.exe -m unittest tools/vi/tests/test_video_ma_giong.py -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add tools/vi/video_ma_parts/giong.py tools/vi/tests/test_video_ma_giong.py
git commit -m "feat(vi): per-scene narration from supplied mp3 or edge-tts with a ledger"
```

---

### Task 5: Phần vẽ (JavaScript) và tám loại cảnh

**Files:**
- Create: `tools/vi/video_ma_parts/runtime/khung-video.js`, `tools/vi/video_ma_parts/runtime/viet-tay.css`
- Create: `tools/vi/video_ma_parts/runtime/canh/{tieu-de,khai-niem,cong-thuc,y-tung-y,quy-trinh,so-sanh,do-thi,thi-nghiem}.js`
- Test: `tools/vi/tests/js/test_canh.js`, `tools/vi/tests/test_video_ma_canh.py`

**Interfaces (JavaScript, toàn cục `THI_VIDEO` và `THI_CANH`):**
- `THI_VIDEO`: `DAN_DAU`, `kep(x,a,b)`, `tienDo(t, batDau, thoiLuong) -> 0..1`, `thoat(s)`, `demKyTu(chu)`, `catDanhDau(chu, n) -> html` (n ký tự đầu hiện, phần còn lại nằm trong `<span class="an">` để bố cục không xê dịch), `thoiGianViet(chu)`, `duongQua(diem, hat) -> d`, `hopQua(x,y,w,h,hat) -> d`, `vongTron(cx,cy,r) -> d`, `muiTen(x1,y1,x2,y2,hat) -> d`, `tao(du) -> {chu, net, tieuDe, gh}`, `khoiDong(du)` (gắn trang, đặt `window.datThoiDiem`, `THI_VIDEO.san = true`, `THI_VIDEO.kiemTran()`, `THI_VIDEO.thoiDiemCuoi()`).
- Một **mục** là `{id, kieu: 'chu'|'net', batDau, thoiLuong, ...}`; `chu` có `chu, x, y, rong, cao, co, can, mau, dong?`; `net` có `d, mau, day`.
- `THI_CANH[loai] = { muc(du) -> mục[], dung?(goc, du), capNhat?(goc, du, t) }`. `du` là dữ liệu từ `lich.du_lieu_canh`.
- `THI_CANH['thi-nghiem']` thêm `thamSoTai(du, t) -> {mã: giá trị}` và `giaTri(du, t) -> {tham, dai}`.

Quy ước bố cục (px, khung 1280×720): tiêu đề cảnh ở y=30 cao 116, gạch chân y=160; nội dung bắt đầu từ y=190. Bút: chữ hiện dần theo số ký tự, nét vẽ dần bằng `stroke-dashoffset` với `pathLength=1`, nét lệch nhẹ xác định theo hạt giống.

- [ ] **Step 1: Viết test Node thất bại**

Tạo `tools/vi/tests/js/test_canh.js`:

```javascript
'use strict';
var test = require('node:test');
var assert = require('node:assert');
var path = require('node:path');
var fs = require('node:fs');

var RT = path.join(__dirname, '..', '..', 'video_ma_parts', 'runtime');
var TN = path.join(__dirname, '..', '..', 'thi_nghiem_parts');
require(path.join(TN, 'runtime', 'khung.js'));
require(path.join(TN, 'mo_hinh', 'li-con-lac-don.js'));
require(path.join(RT, 'khung-video.js'));
var LOAI = ['tieu-de', 'khai-niem', 'cong-thuc', 'y-tung-y', 'quy-trinh', 'so-sanh', 'do-thi', 'thi-nghiem'];
LOAI.forEach(function (l) { require(path.join(RT, 'canh', l + '.js')); });
var V = globalThis.THI_VIDEO;
var C = globalThis.THI_CANH;
var KHAI_BAO = JSON.parse(fs.readFileSync(path.join(TN, 'mo_hinh', 'li-con-lac-don.json'), 'utf8'));

function du(loai, thoiLuong, moc, truong, them) {
  var d = { so: 1, loai: loai, thoiLuong: thoiLuong, danDau: 0.7, moc: moc, truong: truong };
  Object.keys(them || {}).forEach(function (k) { d[k] = them[k]; });
  return d;
}
function tatCa(gh) {
  var s = gh === 2.5;
  return {
    'tieu-de': du('tieu-de', gh, [], { chu: ['Chuẩn độ acid – base'], phu: ['Hoá học 11'] }),
    'khai-niem': du('khai-niem', gh, [], { 'thuat-ngu': ['Chuẩn độ'], 'dinh-nghia': ['Xác định nồng độ một dung dịch.'] }),
    'cong-thuc': du('cong-thuc', gh, s ? [0.7, 1.0] : [2.5, 4.5], { 'bieu-thuc': ['C~M~ = n/V'], 'giai-thich': ['n là số mol', 'V là thể tích'] }),
    'y-tung-y': du('y-tung-y', gh, s ? [0.7, 1.0, 1.3] : [0.7, 3, 5.5], { 'tieu-de': ['Ba bước'], y: ['a', 'b', 'c'] }),
    'quy-trinh': du('quy-trinh', gh, s ? [0.7, 1.0, 1.3] : [0.7, 3.5, 6], { 'tieu-de': ['Quy trình'], buoc: ['Đo', 'Nhỏ', 'Dừng'] }),
    'so-sanh': du('so-sanh', gh, s ? [0.7, 0.9, 1.1, 1.3] : [0.7, 2.5, 4.5, 6.5],
      { 'tieu-de': ['So sánh'], trai: ['Acid'], phai: ['Base'], 'y-trai': ['a', 'b'], 'y-phai': ['c', 'd'] }),
    'do-thi': du('do-thi', gh, s ? [0.7, 1.0, 1.3] : [0.7, 3, 5.5], { 'tieu-de': ['Đồ thị'], 'truc-ngang': ['t'], 'truc-doc': ['v'], diem: ['0, 0', '1, 2', '2, 4'] },
      { diem: [[0, 0], [1, 2], [2, 4]] }),
    'thi-nghiem': du('thi-nghiem', gh, [], { mau: ['li-con-lac-don'] },
      { khaiBao: KHAI_BAO, thamSo: { 'chieu-dai': [[0, 0.4], [6, 1.6]] }, do: ['chu-ki'] })
  };
}
function hienThi(html) {
  return html.replace(/<span class="an">.*?<\/span>/g, '').replace(/<[^>]+>/g, '');
}

test('catDanhDau: che chu, giu the dong mo va thoat ky tu dac biet', function () {
  assert.strictEqual(V.catDanhDau('**ab**c', 1), '<b>a<span class="an">b</span></b><span class="an">c</span>');
  assert.strictEqual(V.catDanhDau('H~2~O', 99), 'H<sub>2</sub>O');
  assert.strictEqual(V.catDanhDau('m/s^2^', 99), 'm/s<sup>2</sup>');
  assert.strictEqual(V.catDanhDau('a < b & "c"', 99), 'a &lt; b &amp; &quot;c&quot;');
  var doc = V.catDanhDau('</script><img src=x onerror=alert(1)>', 99);
  assert.ok(doc.indexOf('<img') === -1 && doc.indexOf('</script') === -1, doc);
  assert.strictEqual(hienThi(V.catDanhDau('**ab**c', 2)), 'ab');
  assert.strictEqual(hienThi(V.catDanhDau('Nhờ ướt', 3)), 'Nhờ');
  assert.strictEqual(V.catDanhDau('abc', 0), '<span class="an">abc</span>');
});

test('demKyTu bo dau danh dau; tienDo kep 0..1', function () {
  assert.strictEqual(V.demKyTu('H~2~SO~4~ **x**'), 'H2SO4 x'.length);
  assert.strictEqual(V.tienDo(0, 1, 2), 0);
  assert.strictEqual(V.tienDo(2, 1, 2), 0.5);
  assert.strictEqual(V.tienDo(9, 1, 2), 1);
  assert.strictEqual(V.tienDo(1, 1, 0), 1);
});

test('duong ve la xac dinh theo hat giong', function () {
  assert.strictEqual(V.duongQua([[0, 0], [100, 0]], 5), V.duongQua([[0, 0], [100, 0]], 5));
  assert.notStrictEqual(V.duongQua([[0, 0], [100, 0]], 5), V.duongQua([[0, 0], [100, 0]], 6));
});

test('moi loai canh: muc xac dinh, t=0 chua hien gi, cuoi canh hien du', function () {
  [12, 2.5].forEach(function (gh) {
    var bo = tatCa(gh);
    LOAI.forEach(function (l) {
      var a = C[l].muc(bo[l]);
      var b = C[l].muc(bo[l]);
      assert.deepStrictEqual(a, b, l);
      assert.ok(a.length > 0, l);
      a.forEach(function (m) {
        if (m.dong) { return; }
        assert.ok(m.batDau > 0, l + ':' + m.id + ' phai bat dau sau t=0');
        assert.strictEqual(V.tienDo(0, m.batDau, m.thoiLuong), 0, l + ':' + m.id);
        assert.ok(m.batDau + m.thoiLuong <= gh - 0.2 + 1e-9, l + ':' + m.id + ' phai xong truoc cuoi canh ' + gh);
        assert.strictEqual(V.tienDo(gh, m.batDau, m.thoiLuong), 1, l + ':' + m.id);
      });
      var ids = a.map(function (m) { return m.id; });
      assert.strictEqual(new Set(ids).size, ids.length, l + ' co id trung');
    });
  });
});

test('y k hien dung tai moc cau k', function () {
  var bo = tatCa(12);
  var ids = function (l, tienTo) {
    var muc = C[l].muc(bo[l]);
    return [0, 1, 2].map(function (k) { return muc.filter(function (m) { return m.id === tienTo + k; })[0]; });
  };
  ids('y-tung-y', 'y-').forEach(function (m, k) { assert.strictEqual(m.batDau, bo['y-tung-y'].moc[k]); });
  ids('quy-trinh', 'buoc-').forEach(function (m, k) { assert.strictEqual(m.batDau, bo['quy-trinh'].moc[k]); });
  ids('do-thi', 'diem-').forEach(function (m, k) { assert.strictEqual(m.batDau, bo['do-thi'].moc[k]); });
  var ss = C['so-sanh'].muc(bo['so-sanh']);
  var at = function (id) { return ss.filter(function (m) { return m.id === id; })[0].batDau; };
  assert.strictEqual(at('y-trai-0'), 0.7);
  assert.strictEqual(at('y-trai-1'), 2.5);
  assert.strictEqual(at('y-phai-0'), 4.5);
  assert.strictEqual(at('y-phai-1'), 6.5);
  var ct = C['cong-thuc'].muc(bo['cong-thuc']);
  var bt = ct.filter(function (m) { return m.id === 'bieu-thuc'; })[0];
  ct.filter(function (m) { return /^giai-thich-/.test(m.id); }).forEach(function (m, k) {
    assert.ok(m.batDau >= bo['cong-thuc'].moc[k]);
    assert.ok(m.batDau >= bt.batDau + bt.thoiLuong);
  });
});

test('do thi: diem nam trong vung ve, cung y thi vao giua', function () {
  var muc = C['do-thi'].muc(tatCa(12)['do-thi']);
  var vong = muc.filter(function (m) { return /^diem-/.test(m.id); });
  assert.strictEqual(vong.length, 3);
  vong.forEach(function (m) {
    var so = m.d.match(/-?\d+(\.\d+)?/g).map(Number);
    var cx = so[0] + 9;
    assert.ok(cx >= 190 - 1 && cx <= 1090 + 1, 'x ngoai vung: ' + cx);
  });
  var bang = tatCa(12)['do-thi'];
  bang.diem = [[0, 5], [1, 5]];
  var muc2 = C['do-thi'].muc(bang).filter(function (m) { return /^diem-/.test(m.id); });
  var y = muc2.map(function (m) { return Number(m.d.match(/-?\d+(\.\d+)?/g)[1]); });
  assert.strictEqual(y[0], y[1]);
  assert.ok(y[0] > 270 && y[0] < 560);
});

test('thi nghiem: noi suy tham so theo moc thoi gian', function () {
  var d = tatCa(12)['thi-nghiem'];
  var T = C['thi-nghiem'];
  assert.strictEqual(T.thamSoTai(d, 0)['chieu-dai'], 0.4);
  assert.ok(Math.abs(T.thamSoTai(d, 3)['chieu-dai'] - 1.0) < 1e-9);
  assert.strictEqual(T.thamSoTai(d, 6)['chieu-dai'], 1.6);
  assert.strictEqual(T.thamSoTai(d, 100)['chieu-dai'], 1.6);
  assert.strictEqual(T.thamSoTai(d, 3)['g'], 9.8);
  d.thamSo = { 'chieu-dai': [[2, 0.4]] };
  assert.strictEqual(T.thamSoTai(d, 1)['chieu-dai'], 1.0);
  assert.strictEqual(T.thamSoTai(d, 2)['chieu-dai'], 0.4);
});

test('thi nghiem: so do la so cua tinh() cua mo hinh, o moi thoi diem', function () {
  var d = tatCa(12)['thi-nghiem'];
  var T = C['thi-nghiem'];
  var M = globalThis.THI_NGHIEM_MO_HINH;
  [0, 1.5, 3, 5, 6, 11].forEach(function (t) {
    var g = T.giaTri(d, t);
    assert.strictEqual(g.dai['chu-ki'], M.tinh(g.tham)['chu-ki']);
  });
  assert.ok(Math.abs(T.giaTri(d, 3).dai['chu-ki'] - 2.007089923) < 5e-4);
});

test('thi nghiem: khong co dong chu nao bat dau tai t=0 ngoai dong so do', function () {
  var muc = C['thi-nghiem'].muc(tatCa(12)['thi-nghiem']);
  var dong = muc.filter(function (m) { return m.dong; });
  assert.ok(dong.length >= 2);
});
```

Tạo `tools/vi/tests/test_video_ma_canh.py`:

```python
"""Chạy bộ test Node của phần vẽ video giải thích (khung-video.js và tám loại cảnh)."""

import shutil
import subprocess
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
HAS_NODE = shutil.which("node") is not None
NEED_NODE = "máy không có Node nên không chạy được JavaScript"


class SceneRuntimeTest(unittest.TestCase):
    @unittest.skipUnless(HAS_NODE, NEED_NODE)
    def test_runtime_passes_its_node_tests(self):
        proc = subprocess.run([shutil.which("node"), "--test", str(TOOLS_VI / "tests" / "js" / "test_canh.js")],
                              capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180)
        self.assertEqual(proc.returncode, 0, proc.stdout[-3000:] + proc.stderr[-1500:])


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Chạy test để thấy lỗi**

Run: `node --test tools/vi/tests/js/test_canh.js`
Expected: FAIL — `Cannot find module .../runtime/khung-video.js`.

- [ ] **Step 3: Viết `khung-video.js`**

Tạo `tools/vi/video_ma_parts/runtime/khung-video.js`:

```javascript
(function (root) {
  'use strict';

  var DAN_DAU = 0.7;
  var TOC_DO_VIET = 20;
  var NS = 'http://www.w3.org/2000/svg';
  var DANH_DAU = /\*\*(.+?)\*\*|~([^~]+)~|\^([^\^]+)\^/g;

  function kep(x, a, b) { return x < a ? a : (x > b ? b : x); }
  function tienDo(t, batDau, thoiLuong) {
    if (thoiLuong <= 0) { return t >= batDau ? 1 : 0; }
    return kep((t - batDau) / thoiLuong, 0, 1);
  }
  function thoat(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function tach(chu) {
    var kq = [];
    var vt = 0;
    var m;
    chu = String(chu);
    DANH_DAU.lastIndex = 0;
    while ((m = DANH_DAU.exec(chu))) {
      if (m.index > vt) { kq.push({ the: '', chu: chu.slice(vt, m.index) }); }
      kq.push(m[1] !== undefined ? { the: 'b', chu: m[1] } : (m[2] !== undefined ? { the: 'sub', chu: m[2] } : { the: 'sup', chu: m[3] }));
      vt = m.index + m[0].length;
    }
    if (vt < chu.length) { kq.push({ the: '', chu: chu.slice(vt) }); }
    return kq;
  }
  function demKyTu(chu) { return tach(chu).reduce(function (n, o) { return n + o.chu.length; }, 0); }
  function catDanhDau(chu, n) {
    var con = n;
    var html = '';
    tach(chu).forEach(function (o) {
      var hien = con > 0 ? o.chu.slice(0, con) : '';
      var an = o.chu.slice(hien.length);
      con -= hien.length;
      var trong = thoat(hien) + (an ? '<span class="an">' + thoat(an) + '</span>' : '');
      html += o.the ? '<' + o.the + '>' + trong + '</' + o.the + '>' : trong;
    });
    return html;
  }
  function thoiGianViet(chu) { return Math.max(0.5, demKyTu(chu) / TOC_DO_VIET); }

  function rng(hat) {
    var a = hat | 0;
    return function () {
      a = a + 0x6D2B79F5 | 0;
      var t = Math.imul(a ^ a >>> 15, 1 | a);
      t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t;
      return ((t ^ t >>> 14) >>> 0) / 4294967296;
    };
  }
  function lam(x) { return Math.round(x * 10) / 10; }
  function duongQua(diem, hat) {
    var r = rng(hat);
    var d = '';
    for (var i = 0; i < diem.length; i++) {
      var p = diem[i];
      if (i === 0) { d += 'M' + lam(p[0]) + ' ' + lam(p[1]); continue; }
      var q = diem[i - 1];
      var dx = p[0] - q[0];
      var dy = p[1] - q[1];
      var dai = Math.sqrt(dx * dx + dy * dy) || 1;
      var nx = -dy / dai;
      var ny = dx / dai;
      var n = Math.max(1, Math.round(dai / 70));
      for (var j = 1; j <= n; j++) {
        var u = j / n;
        var lech = j === n ? 0 : (r() - 0.5) * 4;
        d += ' L' + lam(q[0] + dx * u + nx * lech) + ' ' + lam(q[1] + dy * u + ny * lech);
      }
    }
    return d;
  }
  function hopQua(x, y, w, h, hat) { return duongQua([[x, y], [x + w, y], [x + w, y + h], [x, y + h], [x, y - 2]], hat); }
  function vongTron(cx, cy, r) {
    return 'M' + lam(cx - r) + ' ' + lam(cy) + ' a' + r + ' ' + r + ' 0 1 0 ' + (2 * r) + ' 0 a' + r + ' ' + r + ' 0 1 0 ' + (-2 * r) + ' 0';
  }
  function muiTen(x1, y1, x2, y2, hat) {
    var goc = Math.atan2(y2 - y1, x2 - x1);
    var c = function (lech) { return [x2 - 16 * Math.cos(goc + lech), y2 - 16 * Math.sin(goc + lech)]; };
    var a = c(0.5);
    var b = c(-0.5);
    return duongQua([[x1, y1], [x2, y2]], hat) + ' M' + lam(a[0]) + ' ' + lam(a[1]) + ' L' + x2 + ' ' + y2 + ' L' + lam(b[0]) + ' ' + lam(b[1]);
  }

  function gan(dich, tuy) {
    Object.keys(tuy || {}).forEach(function (k) { dich[k] = tuy[k]; });
    return dich;
  }
  function tao(du) {
    var gh = du.thoiLuong;
    function chu(id, noiDung, x, y, rong, cao, co, batDau, tuy) {
      batDau = Math.min(batDau, gh - 0.6);
      var dai = Math.max(0.3, Math.min(thoiGianViet(noiDung), gh - 0.2 - batDau));
      return gan({ id: id, kieu: 'chu', chu: noiDung, x: x, y: y, rong: rong, cao: cao, co: co, batDau: batDau, thoiLuong: dai, can: 'trai', mau: '' }, tuy);
    }
    function net(id, d, batDau, dai, tuy) {
      batDau = Math.min(batDau, gh - 0.6);
      dai = Math.max(0.2, Math.min(dai, gh - 0.2 - batDau));
      return gan({ id: id, kieu: 'net', d: d, batDau: batDau, thoiLuong: dai, mau: '', day: 4 }, tuy);
    }
    function tieuDe(noiDung, batDau) {
      var c = chu('tieu-de', noiDung, 60, 30, 1160, 116, 40, batDau, { mau: 'nhan' });
      var w = Math.min(1160, Math.max(240, demKyTu(noiDung) * 40 * 0.6));
      return [c, net('gach', duongQua([[60, 160], [60 + w, 160]], 7), c.batDau + c.thoiLuong, 0.4, { mau: 'nhan' })];
    }
    return { chu: chu, net: net, tieuDe: tieuDe, gh: gh };
  }

  function khoiDong(du) {
    var goc = document.getElementById('khung');
    var loai = root.THI_CANH[du.loai];
    var muc = loai.muc(du);
    var svg = document.createElementNS(NS, 'svg');
    svg.setAttribute('viewBox', '0 0 1280 720');
    goc.appendChild(svg);
    var ds = muc.map(function (m) {
      var el;
      if (m.kieu === 'net') {
        el = document.createElementNS(NS, 'path');
        el.setAttribute('d', m.d);
        el.setAttribute('pathLength', '1');
        el.setAttribute('class', 'net ' + m.mau);
        el.style.strokeWidth = String(m.day);
        el.style.strokeDasharray = '1';
        svg.appendChild(el);
      } else {
        el = document.createElement('div');
        el.className = 'chu ' + m.can + ' ' + m.mau;
        el.setAttribute('data-id', m.id);
        el.style.left = m.x + 'px';
        el.style.top = m.y + 'px';
        el.style.width = m.rong + 'px';
        el.style.height = m.cao + 'px';
        el.style.fontSize = m.co + 'px';
        goc.appendChild(el);
      }
      return { m: m, el: el };
    });
    if (loai.dung) { loai.dung(goc, du); }
    function dat(t) {
      ds.forEach(function (o) {
        var p = tienDo(t, o.m.batDau, o.m.thoiLuong);
        if (o.m.kieu === 'net') {
          o.el.style.strokeDashoffset = String(1 - p);
          o.el.style.opacity = p > 0 ? '1' : '0';
        } else if (!o.m.dong) {
          o.el.innerHTML = catDanhDau(o.m.chu, Math.round(p * demKyTu(o.m.chu)));
        }
      });
      if (loai.capNhat) { loai.capNhat(goc, du, t); }
    }
    root.datThoiDiem = dat;
    root.THI_VIDEO.thoiDiemCuoi = function () {
      return muc.reduce(function (cao, m) { return Math.max(cao, m.batDau + m.thoiLuong); }, 0) + 0.3;
    };
    root.THI_VIDEO.kiemTran = function () {
      dat(1e6);
      var loi = [];
      var cacO = goc.querySelectorAll('.chu');
      for (var i = 0; i < cacO.length; i++) {
        var el = cacO[i];
        var r = el.getBoundingClientRect();
        if (el.scrollHeight > el.clientHeight + 1 || el.scrollWidth > el.clientWidth + 1 || r.right > 1281 || r.bottom > 721) {
          loi.push(el.getAttribute('data-id'));
        }
      }
      return loi;
    };
    dat(0);
    root.THI_VIDEO.san = true;
  }

  root.THI_CANH = root.THI_CANH || {};
  root.THI_VIDEO = {
    DAN_DAU: DAN_DAU, kep: kep, tienDo: tienDo, thoat: thoat, demKyTu: demKyTu, catDanhDau: catDanhDau,
    thoiGianViet: thoiGianViet, duongQua: duongQua, hopQua: hopQua, vongTron: vongTron, muiTen: muiTen,
    tao: tao, khoiDong: khoiDong, san: false
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
```

- [ ] **Step 4: Viết `viet-tay.css`**

```css
html, body { margin: 0; padding: 0; background: #fbfaf5; }
#khung { position: relative; width: 1280px; height: 720px; overflow: hidden; background: #fbfaf5;
  font-family: 'Segoe Print', 'Ink Free', 'Comic Sans MS', cursive; color: #1f2937; }
#khung svg { position: absolute; left: 0; top: 0; width: 1280px; height: 720px; }
.chu { position: absolute; overflow: hidden; line-height: 1.3; white-space: normal; overflow-wrap: normal; text-align: left; }
.chu.giua { text-align: center; }
.chu.phai { text-align: right; }
.chu.nhan { color: #1d4ed8; }
.chu.do { color: #b91c1c; }
.chu .an { visibility: hidden; }
path.net { fill: none; stroke: #1f2937; stroke-linecap: round; stroke-linejoin: round; }
path.net.nhan { stroke: #1d4ed8; }
path.net.do { stroke: #b91c1c; }
```

- [ ] **Step 5: Viết sáu loại cảnh chữ**

Tạo các file trong `tools/vi/video_ma_parts/runtime/canh/`. Mỗi file có dạng `(function (root) { ... })(typeof globalThis !== 'undefined' ? globalThis : this);` với `var V = root.THI_VIDEO;` ở đầu và `root.THI_CANH['<loại>'] = {...}`.

`tieu-de.js`:

```javascript
(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  root.THI_CANH['tieu-de'] = {
    muc: function (du) {
      var B = V.tao(du);
      var t = du.truong;
      var kq = [];
      var c = B.chu('chu', t.chu[0], 80, 150, 1120, 340, 60, 0.3, { can: 'giua', mau: 'nhan' });
      kq.push(c);
      kq.push(B.net('gach', V.duongQua([[340, 510], [940, 510]], 3), c.batDau + c.thoiLuong, 0.4, { mau: 'nhan' }));
      if (t.phu) { kq.push(B.chu('phu', t.phu[0], 80, 540, 1120, 90, 34, c.batDau + c.thoiLuong + 0.4, { can: 'giua' })); }
      return kq;
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
```

`khai-niem.js`:

```javascript
(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  root.THI_CANH['khai-niem'] = {
    muc: function (du) {
      var B = V.tao(du);
      var t = du.truong;
      var kq = [B.net('khung', V.hopQua(60, 190, 1160, 450, 11), 0.1, 0.9, {})];
      var tn = B.chu('thuat-ngu', t['thuat-ngu'][0], 100, 215, 1080, 110, 40, 1.0, { mau: 'nhan' });
      kq.push(tn);
      kq.push(B.net('gach', V.duongQua([[100, 335], [700, 335]], 4), tn.batDau + tn.thoiLuong, 0.4, { mau: 'nhan' }));
      kq.push(B.chu('dinh-nghia', t['dinh-nghia'][0], 100, 360, 1080, 260, 32, tn.batDau + tn.thoiLuong + 0.5, {}));
      return kq;
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
```

`cong-thuc.js`:

```javascript
(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  root.THI_CANH['cong-thuc'] = {
    muc: function (du) {
      var B = V.tao(du);
      var t = du.truong;
      var kq = [];
      kq.push(B.net('khung', V.hopQua(100, 190, 1080, 150, 21), 0.1, 0.7, {}));
      var bt = B.chu('bieu-thuc', t['bieu-thuc'][0], 120, 215, 1040, 110, 40, 0.9, { can: 'giua', mau: 'nhan' });
      kq.push(bt);
      (t['giai-thich'] || []).forEach(function (g, k) {
        var bd = Math.max(du.moc[k], bt.batDau + bt.thoiLuong + 0.3);
        kq.push(B.chu('giai-thich-' + k, g, 110, 370 + k * 84, 1060, 76, 28, bd, {}));
      });
      return kq;
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
```

Cảnh công thức không có tiêu đề riêng; khung bắt đầu ở y=190.

`y-tung-y.js`:

```javascript
(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  root.THI_CANH['y-tung-y'] = {
    muc: function (du) {
      var B = V.tao(du);
      var t = du.truong;
      var kq = B.tieuDe(t['tieu-de'][0], 0.2);
      t.y.forEach(function (y, k) {
        var top = 190 + k * 82;
        kq.push(B.net('cham-' + k, V.vongTron(96, top + 22, 9), du.moc[k], 0.3, { mau: 'nhan' }));
        kq.push(B.chu('y-' + k, y, 124, top, 1100, 78, 30, du.moc[k], {}));
      });
      return kq;
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
```

`quy-trinh.js`:

```javascript
(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  root.THI_CANH['quy-trinh'] = {
    muc: function (du) {
      var B = V.tao(du);
      var t = du.truong;
      var kq = B.tieuDe(t['tieu-de'][0], 0.2);
      var n = t.buoc.length;
      var w = (1160 - (n - 1) * 70) / n;
      t.buoc.forEach(function (b, k) {
        var x = 60 + k * (w + 70);
        kq.push(B.net('hop-' + k, V.hopQua(x, 230, w, 300, 20 + k), Math.max(0.1, du.moc[k] - 0.4), 0.5, {}));
        kq.push(B.chu('buoc-' + k, b, x + 14, 250, w - 28, 260, 24, du.moc[k], {}));
        if (k > 0) {
          kq.push(B.net('mui-' + k, V.muiTen(x - 64, 380, x - 6, 380, 30 + k), Math.max(0.1, du.moc[k] - 0.6), 0.3, { mau: 'nhan' }));
        }
      });
      return kq;
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
```

`so-sanh.js`:

```javascript
(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  root.THI_CANH['so-sanh'] = {
    muc: function (du) {
      var B = V.tao(du);
      var t = du.truong;
      var nt = t['y-trai'].length;
      var kq = B.tieuDe(t['tieu-de'][0], 0.2);
      kq.push(B.net('giua', V.duongQua([[640, 190], [640, 660]], 9), 0.3, 0.6, {}));
      kq.push(B.chu('trai', t.trai[0], 60, 190, 540, 70, 34, 0.3, { mau: 'nhan' }));
      kq.push(B.chu('phai', t.phai[0], 680, 190, 540, 70, 34, Math.max(0.3, du.moc[nt] - 0.9), { mau: 'nhan' }));
      t['y-trai'].forEach(function (y, k) { kq.push(B.chu('y-trai-' + k, y, 80, 285 + k * 92, 520, 88, 28, du.moc[k], {})); });
      t['y-phai'].forEach(function (y, k) { kq.push(B.chu('y-phai-' + k, y, 700, 285 + k * 92, 520, 88, 28, du.moc[nt + k], {})); });
      return kq;
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
```

- [ ] **Step 6: Viết `do-thi.js`**

```javascript
(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  var X0 = 190, X1 = 1090, Y0 = 560, Y1 = 270;

  function soVN(v) { return String(Math.round(v * 1000) / 1000).replace('.', ','); }
  function tyLe(v, thap, cao, a, b) { return cao === thap ? (a + b) / 2 : a + (v - thap) * (b - a) / (cao - thap); }

  root.THI_CANH['do-thi'] = {
    muc: function (du) {
      var B = V.tao(du);
      var t = du.truong;
      var xs = du.diem.map(function (p) { return p[0]; });
      var ys = du.diem.map(function (p) { return p[1]; });
      var xMin = Math.min.apply(null, xs), xMax = Math.max.apply(null, xs);
      var yMin = Math.min.apply(null, ys), yMax = Math.max.apply(null, ys);
      var kq = B.tieuDe(t['tieu-de'][0], 0.2);
      kq.push(B.net('truc-x', V.duongQua([[140, 600], [1140, 600]], 1), 0.3, 0.6, {}));
      kq.push(B.net('truc-y', V.duongQua([[140, 600], [140, 230]], 2), 0.4, 0.6, {}));
      kq.push(B.chu('truc-ngang', t['truc-ngang'][0], 700, 650, 440, 40, 24, 1.0, { can: 'phai', mau: 'nhan' }));
      kq.push(B.chu('truc-doc', t['truc-doc'][0], 160, 186, 700, 40, 24, 1.0, { mau: 'nhan' }));
      kq.push(B.chu('x-min', soVN(xMin), X0 - 60, 606, 120, 30, 20, 1.0, { can: 'giua' }));
      kq.push(B.chu('x-max', soVN(xMax), X1 - 60, 606, 120, 30, 20, 1.0, { can: 'giua' }));
      kq.push(B.chu('y-min', soVN(yMin), 20, Y0 - 15, 110, 30, 20, 1.0, { can: 'phai' }));
      kq.push(B.chu('y-max', soVN(yMax), 20, Y1 - 15, 110, 30, 20, 1.0, { can: 'phai' }));
      var truoc = null;
      du.diem.forEach(function (p, k) {
        var px = tyLe(p[0], xMin, xMax, X0, X1);
        var py = tyLe(p[1], yMin, yMax, Y0, Y1);
        if (truoc) { kq.push(B.net('doan-' + k, V.duongQua([truoc, [px, py]], 40 + k), du.moc[k], 0.4, { mau: 'nhan' })); }
        kq.push(B.net('diem-' + k, V.vongTron(px, py, 9), du.moc[k], 0.3, { mau: 'do' }));
        truoc = [px, py];
      });
      return kq;
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
```

Trong test `do-thi`, `m.d` của điểm là `M{cx-9} {cy} a9 9 ...`: `so[0]` là `cx-9` và `so[1]` là `cy`.

- [ ] **Step 7: Viết `thi-nghiem.js`**

```javascript
(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  var RONG = 792, CAO = 462;

  function thamSoTai(du, t) {
    var K = root.THI_NGHIEM_KHUNG;
    var p = K.thamSoMacDinh(du.khaiBao);
    Object.keys(du.thamSo).forEach(function (ma) {
      var ds = du.thamSo[ma];
      var cuoi = ds[ds.length - 1];
      if (t < ds[0][0]) { return; }
      if (t >= cuoi[0]) { p[ma] = cuoi[1]; return; }
      for (var i = 1; i < ds.length; i++) {
        if (t < ds[i][0]) {
          var a = ds[i - 1], b = ds[i];
          p[ma] = a[1] + (b[1] - a[1]) * (t - a[0]) / (b[0] - a[0]);
          return;
        }
      }
    });
    return p;
  }

  function giaTri(du, t) {
    var p = thamSoTai(du, t);
    return { tham: p, dai: root.THI_NGHIEM_MO_HINH.tinh(p) };
  }

  function soThapPhan(buoc) {
    var s = String(buoc);
    var i = s.indexOf('.');
    return i < 0 ? 0 : Math.min(3, s.length - i - 1);
  }
  function tim(ds, ma) { return ds.filter(function (x) { return x.ma === ma; })[0]; }

  root.THI_CANH['thi-nghiem'] = {
    thamSoTai: thamSoTai,
    giaTri: giaTri,
    muc: function (du) {
      var B = V.tao(du);
      var kb = du.khaiBao;
      var kq = B.tieuDe(kb.ten, 0.2);
      kq.push(B.net('khung', V.hopQua(40, 190, 800, 470, 5), 0.1, 0.7, {}));
      kq.push(B.chu('nhan-tham-so', 'Thông số', 880, 190, 360, 40, 26, 0.4, { mau: 'nhan' }));
      Object.keys(du.thamSo).slice(0, 3).forEach(function (ma, k) {
        kq.push(B.chu('ts-' + k, '', 880, 236 + k * 60, 360, 56, 20, 0.4, { dong: true }));
      });
      kq.push(B.chu('nhan-do', 'Số đo', 880, 430, 360, 40, 26, 0.4, { mau: 'nhan' }));
      du.do.slice(0, 3).forEach(function (ma, k) {
        kq.push(B.chu('do-' + k, '', 880, 476 + k * 70, 360, 66, 22, 0.4, { dong: true, mau: 'do' }));
      });
      return kq;
    },
    dung: function (goc) {
      var c = document.createElement('canvas');
      c.id = 'ban-ve';
      c.width = RONG;
      c.height = CAO;
      c.style.position = 'absolute';
      c.style.left = '44px';
      c.style.top = '194px';
      c.style.background = '#ffffff';
      goc.appendChild(c);
    },
    capNhat: function (goc, du, t) {
      var K = root.THI_NGHIEM_KHUNG;
      var M = root.THI_NGHIEM_MO_HINH;
      var g = giaTri(du, t);
      var ctx = goc.querySelector('#ban-ve').getContext('2d');
      ctx.clearRect(0, 0, RONG, CAO);
      var tm = Math.max(0, t - du.danDau);
      if (typeof M.thoiLuong === 'function') { tm = Math.min(tm, M.thoiLuong(g.tham, g.dai)); }
      M.ve(ctx, g.tham, tm, { rong: RONG, cao: CAO }, g.dai);
      Object.keys(du.thamSo).slice(0, 3).forEach(function (ma, k) {
        var ts = tim(du.khaiBao.thamSo, ma);
        goc.querySelector('[data-id="ts-' + k + '"]').innerHTML =
          K.danhDau(ts.ten) + ' = ' + K.dinhDang(g.tham[ma], soThapPhan(ts.buoc)) + ' ' + K.danhDau(ts.donVi || '');
      });
      du.do.slice(0, 3).forEach(function (ma, k) {
        var dl = tim(du.khaiBao.daiLuongDo, ma);
        goc.querySelector('[data-id="do-' + k + '"]').innerHTML =
          K.danhDau(dl.ten) + ' = ' + K.dinhDang(g.dai[ma], dl.chuSo) + ' ' + K.danhDau(dl.donVi || '');
      });
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
```

- [ ] **Step 8: Chạy test đến khi xanh**

Run: `node --test tools/vi/tests/js/test_canh.js` rồi `venv\Scripts\python.exe -m unittest tools/vi/tests/test_video_ma_canh.py`
Expected: PASS. Nếu một khẳng định "xong trước cuối cảnh" trượt ở cảnh 2,5 giây, chỉnh **cách kẹp** trong `tao()` (không nới test).

- [ ] **Step 9: Commit**

```bash
git add tools/vi/video_ma_parts/runtime tools/vi/tests/js/test_canh.js tools/vi/tests/test_video_ma_canh.py
git commit -m "feat(vi): hand-drawn scene runtime and the eight scene types"
```

---

### Task 6: Dựng trang và chụp khung bằng Chromium

**Files:**
- Create: `tools/vi/video_ma_parts/trang.py`, `tools/vi/video_ma_parts/chup.py`
- Test: `tools/vi/tests/test_video_ma_chup.py`

**Interfaces:**
- Consumes: `lich.du_lieu_canh` (dict `du`); `thu_vien.Model` (`.js`, `.khai_bao`); `media.MediaError`.
- Produces (`trang.py`): `dung_trang(du: dict, model=None) -> str` (HTML đầy đủ, mọi thứ nhúng trong trang, không địa chỉ web); `json_nhung(data) -> str` (JSON đã thoát `<`).
- Produces (`chup.py`): `FIX_CHROMIUM`; `trinh_duyet()` (context manager, yield browser; lỗi `MediaError` step `chromium`); `trang_moi(browser) -> page` (viewport 1280×720, scale 1); `mo_trang(page, html)` (nạp trang, chờ font và `THI_VIDEO.san`); `kiem_tran(page, html) -> list[str]` (id các ô chữ tràn khung, rỗng nếu ổn); `chup_canh(page, html, so_khung, fps, thu_muc, so_dau, ghi_log=None) -> int` (ghi `f%06d.png` từ số `so_dau`, trả về số tiếp theo); `chup_cuoi(page, html, duong_dan)` (chụp trạng thái cuối cảnh).

- [ ] **Step 1: Viết test thất bại**

Tạo `tools/vi/tests/test_video_ma_chup.py`:

```python
"""Test dựng trang và chụp khung. Phần Chromium tự bỏ qua nếu máy thiếu Chromium hoặc playwright."""

import importlib.util
import os
import re
import sys
import tempfile
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))

from video_ma_parts import chup, lich, parse, trang  # noqa: E402

META = "tieu-de: T\nmon: Toán\nlop: 8\n"


def co_chromium() -> bool:
    root = os.environ.get("LOCALAPPDATA")
    if not root or importlib.util.find_spec("playwright") is None:
        return False
    base = Path(root) / "ms-playwright"
    return base.is_dir() and (any(base.glob("chromium-*")) or any(base.glob("chromium_headless_shell-*")))


NEED_CHROMIUM = "máy không có Chromium hoặc playwright"


def du_cua(noi_dung: str, giay: float = 6.0, loi: str = "Xin chào các em. Hôm nay học bài mới. Cảm ơn các em."):
    text = f"---\n{META}---\n\n## Cảnh 1\n{noi_dung}loi: {loi}\n"
    canh = parse.parse(text).canh[0]
    giong = lich.GiongInfo(mp3=None, giay=giay, moc_cau=[0.0, 2.0, 4.0], uoc_luong=False, nguon="may")
    plan, _ = lich.dung_lich([canh], [giong])
    return lich.du_lieu_canh(canh, plan[0]), plan[0]


class PageBuildTest(unittest.TestCase):
    def test_page_is_self_contained(self):
        du, _ = du_cua("loai: tieu-de\nchu: Xin chào\n")
        html = trang.dung_trang(du)
        self.assertNotIn("http://", html.replace("http://www.w3.org/2000/svg", ""))
        self.assertNotIn("https://", html)
        self.assertIn("Segoe Print", html)
        self.assertIn("THI_CANH['tieu-de']", html)
        self.assertNotIn("THI_CANH['y-tung-y']", html)
        self.assertIn("khoiDong", html)

    def test_trang_khong_chua_ma_chay_duoc(self):
        du, _ = du_cua('loai: tieu-de\nchu: a < b </script><img src=x onerror=alert(1)> & "c"\n')
        html = trang.dung_trang(du)
        self.assertEqual(html.count("</script>"), html.count("<script>"))
        self.assertNotIn("<img src=x", html)
        self.assertIn("\\u003c/script>", html)

    def test_experiment_page_embeds_khung_and_model(self):
        from thi_nghiem_parts import thu_vien

        model = thu_vien.load("li-con-lac-don", Path("."))
        text = f"---\n{META}---\n\n## Cảnh 1\nloai: thi-nghiem\nmau: li-con-lac-don\nloi: Ok.\n"
        canh = parse.parse(text).canh[0]
        giong = lich.GiongInfo(mp3=None, giay=4.0, moc_cau=[0.0], uoc_luong=False, nguon="may")
        plan, _ = lich.dung_lich([canh], [giong])
        html = trang.dung_trang(lich.du_lieu_canh(canh, plan[0], model), model)
        self.assertIn("THI_NGHIEM_KHUNG", html)
        self.assertIn("THI_NGHIEM_MO_HINH", html)
        self.assertLess(html.index("THI_NGHIEM_KHUNG = "), html.index("THI_NGHIEM_MO_HINH = "))


@unittest.skipUnless(co_chromium(), NEED_CHROMIUM)
class ChromiumTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.cm = chup.trinh_duyet()
        cls.browser = cls.cm.__enter__()
        cls.page = chup.trang_moi(cls.browser)

    @classmethod
    def tearDownClass(cls):
        cls.cm.__exit__(None, None, None)

    def test_normal_scene_does_not_overflow_and_ink_grows(self):
        du, cl = du_cua("loai: y-tung-y\ntieu-de: Ba bước\ny: Một\ny: Hai\ny: Ba\n")
        html = trang.dung_trang(du)
        self.assertEqual(chup.kiem_tran(self.page, html), [])
        with tempfile.TemporaryDirectory() as tmp:
            n = chup.chup_canh(self.page, html, 6, lich.FPS, Path(tmp), 0)
            self.assertEqual(n, 6)
            frames = sorted(Path(tmp).glob("f*.png"))
            self.assertEqual([f.name for f in frames], [f"f{i:06d}.png" for i in range(6)])
            first = frames[0].read_bytes()
            self.assertEqual(first[:8], b"\x89PNG\r\n\x1a\n")
            final = Path(tmp) / "cuoi.png"
            chup.chup_cuoi(self.page, html, final)
            self.assertGreater(final.stat().st_size, frames[0].stat().st_size)

    def test_overflowing_text_is_reported_by_item_id(self):
        du, _ = du_cua("loai: tieu-de\nchu: " + "A" * 80 + "\n")
        self.assertIn("chu", chup.kiem_tran(self.page, trang.dung_trang(du)))

    def test_hostile_text_renders_as_text(self):
        du, _ = du_cua('loai: tieu-de\nchu: a < b </script><img src=x onerror=alert(1)> & "c"\n')
        html = trang.dung_trang(du)
        chup.mo_trang(self.page, html)
        self.page.evaluate("window.datThoiDiem(1e6)")
        self.assertEqual(self.page.evaluate("document.querySelectorAll('#khung img').length"), 0)
        texto = self.page.evaluate("document.querySelector('[data-id=chu]').textContent")
        self.assertIn("a < b </script><img src=x onerror=alert(1)> & \"c\"", re.sub(r"\s+", " ", texto))

    def test_frames_are_repeatable(self):
        du, _ = du_cua("loai: khai-niem\nthuat-ngu: Chu kì\ndinh-nghia: Thời gian một dao động.\n")
        html = trang.dung_trang(du)
        with tempfile.TemporaryDirectory() as tmp:
            chup.chup_canh(self.page, html, 4, lich.FPS, Path(tmp), 0)
            chup.chup_canh(self.page, html, 4, lich.FPS, Path(tmp), 10)
            for i in range(4):
                self.assertEqual((Path(tmp) / f"f{i:06d}.png").read_bytes(), (Path(tmp) / f"f{10 + i:06d}.png").read_bytes())

    def test_experiment_scene_draws_the_model(self):
        from thi_nghiem_parts import thu_vien

        model = thu_vien.load("li-con-lac-don", Path("."))
        text = (f"---\n{META}---\n\n## Cảnh 1\nloai: thi-nghiem\nmau: li-con-lac-don\ntham-so: 0 chieu-dai 0.4\n"
                "tham-so: 5 chieu-dai 1.6\ndo: chu-ki\nloi: Ok.\n")
        canh = parse.parse(text).canh[0]
        giong = lich.GiongInfo(mp3=None, giay=6.0, moc_cau=[0.0], uoc_luong=False, nguon="may")
        plan, _ = lich.dung_lich([canh], [giong])
        html = trang.dung_trang(lich.du_lieu_canh(canh, plan[0], model), model)
        self.assertEqual(chup.kiem_tran(self.page, html), [])
        self.page.evaluate("window.datThoiDiem(3)")
        dong = self.page.evaluate("document.querySelector('[data-id=do-0]').textContent")
        self.assertRegex(dong, r"Chu kì T = \d,\d{3} s")
        self.page.evaluate("window.datThoiDiem(1e6)")
        cuoi = self.page.evaluate("document.querySelector('[data-id=do-0]').textContent")
        self.assertIn("2,539", cuoi)


class ChromiumMissingTest(unittest.TestCase):
    def test_missing_playwright_is_a_chromium_error(self):
        import builtins
        real_import = builtins.__import__

        def no_pw(name, *args, **kwargs):
            if name.startswith("playwright"):
                raise ImportError("no playwright")
            return real_import(name, *args, **kwargs)

        builtins.__import__ = no_pw
        try:
            with self.assertRaises(chup.MediaError) as caught:
                with chup.trinh_duyet():
                    pass
        finally:
            builtins.__import__ = real_import
        self.assertEqual(caught.exception.step, "chromium")
        self.assertIn("pptmaster.ps1", caught.exception.fix)


if __name__ == "__main__":
    unittest.main()
```

Ghi chú: chu kì với `chieu-dai 1.6` và `g 9.8` là `2π·√(1.6/9.8) = 2.5386…`, hiển thị `2,539`.

- [ ] **Step 2: Chạy test để thấy lỗi**

Run: `C:/Users/ADMIN/vmt/v/Scripts/python.exe -m unittest tools/vi/tests/test_video_ma_chup.py`
Expected: FAIL — `ImportError: cannot import name 'chup'`.

- [ ] **Step 3: Viết `trang.py`**

```python
"""Ghép trang HTML cho một cảnh: phong cách + khung + file cảnh + dữ liệu. Mọi thứ nhúng trong trang."""

from __future__ import annotations

import json
from pathlib import Path

RUNTIME = Path(__file__).resolve().parent / "runtime"
NGHIEM = Path(__file__).resolve().parents[1] / "thi_nghiem_parts" / "runtime"


def json_nhung(data) -> str:
    return json.dumps(data, ensure_ascii=False).replace("<", "\\u003c")


def _doc(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig")


def dung_trang(du: dict, model=None) -> str:
    scripts = []
    if du["loai"] == "thi-nghiem":
        scripts.append(_doc(NGHIEM / "khung.js"))
        scripts.append(model.js)
    scripts.append(_doc(RUNTIME / "khung-video.js"))
    scripts.append(_doc(RUNTIME / "canh" / f"{du['loai']}.js"))
    scripts.append(f"window.DU_CANH = {json_nhung(du)};\nTHI_VIDEO.khoiDong(window.DU_CANH);")
    body = "\n".join(f"<script>\n{s}\n</script>" for s in scripts)
    return ("<!doctype html>\n<html lang=\"vi\"><head><meta charset=\"utf-8\">"
            f"<style>\n{_doc(RUNTIME / 'viet-tay.css')}\n</style></head>\n"
            f"<body><div id=\"khung\"></div>\n{body}\n</body></html>\n")
```

Lưu ý: `model.js` không được chứa `</script` (đã bị `thu_vien.check_js` chặn). Nhưng `khung-video.js` và các file cảnh là mã của repo, không chứa chuỗi đó.

- [ ] **Step 4: Viết `chup.py`**

```python
"""Chụp khung bằng Chromium: mở trang một cảnh, gọi datThoiDiem(t), chụp PNG. `playwright` chỉ import khi dùng."""

from __future__ import annotations

import contextlib
import os
from pathlib import Path

from video_parts.media import MediaError

FIX_CHROMIUM = (
    "Cài Chromium bằng: powershell -NoProfile -ExecutionPolicy Bypass -File tools\\vi\\pptmaster.ps1 "
    "-Action tool -Name chromium"
)
VIEWPORT = {"width": 1280, "height": 720}


def _mo(p):
    try:
        return p.chromium.launch(args=["--no-sandbox"])
    except Exception as first:
        base = Path(os.environ.get("LOCALAPPDATA", "")) / "ms-playwright"
        for pattern in ("chromium_headless_shell-*/*/chrome-headless-shell.exe", "chromium-*/*/chrome.exe"):
            for exe in sorted(base.glob(pattern), reverse=True):
                try:
                    return p.chromium.launch(executable_path=str(exe), args=["--no-sandbox"])
                except Exception:
                    continue
        raise MediaError("chromium", f"Không mở được Chromium: {first}", FIX_CHROMIUM) from first


@contextlib.contextmanager
def trinh_duyet():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise MediaError("chromium", "Chưa cài playwright (Chromium).", FIX_CHROMIUM) from exc
    with sync_playwright() as p:
        browser = _mo(p)
        try:
            yield browser
        finally:
            browser.close()


def trang_moi(browser):
    return browser.new_page(viewport=VIEWPORT, device_scale_factor=1)


def mo_trang(page, html: str) -> None:
    page.set_content(html)
    page.wait_for_function("window.THI_VIDEO && window.THI_VIDEO.san === true", timeout=30000)
    page.evaluate("() => document.fonts.ready.then(() => true)")


def kiem_tran(page, html: str) -> list:
    mo_trang(page, html)
    return list(page.evaluate("() => window.THI_VIDEO.kiemTran()"))


def chup_canh(page, html: str, so_khung: int, fps: int, thu_muc: Path, so_dau: int, ghi_log=None) -> int:
    mo_trang(page, html)
    thu_muc.mkdir(parents=True, exist_ok=True)
    for i in range(so_khung):
        page.evaluate("(t) => window.datThoiDiem(t)", i / fps)
        page.screenshot(path=str(thu_muc / f"f{so_dau + i:06d}.png"), type="png")
        if ghi_log is not None and (i + 1) % 60 == 0:
            ghi_log(f"  đã chụp {i + 1}/{so_khung} khung của cảnh này")
    return so_dau + so_khung


def chup_cuoi(page, html: str, duong_dan: Path) -> None:
    mo_trang(page, html)
    page.evaluate("() => window.datThoiDiem(window.THI_VIDEO.thoiDiemCuoi())")
    duong_dan.parent.mkdir(parents=True, exist_ok=True)
    page.screenshot(path=str(duong_dan), type="png")
```

- [ ] **Step 4b: Chạy test đến khi xanh**

Run: `C:/Users/ADMIN/vmt/v/Scripts/python.exe -m unittest tools/vi/tests/test_video_ma_chup.py -v` (phần Chromium chạy) và `venv\Scripts\python.exe -m unittest tools/vi/tests/test_video_ma_chup.py` (phần Chromium bỏ qua, phần trang vẫn chạy).
Expected: PASS ở cả hai. Nếu `test_overflowing_text_is_reported_by_item_id` không báo tràn vì Segoe Print rộng hơn dự tính, tăng số `A` — bài kiểm là "một từ dài không có dấu cách phải bị bắt".

- [ ] **Step 5: Commit**

```bash
git add tools/vi/video_ma_parts/trang.py tools/vi/video_ma_parts/chup.py tools/vi/tests/test_video_ma_chup.py
git commit -m "feat(vi): build scene pages and capture frames with Chromium"
```

---

### Task 7: Ghép hình, tiếng và phụ đề bằng FFmpeg

**Files:**
- Create: `tools/vi/video_ma_parts/ghep.py`
- Test: `tools/vi/tests/test_video_ma_ghep.py`

**Interfaces:**
- Consumes: `lich.CanhLich`, `lich.GiongInfo`, `lich.DAN_DAU`, `lich.FPS`; `video_parts.srt.Cue`, `srt.render_srt`; `video_parts.media` (`MediaError`, `build_audio_concat_text`, `FIX_FFMPEG`).
- Produces: `cues_phu_de(cac_lich) -> list[Cue]`; `lenh_am_canh(mp3, wav, thoi_luong) -> list[str]`; `lenh_video(anh_dir, danh_sach_am, out_mp4, fps, phu_de_tuong_doi=None) -> list[str]`; `ghep_video(thu_muc, cac_lich, cac_giong, phu_de, fps=FPS, run=subprocess.run) -> list[str]` (tên các file đầu ra tương đối: `video.mp4` và `phu-de.srt` khi `phu_de == "file"`). Quy ước thư mục tạm: `thu_muc/.khung/anh/f%06d.png` do `chup` ghi; `ghep_video` tạo `thu_muc/.khung/am-N.wav`, `am.txt`, `phu-de.srt` (khi `hinh`), `video.mp4` rồi chuyển `video.mp4` ra `thu_muc/`. `run` là hàm giả được trong test; FFmpeg chạy với `cwd=thu_muc`.

- [ ] **Step 1: Viết test thất bại**

Tạo `tools/vi/tests/test_video_ma_ghep.py`:

```python
"""Test ghép video: lệnh FFmpeg, phụ đề, danh sách nối tiếng. Không chạy FFmpeg thật."""

import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))

from video_ma_parts import ghep, lich  # noqa: E402
from video_parts import media, srt  # noqa: E402


def canh_lich(so, bat_dau, thoi_luong, giay, cau, moc_giong):
    return lich.CanhLich(so=so, bat_dau=bat_dau, thoi_luong=thoi_luong, so_khung=round(thoi_luong * 15), giay_giong=giay,
                         cau=cau, moc_cau=[lich.DAN_DAU + m for m in moc_giong], moc_cau_giong=moc_giong, uoc_luong=False)


PLAN = [
    canh_lich(1, 0.0, 6.0, 4.0, ["Xin chào **các** em.", "Học H~2~SO~4~ nhé."], [0.0, 2.0]),
    canh_lich(2, 6.0, 3.0, 1.5, ["Tạm biệt."], [0.0]),
]


class SubtitleTest(unittest.TestCase):
    def test_cues_follow_sentence_marks_with_scene_offsets_and_strip_markup(self):
        cues = ghep.cues_phu_de(PLAN)
        self.assertEqual([c.index for c in cues], [1, 2, 3])
        self.assertEqual(cues[0].text, "Xin chào các em.")
        self.assertEqual(cues[1].text, "Học H2SO4 nhé.")
        self.assertAlmostEqual(cues[0].start, 0.7)
        self.assertAlmostEqual(cues[0].end, 2.7)
        self.assertAlmostEqual(cues[1].start, 2.7)
        self.assertAlmostEqual(cues[1].end, 0.7 + 4.0)
        self.assertAlmostEqual(cues[2].start, 6.0 + 0.7)
        self.assertAlmostEqual(cues[2].end, 6.0 + 0.7 + 1.5)

    def test_cues_render_to_valid_srt(self):
        text = srt.render_srt(ghep.cues_phu_de(PLAN))
        self.assertEqual(len(srt.parse_srt(text)), 3)


class CommandTest(unittest.TestCase):
    def test_audio_command_delays_pads_and_cuts_to_scene_length(self):
        cmd = ghep.lenh_am_canh(Path("giong/canh-1.mp3"), Path("am-1.wav"), 6.0)
        joined = " ".join(cmd)
        self.assertIn("adelay=700:all=1", joined)
        self.assertIn("apad=whole_dur=6.000", joined)
        self.assertEqual(cmd[cmd.index("-t") + 1], "6.000")

    def test_video_command_uses_15_fps_frames_and_30_fps_output(self):
        cmd = ghep.lenh_video(Path("anh"), Path("am.txt"), Path(".khung/video.mp4"), 15, ".khung/phu-de.srt")
        self.assertEqual(cmd[cmd.index("-framerate") + 1], "15")
        self.assertEqual(cmd[cmd.index("-r") + 1], "30")
        self.assertIn("libx264", cmd)
        vf = cmd[cmd.index("-vf") + 1]
        self.assertTrue(vf.startswith("subtitles=.khung/phu-de.srt:force_style="), vf)

    def test_video_command_without_burned_subtitles_has_no_filter(self):
        cmd = ghep.lenh_video(Path("anh"), Path("am.txt"), Path(".khung/video.mp4"), 15, None)
        self.assertNotIn("-vf", cmd)


class AssembleTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)

    def project(self, name):
        thu_muc = Path(self.tmp.name) / name
        (thu_muc / ".khung" / "anh").mkdir(parents=True)
        return thu_muc

    def giong(self, thu_muc):
        out = []
        for cl in PLAN:
            mp3 = thu_muc / "giong" / f"canh-{cl.so}.mp3"
            mp3.parent.mkdir(exist_ok=True)
            mp3.write_bytes(b"ID3")
            out.append(lich.GiongInfo(mp3=mp3, giay=cl.giay_giong, moc_cau=[], uoc_luong=False, nguon="may"))
        return out

    def fake_run(self, thu_muc, calls, fail=None):
        def run(cmd, **kwargs):
            calls.append((cmd, kwargs.get("cwd")))
            if cmd[-1].endswith("video.mp4"):
                Path(cmd[-1] if Path(cmd[-1]).is_absolute() else Path(kwargs["cwd"]) / cmd[-1]).write_bytes(b"mp4")
            if fail:
                return subprocess.CompletedProcess(cmd, 1, "", "loi ffmpeg")
            return subprocess.CompletedProcess(cmd, 0, "", "")
        return run

    def test_folder_with_spaces_diacritics_and_apostrophe(self):
        thu_muc = self.project("Bài 5 – Sulfur dioxide (thử) 'a'")
        calls = []
        files = ghep.ghep_video(thu_muc, PLAN, self.giong(thu_muc), "hinh", run=self.fake_run(thu_muc, calls))
        self.assertEqual(files, ["video.mp4"])
        self.assertTrue((thu_muc / "video.mp4").is_file())
        listing = (thu_muc / ".khung" / "am.txt").read_text(encoding="utf-8")
        self.assertEqual(listing.count("file '"), 2)
        self.assertIn("Bài 5 – Sulfur dioxide (thử) '\\''a'\\''", listing)
        self.assertTrue(all(cwd == thu_muc for _, cwd in calls))
        self.assertEqual(len(calls), 3)

    def test_subtitle_modes(self):
        for mode, expect_srt, expect_burn in (("hinh", False, True), ("file", True, False), ("khong", False, False)):
            with self.subTest(mode=mode):
                thu_muc = self.project(f"p-{mode}")
                calls = []
                files = ghep.ghep_video(thu_muc, PLAN, self.giong(thu_muc), mode, run=self.fake_run(thu_muc, calls))
                self.assertEqual((thu_muc / "phu-de.srt").is_file(), expect_srt)
                self.assertEqual("phu-de.srt" in files, expect_srt)
                joined = " ".join(calls[-1][0])
                self.assertEqual("subtitles=" in joined, expect_burn)

    def test_ffmpeg_failure_is_a_dung_error(self):
        thu_muc = self.project("loi")
        with self.assertRaises(media.MediaError) as caught:
            ghep.ghep_video(thu_muc, PLAN, self.giong(thu_muc), "khong", run=self.fake_run(thu_muc, [], fail=True))
        self.assertEqual(caught.exception.step, "dung")
        self.assertIn("loi ffmpeg", str(caught.exception))

    def test_missing_ffmpeg_is_an_ffmpeg_error(self):
        thu_muc = self.project("khong-co")

        def run(cmd, **kwargs):
            raise FileNotFoundError("ffmpeg")

        with self.assertRaises(media.MediaError) as caught:
            ghep.ghep_video(thu_muc, PLAN, self.giong(thu_muc), "khong", run=run)
        self.assertEqual(caught.exception.step, "ffmpeg")

    def test_video_open_in_a_player_is_a_write_error_with_advice(self):
        thu_muc = self.project("dang-mo")
        (thu_muc / "video.mp4").write_bytes(b"cu")
        import os
        real_replace = os.replace

        def deny(src, dst):
            raise PermissionError("dang mo")

        os.replace = deny
        try:
            with self.assertRaises(media.MediaError) as caught:
                ghep.ghep_video(thu_muc, PLAN, self.giong(thu_muc), "khong", run=self.fake_run(thu_muc, []))
        finally:
            os.replace = real_replace
        self.assertEqual(caught.exception.step, "write")
        self.assertIn("trình phát", caught.exception.fix)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Chạy test để thấy lỗi**

Run: `venv\Scripts\python.exe -m unittest tools/vi/tests/test_video_ma_ghep.py`
Expected: FAIL — `ImportError: cannot import name 'ghep'`.

- [ ] **Step 3: Viết `ghep.py`**

```python
"""Ghép khung hình + tiếng + phụ đề thành video.mp4 bằng FFmpeg. Thư mục dự án chỉ được truyền dưới dạng cwd và đường dẫn tuyệt đối."""

from __future__ import annotations

import os
import re
import subprocess
from pathlib import Path

from video_parts import media, srt

from .lich import DAN_DAU, FPS

_MARKUP_RE = re.compile(r"\*\*|~|\^")
STYLE = "FontName=Segoe UI,FontSize=12,Outline=1,Shadow=0,MarginV=24"


def cues_phu_de(cac_lich: list) -> list:
    cues = []
    so = 0
    for cl in cac_lich:
        het = cl.bat_dau + DAN_DAU + cl.giay_giong
        for k, cau in enumerate(cl.cau):
            start = cl.bat_dau + cl.moc_cau[k]
            end = cl.bat_dau + cl.moc_cau[k + 1] if k + 1 < len(cl.cau) else het
            so += 1
            cues.append(srt.Cue(index=so, start=start, end=max(end, start + 0.1), text=_MARKUP_RE.sub("", cau)))
    return cues


def lenh_am_canh(mp3: Path, wav: Path, thoi_luong: float) -> list:
    return [
        "ffmpeg", "-y", "-hide_banner", "-loglevel", "error", "-i", str(mp3),
        "-af", f"adelay={int(round(DAN_DAU * 1000))}:all=1,apad=whole_dur={thoi_luong:.3f}",
        "-t", f"{thoi_luong:.3f}", "-ar", "44100", "-ac", "1", "-c:a", "pcm_s16le", str(wav),
    ]


def lenh_video(anh_dir: Path, danh_sach_am: Path, out_mp4: Path, fps: int, phu_de_tuong_doi) -> list:
    cmd = [
        "ffmpeg", "-y", "-hide_banner", "-loglevel", "error",
        "-framerate", str(fps), "-i", str(anh_dir / "f%06d.png"),
        "-f", "concat", "-safe", "0", "-i", str(danh_sach_am),
    ]
    if phu_de_tuong_doi:
        cmd += ["-vf", f"subtitles={phu_de_tuong_doi}:force_style='{STYLE}'"]
    cmd += ["-r", "30", "-c:v", "libx264", "-preset", "medium", "-crf", "20", "-pix_fmt", "yuv420p",
            "-c:a", "aac", "-b:a", "160k", "-movflags", "+faststart", str(out_mp4)]
    return cmd


def _chay(cmd: list, run, cwd: Path) -> None:
    try:
        proc = run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=7200, cwd=cwd)
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise media.MediaError("ffmpeg", f"Không chạy được FFmpeg: {exc}", media.FIX_FFMPEG) from exc
    if proc.returncode != 0:
        loi = (proc.stderr or "").strip()[-400:]
        raise media.MediaError("dung", f"FFmpeg lỗi (mã {proc.returncode}): {loi}", "Báo nội dung lỗi cho người bảo trì.")


def ghep_video(thu_muc: Path, cac_lich: list, cac_giong: list, phu_de: str, fps: int = FPS, run=subprocess.run) -> list:
    lam = thu_muc / ".khung"
    wavs = []
    for cl, giong in zip(cac_lich, cac_giong):
        wav = lam / f"am-{cl.so}.wav"
        _chay(lenh_am_canh(giong.mp3, wav, cl.thoi_luong), run, thu_muc)
        wavs.append(wav)
    danh_sach = lam / "am.txt"
    danh_sach.write_text(media.build_audio_concat_text(wavs), encoding="utf-8")
    cues = srt.render_srt(cues_phu_de(cac_lich))
    files = ["video.mp4"]
    burn = None
    if phu_de == "hinh":
        (lam / "phu-de.srt").write_text(cues, encoding="utf-8")
        burn = ".khung/phu-de.srt"
    elif phu_de == "file":
        (thu_muc / "phu-de.srt").write_text(cues, encoding="utf-8")
        files.append("phu-de.srt")
    tam = lam / "video.mp4"
    _chay(lenh_video(lam / "anh", danh_sach, tam, fps, burn), run, thu_muc)
    try:
        os.replace(tam, thu_muc / "video.mp4")
    except PermissionError as exc:
        raise media.MediaError("write", "Không ghi được video.mp4 (đang mở ở nơi khác).",
                               "Đóng video.mp4 nếu đang mở trong trình phát rồi chạy lại.") from exc
    return files
```

- [ ] **Step 4: Chạy test đến khi xanh**

Run: `venv\Scripts\python.exe -m unittest tools/vi/tests/test_video_ma_ghep.py -v`
Expected: PASS. (Test `test_video_open_in_a_player...` thay `os.replace` trên module `os` toàn cục trong lúc chạy; nó khôi phục trong `finally`.)

- [ ] **Step 5: Commit**

```bash
git add tools/vi/video_ma_parts/ghep.py tools/vi/tests/test_video_ma_ghep.py
git commit -m "feat(vi): assemble frames, narration and subtitles with FFmpeg"
```

---

### Task 8: Công cụ dòng lệnh `video_ma.py`

**Files:**
- Create: `tools/vi/video_ma.py`
- Test: `tools/vi/tests/test_video_ma_cong_cu.py`

**Interfaces:**
- Consumes: mọi module của `video_ma_parts`; `thi_nghiem_parts.thu_vien`; `video_parts.media`.
- Produces: `python tools/vi/video_ma.py <thư_mục> [--plan-only] [--xem-truoc]`; hàm `main(argv=None) -> int`; `co_ffmpeg() -> bool`, `co_chromium() -> bool` (test giả lập được); `chay(thu_muc, plan_only, xem_truoc, warnings) -> dict`. Dòng JSON: `ready`, `files`, `so_canh`, `thoi_luong_giay`, `phong_cach`, `giong` (`may`/`co-san`/`hon-hop`/`null`), `warnings`, `error` (`{step, message, fix}` hoặc `null`). `error.step` ∈ `input, parse, canh, giong, chromium, ffmpeg, dung, write, internal`. Thư mục tạm `<thư_mục>/.khung/` được **xoá trước khi bắt đầu chụp và luôn xoá khi xong** (kể cả lỗi).

- [ ] **Step 1: Viết test thất bại**

Tạo `tools/vi/tests/test_video_ma_cong_cu.py`:

```python
"""Test công cụ dòng lệnh video_ma.py: mọi nhánh ra đúng một dòng JSON, không gọi giọng thật, Chromium/FFmpeg giả lập."""

import contextlib
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))

import video_ma  # noqa: E402
from video_ma_parts import lich  # noqa: E402
from video_parts import media  # noqa: E402

MAU = (TOOLS_VI / "fixtures" / "video-mau" / "video.md").read_text(encoding="utf-8")
MOT_CANH = "---\ntieu-de: T\nmon: Toán\nlop: 8\n---\n\n## Cảnh 1\nloai: tieu-de\nchu: Xin chào\nloi: Xin chào các em.\n"


def chay(args):
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        code = video_ma.main(args)
    lines = [l for l in out.getvalue().splitlines() if l.strip()]
    return code, lines, err.getvalue()


class CliTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dir = Path(self.tmp.name) / "bai"
        self.dir.mkdir()

    def viet(self, text):
        (self.dir / "video.md").write_text(text, encoding="utf-8")

    def one_json(self, args):
        code, lines, _ = chay(args)
        self.assertEqual(len(lines), 1, lines)
        return code, json.loads(lines[0])

    def test_missing_folder_and_missing_file_are_input_errors(self):
        code, data = self.one_json([str(self.dir / "khong-co")])
        self.assertEqual((code, data["error"]["step"], data["ready"]), (1, "input", False))
        code, data = self.one_json([str(self.dir)])
        self.assertEqual((code, data["error"]["step"]), (1, "input"))
        self.assertIn("video-giai-thich.md", data["error"]["fix"])

    def test_parse_error_names_the_line(self):
        self.viet("---\ntieu-de: T\nmon: Toán\n---\n")
        code, data = self.one_json([str(self.dir)])
        self.assertEqual((code, data["error"]["step"]), (1, "parse"))
        self.assertIn("Dòng 4", data["error"]["message"])

    def test_scene_content_error_is_canh(self):
        self.viet(MOT_CANH.replace("chu: Xin chào", "chu: " + "a" * 95))
        code, data = self.one_json([str(self.dir)])
        self.assertEqual((code, data["error"]["step"]), (1, "canh"))
        self.assertIn("Cảnh 1", data["error"]["message"])

    def test_plan_only_checks_and_writes_nothing(self):
        self.viet(MAU)
        code, data = self.one_json([str(self.dir), "--plan-only"])
        self.assertEqual(code, 0)
        self.assertTrue(data["ready"])
        self.assertEqual(data["so_canh"], 8)
        self.assertEqual(data["files"], [])
        self.assertIsNone(data["error"])
        self.assertEqual(sorted(p.name for p in self.dir.iterdir()), ["video.md"])

    def test_missing_ffmpeg_and_chromium_are_reported_before_any_voice(self):
        self.viet(MOT_CANH)
        with mock.patch.object(video_ma, "co_ffmpeg", return_value=False), \
                mock.patch.object(video_ma.giong, "lay_giong", side_effect=AssertionError("không được gọi giọng")):
            code, data = self.one_json([str(self.dir)])
        self.assertEqual((code, data["error"]["step"]), (1, "ffmpeg"))
        with mock.patch.object(video_ma, "co_ffmpeg", return_value=True), \
                mock.patch.object(video_ma, "co_chromium", return_value=False), \
                mock.patch.object(video_ma.giong, "lay_giong", side_effect=AssertionError("không được gọi giọng")):
            code, data = self.one_json([str(self.dir)])
        self.assertEqual((code, data["error"]["step"]), (1, "chromium"))
        self.assertIn("pptmaster.ps1", data["error"]["fix"])

    def test_voice_error_is_reported_as_giong(self):
        self.viet(MOT_CANH)
        boom = media.MediaError("giong", "mất mạng", "Có mạng rồi chạy lại.")
        with mock.patch.object(video_ma, "co_ffmpeg", return_value=True), \
                mock.patch.object(video_ma, "co_chromium", return_value=True), \
                mock.patch.object(video_ma.giong, "lay_giong", side_effect=boom):
            code, data = self.one_json([str(self.dir)])
        self.assertEqual((code, data["error"]["step"]), (1, "giong"))

    def test_unexpected_error_is_internal(self):
        self.viet(MOT_CANH)
        with mock.patch.object(video_ma, "co_ffmpeg", side_effect=RuntimeError("bất ngờ")):
            code, data = self.one_json([str(self.dir)])
        self.assertEqual((code, data["error"]["step"]), (1, "internal"))
        self.assertIn("bất ngờ", data["error"]["message"])

    def test_success_payload_and_temp_folder_is_cleaned(self):
        self.viet(MOT_CANH)
        giong = lich.GiongInfo(mp3=self.dir / "giong" / "canh-1.mp3", giay=3.0, moc_cau=[0.0], uoc_luong=False, nguon="may")

        class FakeBrowser:
            def __enter__(self_inner):
                return object()

            def __exit__(self_inner, *a):
                return False

        (self.dir / ".khung" / "anh").mkdir(parents=True)
        (self.dir / ".khung" / "anh" / "f000000.png").write_bytes(b"khung-cu")

        def fake_chup(page, html, so_khung, fps, thu_muc, so_dau, ghi_log=None):
            self.assertFalse((thu_muc / "f000000.png").exists(), "khung cũ phải bị dọn trước khi chụp")
            for i in range(so_khung):
                (thu_muc / f"f{so_dau + i:06d}.png").write_bytes(b"x")
            return so_dau + so_khung

        def fake_ghep(thu_muc, cac_lich, cac_giong, phu_de, **kw):
            (thu_muc / "video.mp4").write_bytes(b"mp4")
            return ["video.mp4"]

        with mock.patch.object(video_ma, "co_ffmpeg", return_value=True), \
                mock.patch.object(video_ma, "co_chromium", return_value=True), \
                mock.patch.object(video_ma.giong, "lay_giong", return_value=giong), \
                mock.patch.object(video_ma.chup, "trinh_duyet", return_value=FakeBrowser()), \
                mock.patch.object(video_ma.chup, "trang_moi", return_value=object()), \
                mock.patch.object(video_ma.chup, "kiem_tran", return_value=[]), \
                mock.patch.object(video_ma.chup, "chup_canh", side_effect=fake_chup), \
                mock.patch.object(video_ma.ghep, "ghep_video", side_effect=fake_ghep):
            code, data = self.one_json([str(self.dir)])
        self.assertEqual(code, 0, data)
        self.assertTrue(data["ready"])
        self.assertEqual(data["files"], ["video.mp4"])
        self.assertEqual(data["so_canh"], 1)
        self.assertEqual(data["giong"], "may")
        self.assertEqual(data["phong_cach"], "viet-tay")
        self.assertAlmostEqual(data["thoi_luong_giay"], lich.thoi_luong_canh(3.0), delta=0.01)
        self.assertFalse((self.dir / ".khung").exists())

    def test_thu_muc_khung_tam_duoc_don(self):
        self.test_success_payload_and_temp_folder_is_cleaned()

    def test_overflow_is_a_canh_error_before_capturing(self):
        self.viet(MOT_CANH)
        giong = lich.GiongInfo(mp3=self.dir / "x.mp3", giay=3.0, moc_cau=[0.0], uoc_luong=False, nguon="co-san")

        class FakeBrowser:
            def __enter__(self_inner):
                return object()

            def __exit__(self_inner, *a):
                return False

        with mock.patch.object(video_ma, "co_ffmpeg", return_value=True), \
                mock.patch.object(video_ma, "co_chromium", return_value=True), \
                mock.patch.object(video_ma.giong, "lay_giong", return_value=giong), \
                mock.patch.object(video_ma.chup, "trinh_duyet", return_value=FakeBrowser()), \
                mock.patch.object(video_ma.chup, "trang_moi", return_value=object()), \
                mock.patch.object(video_ma.chup, "kiem_tran", return_value=["chu"]), \
                mock.patch.object(video_ma.chup, "chup_canh", side_effect=AssertionError("không được chụp")):
            code, data = self.one_json([str(self.dir)])
        self.assertEqual((code, data["error"]["step"]), (1, "canh"))
        self.assertIn("chu", data["error"]["message"])
        self.assertIn("Cảnh 1", data["error"]["message"])
        self.assertFalse((self.dir / ".khung").exists())

    def test_preview_needs_no_voice_and_no_ffmpeg(self):
        self.viet(MAU)

        class FakeBrowser:
            def __enter__(self_inner):
                return object()

            def __exit__(self_inner, *a):
                return False

        def fake_cuoi(page, html, duong_dan):
            duong_dan.parent.mkdir(parents=True, exist_ok=True)
            duong_dan.write_bytes(b"png")

        with mock.patch.object(video_ma, "co_ffmpeg", return_value=False), \
                mock.patch.object(video_ma, "co_chromium", return_value=True), \
                mock.patch.object(video_ma.giong, "lay_giong", side_effect=AssertionError("không được gọi giọng")), \
                mock.patch.object(video_ma.chup, "trinh_duyet", return_value=FakeBrowser()), \
                mock.patch.object(video_ma.chup, "trang_moi", return_value=object()), \
                mock.patch.object(video_ma.chup, "kiem_tran", return_value=[]), \
                mock.patch.object(video_ma.chup, "chup_cuoi", side_effect=fake_cuoi):
            code, data = self.one_json([str(self.dir), "--xem-truoc"])
        self.assertEqual(code, 0, data)
        self.assertEqual(data["files"], [f"xem-truoc/canh-{i}.png" for i in range(1, 9)])
        self.assertTrue((self.dir / "xem-truoc" / "canh-8.png").is_file())

    def test_json_line_is_valid_even_with_vietnamese_text(self):
        self.viet(MOT_CANH.replace("Xin chào các em.", "Nhờ ướt nhẫm quyết định."))
        code, lines, _ = chay([str(self.dir), "--plan-only"])
        self.assertEqual(len(lines), 1)
        json.loads(lines[0])


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Chạy test để thấy lỗi**

Run: `venv\Scripts\python.exe -m unittest tools/vi/tests/test_video_ma_cong_cu.py`
Expected: FAIL — `ModuleNotFoundError: video_ma`.

- [ ] **Step 3: Viết `video_ma.py`**

```python
#!/usr/bin/env python3
"""Dựng video giải thích kiểu viết tay từ video.md (giọng đọc, cảnh vẽ bằng mã, phụ đề).

  python tools/vi/video_ma.py <thư_mục> [--plan-only] [--xem-truoc]

stdout đúng một dòng JSON. Hướng dẫn: docs/vi/tro-ly/video-giai-thich.md
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import os
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from thi_nghiem_parts import thu_vien  # noqa: E402
from video_ma_parts import chup, ghep, giong, kiem, lich, parse, trang  # noqa: E402
from video_parts import media  # noqa: E402

FIX_INPUT = "Viết video.md trong thư mục dự án (xem docs/vi/tro-ly/video-giai-thich.md) rồi chạy lại."
FIX_INTERNAL = "Lỗi ngoài dự kiến; dán nguyên thông báo này cho người bảo trì."


def log(text: str) -> None:
    print(text, file=sys.stderr, flush=True)


def emit(payload: dict) -> None:
    text = json.dumps(payload, ensure_ascii=False) + "\n"
    try:
        sys.stdout.write(text)
    except UnicodeEncodeError:
        sys.stdout.buffer.write(text.encode("utf-8", errors="replace"))
    sys.stdout.flush()


def co_ffmpeg() -> bool:
    return shutil.which("ffmpeg") is not None and shutil.which("ffprobe") is not None


def co_chromium() -> bool:
    root = os.environ.get("LOCALAPPDATA")
    if not root:
        return False
    base = Path(root) / "ms-playwright"
    if not (base.is_dir() and (any(base.glob("chromium-*")) or any(base.glob("chromium_headless_shell-*")))):
        return False
    return importlib.util.find_spec("playwright") is not None


def _mo_hinh(video: parse.Video, thu_muc: Path) -> dict:
    return {c.so: thu_vien.load(c.truong["mau"][0], thu_muc) for c in video.canh if c.loai == "thi-nghiem"}


def _trang(video, cac_lich, models) -> list:
    return [trang.dung_trang(lich.du_lieu_canh(c, cl, models.get(c.so)), models.get(c.so)) for c, cl in zip(video.canh, cac_lich)]


def _kiem_tran_tat_ca(page, video, trang_html) -> None:
    for canh, html in zip(video.canh, trang_html):
        tran = chup.kiem_tran(page, html)
        if tran:
            raise kiem.CanhError(canh.so, f"chữ ở mục `{', '.join(tran)}` tràn khung. Rút ngắn nội dung hoặc chia thành hai cảnh.")


def _xem_truoc(video: parse.Video, thu_muc: Path, warnings: list) -> dict:
    gia = [lich.GiongInfo(mp3=None, giay=8.0, moc_cau=[], uoc_luong=True, nguon="may") for _ in video.canh]
    cac_lich, _ = lich.dung_lich(video.canh, gia, kiem_moc=False)
    trang_html = _trang(video, cac_lich, _mo_hinh(video, thu_muc))
    ra = thu_muc / "xem-truoc"
    files = []
    with chup.trinh_duyet() as browser:
        page = chup.trang_moi(browser)
        _kiem_tran_tat_ca(page, video, trang_html)
        for canh, html in zip(video.canh, trang_html):
            chup.chup_cuoi(page, html, ra / f"canh-{canh.so}.png")
            files.append(f"xem-truoc/canh-{canh.so}.png")
    return {"files": files, "so_canh": len(video.canh), "thoi_luong_giay": None,
            "phong_cach": video.meta["phong-cach"], "giong": None}


def _dung(video: parse.Video, thu_muc: Path, warnings: list) -> dict:
    if not co_ffmpeg():
        raise media.MediaError("ffmpeg", "Chưa có FFmpeg.", media.FIX_FFMPEG)
    if not co_chromium():
        raise media.MediaError("chromium", "Chưa cài Chromium hoặc playwright.", chup.FIX_CHROMIUM)
    cac_giong = [giong.lay_giong(c.so, c.loi, thu_muc / "giong", video.meta["giong"], video.meta["toc-do"]) for c in video.canh]
    cac_lich, canh_bao = lich.dung_lich(video.canh, cac_giong)
    warnings.extend(canh_bao)
    trang_html = _trang(video, cac_lich, _mo_hinh(video, thu_muc))
    lam = thu_muc / ".khung"
    shutil.rmtree(lam, ignore_errors=True)
    anh = lam / "anh"
    anh.mkdir(parents=True)
    try:
        with chup.trinh_duyet() as browser:
            page = chup.trang_moi(browser)
            _kiem_tran_tat_ca(page, video, trang_html)
            so = 0
            for canh, cl, html in zip(video.canh, cac_lich, trang_html):
                log(f"Chụp cảnh {canh.so}/{len(video.canh)} ({cl.so_khung} khung)...")
                so = chup.chup_canh(page, html, cl.so_khung, lich.FPS, anh, so, log)
        log("Ghép video bằng FFmpeg...")
        files = ghep.ghep_video(thu_muc, cac_lich, cac_giong, video.meta["phu-de"])
    finally:
        shutil.rmtree(lam, ignore_errors=True)
    nguon = {g.nguon for g in cac_giong}
    return {"files": files, "so_canh": len(video.canh), "thoi_luong_giay": round(sum(cl.thoi_luong for cl in cac_lich), 2),
            "phong_cach": video.meta["phong-cach"], "giong": nguon.pop() if len(nguon) == 1 else "hon-hop"}


def chay(thu_muc: Path, plan_only: bool, xem_truoc: bool, warnings: list) -> dict:
    md = thu_muc / "video.md"
    if not thu_muc.is_dir() or not md.is_file():
        raise media.MediaError("input", f"Không thấy {md}.", FIX_INPUT)
    video = parse.parse(md.read_text(encoding="utf-8-sig"))
    warnings.extend(kiem.kiem(video, thu_muc))
    if plan_only:
        return {"files": [], "so_canh": len(video.canh), "thoi_luong_giay": None,
                "phong_cach": video.meta["phong-cach"], "giong": None}
    if xem_truoc:
        if not co_chromium():
            raise media.MediaError("chromium", "Chưa cài Chromium hoặc playwright.", chup.FIX_CHROMIUM)
        return _xem_truoc(video, thu_muc, warnings)
    return _dung(video, thu_muc, warnings)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Dựng video giải thích kiểu viết tay từ video.md")
    ap.add_argument("thu_muc")
    ap.add_argument("--plan-only", action="store_true")
    ap.add_argument("--xem-truoc", action="store_true")
    try:
        args = ap.parse_args(argv)
    except SystemExit:
        emit({"ready": False, "files": [], "so_canh": 0, "thoi_luong_giay": None, "phong_cach": None, "giong": None, "warnings": [],
              "error": {"step": "input", "message": "Sai tham số dòng lệnh.", "fix": "Xem: python tools/vi/video_ma.py --help"}})
        return 1
    warnings: list = []
    base = {"ready": False, "files": [], "so_canh": 0, "thoi_luong_giay": None, "phong_cach": None, "giong": None}
    try:
        kq = chay(Path(args.thu_muc), args.plan_only, args.xem_truoc, warnings)
        emit({**base, **kq, "ready": True, "warnings": warnings, "error": None})
        return 0
    except parse.ParseError as exc:
        error = {"step": "parse", "message": str(exc), "fix": "Sửa đúng dòng đó trong video.md rồi chạy lại."}
    except kiem.CanhError as exc:
        error = {"step": "canh", "message": str(exc), "fix": "Rút gọn hoặc sửa nội dung cảnh đó theo thông báo."}
    except media.MediaError as exc:
        error = {"step": exc.step, "message": str(exc), "fix": exc.fix}
    except OSError as exc:
        error = {"step": "write", "message": f"Không ghi được file: {exc}", "fix": "Đóng file đang mở và kiểm tra ổ đĩa rồi chạy lại."}
    except Exception as exc:  # noqa: BLE001
        error = {"step": "internal", "message": f"{type(exc).__name__}: {exc}", "fix": FIX_INTERNAL}
    emit({**base, "warnings": warnings, "error": error})
    return 1


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: Chạy test đến khi xanh**

Run: `venv\Scripts\python.exe -m unittest tools/vi/tests/test_video_ma_cong_cu.py -v`
Expected: PASS. Ghi chú: `test_missing_ffmpeg...` phải thấy `ffmpeg` được kiểm **trước** `lay_giong` — `_dung` đã xếp như vậy.

- [ ] **Step 5: Chạy toàn bộ test mới và cũ liên quan**

Run: `venv\Scripts\python.exe -m unittest tools/vi/tests/test_video_ma_parse.py tools/vi/tests/test_video_ma_lich.py tools/vi/tests/test_video_ma_giong.py tools/vi/tests/test_video_ma_canh.py tools/vi/tests/test_video_ma_chup.py tools/vi/tests/test_video_ma_ghep.py tools/vi/tests/test_video_ma_cong_cu.py tools/vi/tests/test_video.py`
Expected: PASS (test Chromium bỏ qua ở venv của repo).

- [ ] **Step 6: Commit**

```bash
git add tools/vi/video_ma.py tools/vi/tests/test_video_ma_cong_cu.py
git commit -m "feat(vi): add the video_ma command that builds explainer videos"
```

---

### Task 9: Dựng thật đầu-cuối và kiểm bằng mắt

**Files:**
- Create: `tools/vi/tests/test_video_ma_tich_hop.py`
- Modify (nếu lộ lỗi): file tương ứng ở Task 5–8, kèm test.

**Interfaces:**
- Consumes: CLI `video_ma.main`; `ffmpeg`, `ffprobe` thật; Chromium thật.
- Produces: bằng chứng dựng thật để ghi vào biên bản ở Task 11.

- [ ] **Step 1: Viết test tích hợp**

Tạo `tools/vi/tests/test_video_ma_tich_hop.py`:

```python
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
        thu_muc, code, data = self.dung("Bài 5 – Sulfur dioxide (thử)", "hinh")
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
```

- [ ] **Step 2: Chạy test tích hợp bằng Python có playwright**

Run: `C:/Users/ADMIN/vmt/v/Scripts/python.exe -m unittest tools/vi/tests/test_video_ma_tich_hop.py -v`
Expected: PASS. Cảnh báo: mở Chromium ẩn và chạy FFmpeg vài chục giây, không mở cửa sổ nào. Nếu lỗi, sửa đúng file gây lỗi (kèm test ở task đó), không nới tích hợp. Trên `venv` của repo test này tự bỏ qua.

- [ ] **Step 3: Xem trước cả tám loại cảnh và xem bằng mắt**

Run:
```bash
mkdir -p "projects/_video/_thu_con_lac"
cp tools/vi/fixtures/video-mau/video.md projects/_video/_thu_con_lac/video.md
C:/Users/ADMIN/vmt/v/Scripts/python.exe tools/vi/video_ma.py projects/_video/_thu_con_lac --xem-truoc
```
Expected: một dòng JSON `ready: true`, `files` gồm 8 ảnh `xem-truoc/canh-1.png` … `canh-8.png`. Dùng công cụ Read xem **từng** ảnh. Kiểm: chữ tiếng Việt đủ dấu, không chữ nào chồng lên nhau hay tràn khung, gạch chân nằm dưới tiêu đề, đồ thị có đủ 4 điểm nối nét và nhãn hai trục, cảnh 8 có con lắc trong khung vẽ và số đo ở cột phải. Ghi mọi lỗi nhìn thấy; sửa toạ độ trong file cảnh tương ứng (kèm chạy lại `node --test`), rồi chạy lại đến khi cả tám ảnh đạt. Ảnh thử là file tạm trong `projects/` (không commit).

- [ ] **Step 4: Dựng video mẫu đầy đủ và kiểm khung giữa cảnh**

Run:
```bash
mkdir -p projects/_video/_thu_con_lac/giong
for n in 1 2 3 4 5 6 7 8; do ffmpeg -y -hide_banner -loglevel error -f lavfi -i "sine=frequency=440:duration=8" -q:a 9 "projects/_video/_thu_con_lac/giong/canh-$n.mp3"; done
C:/Users/ADMIN/vmt/v/Scripts/python.exe tools/vi/video_ma.py projects/_video/_thu_con_lac
```
(Các lệnh `ffmpeg` chỉ tạo tiếng giả để không cần mạng.)
Expected: `ready: true`, `video.mp4` ra. Trích khung giữa mỗi cảnh: `ffmpeg -y -ss <giây> -i projects/_video/_thu_con_lac/video.mp4 -frames:v 1 khung.png` với các mốc giữa cảnh 3 (đang viết ý), cảnh 7 (đang nối điểm) và cảnh 8 (con lắc đang lắc); xem bằng Read. Kiểm: nét đang vẽ dở, chữ đang viết dở đúng vị trí, phụ đề in ở đáy, con lắc chuyển động (hai khung cách nhau 0,5 giây phải khác nhau). Ghi thời gian dựng và kích thước file để đưa vào biên bản. Không commit gì trong `projects/`.

- [ ] **Step 5: Commit**

```bash
git add tools/vi/tests/test_video_ma_tich_hop.py
git commit -m "test(vi): end-to-end build of an explainer video with synthetic narration"
```
(Nếu Step 3–4 buộc sửa file cảnh, commit chúng riêng với thông điệp `fix(vi): ...` mô tả lỗi bố cục thật đã gặp.)

---

### Task 10: Lớp hướng dẫn: tài liệu, AGENTS.vi.md, luật Antigravity, test

**Files:**
- Create: `docs/vi/tro-ly/video-giai-thich.md`, `docs/vi/tro-ly/canh-video.md`, `docs/vi/video-giai-thich.md`
- Modify: `AGENTS.vi.md`, `.agents/rules/ppt-master-vi.md`, `docs/vi/tro-ly/quy-trinh-hoi.md`, `docs/vi/tro-ly/mau-brief.md`, `docs/vi/xu-ly-loi.md`, `docs/vi/cau-lenh-mau.md`, `docs/vi/bat-dau-nhanh.md`, `README.md`
- Modify: `tools/vi/tests/test_vi_layer.py`

**Interfaces:**
- Consumes: CLI ở Task 8; ngữ pháp ở Task 2; danh mục cảnh ở Task 5.
- Produces: loại việc thứ 10 "Video giải thích" (mục 15 của `AGENTS.vi.md`), chữ "10 loại", bảng 10 dòng khớp giữa `quy-trinh-hoi.md` và file luật Antigravity.

**Mẫu để làm theo:** gói thí nghiệm ảo đã làm đúng việc này ở commit `918820bf` (agent) và `cba86b62` (thầy cô). Chạy `git show 918820bf` và `git show cba86b62` trước, rồi lặp lại từng thay đổi cho loại việc mới. Nội dung riêng của gói này ở các bước dưới.

- [ ] **Step 1: Viết test thất bại (thêm vào `test_vi_layer.py`)**

Thêm hằng và lớp test (sao theo `ExperimentGuideTest`, dùng các hàm sẵn có `read`, `section`, `h2_headings`, `numbered_items`, `task_table_rows`, `LINK_RE`, `REPO_ROOT`):

```python
AGENTS_VI_EXPLAINER_HEADING = "## 15. Làm video giải thích"
EXPLAINER_GUIDE = "docs/vi/tro-ly/video-giai-thich.md"
SCENE_GUIDE = "docs/vi/tro-ly/canh-video.md"
EXPLAINER_COMMAND = r"python tools\vi\video_ma.py"
EXPLAINER_GUIDE_HEADINGS = (
    "## Khi nào dùng",
    "## Câu hỏi bắt buộc",
    "## Câu hỏi tuỳ chọn",
    "## Tạo nhanh",
    "## Cấu trúc video.md",
    "## Đầu ra",
    "## Ghi vào brief",
)


class ExplainerVideoGuideTest(unittest.TestCase):
    def test_guide_has_its_own_sections_in_order(self):
        self.assertEqual(h2_headings(read(EXPLAINER_GUIDE)), list(EXPLAINER_GUIDE_HEADINGS))

    def test_guide_questions_are_limited_and_have_suggestions(self):
        items = numbered_items(section(read(EXPLAINER_GUIDE), "## Câu hỏi bắt buộc"))
        self.assertTrue(1 <= len(items) <= 7, f"{len(items)} câu")
        for item in items:
            self.assertIn("Gợi ý:", item)
        quick = numbered_items(section(read(EXPLAINER_GUIDE), "## Tạo nhanh"))
        self.assertTrue(2 <= len(quick) <= 3, f"{len(quick)} câu")

    def test_guide_example_parses_with_the_real_reader(self):
        from video_ma_parts import kiem, parse

        body = section(read(EXPLAINER_GUIDE), "## Cấu trúc video.md")
        blocks = re.findall(r"```[a-z]*\n(---\n.*?)```", body, re.S)
        self.assertGreaterEqual(len(blocks), 1)
        for block in blocks:
            video = parse.parse(block)
            self.assertEqual(kiem.kiem(video, REPO_ROOT), [])

    def test_scene_guide_lists_every_scene_type_and_field(self):
        from video_ma_parts import parse

        text = read(SCENE_GUIDE)
        for loai, (required, optional, repeated) in parse.SCENE_SPEC.items():
            self.assertIn(f"`{loai}`", text)
            for key in (*required, *optional, *repeated):
                with self.subTest(loai=loai, key=key):
                    self.assertIn(f"`{key}`", text)
        for phrase in ("li-con-lac-don", "hoa-chuan-do", "toan-ham-so", "tham-so", "Không chèn địa chỉ web"):
            self.assertIn(phrase, text)

    def test_guide_states_the_grammar_and_limits(self):
        body = section(read(EXPLAINER_GUIDE), "## Cấu trúc video.md")
        for phrase in ("tieu-de", "mon", "lop", "phong-cach", "giong", "toc-do", "phu-de", "## Cảnh", "loai:", "loi:",
                       "canh-1.mp3", "Không chèn địa chỉ web", "H~2~SO~4~"):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, body)

    def test_guide_names_outputs_and_forbids_hand_editing_frames(self):
        body = section(read(EXPLAINER_GUIDE), "## Đầu ra")
        for phrase in (EXPLAINER_COMMAND, "video.mp4", "phu-de.srt", "xem-truoc", "--xem-truoc", "--plan-only",
                       "projects\\_video\\", "Không tự chạy FFmpeg", "Không viết HTML"):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, body)

    def test_guide_skips_the_pptx_only_steps(self):
        body = section(read(EXPLAINER_GUIDE), "## Ghi vào brief")
        for phrase in ("projects/_video/", "brief.md", "import-sources", "dòng chốt"):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, body)

    def test_guides_have_no_markdown_links(self):
        for name in (EXPLAINER_GUIDE, SCENE_GUIDE):
            self.assertEqual(LINK_RE.findall(read(name)), [], name)

    def test_agents_vi_section_15_is_last_and_routes_the_word_video(self):
        headings = h2_headings(read("AGENTS.vi.md"))
        self.assertEqual(headings[-1], AGENTS_VI_EXPLAINER_HEADING)
        body = section(read("AGENTS.vi.md"), AGENTS_VI_EXPLAINER_HEADING)
        for phrase in (f"({EXPLAINER_GUIDE})", f"({SCENE_GUIDE})", EXPLAINER_COMMAND, "--plan-only", "--xem-truoc",
                       "`ready`", "error.step", "Không chạy `project_manager.py init`", "không tạo SVG", "không chạm `skills/`",
                       "không commit gì trong `projects/`", "Không tự cài phần mềm", "chromium", "edge-tts"):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, body)
        for step in ("input", "parse", "canh", "giong", "chromium", "ffmpeg", "dung", "write", "internal"):
            self.assertIn(f"| `{step}` |", body)
        assistant = section(read("AGENTS.vi.md"), AGENTS_VI_ASSISTANT_HEADING)
        for phrase in ("video giải thích", "video viết tay", "video whiteboard", "Thầy cô muốn làm video từ bài giảng slide đã có"):
            self.assertIn(phrase, read("AGENTS.vi.md"))

    def test_rule_file_carries_the_gate_inline_and_stays_under_the_cap(self):
        rule = read(".agents/rules/ppt-master-vi.md")
        self.assertLess(len(rule), 12000)
        for phrase in ("10 loại", "Video giải thích", "video_ma.py", "Thầy cô muốn làm video từ bài giảng slide đã có"):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, rule)
```

Đồng thời sửa các test đã có để khớp trạng thái mới (đúng như commit `918820bf` đã làm cho loại thứ 9):
- `test_antigravity_rule_task_table_matches_common_rules`: `len(common)` từ 9 lên **10**.
- `test_agents_vi_keeps_the_five_task_sections_last_in_order` → đổi tên `..._six_task_sections_...` và dịch chỉ số: `headings[-6]` = `AGENTS_VI_ASSISTANT_HEADING`, `[-5]` video slide, `[-4]` đề, `[-3]` giáo án, `[-2]` thí nghiệm, `[-1]` `AGENTS_VI_EXPLAINER_HEADING`.
- Mọi `h2_headings(...)[-5]` trỏ mục 10 (`test_agents_vi_environment_section...` và `SelfInstallGuideTest`) đổi thành `[-6]`.
- `test_task_type_count_matches_the_table`: `"9 loại"` thành `"10 loại"`, danh sách `stale` thêm `"9 loại"`, ở cả hai đoạn (`AGENTS.vi.md` mục 10 và `quy-trinh-hoi.md` mục "Khi nào áp dụng").

- [ ] **Step 2: Chạy để thấy lỗi**

Run: `venv\Scripts\python.exe -m unittest tools/vi/tests/test_vi_layer.py`
Expected: FAIL (thiếu file, thiếu mục 15, sai số loại).

- [ ] **Step 3: Viết `docs/vi/tro-ly/video-giai-thich.md` (cho AI)**

Nội dung bắt buộc (tiếng Việt, không có liên kết Markdown, đúng thứ tự tiêu đề ở `EXPLAINER_GUIDE_HEADINGS`):

- `## Khi nào dùng`: khi thầy cô muốn video giải thích một bài từ nội dung chữ (video viết tay, whiteboard, hoạt hình chữ), khác video từ slide (loại "Video bài giảng"). Câu chỉ có "làm video"/"xuất video" thì hỏi đúng một câu: "Thầy cô muốn làm video từ bài giảng slide đã có, hay dựng video giải thích mới từ nội dung chữ?". Nói rõ không dùng cho video AI sinh (Veo, Flow), không có kiểu Vox hay khổ dọc ở bản này.
- `## Câu hỏi bắt buộc`: bảy câu đánh số, mỗi câu có `Gợi ý:` — (1) môn, lớp, bài và nội dung cần giải thích (xin dán nội dung hoặc dàn ý); (2) học sinh cần hiểu hoặc làm được gì sau video; (3) độ dài mong muốn (gợi ý 2–4 phút, 6–10 cảnh); (4) giọng nam hay nữ, tốc độ chậm/vừa/nhanh (gợi ý nữ, vừa); (5) có muốn chèn cảnh thí nghiệm ảo không, liệt kê 8 mẫu (con lắc đơn, ném xiên, mạch Ohm; chuẩn độ, cân bằng NO2, tốc độ phản ứng; hàm số, xác suất); (6) phụ đề in lên hình hay file rời (gợi ý in lên hình); (7) có file giọng thu sẵn không (gợi ý không, dùng giọng máy, cần mạng).
- `## Câu hỏi tuỳ chọn`: nhấn mạnh video dài, tên video, đối tượng học sinh.
- `## Tạo nhanh`: hai đến ba câu: câu 1 và câu 3 của mục bắt buộc; vẫn ghi brief.
- `## Cấu trúc video.md`: ngữ pháp đúng `parse.py`: khối `---`, khoá `tieu-de`, `mon`, `lop` bắt buộc, `phong-cach` (mặc định `viet-tay`), `giong` (`nu`/`nam`), `toc-do` (`cham`/`vua`/`nhanh`), `phu-de` (`hinh`/`file`/`khong`); `## Cảnh N` đánh số liên tiếp; `loai:` và `loi:` mỗi cảnh; quy ước chữ `H~2~SO~4~`, `m/s^2^`, `**in đậm**`; "Không chèn địa chỉ web"; giới hạn chữ từng trường; cách dùng `giong/canh-1.mp3` (giọng thu sẵn, không cần mạng); một ví dụ đầy đủ bằng **khối mã bắt đầu bằng dòng `---`** (ví dụ 3–4 cảnh, dùng nội dung hợp lệ, qua `parse.parse` và `kiem.kiem` không cảnh báo). Cách chia bài thành cảnh: mỗi cảnh một ý, lời 1–4 câu, câu đầu nêu ý, mỗi ý của cảnh danh sách ứng với một câu.
- `## Đầu ra`: thư mục `projects\_video\<tên_video>\`; lệnh `python tools\vi\video_ma.py projects\_video\<tên_video>`; `--plan-only` (chỉ kiểm), `--xem-truoc` (mỗi cảnh một ảnh ở `xem-truoc`, nên chạy trước khi dựng thật và xem ảnh); đầu ra `video.mp4`, `phu-de.srt` (khi `phu-de: file`), `giong\canh-N.mp3`; đọc dòng JSON; báo đường dẫn, thời lượng, giọng; đọc nguyên văn `warnings`; bảng `error.step` (mọi bước, cột thứ hai là xử lý); "Không tự chạy FFmpeg", "Không viết HTML hay ảnh cảnh bằng tay"; dựng thật mất vài phút (video 5 phút khoảng 4 phút chụp khung), báo trước.
- `## Ghi vào brief`: ghi vào `projects/_video/<tên>/brief.md` theo `mau-brief.md`; không chạy `import-sources`; giữ dòng chốt cách xác nhận như các loại việc khác.

Xem `docs/vi/tro-ly/thi-nghiem-ao.md` để bám giọng văn và bố cục; đừng chép câu chữ riêng của thí nghiệm ảo.

- [ ] **Step 4: Viết `docs/vi/tro-ly/canh-video.md`**

Danh mục tám loại cảnh, mỗi loại một mục: mã loại trong dấu huyền, các trường (mọi khoá trong `parse.SCENE_SPEC`, mỗi khoá trong dấu huyền), giới hạn (lấy đúng từ `kiem.LIMITS`), cách hiện, một ví dụ cảnh ngắn. Loại `thi-nghiem`: liệt kê đủ tám mã mẫu (`li-nem-xien`, `li-con-lac-don`, `li-mach-ohm`, `hoa-chuan-do`, `hoa-can-bang-no2`, `hoa-toc-do`, `toan-ham-so`, `toan-xac-suat`), cách dùng `tham-so: <giây> <mã> <giá trị>` (nội suy tuyến tính, trước mốc đầu giữ mặc định, sau mốc cuối giữ giá trị cuối; tối đa 3 tham số, 3 đại lượng đo; mã tham số và đại lượng đo của từng mẫu xem `docs/vi/tro-ly/mo-hinh-thi-nghiem.md`); nhắc "Không chèn địa chỉ web". Không liên kết Markdown.

- [ ] **Step 5: Viết `docs/vi/video-giai-thich.md` (cho thầy cô)**

Ngắn (một trang): video giải thích là gì, thầy cô cần đưa gì (nội dung bài), AI hỏi vài câu rồi dựng, cần gì trên máy (Chromium 150–300 MB tải một lần và mạng để tạo giọng máy; hoặc giọng thu sẵn `giong\canh-N.mp3` thì không cần mạng), video dài 5 phút dựng khoảng 4–5 phút, cách sửa một cảnh (sửa `video.md` rồi chạy lại: chỉ cảnh đổi bị tạo giọng lại), lưu ý kịch bản và video không lên GitHub. Theo phong cách `docs/vi/thi-nghiem-ao.md`.

- [ ] **Step 6: Sửa `AGENTS.vi.md`**

Ba thay đổi trong các mục đã có:
1. Mục 3: thêm vào danh sách kích hoạt các cụm `"video giải thích", "video viết tay", "video whiteboard", "video hoạt hình chữ"` (giữ nguyên các cụm khác).
2. Mục 10: đổi "một trong 9 loại việc dưới đây" thành "một trong 10 loại việc dưới đây"; đổi "không thuộc 9 loại" thành "không thuộc 10 loại"; thêm dòng bảng `| Video giải thích dựng bằng mã | [docs/vi/tro-ly/video-giai-thich.md](docs/vi/tro-ly/video-giai-thich.md) |` ngay sau dòng "Thí nghiệm ảo"; đoạn "bước xác nhận của upstream vẫn bắt buộc, trừ khi ..." thêm loại "Video giải thích" (mục 15) vào danh sách loại không có bước xác nhận của upstream và đổi "ba loại việc đó" thành "bốn loại việc đó", "xem mục 12, mục 13 và mục 14" thành "xem mục 12, mục 13, mục 14 và mục 15".
3. Mục 11: thêm một dòng đầu mục: "Câu chỉ có "làm video" hoặc "xuất video" (không nói rõ từ slide hay video giải thích): hỏi đúng một câu "Thầy cô muốn làm video từ bài giảng slide đã có, hay dựng video giải thích mới từ nội dung chữ?". Trả lời slide thì làm theo mục này; trả lời video mới thì làm theo mục 15."

Rồi thêm mục cuối `## 15. Làm video giải thích` (đây phải là mục cuối cùng của file), theo mẫu mục 14, gồm: đọc `docs/vi/tro-ly/video-giai-thich.md` và `docs/vi/tro-ly/canh-video.md` (viết dạng `[docs/vi/tro-ly/video-giai-thich.md](docs/vi/tro-ly/video-giai-thich.md)` và `[docs/vi/tro-ly/canh-video.md](docs/vi/tro-ly/canh-video.md)` để test tìm được `(đường_dẫn)`); như mục 4: dùng `venv\Scripts\python.exe` nếu có; các bước: 1. hỏi một lượt theo file hướng dẫn, chờ trả lời; 2. tạo `projects/_video/<tên_video>/` và viết `video.md`; 3. chạy `python tools\vi\video_ma.py projects\_video\<tên_video> --plan-only` rồi `--xem-truoc` và xem ảnh từng cảnh, sửa nội dung nếu chữ chồng hay tràn; 4. chạy `python tools\vi\video_ma.py projects\_video\<tên_video>`; 5. đọc dòng JSON: `ready` là `true` thì báo đường dẫn `video.mp4`, thời lượng, nguồn giọng, đọc nguyên văn `warnings`; bảng `error.step` (chín dòng, mỗi dòng dạng ``| `input` | ... |``, xử lý theo bảng ở spec mục 10: `chromium` thì hỏi thầy cô trước rồi chạy `powershell -NoProfile -ExecutionPolicy Bypass -File tools\vi\pptmaster.ps1 -Action tool -Name chromium` vì tải 150–300 MB; `giong` thì kiểm mạng hoặc đặt sẵn `giong/canh-N.mp3` do thầy cô đưa, thử lại tối đa một lần; `ffmpeg` thì làm theo mục "Công cụ tuỳ chọn" của `docs/vi/cai-dat-bang-ai.md`); "Điều cấm": không tự sửa số liệu, công thức hay lời giảng của thầy cô; không chạy `project_manager.py init`; không tạo SVG; không chạm `skills/`; không commit gì trong `projects/`; "Không tự cài phần mềm nào khác, không tự chạy FFmpeg theo cách riêng"; cần mạng cho giọng edge-tts (dựng không mạng thì dùng file giọng có sẵn); video dựng mất vài phút, báo trước một dòng.

- [ ] **Step 7: Sửa file luật Antigravity và các tài liệu còn lại**

Mở `git show 918820bf -- .agents/rules/ppt-master-vi.md docs/vi/tro-ly/quy-trinh-hoi.md docs/vi/tro-ly/mau-brief.md` và lặp lại từng thay đổi cho loại thứ 10:
- `.agents/rules/ppt-master-vi.md`: "9 loại" thành "10 loại"; thêm dòng bảng loại việc **giống hệt** dòng đã thêm vào `quy-trinh-hoi.md`; thêm vào phần "xem mục" dòng cho video giải thích: câu hỏi cổng "Thầy cô muốn làm video từ bài giảng slide đã có, hay dựng video giải thích mới từ nội dung chữ?" khi câu chỉ có "làm video"; nêu lệnh `python tools\vi\video_ma.py` và tên file hướng dẫn. Viết **trực tiếp trong file** (Antigravity không chép nội dung file nhắc bằng `@`). File phải còn dưới 12.000 ký tự (test đã kiểm).
- `docs/vi/tro-ly/quy-trinh-hoi.md`: "9 loại" thành "10 loại" trong "Khi nào áp dụng"; thêm dòng bảng `| Video giải thích dựng bằng mã | ... |` cùng nội dung cột với `AGENTS.vi.md`; nếu file có bảng loại việc thứ hai khác, thêm dòng tương ứng.
- `docs/vi/tro-ly/mau-brief.md`: thêm dòng cho loại mới theo đúng cách commit `918820bf` làm cho thí nghiệm ảo.
- `docs/vi/xu-ly-loi.md`: thêm mục "Video giải thích" (theo mẫu mục thí nghiệm ảo của commit `cba86b62`): bảng chín `error.step`, cách xử lý mỗi bước, lưu ý đường dẫn quá dài khi cài playwright, giọng cần mạng, chữ tràn khung, file `video.mp4` đang mở.
- `docs/vi/cau-lenh-mau.md`, `docs/vi/bat-dau-nhanh.md`, `README.md`: thêm câu lệnh mẫu ("Làm video giải thích bài Con lắc đơn bằng kiểu viết tay", "Làm video viết tay giải thích phản ứng trao đổi ion"), một đoạn giới thiệu và liên kết tới `docs/vi/video-giai-thich.md`, theo đúng vị trí và giọng của phần thí nghiệm ảo.

- [ ] **Step 8: Chạy toàn bộ test lớp Việt đến khi xanh**

Run: `venv\Scripts\python.exe -m unittest discover -s tools/vi/tests -t .`
Expected: PASS toàn bộ (test cần Chromium/FFmpeg tự bỏ qua trên `venv` của repo). Không sửa test để cho qua: nếu một khẳng định cũ trượt vì dịch chỉ số tiêu đề, sửa **theo Step 1 đã liệt kê**; phần còn lại phải sửa tài liệu.

- [ ] **Step 9: Commit (hai commit)**

```bash
git add AGENTS.vi.md .agents/rules/ppt-master-vi.md docs/vi/tro-ly/video-giai-thich.md docs/vi/tro-ly/canh-video.md docs/vi/tro-ly/quy-trinh-hoi.md docs/vi/tro-ly/mau-brief.md tools/vi/tests/test_vi_layer.py
git commit -m "docs(vi): add the explainer video task type for agents"
git add docs/vi/video-giai-thich.md docs/vi/xu-ly-loi.md docs/vi/cau-lenh-mau.md docs/vi/bat-dau-nhanh.md README.md
git commit -m "docs(vi): document explainer videos for teachers"
```

---

### Task 11: Biên bản nghiệm thu và ghi phiên bản

**Files:**
- Modify: `docs/vi/phat-trien/2026-09-25-video-ma-kiem-thu.md`, `CHANGELOG-VI.md`
- Modify (theo cách phát hành trước): các file phiên bản. Chạy `git show 9a7c2048 --stat` để biết những file nào được sửa khi phát hành 6.3.2-vi.8 và làm tương tự cho `6.3.2-vi.9`.

**Interfaces:**
- Consumes: kết quả các task trước.
- Produces: biên bản ghi kết quả thật; mục `6.3.2-vi.9` trong `CHANGELOG-VI.md`; nhánh sẵn sàng để chủ repo duyệt.

- [ ] **Step 1: Chạy toàn bộ test hai lần**

Run: `venv\Scripts\python.exe -m unittest discover -s tools/vi/tests -t .` rồi `C:/Users/ADMIN/vmt/v/Scripts/python.exe -m unittest discover -s tools/vi/tests -t . -k video_ma`
Expected: lần một PASS (Chromium/FFmpeg bỏ qua), lần hai PASS gồm cả tích hợp và Chromium. Ghi số test và số bỏ qua.

- [ ] **Step 2: Điền biên bản**

Thay dòng `(Điền ở Task 11.)` ở mục 2 bằng: kết quả hai lần chạy test (số test, bỏ qua bao nhiêu và vì sao); bảng tám loại cảnh với kết quả nhìn ảnh xem trước (đạt/chỉnh gì); thông số video mẫu dựng thật ở Task 9 (thời lượng, dung lượng, thời gian dựng, `ffprobe` 1280×720 30 khung/giây có tiếng); các khung giữa cảnh đã xem (viết tay dở, con lắc chuyển động); lỗi thật đã gặp khi dựng và cách sửa; các việc **chưa** kiểm: giọng edge-tts thật (không gọi trong test), cài Chromium từ đầu trên máy sạch, chạy trong Antigravity thật (chủ repo tự thử bằng câu "Làm video giải thích bài Con lắc đơn"). Trung thực: không ghi "đạt" cho điều chưa chạy.

- [ ] **Step 3: Ghi CHANGELOG và phiên bản**

Thêm mục `6.3.2-vi.9` vào `CHANGELOG-VI.md` (theo mẫu mục vi.8): loại việc thứ 10 "Video giải thích" kiểu viết tay, tám loại cảnh, cảnh thí nghiệm dùng tám mô hình đã kiểm, giọng edge-tts hoặc file có sẵn, phụ đề in hoặc file rời; nêu rõ yêu cầu Chromium (tải một lần) và mạng cho giọng máy; ngoài phạm vi (Vox, khổ dọc, video AI sinh). Cập nhật các file phiên bản đúng như commit `9a7c2048`. **Không** tạo tag, **không** push, **không** phát hành.

- [ ] **Step 4: Commit**

```bash
git add docs/vi/phat-trien/2026-09-25-video-ma-kiem-thu.md CHANGELOG-VI.md
git commit -m "docs(vi): record the explainer video acceptance run and changelog"
```

- [ ] **Step 5: Báo chủ repo**

Báo: nhánh `feat/vi-video-ma` đã sẵn sàng; nêu 3 việc chủ repo cần tự làm (xem ảnh xem trước tám cảnh và một video mẫu; thử một lần với giọng edge-tts thật khi có mạng; thử trong Antigravity). Hỏi trước khi merge, push hay phát hành v6.3.2-vi.9.

---

## Self-review

**Phủ spec:** tiêu chí 1–10 (mục 1) → Task 8–9 (đầu ra 1280×720, tám cảnh, cảnh thí nghiệm ăn số `tinh()` ở test Node, dấu tiếng Việt ở Task 1/9, mốc ý ở Task 3/5, giọng có sẵn ở Task 4, chặn lỗi trước khi dựng ở Task 2/8, Antigravity ở Task 10, không lên GitHub: `projects/` đã bị bỏ qua và test Task 10, toàn bộ test xanh ở Task 11). Mục 3 Q1–Q12 → Task 5 (Q5 hàm xác định), Task 4 (Q7, Q8), Task 8 (Q9 lỗi thay vì cắt chữ, Q10 hỏi trước khi cài), Task 7 (Q11 không sửa `video.py`). Mục 5 ngữ pháp → Task 2. Mục 6, 7 → Task 5. Mục 8 → Task 3, 4, 7 (đã sửa ở Task 1). Mục 9 → Task 6, 7. Mục 10 → Task 8. Mục 11 → Task 10. Mục 12 → Task 1 (đo xong). Mục 13 → test ở mọi task. Mục 14 không làm. Mục 15 rủi ro: tốc độ (đã đo), font (đã kiểm), mạng (Task 4/8), Chromium (Task 8), nội dung thầy cô (docs), mốc ước lượng (Task 3 cảnh báo).

**Quét chỗ trống:** không có "TBD/TODO". Hai chỗ cố ý chỉ dẫn "làm như commit mẫu" ở Task 10 (file luật, `mau-brief`, `xu-ly-loi`, `cau-lenh-mau`, `bat-dau-nhanh`, `README`) vì nội dung phụ thuộc văn bản hiện có của từng file; mỗi chỗ nêu rõ việc cần thêm và test khoá kết quả.

**Nhất quán kiểu:** `GiongInfo(mp3, giay, moc_cau, uoc_luong, nguon)` dùng đồng nhất ở Task 3, 4, 7, 8 (test Task 4 tạo qua `lay_giong`, Task 8 tạo tay đúng cùng trường); `CanhLich` có `moc_cau` (đã cộng 0,7) và `moc_cau_giong` (không cộng) dùng ở Task 3, 7 đúng nghĩa; `chup_canh(page, html, so_khung, fps, thu_muc, so_dau, ghi_log=None)` khớp Task 6/8; `ghep_video(thu_muc, cac_lich, cac_giong, phu_de, fps=FPS, run=...)` khớp Task 7/8; `du["moc"]` là thời gian cảnh, các file cảnh dùng nguyên giá trị đó; `thamSoTai` dùng `t < mốc đầu` (đã sửa so với bản nháp `<=`).

**Review Focus:** mỗi dòng có test trong task chủ: (1) `test_loi_mot_cau_nhieu_y` (Task 3); (2) `catDanhDau` Node + `test_trang_khong_chua_ma_chay_duoc` + `test_hostile_text_renders_as_text` (Task 5, 6); (3) các test rỗng/hỏng/dở dang (Task 4); (4) `test_folder_with_spaces_diacritics_and_apostrophe` (Task 7) và dựng thật `Bài 5 – Sulfur dioxide (thử)` (Task 9); (5) `test_edited_text_regenerates_only_that_scene`, `test_teacher_supplied_mp3_is_used_and_never_overwritten` (Task 4), `test_thu_muc_khung_tam_duoc_don` (Task 8).
