# Kế hoạch thực thi: Soạn giáo án tích hợp năng lực số và năng lực AI

> **Dành cho AI thực thi:** BẮT BUỘC dùng skill `superpowers:subagent-driven-development` (khuyến nghị) hoặc `superpowers:executing-plans` để làm từng task một. Các bước dùng ô đánh dấu `- [ ]`.

**Mục tiêu:** Thêm loại việc thứ 8 cho lớp Việt — soạn Kế hoạch bài dạy theo Công văn 5512 có tích hợp năng lực số và năng lực AI — và xuất ra file Word đúng thể thức thầy cô đang dùng.

**Kiến trúc:** AI viết một file nguồn `giao-an.md`. `tools/vi/giao_an.py` đọc file đó, kiểm cấu trúc 5512 và kiểm mã năng lực, rồi dựng `giao-an.docx` bằng `python-docx` trên tầng dùng chung `tools/vi/word_parts/`. Bảng mã năng lực chỉ có một nguồn duy nhất là `docs/vi/tro-ly/nang-luc-so-va-ai.md`; module kiểm mã đọc trực tiếp từ file đó nên không thể vênh.

**Công nghệ:** Python 3.10+ (thư viện chuẩn cho phần đọc và kiểm, `python-docx>=1.1.0` cho phần dựng Word), `unittest`.

**Spec:** `docs/vi/phat-trien/2026-09-13-giao-an-nls-ai-design.md`

**Phụ thuộc:** Làm **sau** gói đề thi `v6.3.2-vi.5`. Gói đó tạo ra `tools/vi/word_parts/inline.py`, `tools/vi/word_parts/base.py` và `tools/vi/requirements-vi.txt`. Nếu ba thứ đó chưa có thì dừng và báo, không tự tạo lại.

## Global Constraints

- Không sửa bất cứ gì trong `skills/`, `LICENSE`, `SPONSORS*.md`, và không sửa frontmatter của SKILL.md.
- Không bao giờ sửa, bỏ qua hay tìm cách "sửa chữa" `skills/ppt-master/scripts/attribution_guard.py`.
- Không sửa `requirements.txt` ở gốc repo và `skills/ppt-master/requirements.txt` — cả hai thuộc upstream.
- Không thêm thư viện mới; gói này dùng đúng `python-docx` mà gói vi.5 đã khai.
- Giáo án, đề thi và SGK của thầy cô đặt ở `projects/_giao-an/`; `projects/*` đã bị gitignore và không bao giờ được commit.
- Không copy bất kỳ file nào từ thư mục tài nguyên riêng của chủ repo vào repo, trừ nội dung khung năng lực đã được chủ repo cho phép công bố.
- Không tạo, không commit file `.env`.
- Mọi đường dẫn nhận từ dòng lệnh phải `.expanduser().resolve()` trước khi dùng.
- stdout của `giao_an.py` là **đúng một dòng JSON**; mọi tiến trình và log đi ra stderr.
- `error.step` chỉ thuộc `{input, parse, framework, docx, write, internal}`.
- Ghi chú nội bộ (`## CAN SOAT`) **không bao giờ** được in vào `giao-an.docx`; nó ra file `can-soat.md` riêng.
- File trong `docs/vi/tro-ly/` không được dùng liên kết Markdown.
- Test không được mở cửa sổ Word, PowerPoint hay trình duyệt.
- Chạy test bằng: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests`
- Commit message phải có dạng: dòng tiêu đề, một dòng trống, rồi trailer `Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>`. Kiểm lại bằng `git log -1 --format=%B` trước khi báo xong.
- Mọi chuỗi hiển thị cho thầy cô viết bằng tiếng Việt; tên hàm, tên biến, tên file viết bằng tiếng Anh.

---

### Task 1: Bảng mã năng lực và bộ đọc bảng mã

**Files:**
- Create: `docs/vi/tro-ly/nang-luc-so-va-ai.md`
- Create: `tools/vi/giao_an_parts/__init__.py`
- Create: `tools/vi/giao_an_parts/frameworks.py`
- Create: `tools/vi/tests/test_giao_an.py`

**Interfaces:**
- Consumes: không có.
- Produces:
  - `frameworks.NO_FRAMEWORK = "(chưa có khung mã cho môn này)"`, `frameworks.NO_CODE = "(chưa có mã)"`
  - `frameworks.FrameworkError(message, fix)` — thuộc tính `.message`, `.fix`
  - `frameworks.normalise(text: str) -> str` — bỏ dấu, hạ chữ thường, gộp khoảng trắng
  - `frameworks.Frameworks` — trường `nls: list[str]`, `ai: dict[str, list[str]]`, `index: dict[tuple[str, str], str]`; phương thức `framework_for(subject, grade, track="") -> str | None`, `ai_codes_for(subject, grade, track="") -> list[str]`
  - `frameworks.load(text: str) -> Frameworks`, `frameworks.load_file(path=DOC_PATH) -> Frameworks`
  - `frameworks.DOC_PATH` — đường dẫn tới file bảng mã

- [ ] **Step 1: Viết file bảng mã**

Tạo `docs/vi/tro-ly/nang-luc-so-va-ai.md`. **Khuôn máy đọc được, không được đổi**: mỗi mã là một dòng `` - `<mã>` — <mô tả> ``; mỗi khung AI là một mục h3 `### <môn> — <phạm vi> — khung `<tên khung>``. Dấu gạch dài là `—` (em dash), không phải `-`. Không dùng liên kết Markdown.

```markdown
# Khung năng lực số và năng lực AI

File dành cho AI. Đọc cùng docs/vi/tro-ly/giao-an.md.

Khung do Lương Hải Anh — 2Anh AI Education biên soạn, dựa trên Bảng mã Năng lực số của Bộ GD&ĐT (mức nâng cao NC1 cho THPT) và Phụ lục III & IV về triển khai giáo dục AI. Đây là nguồn mã duy nhất của bộ công cụ: `tools/vi/giao_an.py` đọc mã trực tiếp từ file này, nên sửa mã ở đây là sửa cho cả AI và cho bộ kiểm tra.

## Nguyên tắc tích hợp

- AI là công cụ hỗ trợ tư duy, không thay việc học. Học sinh phải là người viết câu lệnh và là người đánh giá kết quả.
- Mỗi hoạt động dùng AI để dự đoán đều phải có bước đối chứng: thí nghiệm thật, thí nghiệm ảo, hoặc tra cứu nguồn tin cậy.
- Phải có ít nhất một hoạt động yêu cầu học sinh chỉ ra chỗ AI nói sai hoặc nói thiếu.
- Năng lực số và năng lực AI phải nằm trong tiến trình dạy học, không chỉ nằm ở mục tiêu.
- Không đặt mã chỉ báo mới. Môn chưa có khung thì ghi `(chưa có khung mã cho môn này)` ở mục tiêu và `(chưa có mã)` ở hoạt động.

## Khung năng lực số (mọi môn)

- `1.1.NC1a` — Tìm kiếm và chọn lọc dữ liệu, thông tin, nội dung số phục vụ học tập.
- `1.1.NC1b` — Đánh giá độ tin cậy và độ chính xác của nguồn dữ liệu số.
- `1.2.NC1a` — Phân tích, so sánh, đối chiếu các nguồn dữ liệu khoa học số.
- `1.3.NC1a` — Tổ chức, phân loại, lưu trữ và truy xuất thông tin, dữ liệu trong môi trường số.
- `2.1.NC1a` — Tương tác, trao đổi ý kiến qua nền tảng học tập trực tuyến.
- `2.2.NC1a` — Chia sẻ học liệu số và làm việc nhóm trực tuyến an toàn.
- `3.1.NC1a` — Dùng phần mềm chuyên dụng tạo sơ đồ tư duy, báo cáo, infographic, hình ảnh 3D.
- `3.2.NC1a` — Chỉnh sửa, tích hợp nội dung số đa phương tiện.
- `4.1.NC1a` — Bảo vệ thiết bị, dữ liệu cá nhân, tuân thủ bản quyền và an toàn số.
- `5.1.NC1a` — Dùng thiết bị số, phần mềm mô phỏng, thí nghiệm ảo để giải quyết vấn đề học tập và thực tiễn.
- `5.3.NC1a` — Đổi mới cách học bằng cách áp dụng giải pháp công nghệ số.

## Khung năng lực AI theo môn

### Hoá học — lớp 10 — khung `AI-H10`

- `AI-H10` — Dùng AI trực quan hoá phân tử 3D theo VSEPR và lai hoá AO; mô phỏng tốc độ phản ứng theo Arrhenius; giải thích vai trò của dữ liệu huấn luyện; nhận diện orbital p, d, f và dự đoán góc liên kết.

### Hoá học — lớp 11 — khung `AI-H11`

- `AI-H11.1` — Khai thác và nhận diện: tra cứu PubChem, NIST; phân tích phổ IR để nhận diện nhóm chức, phổ MS để xác định phân tử khối.
- `AI-H11.2` — Mô hình hoá và dự đoán: dự đoán chiều chuyển dịch cân bằng theo Le Chatelier, hướng cộng theo Markovnikov, phản ứng tách theo Zaitsev, vị trí thế trên nhân thơm.
- `AI-H11.3` — Thiết kế và tối ưu hoá: dùng prompt engineering gợi ý sơ đồ tổng hợp hữu cơ, tối ưu thông số Haber-Bosch và Contact process, tối ưu quy trình STEM.
- `AI-H11.4` — Đánh giá và phản biện: kiểm chứng kết quả AI với lý thuyết và thực nghiệm, nhận diện rủi ro môi trường và tuân thủ đạo đức AI.

### Hoá học — lớp 12 — khung `AI-H12`

- `AI-H12` — Dùng AI tính suất điện động pin Galvani, dự đoán sản phẩm điện phân, mô tả cấu trúc phức chất; dự đoán tính chất polymer và giải pháp tái chế; viết prompt tra cứu hằng số.

### Hoá học — hệ chuyên — khung `AI-H-Chuyên`

- `AI-H-Chuyên` — Phân tích phổ NMR, MS, IR phức tạp; mô phỏng cơ chế SN1, SN2, E1, E2, cộng ái nhân, cộng ái điện tử; dùng Python và machine learning xử lý dữ liệu động hoá học, xác định bậc phản ứng.

## Môn chưa có khung mã AI

Môn nào không có mục khung ở trên thì coi như chưa có khung. Khi đó:

- Mục `### Nang luc AI` trong giáo án có dòng đầu là đúng chuỗi `- (chưa có khung mã cho môn này)`, các dòng sau viết hoạt động AI bằng lời, không có mã.
- Hoạt động AI trong tiến trình ghi `ai: (chưa có mã)`.
- Tuyệt đối không tự đặt mã mới kiểu `AI-P11` hay `AI-S12`. Đặt mã chỉ báo là việc của Bộ và của tổ chuyên môn.
- Tổ nào đã có khung riêng thì thầy cô gửi file khung; lúc đó thêm một mục khung mới vào file này theo đúng khuôn ở trên.

## Gợi ý hoạt động AI cho môn chưa có khung

- Toán: dùng AI sinh phản ví dụ rồi học sinh kiểm chứng bằng lập luận; dùng AI giải rồi tìm chỗ sai trong lời giải.
- Ngữ văn: dùng AI viết một đoạn theo yêu cầu rồi học sinh nhận xét giọng điệu, dẫn chứng, chỗ bịa; so bản AI với bản của mình.
- Lịch sử, Địa lí: dùng AI tóm tắt một nguồn rồi học sinh đối chiếu với tư liệu gốc, chỉ ra chỗ thiếu bối cảnh.
- Vật lí, Sinh học: dùng AI dự đoán kết quả thí nghiệm rồi làm thí nghiệm thật hoặc mô phỏng để kiểm chứng.
- Tin học: dùng AI sinh mã rồi học sinh chạy, tìm lỗi và giải thích vì sao sai.
- Ngoại ngữ: dùng AI sửa bài viết rồi học sinh giải thích từng chỗ sửa, giữ lại chỗ không đồng ý.
```

- [ ] **Step 2: Viết test trước**

Tạo `tools/vi/tests/test_giao_an.py`:

```python
"""Test cho lớp soạn giáo án tích hợp năng lực số và năng lực AI của bản Việt."""

import sys
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT / "tools" / "vi"))

from giao_an_parts import frameworks  # noqa: E402

NLS_EXPECTED = (
    "1.1.NC1a", "1.1.NC1b", "1.2.NC1a", "1.3.NC1a",
    "2.1.NC1a", "2.2.NC1a",
    "3.1.NC1a", "3.2.NC1a",
    "4.1.NC1a",
    "5.1.NC1a", "5.3.NC1a",
)


class FrameworkLoadTest(unittest.TestCase):
    def setUp(self):
        self.frameworks = frameworks.load_file()

    def test_document_is_the_only_source_of_codes(self):
        self.assertTrue(frameworks.DOC_PATH.is_file(), frameworks.DOC_PATH)
        text = frameworks.DOC_PATH.read_text(encoding="utf-8")
        for code in NLS_EXPECTED:
            self.assertIn(f"`{code}`", text)

    def test_every_digital_competence_code_is_read(self):
        self.assertEqual(tuple(self.frameworks.nls), NLS_EXPECTED)

    def test_four_chemistry_ai_frameworks_are_read(self):
        self.assertEqual(
            sorted(self.frameworks.ai),
            sorted(["AI-H10", "AI-H11", "AI-H12", "AI-H-Chuyên"]),
        )

    def test_grade_eleven_framework_has_four_codes(self):
        self.assertEqual(
            self.frameworks.ai["AI-H11"],
            ["AI-H11.1", "AI-H11.2", "AI-H11.3", "AI-H11.4"],
        )

    def test_framework_is_chosen_by_subject_and_grade(self):
        self.assertEqual(self.frameworks.framework_for("Hoá học", "11"), "AI-H11")
        self.assertEqual(self.frameworks.framework_for("Hoá học", "10"), "AI-H10")
        self.assertEqual(self.frameworks.framework_for("Hoá học", "12"), "AI-H12")

    def test_specialised_track_uses_its_own_framework(self):
        self.assertEqual(self.frameworks.framework_for("Hoá học", "11", "chuyen"), "AI-H-Chuyên")
        self.assertEqual(self.frameworks.framework_for("Hoá học", "10", "hệ chuyên"), "AI-H-Chuyên")

    def test_subject_name_matching_ignores_case_and_diacritics(self):
        self.assertEqual(self.frameworks.framework_for("HOA HOC", "11"), "AI-H11")
        self.assertEqual(self.frameworks.framework_for("  hoá   học ", "11"), "AI-H11")

    def test_subject_without_a_framework_returns_none(self):
        self.assertIsNone(self.frameworks.framework_for("Ngữ văn", "11"))
        self.assertEqual(self.frameworks.ai_codes_for("Ngữ văn", "11"), [])

    def test_unknown_grade_returns_none(self):
        self.assertIsNone(self.frameworks.framework_for("Hoá học", "9"))

    def test_load_rejects_a_document_without_codes(self):
        with self.assertRaises(frameworks.FrameworkError):
            frameworks.load("# Trống\n\n## Khung năng lực số (mọi môn)\n")

    def test_normalise_strips_diacritics_and_case(self):
        self.assertEqual(frameworks.normalise("Hoá  Học"), "hoa hoc")
        self.assertEqual(frameworks.normalise("Hệ Chuyên"), "he chuyen")


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 3: Chạy test để thấy nó fail**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k FrameworkLoadTest`
Kỳ vọng: FAIL với `ModuleNotFoundError: No module named 'giao_an_parts'`

- [ ] **Step 4: Tạo gói con**

Tạo `tools/vi/giao_an_parts/__init__.py` với đúng nội dung một dòng sau:

```python
"""Các phần của công cụ soạn giáo án tích hợp năng lực số và năng lực AI (lớp Việt)."""
```

- [ ] **Step 5: Viết `frameworks.py`**

Tạo `tools/vi/giao_an_parts/frameworks.py`:

```python
"""Đọc khung năng lực số và năng lực AI từ file tài liệu, và kiểm mã của một giáo án.

Nguồn mã duy nhất là docs/vi/tro-ly/nang-luc-so-va-ai.md; module này không khai lại mã nào.
Chỉ dùng thư viện chuẩn Python.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
DOC_PATH = REPO_ROOT / "docs" / "vi" / "tro-ly" / "nang-luc-so-va-ai.md"
NLS_HEADING = "## Khung năng lực số (mọi môn)"
AI_HEADING = "## Khung năng lực AI theo môn"
NO_FRAMEWORK = "(chưa có khung mã cho môn này)"
NO_CODE = "(chưa có mã)"
SPECIAL_TRACK = "hệ chuyên"
TRACK_ALIASES = ("chuyen", "he chuyen")
FIX_RELOAD = "Tải lại bản đầy đủ của bộ công cụ; đừng sửa khuôn của nang-luc-so-va-ai.md."

_CODE_RE = re.compile(r"^-\s+`([^`]+)`")
_FRAME_RE = re.compile(r"^###\s+(.+?)\s+—\s+(.+?)\s+—\s+khung\s+`([^`]+)`\s*$")


class FrameworkError(Exception):
    """Mã năng lực sai, hoặc không đọc được bảng mã."""

    def __init__(self, message: str, fix: str = "") -> None:
        super().__init__(message)
        self.message = message
        self.fix = fix


def normalise(text: str) -> str:
    """Bỏ dấu, hạ chữ thường, gộp khoảng trắng — để so tên môn không lệ thuộc cách viết."""
    decomposed = unicodedata.normalize("NFD", text)
    without_marks = "".join(ch for ch in decomposed if unicodedata.category(ch) != "Mn")
    return " ".join(without_marks.lower().split())


@dataclass
class Frameworks:
    nls: list[str] = field(default_factory=list)
    ai: dict[str, list[str]] = field(default_factory=dict)
    index: dict[tuple[str, str], str] = field(default_factory=dict)

    def framework_for(self, subject: str, grade: str, track: str = "") -> str | None:
        scope = SPECIAL_TRACK if normalise(track) in TRACK_ALIASES else f"lớp {grade}"
        return self.index.get((normalise(subject), normalise(scope)))

    def ai_codes_for(self, subject: str, grade: str, track: str = "") -> list[str]:
        name = self.framework_for(subject, grade, track)
        return list(self.ai.get(name, [])) if name else []


def load(text: str) -> Frameworks:
    frameworks = Frameworks()
    section = ""
    current = ""
    for raw in text.splitlines():
        line = raw.strip()
        if line.startswith("## "):
            section = line
            current = ""
            continue
        if section == NLS_HEADING:
            match = _CODE_RE.match(line)
            if match:
                frameworks.nls.append(match.group(1))
            continue
        if section == AI_HEADING:
            frame = _FRAME_RE.match(line)
            if frame:
                subject, scope, name = frame.groups()
                current = name
                frameworks.ai.setdefault(name, [])
                frameworks.index[(normalise(subject), normalise(scope))] = name
                continue
            match = _CODE_RE.match(line)
            if match and current:
                frameworks.ai[current].append(match.group(1))
    if not frameworks.nls:
        raise FrameworkError(f"Không đọc được mã năng lực số nào trong {DOC_PATH.name}", FIX_RELOAD)
    if not frameworks.ai:
        raise FrameworkError(f"Không đọc được khung năng lực AI nào trong {DOC_PATH.name}", FIX_RELOAD)
    return frameworks


def load_file(path: Path = DOC_PATH) -> Frameworks:
    try:
        return load(path.read_text(encoding="utf-8"))
    except OSError as exc:
        raise FrameworkError(f"Không đọc được {path}: {exc}", FIX_RELOAD) from None
```

- [ ] **Step 6: Chạy test để thấy nó pass**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k FrameworkLoadTest`
Kỳ vọng: PASS, 11 test. Nếu `test_every_digital_competence_code_is_read` fail, kiểm lại dấu gạch trong file bảng mã phải là `—` chứ không phải `-`.

- [ ] **Step 7: Chạy cả bộ test**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests`
Kỳ vọng: OK.

- [ ] **Step 8: Commit**

```bash
git add docs/vi/tro-ly/nang-luc-so-va-ai.md tools/vi/giao_an_parts/__init__.py tools/vi/giao_an_parts/frameworks.py tools/vi/tests/test_giao_an.py
git commit -m "feat(vi): read the digital and AI competence frameworks from one document"
```

---

### Task 2: Đọc file nguồn `giao-an.md` (`parse.py`)

**Files:**
- Create: `tools/vi/giao_an_parts/parse.py`
- Modify: `tools/vi/tests/test_giao_an.py` (thêm `ParseTest`, `LessonMathTest`)

**Interfaces:**
- Consumes: `frameworks.NO_CODE`, `frameworks.NO_FRAMEWORK`.
- Produces:
  - `parse.ParseError(line_no, message)` — thuộc tính `.line_no`, `.message`; `str(exc)` mở đầu bằng `"Dòng <n>: "`
  - `parse.Activity` — trường `title, line_no, minutes, muc_tieu, noi_dung, san_pham, chuyen_giao, thuc_hien, bao_cao, ket_luan, nls, ai` (các trường văn bản là `list[str]`)
  - `parse.Worksheet` — `title: str`, `items: list[str]`
  - `parse.Rubric` — `title: str`, `levels: list[str]` (đúng ba phần tử)
  - `parse.Lesson` — `meta, knowledge, general, specific, nls_lines, ai_lines, qualities, teacher_tools, student_tools, activities, worksheets, rubrics, review_notes, warnings, ai_no_framework`; thuộc tính `nls_codes`, `ai_codes`, `minutes`, `periods`
  - `parse.parse_lesson(text: str) -> Lesson`, `parse.parse_file(path: Path) -> Lesson`

- [ ] **Step 1: Viết test trước**

Thêm vào `tools/vi/tests/test_giao_an.py`, trước khối `if __name__`:

```python
VALID_SOURCE = """---
school: TRƯỜNG THPT CHUYÊN NGUYỄN TRÃI
department: TỔ HOÁ HỌC
teacher: Lương Hải Anh
subject: Hoá học
grade: 11
lesson: Bài 5. Ammonia và một số hợp chất ammonium
periods: 2
week: 5-6
period_numbers: 9-10
school_year: 2026-2027
---

## MUC TIEU

### Kien thuc
- Trình bày được tính chất hoá học của NH~3~ và muối ammonium.

### Nang luc chung
- Tự chủ và tự học: đọc SGK và tài liệu số để rút ra kết luận.

### Nang luc dac thu
- Nhận thức hoá học: giải thích tính base yếu của NH~3~.

### Nang luc so
- 1.1.NC1a — Tìm kiếm, chọn lọc dữ liệu về ứng dụng của ammonia trong sản xuất phân bón.
- 5.1.NC1a — Dùng thí nghiệm ảo PhET kiểm chứng chuyển dịch cân bằng.

### Nang luc AI
- AI-H11.2 — Viết prompt để AI dự đoán chiều chuyển dịch cân bằng khi tăng áp suất.
- AI-H11.4 — Đối chiếu kết quả AI với thí nghiệm, chỉ ra chỗ AI nói sai.

### Pham chat
- Trung thực: ghi đúng số liệu thí nghiệm, kể cả khi lệch dự đoán.

## THIET BI

### Giao vien
- Bộ dụng cụ điều chế NH~3~, máy chiếu, phiếu học tập 1.

### Hoc sinh
- SGK, vở, điện thoại có mạng theo nhóm.

## TIEN TRINH

### Mở đầu
thoi-luong: 10
muc-tieu: Tạo tình huống có vấn đề về mùi khai của ammonia.
noi-dung: Học sinh quan sát video và nêu dự đoán.
san-pham: Câu trả lời của học sinh.
chuyen-giao: Giáo viên chiếu video và nêu câu hỏi.
thuc-hien: Học sinh thảo luận cặp đôi trong 3 phút.
bao-cao: Hai nhóm trình bày dự đoán.
ket-luan: Giáo viên chốt vấn đề của bài học.

### Hình thành kiến thức — cân bằng tổng hợp ammonia
thoi-luong: 35
muc-tieu: Dự đoán và kiểm chứng chiều chuyển dịch cân bằng.
noi-dung: Học sinh viết prompt cho AI dự đoán chiều chuyển dịch khi tăng áp suất.
noi-dung: Sau đó dùng PhET để kiểm chứng và hoàn thành Phiếu học tập 1.
san-pham: Đoạn chat với AI và bảng số liệu trong phiếu học tập 1.
nls: 1.1.NC1a, 5.1.NC1a
ai: AI-H11.2, AI-H11.4
chuyen-giao: Giáo viên giao mẫu prompt và địa chỉ mô phỏng PhET.
thuc-hien: Học sinh làm theo nhóm 4, giáo viên quan sát.
bao-cao: Mỗi nhóm nêu một chỗ AI trả lời chưa đúng.
ket-luan: Giáo viên chuẩn hoá nguyên lý Le Chatelier.

### Luyện tập
thoi-luong: 30
muc-tieu: Vận dụng tính chất của muối ammonium.
noi-dung: Học sinh làm bài tập trong phiếu học tập 1.
san-pham: Bài giải trong phiếu.
chuyen-giao: Giáo viên phát phiếu.
thuc-hien: Học sinh làm cá nhân.
bao-cao: Ba học sinh lên bảng.
ket-luan: Giáo viên nhận xét và chữa bài.

### Vận dụng
thoi-luong: 15
muc-tieu: Liên hệ sản xuất phân bón ở địa phương.
noi-dung: Học sinh đề xuất cách bảo quản phân đạm.
san-pham: Đoạn trả lời ngắn.
chuyen-giao: Giáo viên giao nhiệm vụ về nhà.
thuc-hien: Học sinh làm ở nhà.
bao-cao: Nộp qua lớp học trực tuyến.
ket-luan: Giáo viên nhận xét ở tiết sau.

## PHIEU HOC TAP

### Phiếu học tập 1
- Câu 1. Viết phương trình tổng hợp NH~3~ từ N~2~ và H~2~.
- Câu 2. Dự đoán chiều chuyển dịch cân bằng khi tăng áp suất.

## RUBRIC

### Kĩ năng viết prompt
- Mức 1: Viết được câu lệnh nhưng thiếu dữ kiện.
- Mức 2: Viết được câu lệnh đủ dữ kiện.
- Mức 3: Viết được câu lệnh đủ dữ kiện và có ràng buộc rõ ràng.

### Kĩ năng kiểm chứng và phản biện AI
- Mức 1: Chấp nhận kết quả AI mà không kiểm chứng.
- Mức 2: Kiểm chứng được bằng thí nghiệm ảo.
- Mức 3: Kiểm chứng và chỉ ra được chỗ AI nói sai kèm lý do.

## CAN SOAT
- Thầy cô xác nhận giúp em địa chỉ mô phỏng PhET dùng được ở phòng máy của trường.
"""


def source_without(prefix: str) -> str:
    kept = [line for line in VALID_SOURCE.splitlines() if not line.startswith(prefix)]
    return "\n".join(kept) + "\n"


class ParseTest(unittest.TestCase):
    def test_valid_source_reads_every_section(self):
        lesson = parse.parse_lesson(VALID_SOURCE)
        self.assertEqual(lesson.meta["subject"], "Hoá học")
        self.assertEqual(lesson.meta["grade"], "11")
        self.assertEqual(len(lesson.activities), 4)
        self.assertEqual(len(lesson.worksheets), 1)
        self.assertEqual(len(lesson.rubrics), 2)
        self.assertEqual(len(lesson.review_notes), 1)
        self.assertFalse(lesson.ai_no_framework)

    def test_codes_are_read_from_the_objective_lines(self):
        lesson = parse.parse_lesson(VALID_SOURCE)
        self.assertEqual(lesson.nls_codes, ["1.1.NC1a", "5.1.NC1a"])
        self.assertEqual(lesson.ai_codes, ["AI-H11.2", "AI-H11.4"])

    def test_repeated_key_becomes_two_paragraphs(self):
        activity = parse.parse_lesson(VALID_SOURCE).activities[1]
        self.assertEqual(len(activity.noi_dung), 2)
        self.assertIn("PhET", activity.noi_dung[1])

    def test_activity_codes_are_split_on_commas(self):
        activity = parse.parse_lesson(VALID_SOURCE).activities[1]
        self.assertEqual(activity.nls, ["1.1.NC1a", "5.1.NC1a"])
        self.assertEqual(activity.ai, ["AI-H11.2", "AI-H11.4"])

    def test_rubric_levels_are_stripped_of_their_prefix(self):
        rubric = parse.parse_lesson(VALID_SOURCE).rubrics[0]
        self.assertEqual(len(rubric.levels), 3)
        self.assertTrue(rubric.levels[0].startswith("Viết được câu lệnh nhưng"))

    def test_missing_meta_key_names_it(self):
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_lesson(source_without("subject:"))
        self.assertIn("subject", caught.exception.message)

    def test_grade_must_be_digits(self):
        broken = VALID_SOURCE.replace("grade: 11", "grade: mười một")
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_lesson(broken)
        self.assertIn("grade", caught.exception.message)

    def test_periods_must_be_a_positive_number(self):
        broken = VALID_SOURCE.replace("periods: 2", "periods: 0")
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_lesson(broken)
        self.assertIn("periods", caught.exception.message)

    def test_activity_missing_any_of_the_eight_keys_fails(self):
        """Bốn thành tố 5512 và bốn bước tổ chức là thứ người kiểm tra soi trước tiên."""
        for key in ("thoi-luong", "muc-tieu", "noi-dung", "san-pham",
                    "chuyen-giao", "thuc-hien", "bao-cao", "ket-luan"):
            with self.subTest(key=key):
                dropped = False
                kept = []
                for line in VALID_SOURCE.splitlines():
                    if not dropped and line.startswith(f"{key}:"):
                        dropped = True
                        continue
                    kept.append(line)
                self.assertTrue(dropped, f"không tìm thấy dòng {key}:")
                with self.assertRaises(parse.ParseError) as caught:
                    parse.parse_lesson("\n".join(kept) + "\n")
                self.assertIn(key, caught.exception.message)

    def test_activity_missing_muc_tieu_names_the_key_and_line(self):
        broken = VALID_SOURCE.replace("muc-tieu: Tạo tình huống có vấn đề về mùi khai của ammonia.\n", "", 1)
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_lesson(broken)
        self.assertIn("muc-tieu", caught.exception.message)
        self.assertGreater(caught.exception.line_no, 1)

    def test_activity_time_must_be_positive(self):
        broken = VALID_SOURCE.replace("thoi-luong: 10", "thoi-luong: 0", 1)
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_lesson(broken)
        self.assertIn("thoi-luong", caught.exception.message)

    def test_unknown_activity_key_fails(self):
        broken = VALID_SOURCE.replace("bao-cao: Hai nhóm trình bày dự đoán.", "ghi-chu: abc", 1)
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_lesson(broken)
        self.assertIn("ghi-chu", caught.exception.message)

    def test_sections_out_of_order_fail(self):
        broken = VALID_SOURCE.replace("## THIET BI", "## RUBRIC-TAM", 1)
        with self.assertRaises(parse.ParseError):
            parse.parse_lesson(broken)

    def test_duplicate_section_fails(self):
        broken = VALID_SOURCE + "\n## RUBRIC\n\n### Thêm\n- Mức 1: a\n- Mức 2: b\n- Mức 3: c\n"
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_lesson(broken)
        self.assertIn("hai lần", caught.exception.message)

    def test_review_section_must_be_last(self):
        broken = VALID_SOURCE.replace("## RUBRIC", "## CAN SOAT\n- ghi chú\n\n## RUBRIC", 1)
        with self.assertRaises(parse.ParseError):
            parse.parse_lesson(broken)

    def test_objective_needs_all_six_subsections_in_order(self):
        broken = VALID_SOURCE.replace("### Pham chat\n- Trung thực: ghi đúng số liệu thí nghiệm, kể cả khi lệch dự đoán.\n", "")
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_lesson(broken)
        self.assertIn("Pham chat", caught.exception.message)

    def test_empty_digital_competence_section_fails(self):
        broken = source_without("- 1.1.NC1a").replace("- 5.1.NC1a — Dùng thí nghiệm ảo PhET kiểm chứng chuyển dịch cân bằng.\n", "")
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_lesson(broken)
        self.assertIn("Nang luc so", caught.exception.message)

    def test_no_activity_carrying_a_digital_code_fails(self):
        broken = VALID_SOURCE.replace("nls: 1.1.NC1a, 5.1.NC1a\n", "")
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_lesson(broken)
        self.assertIn("nls", caught.exception.message)
        self.assertIn("tiến trình", caught.exception.message)

    def test_no_activity_carrying_an_ai_code_fails(self):
        broken = VALID_SOURCE.replace("ai: AI-H11.2, AI-H11.4\n", "")
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_lesson(broken)
        self.assertIn("ai", caught.exception.message)
        self.assertIn("tiến trình", caught.exception.message)

    def test_exam_with_no_activity_fails(self):
        broken = VALID_SOURCE.split("## TIEN TRINH")[0] + "## TIEN TRINH\n\n" + VALID_SOURCE.split("## PHIEU HOC TAP")[1]
        with self.assertRaises(parse.ParseError):
            parse.parse_lesson(broken)

    def test_rubric_needs_at_least_two_criteria(self):
        broken = VALID_SOURCE.split("### Kĩ năng kiểm chứng và phản biện AI")[0] + "\n## CAN SOAT\n- ghi chú\n"
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_lesson(broken)
        self.assertIn("hai tiêu chí", caught.exception.message)

    def test_rubric_criterion_needs_exactly_three_levels(self):
        broken = VALID_SOURCE.replace("- Mức 3: Viết được câu lệnh đủ dữ kiện và có ràng buộc rõ ràng.\n", "")
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_lesson(broken)
        self.assertIn("ba mức", caught.exception.message)

    def test_rubric_levels_must_be_in_order(self):
        broken = VALID_SOURCE.replace("- Mức 1: Viết được câu lệnh nhưng thiếu dữ kiện.", "- Mức 2: Viết được câu lệnh nhưng thiếu dữ kiện.", 1)
        with self.assertRaises(parse.ParseError) as caught:
            parse.parse_lesson(broken)
        self.assertIn("Mức 1:", caught.exception.message)

    def test_subject_without_a_framework_declares_the_marker(self):
        broken = VALID_SOURCE.replace("subject: Hoá học", "subject: Ngữ văn")
        broken = broken.replace(
            "- AI-H11.2 — Viết prompt để AI dự đoán chiều chuyển dịch cân bằng khi tăng áp suất.\n"
            "- AI-H11.4 — Đối chiếu kết quả AI với thí nghiệm, chỉ ra chỗ AI nói sai.\n",
            f"- {frameworks.NO_FRAMEWORK}\n- Học sinh nhận xét đoạn văn do AI viết.\n",
        )
        broken = broken.replace("ai: AI-H11.2, AI-H11.4", f"ai: {frameworks.NO_CODE}")
        lesson = parse.parse_lesson(broken)
        self.assertTrue(lesson.ai_no_framework)
        self.assertEqual(lesson.ai_codes, [])
        self.assertEqual(lesson.activities[1].ai, [frameworks.NO_CODE])


class LessonMathTest(unittest.TestCase):
    def test_total_time_matches_two_periods(self):
        lesson = parse.parse_lesson(VALID_SOURCE)
        self.assertEqual(lesson.minutes, 90)
        self.assertEqual(lesson.periods, 2)
        self.assertEqual([warning for warning in lesson.warnings if "thời lượng" in warning], [])

    def test_time_mismatch_is_a_warning(self):
        broken = VALID_SOURCE.replace("thoi-luong: 30", "thoi-luong: 20", 1)
        lesson = parse.parse_lesson(broken)
        self.assertTrue(any("thời lượng" in warning for warning in lesson.warnings))

    def test_missing_worksheet_section_is_a_warning(self):
        broken = VALID_SOURCE.split("## PHIEU HOC TAP")[0] + "## RUBRIC" + VALID_SOURCE.split("## RUBRIC")[1]
        lesson = parse.parse_lesson(broken)
        self.assertTrue(any("PHIEU HOC TAP" in warning for warning in lesson.warnings))

    def test_worksheet_mentioned_but_not_defined_is_a_warning(self):
        broken = VALID_SOURCE.replace("### Phiếu học tập 1", "### Phiếu học tập 2")
        lesson = parse.parse_lesson(broken)
        self.assertTrue(any("Phiếu học tập 1" in warning for warning in lesson.warnings))
```

Đổi dòng import ở đầu file thành:

```python
from giao_an_parts import frameworks, parse  # noqa: E402
```

- [ ] **Step 2: Chạy test để thấy nó fail**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k ParseTest`
Kỳ vọng: FAIL với `ImportError: cannot import name 'parse'`

- [ ] **Step 3: Viết `parse.py`**

Tạo `tools/vi/giao_an_parts/parse.py`:

```python
"""Đọc file nguồn giao-an.md thành cấu trúc, và kiểm cấu trúc theo Công văn 5512.

Ngữ pháp của giao-an.md nằm trong docs/vi/tro-ly/giao-an.md.
Phần kiểm mã năng lực nằm ở frameworks.py, không ở đây.
Chỉ dùng thư viện chuẩn Python.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from .frameworks import NO_FRAMEWORK

OBJECTIVES = "## MUC TIEU"
EQUIPMENT = "## THIET BI"
PROCESS = "## TIEN TRINH"
WORKSHEETS = "## PHIEU HOC TAP"
RUBRIC = "## RUBRIC"
REVIEW = "## CAN SOAT"
SECTIONS = (OBJECTIVES, EQUIPMENT, PROCESS, WORKSHEETS, RUBRIC, REVIEW)
REQUIRED_SECTIONS = (OBJECTIVES, EQUIPMENT, PROCESS, RUBRIC)
OBJECTIVE_PARTS = (
    "Kien thuc",
    "Nang luc chung",
    "Nang luc dac thu",
    "Nang luc so",
    "Nang luc AI",
    "Pham chat",
)
EQUIPMENT_PARTS = ("Giao vien", "Hoc sinh")
ACTIVITY_KEYS = (
    "thoi-luong",
    "muc-tieu",
    "noi-dung",
    "san-pham",
    "chuyen-giao",
    "thuc-hien",
    "bao-cao",
    "ket-luan",
)
ACTIVITY_OPTIONAL = ("nls", "ai")
LEVEL_PREFIXES = ("Mức 1:", "Mức 2:", "Mức 3:")
META_REQUIRED = ("school", "subject", "grade", "lesson", "periods")
META_OPTIONAL = ("department", "teacher", "track", "week", "period_numbers", "school_year")
MINUTES_PER_PERIOD = 45

_CODE_LINE_RE = re.compile(r"^(\S+)\s+—\s+")
_WORKSHEET_MENTION_RE = re.compile(r"(?:[Pp]hiếu học tập|PHT)\s*(\d+)")
_NUMBER_RE = re.compile(r"(\d+)")


class ParseError(Exception):
    """Lỗi cấu trúc trong giao-an.md; luôn kèm số dòng để AI sửa đúng chỗ."""

    def __init__(self, line_no: int, message: str) -> None:
        super().__init__(f"Dòng {line_no}: {message}")
        self.line_no = line_no
        self.message = message


@dataclass
class Block:
    section: str
    title: str
    line_no: int
    body: list[tuple[int, str]] = field(default_factory=list)


@dataclass
class Activity:
    title: str
    line_no: int
    minutes: int
    muc_tieu: list[str]
    noi_dung: list[str]
    san_pham: list[str]
    chuyen_giao: list[str]
    thuc_hien: list[str]
    bao_cao: list[str]
    ket_luan: list[str]
    nls: list[str]
    ai: list[str]


@dataclass
class Worksheet:
    title: str
    items: list[str]


@dataclass
class Rubric:
    title: str
    levels: list[str]


@dataclass
class Lesson:
    meta: dict[str, str]
    knowledge: list[str]
    general: list[str]
    specific: list[str]
    nls_lines: list[str]
    ai_lines: list[str]
    qualities: list[str]
    teacher_tools: list[str]
    student_tools: list[str]
    activities: list[Activity]
    worksheets: list[Worksheet]
    rubrics: list[Rubric]
    review_notes: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    ai_no_framework: bool = False

    @property
    def nls_codes(self) -> list[str]:
        return _codes_from_lines(self.nls_lines)

    @property
    def ai_codes(self) -> list[str]:
        if self.ai_no_framework:
            return []
        return _codes_from_lines(self.ai_lines)

    @property
    def minutes(self) -> int:
        return sum(activity.minutes for activity in self.activities)

    @property
    def periods(self) -> int:
        return int(self.meta["periods"])


def _codes_from_lines(lines: list[str]) -> list[str]:
    codes: list[str] = []
    for line in lines:
        match = _CODE_LINE_RE.match(line)
        if match and match.group(1) not in codes:
            codes.append(match.group(1))
    return codes


def split_codes(values: list[str]) -> list[str]:
    codes: list[str] = []
    for value in values:
        pieces = value.split(",") if "," in value and not value.startswith("(") else [value]
        for piece in pieces:
            code = piece.strip()
            if code and code not in codes:
                codes.append(code)
    return codes


def parse_meta(lines: list[str]) -> tuple[dict[str, str], int]:
    if not lines or lines[0].strip() != "---":
        raise ParseError(1, "file phải mở đầu bằng một dòng '---' rồi tới khối thông tin bài dạy")
    allowed = META_REQUIRED + META_OPTIONAL
    meta: dict[str, str] = {}
    for index in range(1, len(lines)):
        line_no = index + 1
        stripped = lines[index].strip()
        if stripped == "---":
            missing = [key for key in META_REQUIRED if not meta.get(key)]
            if missing:
                raise ParseError(line_no, "khối thông tin bài dạy thiếu: " + ", ".join(missing))
            if not meta["grade"].isdigit():
                raise ParseError(line_no, f"'grade' phải là chữ số, gặp {meta['grade']!r}")
            if not meta["periods"].isdigit() or int(meta["periods"]) <= 0:
                raise ParseError(line_no, f"'periods' phải là số tiết lớn hơn 0, gặp {meta['periods']!r}")
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
                f"khoá {key!r} không dùng trong khối thông tin bài dạy; các khoá hợp lệ: "
                + ", ".join(allowed),
            )
        meta[key] = value.strip()
    raise ParseError(len(lines), "khối thông tin bài dạy chưa được đóng bằng dòng '---'")


def split_blocks(lines: list[str], start: int) -> tuple[list[Block], list[str], dict[str, int]]:
    blocks: list[Block] = []
    review_notes: list[str] = []
    section_lines: dict[str, int] = {}
    seen: list[str] = []
    section = ""
    for index in range(start, len(lines)):
        line_no = index + 1
        stripped = lines[index].strip()
        if not stripped:
            continue
        if stripped.startswith("## "):
            if stripped not in SECTIONS:
                raise ParseError(
                    line_no, f"mục {stripped!r} không hợp lệ; chỉ dùng: " + ", ".join(SECTIONS)
                )
            if stripped in seen:
                raise ParseError(line_no, f"mục {stripped!r} xuất hiện hai lần")
            if seen and SECTIONS.index(stripped) < SECTIONS.index(seen[-1]):
                raise ParseError(line_no, "các mục phải theo thứ tự: " + ", ".join(SECTIONS))
            seen.append(stripped)
            section = stripped
            section_lines[stripped] = line_no
            continue
        if not section:
            raise ParseError(
                line_no, f"dòng {stripped!r} nằm ngoài mục nào; mục đầu tiên phải là {OBJECTIVES!r}"
            )
        if section == REVIEW:
            if not stripped.startswith("- "):
                raise ParseError(
                    line_no, f"dòng trong {REVIEW!r} phải bắt đầu bằng '- ', gặp {stripped!r}"
                )
            review_notes.append(stripped[2:].strip())
            continue
        if stripped.startswith("### "):
            blocks.append(Block(section=section, title=stripped[4:].strip(), line_no=line_no))
            continue
        if not blocks or blocks[-1].section != section:
            raise ParseError(
                line_no,
                f"dòng {stripped!r} nằm ngoài mục con nào; mỗi mục con mở đầu bằng '### '",
            )
        blocks[-1].body.append((line_no, stripped))
    missing = [name for name in REQUIRED_SECTIONS if name not in seen]
    if missing:
        raise ParseError(max(len(lines), 1), "giáo án thiếu mục: " + ", ".join(missing))
    return blocks, review_notes, section_lines


def bullets(block: Block, *, required: bool = True) -> list[str]:
    items: list[str] = []
    for line_no, text in block.body:
        if not text.startswith("- "):
            raise ParseError(
                line_no, f"dòng trong mục {block.title!r} phải bắt đầu bằng '- ', gặp {text!r}"
            )
        items.append(text[2:].strip())
    if required and not items:
        raise ParseError(block.line_no, f"mục {block.title!r} không có dòng nào")
    return items


def build_activity(block: Block) -> Activity:
    values: dict[str, list[str]] = {}
    first_line: dict[str, int] = {}
    for line_no, text in block.body:
        key, separator, value = text.partition(":")
        if not separator:
            raise ParseError(line_no, f"dòng phải dạng 'khoá: giá trị', gặp {text!r}")
        key = key.strip()
        if key not in ACTIVITY_KEYS + ACTIVITY_OPTIONAL:
            raise ParseError(
                line_no,
                f"khoá {key!r} không dùng được trong hoạt động; các khoá hợp lệ: "
                + ", ".join(ACTIVITY_KEYS + ACTIVITY_OPTIONAL),
            )
        if not value.strip():
            raise ParseError(line_no, f"khoá {key!r} không có nội dung")
        values.setdefault(key, []).append(value.strip())
        first_line.setdefault(key, line_no)
    missing = [key for key in ACTIVITY_KEYS if key not in values]
    if missing:
        raise ParseError(
            block.line_no, f"hoạt động {block.title!r} thiếu khoá: " + ", ".join(missing)
        )
    minutes = values["thoi-luong"][0]
    if not minutes.isdigit() or int(minutes) <= 0:
        raise ParseError(
            first_line["thoi-luong"],
            f"'thoi-luong' phải là số phút lớn hơn 0, gặp {minutes!r}",
        )
    return Activity(
        title=block.title,
        line_no=block.line_no,
        minutes=int(minutes),
        muc_tieu=values["muc-tieu"],
        noi_dung=values["noi-dung"],
        san_pham=values["san-pham"],
        chuyen_giao=values["chuyen-giao"],
        thuc_hien=values["thuc-hien"],
        bao_cao=values["bao-cao"],
        ket_luan=values["ket-luan"],
        nls=split_codes(values.get("nls", [])),
        ai=split_codes(values.get("ai", [])),
    )


def build_rubric(block: Block) -> Rubric:
    items = bullets(block)
    if len(items) != 3:
        raise ParseError(
            block.line_no,
            f"tiêu chí rubric {block.title!r} phải có đúng ba mức, gặp {len(items)}",
        )
    levels = []
    for item, prefix in zip(items, LEVEL_PREFIXES):
        if not item.startswith(prefix):
            raise ParseError(
                block.line_no,
                f"các dòng của tiêu chí {block.title!r} phải mở đầu lần lượt bằng "
                + ", ".join(repr(value) for value in LEVEL_PREFIXES),
            )
        levels.append(item[len(prefix):].strip())
    return Rubric(title=block.title, levels=levels)


def _named_blocks(blocks: list[Block], section: str, expected: tuple[str, ...], line_no: int) -> dict[str, Block]:
    found = [block for block in blocks if block.section == section]
    if [block.title for block in found] != list(expected):
        raise ParseError(
            line_no,
            f"mục {section!r} phải có đúng các mục con theo thứ tự: " + ", ".join(expected),
        )
    return {block.title: block for block in found}


def _worksheet_warnings(activities: list[Activity], worksheets: list[Worksheet]) -> list[str]:
    warnings: list[str] = []
    if not worksheets:
        warnings.append(f"Giáo án chưa có mục {WORKSHEETS!r}.")
        return warnings
    defined = set()
    for sheet in worksheets:
        match = _NUMBER_RE.search(sheet.title)
        if match:
            defined.add(match.group(1))
    mentioned = set()
    for activity in activities:
        texts = (
            activity.noi_dung
            + activity.san_pham
            + activity.chuyen_giao
            + activity.thuc_hien
            + activity.bao_cao
        )
        for text in texts:
            for match in _WORKSHEET_MENTION_RE.finditer(text):
                mentioned.add(match.group(1))
    for number in sorted(mentioned - defined):
        warnings.append(
            f"Hoạt động nhắc Phiếu học tập {number} nhưng mục {WORKSHEETS!r} không có phiếu đó."
        )
    return warnings


def parse_lesson(text: str) -> Lesson:
    lines = text.splitlines()
    meta, start = parse_meta(lines)
    blocks, review_notes, section_lines = split_blocks(lines, start)

    objectives = _named_blocks(blocks, OBJECTIVES, OBJECTIVE_PARTS, section_lines[OBJECTIVES])
    equipment = _named_blocks(blocks, EQUIPMENT, EQUIPMENT_PARTS, section_lines[EQUIPMENT])
    ai_lines = bullets(objectives["Nang luc AI"])
    ai_no_framework = ai_lines[0] == NO_FRAMEWORK

    activities = [build_activity(block) for block in blocks if block.section == PROCESS]
    if not activities:
        raise ParseError(section_lines[PROCESS], f"mục {PROCESS!r} không có hoạt động nào")

    worksheets = [
        Worksheet(title=block.title, items=bullets(block))
        for block in blocks
        if block.section == WORKSHEETS
    ]
    rubrics = [build_rubric(block) for block in blocks if block.section == RUBRIC]
    if len(rubrics) < 2:
        raise ParseError(
            section_lines[RUBRIC], f"mục {RUBRIC!r} phải có ít nhất hai tiêu chí, gặp {len(rubrics)}"
        )

    if not any(activity.nls for activity in activities):
        raise ParseError(
            section_lines[PROCESS],
            "không hoạt động nào có dòng 'nls:'; năng lực số phải nằm trong tiến trình dạy học, "
            "không chỉ ở mục tiêu",
        )
    if not any(activity.ai for activity in activities):
        raise ParseError(
            section_lines[PROCESS],
            "không hoạt động nào có dòng 'ai:'; năng lực AI phải nằm trong tiến trình dạy học, "
            "không chỉ ở mục tiêu",
        )

    lesson = Lesson(
        meta=meta,
        knowledge=bullets(objectives["Kien thuc"]),
        general=bullets(objectives["Nang luc chung"]),
        specific=bullets(objectives["Nang luc dac thu"]),
        nls_lines=bullets(objectives["Nang luc so"]),
        ai_lines=ai_lines,
        qualities=bullets(objectives["Pham chat"]),
        teacher_tools=bullets(equipment["Giao vien"]),
        student_tools=bullets(equipment["Hoc sinh"]),
        activities=activities,
        worksheets=worksheets,
        rubrics=rubrics,
        review_notes=review_notes,
        ai_no_framework=ai_no_framework,
    )

    expected_minutes = lesson.periods * MINUTES_PER_PERIOD
    if lesson.minutes != expected_minutes:
        lesson.warnings.append(
            f"Tổng thời lượng các hoạt động là {lesson.minutes} phút, khác "
            f"{lesson.periods} tiết × {MINUTES_PER_PERIOD} = {expected_minutes} phút."
        )
    lesson.warnings.extend(_worksheet_warnings(activities, worksheets))
    return lesson


def parse_file(path: Path) -> Lesson:
    return parse_lesson(path.read_text(encoding="utf-8-sig"))
```

- [ ] **Step 4: Chạy test để thấy nó pass**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k ParseTest`
rồi `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k LessonMathTest`
Kỳ vọng: PASS cả hai.

- [ ] **Step 5: Chạy cả bộ test**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests`
Kỳ vọng: OK.

- [ ] **Step 6: Commit**

```bash
git add tools/vi/giao_an_parts/parse.py tools/vi/tests/test_giao_an.py
git commit -m "feat(vi): read the lesson plan source and enforce the 5512 structure"
```

---

### Task 3: Kiểm mã năng lực (`frameworks.validate`)

**Files:**
- Modify: `tools/vi/giao_an_parts/frameworks.py` (thêm `validate`)
- Modify: `tools/vi/tests/test_giao_an.py` (thêm `FrameworkValidateTest`)

**Interfaces:**
- Consumes: `parse.Lesson` (kiểu vịt: chỉ dùng `meta`, `nls_codes`, `ai_codes`, `ai_no_framework`, `activities`).
- Produces: `frameworks.validate(lesson, frameworks: Frameworks) -> None`, raise `FrameworkError`.

- [ ] **Step 1: Viết test trước**

Thêm vào `tools/vi/tests/test_giao_an.py`:

```python
class FrameworkValidateTest(unittest.TestCase):
    def setUp(self):
        self.frameworks = frameworks.load_file()

    def lesson_from(self, text: str):
        return parse.parse_lesson(text)

    def test_valid_lesson_passes(self):
        frameworks.validate(self.lesson_from(VALID_SOURCE), self.frameworks)

    def test_unknown_digital_code_is_rejected(self):
        broken = VALID_SOURCE.replace("1.1.NC1a —", "9.9.NC9z —")
        broken = broken.replace("nls: 1.1.NC1a,", "nls: 9.9.NC9z,")
        with self.assertRaises(frameworks.FrameworkError) as caught:
            frameworks.validate(self.lesson_from(broken), self.frameworks)
        self.assertIn("9.9.NC9z", caught.exception.message)
        self.assertIn("1.1.NC1a", caught.exception.fix)

    def test_ai_code_outside_the_grade_framework_is_rejected(self):
        broken = VALID_SOURCE.replace("AI-H11.2 —", "AI-H12 —")
        broken = broken.replace("ai: AI-H11.2,", "ai: AI-H12,")
        with self.assertRaises(frameworks.FrameworkError) as caught:
            frameworks.validate(self.lesson_from(broken), self.frameworks)
        self.assertIn("AI-H12", caught.exception.message)
        self.assertIn("AI-H11", caught.exception.message)

    def test_subject_without_a_framework_must_use_the_marker(self):
        broken = VALID_SOURCE.replace("subject: Hoá học", "subject: Ngữ văn")
        with self.assertRaises(frameworks.FrameworkError) as caught:
            frameworks.validate(self.lesson_from(broken), self.frameworks)
        self.assertIn(frameworks.NO_FRAMEWORK, caught.exception.fix)

    def test_subject_with_a_framework_must_not_use_the_marker(self):
        broken = VALID_SOURCE.replace(
            "- AI-H11.2 — Viết prompt để AI dự đoán chiều chuyển dịch cân bằng khi tăng áp suất.\n"
            "- AI-H11.4 — Đối chiếu kết quả AI với thí nghiệm, chỉ ra chỗ AI nói sai.\n",
            f"- {frameworks.NO_FRAMEWORK}\n- Học sinh đối chiếu kết quả AI.\n",
        ).replace("ai: AI-H11.2, AI-H11.4", f"ai: {frameworks.NO_CODE}")
        with self.assertRaises(frameworks.FrameworkError) as caught:
            frameworks.validate(self.lesson_from(broken), self.frameworks)
        self.assertIn("AI-H11", caught.exception.message)

    def test_marker_subject_passes_with_no_code_activities(self):
        text = VALID_SOURCE.replace("subject: Hoá học", "subject: Ngữ văn")
        text = text.replace(
            "- AI-H11.2 — Viết prompt để AI dự đoán chiều chuyển dịch cân bằng khi tăng áp suất.\n"
            "- AI-H11.4 — Đối chiếu kết quả AI với thí nghiệm, chỉ ra chỗ AI nói sai.\n",
            f"- {frameworks.NO_FRAMEWORK}\n- Học sinh nhận xét đoạn văn do AI viết.\n",
        )
        text = text.replace("ai: AI-H11.2, AI-H11.4", f"ai: {frameworks.NO_CODE}")
        frameworks.validate(self.lesson_from(text), self.frameworks)

    def test_marker_subject_may_not_carry_codes(self):
        text = VALID_SOURCE.replace("subject: Hoá học", "subject: Ngữ văn")
        text = text.replace(
            "- AI-H11.2 — Viết prompt để AI dự đoán chiều chuyển dịch cân bằng khi tăng áp suất.\n",
            f"- {frameworks.NO_FRAMEWORK}\n",
        )
        with self.assertRaises(frameworks.FrameworkError) as caught:
            frameworks.validate(self.lesson_from(text), self.frameworks)
        self.assertIn("AI-H11.4", caught.exception.message)

    def test_activity_code_missing_from_the_objectives_is_rejected(self):
        broken = VALID_SOURCE.replace("nls: 1.1.NC1a, 5.1.NC1a", "nls: 1.1.NC1a, 3.1.NC1a")
        with self.assertRaises(frameworks.FrameworkError) as caught:
            frameworks.validate(self.lesson_from(broken), self.frameworks)
        self.assertIn("3.1.NC1a", caught.exception.message)
        self.assertIn("Nang luc so", caught.exception.message)

    def test_activity_ai_code_missing_from_the_objectives_is_rejected(self):
        broken = VALID_SOURCE.replace("ai: AI-H11.2, AI-H11.4", "ai: AI-H11.2, AI-H11.3")
        with self.assertRaises(frameworks.FrameworkError) as caught:
            frameworks.validate(self.lesson_from(broken), self.frameworks)
        self.assertIn("AI-H11.3", caught.exception.message)
        self.assertIn("Nang luc AI", caught.exception.message)

    def test_specialised_track_uses_the_specialised_codes(self):
        text = VALID_SOURCE.replace("grade: 11", "grade: 11\ntrack: chuyen")
        text = text.replace("- AI-H11.2 —", "- AI-H-Chuyên —")
        text = text.replace("- AI-H11.4 — Đối chiếu kết quả AI với thí nghiệm, chỉ ra chỗ AI nói sai.\n", "")
        text = text.replace("ai: AI-H11.2, AI-H11.4", "ai: AI-H-Chuyên")
        frameworks.validate(self.lesson_from(text), self.frameworks)
```

- [ ] **Step 2: Chạy test để thấy nó fail**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k FrameworkValidateTest`
Kỳ vọng: FAIL với `AttributeError: module 'giao_an_parts.frameworks' has no attribute 'validate'`

- [ ] **Step 3: Thêm `validate` vào `frameworks.py`**

Thêm vào cuối `tools/vi/giao_an_parts/frameworks.py`:

```python
def validate(lesson, frameworks: Frameworks) -> None:
    """Kiểm mã NLS và mã AI của một giáo án. Lỗi thì nêu rõ mã sai và cách sửa."""
    unknown_nls = [code for code in lesson.nls_codes if code not in frameworks.nls]
    if unknown_nls:
        raise FrameworkError(
            "Mã năng lực số không có trong bảng của Bộ: " + ", ".join(unknown_nls),
            "Dùng một trong các mã: " + ", ".join(frameworks.nls),
        )

    subject = lesson.meta["subject"]
    grade = lesson.meta["grade"]
    track = lesson.meta.get("track", "")
    name = frameworks.framework_for(subject, grade, track)
    allowed = frameworks.ai_codes_for(subject, grade, track)

    if name is None:
        if not lesson.ai_no_framework:
            raise FrameworkError(
                f"Môn {subject} lớp {grade} chưa có khung mã AI trong {DOC_PATH.name}",
                f"Ghi dòng đầu của mục '### Nang luc AI' là '- {NO_FRAMEWORK}', "
                f"và dùng 'ai: {NO_CODE}' cho hoạt động AI.",
            )
        stray = [code for code in _activity_codes(lesson, "ai") if code != NO_CODE]
        if stray:
            raise FrameworkError(
                "Môn này chưa có khung mã AI nhưng hoạt động vẫn ghi mã: " + ", ".join(stray),
                f"Bỏ các mã đó; dùng 'ai: {NO_CODE}'.",
            )
    else:
        if lesson.ai_no_framework:
            raise FrameworkError(
                f"Môn {subject} lớp {grade} đã có khung {name} nên không được ghi "
                f"'{NO_FRAMEWORK}'",
                "Dùng một trong các mã: " + ", ".join(allowed),
            )
        unknown_ai = [code for code in lesson.ai_codes if code not in allowed]
        if unknown_ai:
            raise FrameworkError(
                f"Mã AI không thuộc khung {name} của môn {subject} lớp {grade}: "
                + ", ".join(unknown_ai),
                "Dùng một trong các mã: " + ", ".join(allowed),
            )

    for activity in lesson.activities:
        for code in activity.nls:
            if code not in lesson.nls_codes:
                raise FrameworkError(
                    f"Hoạt động {activity.title!r} dùng mã {code} chưa khai ở mục "
                    "'### Nang luc so'",
                    "Thêm mã đó vào mục tiêu, hoặc bỏ khỏi hoạt động.",
                )
        for code in activity.ai:
            if code == NO_CODE:
                continue
            if code not in lesson.ai_codes:
                raise FrameworkError(
                    f"Hoạt động {activity.title!r} dùng mã {code} chưa khai ở mục "
                    "'### Nang luc AI'",
                    "Thêm mã đó vào mục tiêu, hoặc bỏ khỏi hoạt động.",
                )


def _activity_codes(lesson, attribute: str) -> list[str]:
    codes: list[str] = []
    for activity in lesson.activities:
        for code in getattr(activity, attribute):
            if code not in codes:
                codes.append(code)
    return codes
```

- [ ] **Step 4: Chạy test để thấy nó pass**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k FrameworkValidateTest`
Kỳ vọng: PASS.

- [ ] **Step 5: Chạy cả bộ test**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests`
Kỳ vọng: OK.

- [ ] **Step 6: Commit**

```bash
git add tools/vi/giao_an_parts/frameworks.py tools/vi/tests/test_giao_an.py
git commit -m "feat(vi): validate competence codes against the framework document"
```

---

### Task 4: Cắt phần SGK (`sgk.py`)

**Files:**
- Create: `tools/vi/giao_an_parts/sgk.py`
- Modify: `tools/vi/tests/test_giao_an.py` (thêm `SgkTest`)

**Interfaces:**
- Consumes: `frameworks.normalise`.
- Produces:
  - `sgk.SgkError(message, fix)` — thuộc tính `.message`, `.fix`
  - `sgk.MAX_KB = 200`
  - `sgk.headings(lines: list[str]) -> list[tuple[int, str]]`
  - `sgk.extract(text: str, query: str) -> tuple[str, str, list[str]]` trả `(tiêu đề đã khớp, nội dung, cảnh báo)`

- [ ] **Step 1: Viết test trước**

Thêm vào `tools/vi/tests/test_giao_an.py`:

```python
SGK_SAMPLE = """# Chương 2. Nitrogen và sulfur

Mở đầu chương.

## Bài 4. Nitrogen

Nội dung bài 4.

## Bài 5. Ammonia và một số hợp chất ammonium

Nội dung bài 5, đoạn một.

Nội dung bài 5, đoạn hai.

## Bài 6. Một số hợp chất với oxygen của nitrogen

Nội dung bài 6.
"""


class SgkTest(unittest.TestCase):
    def test_headings_are_found(self):
        titles = [title for _, title in sgk.headings(SGK_SAMPLE.splitlines())]
        self.assertIn("Bài 5. Ammonia và một số hợp chất ammonium", titles)
        self.assertEqual(len(titles), 4)

    def test_extract_takes_only_the_requested_lesson(self):
        heading, body, warnings = sgk.extract(SGK_SAMPLE, "Bài 5")
        self.assertTrue(heading.startswith("Bài 5."))
        self.assertIn("đoạn một", body)
        self.assertIn("đoạn hai", body)
        self.assertNotIn("Nội dung bài 6", body)
        self.assertNotIn("Nội dung bài 4", body)
        self.assertEqual(warnings, [])

    def test_extract_ignores_case_and_diacritics(self):
        heading, _, _ = sgk.extract(SGK_SAMPLE, "bai 5")
        self.assertTrue(heading.startswith("Bài 5."))

    def test_unknown_lesson_names_the_nearest_headings(self):
        with self.assertRaises(sgk.SgkError) as caught:
            sgk.extract(SGK_SAMPLE, "Bài 99")
        self.assertIn("Bài 99", caught.exception.message)
        self.assertIn("Bài 4.", caught.exception.fix)

    def test_file_without_headings_is_rejected(self):
        with self.assertRaises(sgk.SgkError) as caught:
            sgk.extract("chỉ là văn bản thường\nkhông có tiêu đề\n", "Bài 5")
        self.assertIn("tiêu đề", caught.exception.message)

    def test_plain_bai_lines_count_as_headings(self):
        text = "BÀI 5. AMMONIA\n\nnội dung\n\nBÀI 6. NITRIC ACID\n\nkhác\n"
        heading, body, _ = sgk.extract(text, "Bài 5")
        self.assertIn("AMMONIA", heading)
        self.assertNotIn("NITRIC", body)

    def test_oversized_slice_warns(self):
        filler = "x" * 1024
        text = "## Bài 5. Ammonia\n\n" + "\n".join([filler] * 220) + "\n"
        _, _, warnings = sgk.extract(text, "Bài 5")
        self.assertTrue(any("KB" in warning for warning in warnings))

    def test_several_matches_warn_and_take_the_first(self):
        text = "## Bài 5. Ammonia\n\nmột\n\n## Bài 5. Ammonia (tiếp)\n\nhai\n"
        heading, body, warnings = sgk.extract(text, "Bài 5")
        self.assertEqual(heading, "Bài 5. Ammonia")
        self.assertIn("một", body)
        self.assertTrue(any("khớp" in warning for warning in warnings))
```

Đổi dòng import ở đầu file thành:

```python
from giao_an_parts import frameworks, parse, sgk  # noqa: E402
```

- [ ] **Step 2: Chạy test để thấy nó fail**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k SgkTest`
Kỳ vọng: FAIL với `ImportError: cannot import name 'sgk'`

- [ ] **Step 3: Viết `sgk.py`**

Tạo `tools/vi/giao_an_parts/sgk.py`:

```python
"""Cắt phần của một bài ra khỏi file SGK đã chuyển sang Markdown.

Mục đích: không bao giờ phải nạp cả quyển SGK (19–22 MB) vào ngữ cảnh.
Chỉ dùng thư viện chuẩn Python.
"""

from __future__ import annotations

import re

from .frameworks import normalise

MAX_KB = 200
_BAI_RE = re.compile(r"^\**\s*bai\s+\d+\b")


class SgkError(Exception):
    """Không cắt được phần của bài khỏi file SGK."""

    def __init__(self, message: str, fix: str = "") -> None:
        super().__init__(message)
        self.message = message
        self.fix = fix


def headings(lines: list[str]) -> list[tuple[int, str]]:
    """Trả về (chỉ số dòng, tiêu đề) cho mọi tiêu đề Markdown và mọi dòng mở đầu bằng 'Bài <số>'."""
    found: list[tuple[int, str]] = []
    for index, raw in enumerate(lines):
        text = raw.strip()
        if not text:
            continue
        if text.startswith("#"):
            found.append((index, text.lstrip("# ").strip()))
        elif _BAI_RE.match(normalise(text)):
            found.append((index, text.strip("*# ").strip()))
    return found


def extract(text: str, query: str) -> tuple[str, str, list[str]]:
    lines = text.splitlines()
    found = headings(lines)
    if not found:
        raise SgkError(
            "File SGK không có dòng tiêu đề nào để cắt theo",
            "Kiểm lại file Markdown đã chuyển từ PDF; có thể bước chuyển đổi đã thất bại.",
        )
    target = normalise(query)
    matches = [item for item in found if target in normalise(item[1])]
    if not matches:
        nearest = "; ".join(title for _, title in found[:3])
        raise SgkError(
            f"Không tìm thấy tiêu đề nào chứa {query!r}",
            f"Ba tiêu đề đầu tiên trong file: {nearest}. Sửa lại --bai cho khớp.",
        )
    start_index, heading = matches[0]
    following = [index for index, _ in found if index > start_index]
    end_index = following[0] if following else len(lines)
    body = "\n".join(lines[start_index:end_index]).strip() + "\n"
    warnings: list[str] = []
    size_kb = len(body.encode("utf-8")) / 1024
    if size_kb > MAX_KB:
        warnings.append(
            f"Phần cắt ra nặng {size_kb:.1f} KB, lớn hơn {MAX_KB} KB; có thể đã cắt sang bài sau."
        )
    if len(matches) > 1:
        warnings.append(
            f"Có {len(matches)} tiêu đề khớp {query!r}; đã lấy tiêu đề đầu tiên: {heading}"
        )
    return heading, body, warnings
```

- [ ] **Step 4: Chạy test để thấy nó pass**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k SgkTest`
Kỳ vọng: PASS, 8 test.

- [ ] **Step 5: Chạy cả bộ test**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests`
Kỳ vọng: OK.

- [ ] **Step 6: Commit**

```bash
git add tools/vi/giao_an_parts/sgk.py tools/vi/tests/test_giao_an.py
git commit -m "feat(vi): cut one lesson out of a converted textbook"
```

---

### Task 5: Dựng file Word (`docx_build.py`)

**Files:**
- Create: `tools/vi/giao_an_parts/docx_build.py`
- Modify: `tools/vi/tests/test_giao_an.py` (thêm `DocxBuildTest`)

**Interfaces:**
- Consumes: `word_parts.base.new_document`, `word_parts.base.write`, `word_parts.base.cell_text`, `word_parts.base.clear_borders`, `word_parts.base.grid_borders`, `word_parts.base.set_widths`, `word_parts.base.fill_cell`; `parse.Lesson`, `parse.Activity`.
- Produces:
  - `docx_build.FILENAME = "giao-an.docx"`, `docx_build.REVIEW_FILENAME = "can-soat.md"`
  - `docx_build.build(lesson, folder: Path) -> Path`
  - `docx_build.write_review(lesson, folder: Path, warnings: list[str]) -> Path`

Nếu `tools/vi/word_parts/base.py` chưa tồn tại thì dừng và báo BLOCKED: gói đề thi `v6.3.2-vi.5` phải xong trước.

- [ ] **Step 1: Viết test trước**

Thêm vào `tools/vi/tests/test_giao_an.py`:

```python
import tempfile
import zipfile


def document_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        return archive.read("word/document.xml").decode("utf-8")


class DocxBuildTest(unittest.TestCase):
    def setUp(self):
        self.lesson = parse.parse_lesson(VALID_SOURCE)
        self.tmp = tempfile.TemporaryDirectory()
        self.folder = Path(self.tmp.name)
        self.addCleanup(self.tmp.cleanup)

    def test_page_uses_a4_and_the_school_margins(self):
        from docx import Document

        path = docx_build.build(self.lesson, self.folder)
        section = Document(str(path)).sections[0]
        # Word lưu khổ giấy theo twip nên đọc lại lệch vài trăm EMU (< 0,001 cm); so tới 0,01 cm.
        for attribute, expected_cm in (
            ("page_width", 21.0),
            ("page_height", 29.7),
            ("top_margin", 2.0),
            ("bottom_margin", 2.0),
            ("left_margin", 2.5),
            ("right_margin", 2.0),
        ):
            with self.subTest(attribute=attribute):
                self.assertAlmostEqual(getattr(section, attribute).cm, expected_cm, places=2)

    def test_body_font_is_times_new_roman_fourteen_spaced_one_point_three(self):
        from docx import Document
        from docx.shared import Pt

        path = docx_build.build(self.lesson, self.folder)
        normal = Document(str(path)).styles["Normal"]
        self.assertEqual(normal.font.name, "Times New Roman")
        self.assertEqual(normal.font.size, Pt(14))
        self.assertAlmostEqual(normal.paragraph_format.line_spacing, 1.3, places=3)

    def test_five_twelve_sections_are_present(self):
        path = docx_build.build(self.lesson, self.folder)
        xml = document_xml(path)
        for heading in (
            "I. MỤC TIÊU",
            "II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU",
            "III. TIẾN TRÌNH DẠY HỌC",
            "IV. PHỤ LỤC",
        ):
            with self.subTest(heading=heading):
                self.assertIn(heading, xml)

    def test_every_activity_shows_the_four_elements_and_four_steps(self):
        path = docx_build.build(self.lesson, self.folder)
        xml = document_xml(path)
        for label in ("a) Mục tiêu", "b) Nội dung", "c) Sản phẩm", "d) Tổ chức thực hiện"):
            self.assertIn(label, xml)
        for step in (
            "Bước 1. Chuyển giao nhiệm vụ",
            "Bước 2. Học sinh thực hiện nhiệm vụ",
            "Bước 3. Báo cáo, thảo luận",
            "Bước 4. Kết luận, nhận định",
        ):
            self.assertIn(step, xml)

    def test_activities_are_numbered_in_file_order(self):
        path = docx_build.build(self.lesson, self.folder)
        xml = document_xml(path)
        self.assertIn("Hoạt động 1. Mở đầu (10 phút)", xml)
        self.assertIn("Hoạt động 4. Vận dụng (15 phút)", xml)

    def test_activity_codes_are_printed(self):
        path = docx_build.build(self.lesson, self.folder)
        xml = document_xml(path)
        self.assertIn("Năng lực số: 1.1.NC1a, 5.1.NC1a", xml)
        self.assertIn("Năng lực AI: AI-H11.2, AI-H11.4", xml)

    def test_worksheet_items_are_separate_left_aligned_paragraphs(self):
        from docx import Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        path = docx_build.build(self.lesson, self.folder)
        document = Document(str(path))
        items = [
            paragraph for paragraph in document.paragraphs
            if paragraph.text.startswith("Câu 1.") or paragraph.text.startswith("Câu 2.")
        ]
        self.assertEqual(len(items), 2, [paragraph.text for paragraph in document.paragraphs])
        for paragraph in items:
            self.assertEqual(paragraph.alignment, WD_ALIGN_PARAGRAPH.LEFT)

    def test_rubric_is_a_four_column_table(self):
        from docx import Document

        path = docx_build.build(self.lesson, self.folder)
        tables = Document(str(path)).tables
        rubric = tables[-1]
        self.assertEqual(len(rubric.columns), 4)
        self.assertEqual(len(rubric.rows), 3)
        self.assertEqual(rubric.rows[0].cells[3].text, "Mức 3")

    def test_chemical_subscripts_are_real_subscripts(self):
        path = docx_build.build(self.lesson, self.folder)
        self.assertIn('w:val="subscript"', document_xml(path))

    def test_internal_notes_never_reach_the_lesson_plan(self):
        source = VALID_SOURCE.replace(
            "- Thầy cô xác nhận giúp em địa chỉ mô phỏng PhET dùng được ở phòng máy của trường.",
            "- DAU-HIEU-GHI-CHU-NOI-BO",
        )
        lesson = parse.parse_lesson(source)
        path = docx_build.build(lesson, self.folder)
        xml = document_xml(path)
        self.assertNotIn("DAU-HIEU-GHI-CHU-NOI-BO", xml)
        self.assertNotIn("Cần thầy cô soát", xml)

    def test_review_file_carries_notes_and_warnings(self):
        source = VALID_SOURCE.replace("thoi-luong: 30", "thoi-luong: 20", 1)
        lesson = parse.parse_lesson(source)
        path = docx_build.write_review(lesson, self.folder, lesson.warnings)
        text = path.read_text(encoding="utf-8")
        self.assertEqual(path.name, "can-soat.md")
        self.assertIn("PhET", text)
        self.assertIn("thời lượng", text)

    def test_review_file_says_so_when_nothing_needs_review(self):
        source = VALID_SOURCE.split("## CAN SOAT")[0]
        lesson = parse.parse_lesson(source)
        path = docx_build.write_review(lesson, self.folder, [])
        self.assertIn("Không có mục nào cần soát", path.read_text(encoding="utf-8"))
```

Đổi dòng import ở đầu file thành:

```python
from giao_an_parts import docx_build, frameworks, parse, sgk  # noqa: E402
```

- [ ] **Step 2: Chạy test để thấy nó fail**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k DocxBuildTest`
Kỳ vọng: FAIL với `ImportError: cannot import name 'docx_build'`

- [ ] **Step 3: Viết `docx_build.py`**

Tạo `tools/vi/giao_an_parts/docx_build.py`:

```python
"""Dựng file Word Kế hoạch bài dạy theo Công văn 5512, và file ghi chú nội bộ.

Thể thức: A4 dọc, lề 2/2/2,5/2 cm, Times New Roman 14pt, giãn dòng 1,3.
Ghi chú nội bộ không bao giờ vào file giáo án; nó ra can-soat.md.
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

from word_parts import base
from word_parts.base import cell_text, clear_borders, fill_cell, grid_borders, set_widths, write

from .parse import Activity, Lesson

FILENAME = "giao-an.docx"
REVIEW_FILENAME = "can-soat.md"
MARGINS = (2.0, 2.0, 2.5, 2.0)
FONT_SIZE = 14
TABLE_SIZE = 12
LINE_SPACING = 1.3
NO_REVIEW = "Không có mục nào cần soát."
STEPS = (
    ("chuyen_giao", "Bước 1. Chuyển giao nhiệm vụ"),
    ("thuc_hien", "Bước 2. Học sinh thực hiện nhiệm vụ"),
    ("bao_cao", "Bước 3. Báo cáo, thảo luận"),
    ("ket_luan", "Bước 4. Kết luận, nhận định"),
)
OBJECTIVE_GROUPS = (
    ("a) Năng lực chung", "general"),
    ("b) Năng lực đặc thù", "specific"),
    ("c) Năng lực số", "nls_lines"),
    ("d) Năng lực trí tuệ nhân tạo (AI)", "ai_lines"),
)


def _document() -> Document:
    return base.new_document(
        margins_cm=MARGINS,
        size_pt=FONT_SIZE,
        line_spacing=LINE_SPACING,
        space_after_pt=4,
    )


def _paragraph(document: Document, text: str, alignment, *, bold: bool = False,
               italic: bool = False, space_before: int = 0):
    paragraph = document.add_paragraph()
    paragraph.alignment = alignment
    if space_before:
        paragraph.paragraph_format.space_before = Pt(space_before)
    write(paragraph, text, bold=bold, italic=italic)
    return paragraph


def _justified(document: Document, text: str, **kwargs):
    return _paragraph(document, text, WD_ALIGN_PARAGRAPH.JUSTIFY, **kwargs)


def _left(document: Document, text: str, **kwargs):
    return _paragraph(document, text, WD_ALIGN_PARAGRAPH.LEFT, **kwargs)


def _centre(document: Document, text: str, **kwargs):
    return _paragraph(document, text, WD_ALIGN_PARAGRAPH.CENTER, **kwargs)


def _shrink(table) -> None:
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(TABLE_SIZE)


def _add_header(document: Document, lesson: Lesson) -> None:
    table = document.add_table(rows=1, cols=2)
    clear_borders(table)
    set_widths(table, [Cm(8.0), Cm(8.5)])
    left_lines = [lesson.meta["school"]]
    if lesson.meta.get("department"):
        left_lines.append(lesson.meta["department"])
    fill_cell(table.rows[0].cells[0], left_lines)
    right_lines = []
    if lesson.meta.get("teacher"):
        right_lines.append(f"Giáo viên: {lesson.meta['teacher']}")
    if lesson.meta.get("school_year"):
        right_lines.append(f"Năm học: {lesson.meta['school_year']}")
    fill_cell(table.rows[0].cells[1], right_lines or [""], bold_first=False)
    _shrink(table)


def _add_title(document: Document, lesson: Lesson) -> None:
    _centre(document, "KẾ HOẠCH BÀI DẠY", bold=True, space_before=8)
    _centre(document, lesson.meta["lesson"], bold=True)
    detail = f"Môn: {lesson.meta['subject']} — Lớp {lesson.meta['grade']}"
    if lesson.meta.get("track"):
        detail += " (hệ chuyên)"
    _centre(document, detail)
    schedule = f"Thời lượng: {lesson.meta['periods']} tiết"
    if lesson.meta.get("period_numbers"):
        schedule += f" — Tiết thứ {lesson.meta['period_numbers']}"
    if lesson.meta.get("week"):
        schedule += f" — Tuần {lesson.meta['week']}"
    _centre(document, schedule)


def _add_objectives(document: Document, lesson: Lesson) -> None:
    _justified(document, "I. MỤC TIÊU", bold=True, space_before=8)
    _justified(document, "1. Về kiến thức", bold=True)
    for item in lesson.knowledge:
        _justified(document, f"- {item}")
    _justified(document, "2. Về năng lực", bold=True)
    for label, attribute in OBJECTIVE_GROUPS:
        _justified(document, label, italic=True)
        for item in getattr(lesson, attribute):
            _justified(document, f"- {item}")
    _justified(document, "3. Về phẩm chất", bold=True)
    for item in lesson.qualities:
        _justified(document, f"- {item}")


def _add_equipment(document: Document, lesson: Lesson) -> None:
    _justified(document, "II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU", bold=True, space_before=8)
    _justified(document, "1. Chuẩn bị của giáo viên", bold=True)
    for item in lesson.teacher_tools:
        _justified(document, f"- {item}")
    _justified(document, "2. Chuẩn bị của học sinh", bold=True)
    for item in lesson.student_tools:
        _justified(document, f"- {item}")


def _add_activity(document: Document, activity: Activity, number: int) -> None:
    _justified(
        document,
        f"Hoạt động {number}. {activity.title} ({activity.minutes} phút)",
        bold=True,
        space_before=8,
    )
    _justified(document, "a) Mục tiêu", italic=True)
    for text in activity.muc_tieu:
        _justified(document, text)
    if activity.nls:
        _justified(document, "Năng lực số: " + ", ".join(activity.nls))
    if activity.ai:
        _justified(document, "Năng lực AI: " + ", ".join(activity.ai))
    _justified(document, "b) Nội dung", italic=True)
    for text in activity.noi_dung:
        _justified(document, text)
    _justified(document, "c) Sản phẩm", italic=True)
    for text in activity.san_pham:
        _justified(document, text)
    _justified(document, "d) Tổ chức thực hiện", italic=True)
    for attribute, label in STEPS:
        _justified(document, label, bold=True)
        for text in getattr(activity, attribute):
            _justified(document, text)


def _add_appendix(document: Document, lesson: Lesson) -> None:
    _justified(document, "IV. PHỤ LỤC", bold=True, space_before=8)
    for sheet in lesson.worksheets:
        _left(document, sheet.title, bold=True, space_before=6)
        for item in sheet.items:
            _left(document, item)
    _justified(document, "Rubric đánh giá năng lực số và năng lực AI", bold=True, space_before=8)
    table = document.add_table(rows=len(lesson.rubrics) + 1, cols=4)
    grid_borders(table)
    set_widths(table, [Cm(4.5), Cm(4.0), Cm(4.0), Cm(4.0)])
    for cell, text in zip(table.rows[0].cells, ("Tiêu chí", "Mức 1", "Mức 2", "Mức 3")):
        cell_text(cell, text, bold=True)
    for row_index, rubric in enumerate(lesson.rubrics, start=1):
        cells = table.rows[row_index].cells
        cell_text(cells[0], rubric.title)
        for column, level in enumerate(rubric.levels, start=1):
            cell_text(cells[column], level)
    _shrink(table)


def build(lesson: Lesson, folder: Path) -> Path:
    document = _document()
    _add_header(document, lesson)
    _add_title(document, lesson)
    _add_objectives(document, lesson)
    _add_equipment(document, lesson)
    _justified(document, "III. TIẾN TRÌNH DẠY HỌC", bold=True, space_before=8)
    for number, activity in enumerate(lesson.activities, start=1):
        _add_activity(document, activity, number)
    _add_appendix(document, lesson)
    path = folder / FILENAME
    document.save(str(path))
    return path


def write_review(lesson: Lesson, folder: Path, warnings: list[str]) -> Path:
    path = folder / REVIEW_FILENAME
    lines = [
        "# Cần thầy cô soát",
        "",
        f"Bài: {lesson.meta['lesson']}",
        f"Môn: {lesson.meta['subject']} — Lớp {lesson.meta['grade']}",
        "",
        "File này là ghi chú nội bộ, không nằm trong giáo án nộp cho trường.",
        "",
    ]
    items = list(lesson.review_notes) + list(warnings)
    lines += [f"- {item}" for item in items] if items else [NO_REVIEW]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path
```

- [ ] **Step 4: Chạy test để thấy nó pass**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k DocxBuildTest`
Kỳ vọng: PASS. Nếu `test_worksheet_items_are_separate_left_aligned_paragraphs` fail vì tìm được nhiều hơn hai đoạn, kiểm lại `_add_appendix` không nhân đôi danh sách phiếu.

- [ ] **Step 5: Chạy cả bộ test**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests`
Kỳ vọng: OK.

- [ ] **Step 6: Commit**

```bash
git add tools/vi/giao_an_parts/docx_build.py tools/vi/tests/test_giao_an.py
git commit -m "feat(vi): build the 5512 lesson plan Word file"
```

---

### Task 6: Điểm vào dòng lệnh (`giao_an.py`)

**Files:**
- Create: `tools/vi/giao_an.py`
- Modify: `tools/vi/tests/test_giao_an.py` (thêm `CliTest`)

**Interfaces:**
- Consumes: `parse.parse_lesson`, `parse.ParseError`, `frameworks.load_file`, `frameworks.validate`, `frameworks.FrameworkError`, `sgk.extract`, `sgk.SgkError`, `docx_build.build`, `docx_build.write_review`.
- Produces:
  - `giao_an.main(argv: list[str] | None = None) -> int`
  - `giao_an.load_docx_build()` — import muộn, để test thay được
  - stdout: đúng một dòng JSON.

- [ ] **Step 1: Viết test trước**

Thêm vào `tools/vi/tests/test_giao_an.py`:

```python
import contextlib
import io
import json
from unittest import mock

import giao_an  # noqa: E402


class CliTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.folder = Path(self.tmp.name) / "giáo án hoá 11"
        self.folder.mkdir(parents=True)
        self.addCleanup(self.tmp.cleanup)

    def write_source(self, text: str = VALID_SOURCE) -> None:
        (self.folder / "giao-an.md").write_text(text, encoding="utf-8")

    def run_cli(self, *args: str) -> tuple[int, dict]:
        out = io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(io.StringIO()):
            code = giao_an.main(list(args))
        printed = out.getvalue().strip().splitlines()
        self.assertEqual(len(printed), 1, printed)
        return code, json.loads(printed[0])

    def test_export_writes_the_plan_and_the_review_file(self):
        self.write_source()
        code, data = self.run_cli("xuat", str(self.folder))
        self.assertEqual(code, 0, data)
        self.assertTrue(data["ready"])
        self.assertEqual([Path(name).name for name in data["files"]], ["giao-an.docx", "can-soat.md"])
        self.assertEqual(data["activities"], 4)
        self.assertEqual(data["periods"], 2)
        self.assertEqual(data["minutes"], 90)
        self.assertEqual(data["nls_codes"], ["1.1.NC1a", "5.1.NC1a"])
        self.assertEqual(data["ai_codes"], ["AI-H11.2", "AI-H11.4"])
        self.assertIsNone(data["error"])

    def test_export_works_with_a_vietnamese_folder_name(self):
        self.write_source()
        code, data = self.run_cli("xuat", str(self.folder))
        self.assertEqual(code, 0, data)
        for name in data["files"]:
            self.assertTrue(Path(name).is_file(), name)

    def test_plan_only_writes_nothing(self):
        self.write_source()
        code, data = self.run_cli("xuat", str(self.folder), "--plan-only")
        self.assertEqual(code, 0, data)
        self.assertEqual(data["files"], [])
        self.assertEqual(list(self.folder.glob("*.docx")), [])

    def test_missing_source_reports_input_step(self):
        code, data = self.run_cli("xuat", str(self.folder))
        self.assertEqual(code, 1)
        self.assertEqual(data["error"]["step"], "input")
        self.assertIn("giao-an.md", data["error"]["message"])

    def test_structure_error_reports_parse_step_with_the_line(self):
        self.write_source(VALID_SOURCE.replace("bao-cao: Hai nhóm trình bày dự đoán.\n", ""))
        code, data = self.run_cli("xuat", str(self.folder))
        self.assertEqual(code, 1)
        self.assertEqual(data["error"]["step"], "parse")
        self.assertIn("Dòng", data["error"]["message"])

    def test_code_error_reports_framework_step(self):
        self.write_source(VALID_SOURCE.replace("1.1.NC1a", "9.9.NC9z"))
        code, data = self.run_cli("xuat", str(self.folder))
        self.assertEqual(code, 1)
        self.assertEqual(data["error"]["step"], "framework")
        self.assertIn("9.9.NC9z", data["error"]["message"])

    def test_missing_python_docx_reports_docx_step(self):
        self.write_source()
        with mock.patch.object(giao_an, "load_docx_build", side_effect=ImportError("No module named 'docx'")):
            code, data = self.run_cli("xuat", str(self.folder))
        self.assertEqual(code, 1)
        self.assertEqual(data["error"]["step"], "docx")
        self.assertIn("requirements-vi.txt", data["error"]["fix"])

    def test_write_failure_reports_write_step(self):
        self.write_source()
        with mock.patch.object(giao_an, "load_docx_build") as loader:
            loader.return_value.build.side_effect = OSError("file đang mở trong Word")
            code, data = self.run_cli("xuat", str(self.folder))
        self.assertEqual(code, 1)
        self.assertEqual(data["error"]["step"], "write")
        self.assertIn("Word", data["error"]["fix"])

    def test_unexpected_failure_still_prints_one_json_line(self):
        self.write_source()
        with mock.patch.object(giao_an, "load_docx_build") as loader:
            loader.return_value.build.side_effect = ValueError("lỗi ngoài dự kiến")
            code, data = self.run_cli("xuat", str(self.folder))
        self.assertEqual(code, 1)
        self.assertEqual(data["error"]["step"], "internal")

    def test_time_mismatch_surfaces_as_a_warning(self):
        self.write_source(VALID_SOURCE.replace("thoi-luong: 30", "thoi-luong: 20", 1))
        code, data = self.run_cli("xuat", str(self.folder))
        self.assertEqual(code, 0, data)
        self.assertTrue(any("thời lượng" in warning for warning in data["warnings"]))

    def test_second_run_warns_about_overwriting(self):
        self.write_source()
        self.run_cli("xuat", str(self.folder))
        code, data = self.run_cli("xuat", str(self.folder))
        self.assertEqual(code, 0, data)
        self.assertTrue(any("Ghi đè" in warning for warning in data["warnings"]))

    def test_cut_textbook_writes_the_slice(self):
        source = self.folder / "sgk.md"
        source.write_text(SGK_SAMPLE, encoding="utf-8")
        code, data = self.run_cli("trich-sgk", str(source), "--bai", "Bài 5")
        self.assertEqual(code, 0, data)
        self.assertTrue(data["heading"].startswith("Bài 5."))
        self.assertEqual(len(data["files"]), 1)
        written = Path(data["files"][0])
        self.assertTrue(written.is_file())
        self.assertIn("đoạn hai", written.read_text(encoding="utf-8"))
        self.assertGreater(data["size_kb"], 0)

    def test_cut_textbook_honours_the_output_path(self):
        source = self.folder / "sgk.md"
        source.write_text(SGK_SAMPLE, encoding="utf-8")
        target = self.folder / "phần bài 5.md"
        code, data = self.run_cli("trich-sgk", str(source), "--bai", "Bài 5", "--ra", str(target))
        self.assertEqual(code, 0, data)
        self.assertEqual(Path(data["files"][0]).name, target.name)

    def test_cut_textbook_missing_file_reports_input_step(self):
        code, data = self.run_cli("trich-sgk", str(self.folder / "khong-co.md"), "--bai", "Bài 5")
        self.assertEqual(code, 1)
        self.assertEqual(data["error"]["step"], "input")

    def test_cut_textbook_unknown_lesson_reports_parse_step(self):
        source = self.folder / "sgk.md"
        source.write_text(SGK_SAMPLE, encoding="utf-8")
        code, data = self.run_cli("trich-sgk", str(source), "--bai", "Bài 99")
        self.assertEqual(code, 1)
        self.assertEqual(data["error"]["step"], "parse")
        self.assertIn("Bài 4.", data["error"]["fix"])

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
        with mock.patch.object(giao_an.sys, "stdout", stream):
            giao_an.emit({"ready": False, "warnings": ["Thiếu phiếu học tập"]})
        data = json.loads(stream.buffer.getvalue().decode("utf-8").strip())
        self.assertEqual(data["warnings"], ["Thiếu phiếu học tập"])
```

- [ ] **Step 2: Chạy test để thấy nó fail**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k CliTest`
Kỳ vọng: FAIL với `ModuleNotFoundError: No module named 'giao_an'`

- [ ] **Step 3: Viết `giao_an.py`**

Tạo `tools/vi/giao_an.py`:

```python
#!/usr/bin/env python3
"""Soạn Kế hoạch bài dạy (giáo án) tích hợp năng lực số và năng lực AI.

Cách dùng:
    python tools/vi/giao_an.py xuat <thư_mục_giáo_án> [--plan-only]
    python tools/vi/giao_an.py trich-sgk <file_sgk.md> --bai "<tên bài>" [--ra <file.md>]

stdout: đúng một dòng JSON. Tiến trình đi ra stderr.
Mã thoát: 0 khi xong, 1 khi lỗi.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from giao_an_parts import frameworks, parse, sgk  # noqa: E402

SOURCE_NAME = "giao-an.md"
SLICE_NAME = "sgk-trich.md"
MAX_PATH = 200
FIX_SOURCE = (
    "Viết file giao-an.md trong thư mục giáo án theo docs/vi/tro-ly/giao-an.md rồi chạy lại."
)
FIX_PARSE = "Sửa đúng dòng đó trong giao-an.md theo docs/vi/tro-ly/giao-an.md rồi chạy lại."
FIX_DOCX = (
    "Cài thư viện bằng: python -m pip install -r tools/vi/requirements-vi.txt "
    "(hoặc chạy lại CAI-DAT.bat)"
)
FIX_WRITE = "Đóng file Word đang mở rồi chạy lại; kiểm tra ổ đĩa còn trống."
FIX_INTERNAL = "Gửi nguyên dòng error.message cho người bảo trì."


def log(text: str) -> None:
    print(text, file=sys.stderr, flush=True)


def emit(payload: dict) -> None:
    text = json.dumps(payload, ensure_ascii=False) + "\n"
    try:
        sys.stdout.write(text)
    except UnicodeEncodeError:
        sys.stdout.buffer.write(text.encode("utf-8", errors="replace"))
    sys.stdout.flush()


def result(*, ready: bool, files=(), warnings=(), error: dict | None = None, **extra) -> dict:
    payload = {
        "ready": ready,
        "files": [str(path) for path in files],
        "warnings": list(warnings),
        "error": error,
    }
    payload.update(extra)
    return payload


def failure(step: str, message: str, fix: str, **rest) -> dict:
    return result(ready=False, error={"step": step, "message": message, "fix": fix}, **rest)


def load_docx_build():
    """Import muộn để thiếu python-docx vẫn báo được lỗi dạng JSON."""
    from giao_an_parts import docx_build

    return docx_build


def configure_streams() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


def _lesson_extra(lesson) -> dict:
    return {
        "activities": len(lesson.activities),
        "periods": lesson.periods,
        "minutes": lesson.minutes,
        "nls_codes": lesson.nls_codes,
        "ai_codes": lesson.ai_codes,
    }


def command_export(args) -> int:
    folder = args.folder.expanduser().resolve()
    if not folder.is_dir():
        emit(failure("input", f"Không có thư mục giáo án: {folder}", FIX_SOURCE))
        return 1
    source = folder / SOURCE_NAME
    if not source.is_file():
        emit(failure("input", f"Không có file {SOURCE_NAME} trong {folder}", FIX_SOURCE))
        return 1
    try:
        text = source.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError) as exc:
        emit(failure("input", f"Không đọc được {source}: {exc}",
                     "Lưu lại giao-an.md bằng bảng mã UTF-8 rồi chạy lại"))
        return 1

    try:
        lesson = parse.parse_lesson(text)
    except parse.ParseError as exc:
        emit(failure("parse", str(exc), FIX_PARSE))
        return 1

    try:
        frameworks.validate(lesson, frameworks.load_file())
    except frameworks.FrameworkError as exc:
        emit(failure("framework", exc.message, exc.fix, **_lesson_extra(lesson)))
        return 1

    extra = _lesson_extra(lesson)
    paper_warnings = list(lesson.warnings)
    warnings = list(paper_warnings)
    if len(str(folder)) > MAX_PATH:
        warnings.append(
            f"Đường dẫn thư mục giáo án dài {len(str(folder))} ký tự; Windows có thể không ghi "
            "được file. Chuyển bộ công cụ ra ổ đĩa gần gốc, ví dụ D:\\PPTmaster."
        )

    if args.plan_only:
        log("Chỉ kiểm giao-an.md, không ghi file.")
        emit(result(ready=True, warnings=warnings, **extra))
        return 0

    try:
        docx_build = load_docx_build()
    except ImportError as exc:
        emit(failure("docx", f"Chưa cài thư viện python-docx ({exc})", FIX_DOCX,
                     warnings=warnings, **extra))
        return 1

    for name in ("giao-an.docx", "can-soat.md"):
        if (folder / name).exists():
            warnings.append(f"Ghi đè file có sẵn: {name}")

    try:
        plan_path = docx_build.build(lesson, folder)
        review_path = docx_build.write_review(lesson, folder, paper_warnings)
    except OSError as exc:
        emit(failure("write", f"Không ghi được file: {exc}", FIX_WRITE,
                     warnings=warnings, **extra))
        return 1
    except Exception as exc:  # noqa: BLE001 - stdout không bao giờ được để trống
        emit(failure("internal", f"Lỗi ngoài dự kiến khi dựng file Word: {exc}", FIX_INTERNAL,
                     warnings=warnings, **extra))
        return 1

    log("Đã xuất giáo án và file ghi chú.")
    emit(result(ready=True, files=[plan_path, review_path], warnings=warnings, **extra))
    return 0


def command_cut(args) -> int:
    source = args.sgk.expanduser().resolve()
    if not source.is_file():
        emit(failure("input", f"Không có file SGK: {source}",
                     "Chuyển SGK PDF sang Markdown một lần bằng "
                     "skills/ppt-master/scripts/source_to_md.py rồi chạy lại."))
        return 1
    try:
        text = source.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError) as exc:
        emit(failure("input", f"Không đọc được {source}: {exc}",
                     "Lưu lại file SGK bằng bảng mã UTF-8 rồi chạy lại"))
        return 1
    try:
        heading, body, warnings = sgk.extract(text, args.bai)
    except sgk.SgkError as exc:
        emit(failure("parse", exc.message, exc.fix))
        return 1
    target = (args.ra or source.parent / SLICE_NAME).expanduser().resolve()
    try:
        target.write_text(body, encoding="utf-8")
    except OSError as exc:
        emit(failure("write", f"Không ghi được {target}: {exc}", FIX_WRITE))
        return 1
    size_kb = round(len(body.encode("utf-8")) / 1024, 1)
    log(f"Đã cắt phần {heading!r} ra {target}.")
    emit(result(ready=True, files=[target], warnings=warnings, heading=heading, size_kb=size_kb))
    return 0


def main(argv: list[str] | None = None) -> int:
    configure_streams()
    parser = argparse.ArgumentParser(description="Soạn giáo án tích hợp năng lực số và năng lực AI")
    commands = parser.add_subparsers(dest="command", required=True)

    export = commands.add_parser("xuat", help="Xuất giáo án từ giao-an.md")
    export.add_argument("folder", type=Path, help="Thư mục giáo án, chứa file giao-an.md")
    export.add_argument("--plan-only", action="store_true", help="Chỉ kiểm, không ghi file")
    export.set_defaults(handler=command_export)

    cut = commands.add_parser("trich-sgk", help="Cắt phần một bài ra khỏi SGK đã chuyển Markdown")
    cut.add_argument("sgk", type=Path, help="File SGK dạng Markdown")
    cut.add_argument("--bai", required=True, help="Tên hoặc số bài, ví dụ \"Bài 5\"")
    cut.add_argument("--ra", type=Path, default=None, help="Đường dẫn file kết quả")
    cut.set_defaults(handler=command_cut)

    args = parser.parse_args(argv)
    try:
        return args.handler(args)
    except Exception as exc:  # noqa: BLE001 - stdout không bao giờ được để trống
        emit(failure("internal", f"Lỗi ngoài dự kiến: {exc}", FIX_INTERNAL))
        return 1


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: Chạy test để thấy nó pass**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k CliTest`
Kỳ vọng: PASS.

- [ ] **Step 5: Chạy cả bộ test**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests`
Kỳ vọng: OK.

- [ ] **Step 6: Commit**

```bash
git add tools/vi/giao_an.py tools/vi/tests/test_giao_an.py
git commit -m "feat(vi): add the giao_an command with a one-line JSON contract"
```

---

### Task 7: Hướng dẫn cho AI

**Files:**
- Create: `docs/vi/tro-ly/giao-an.md`
- Modify: `AGENTS.vi.md` (mục 3, bảng mục 10, thêm mục 13)
- Modify: `docs/vi/tro-ly/quy-trinh-hoi.md` (dòng thứ 8, ca mơ hồ "giáo án")
- Modify: `docs/vi/tro-ly/mau-brief.md` (thêm loại việc thứ 8)
- Modify: `tools/vi/tests/test_vi_layer.py`

**Interfaces:**
- Consumes: hợp đồng `giao-an.md` ở Task 2, bảng mã ở Task 1, lệnh và JSON ở Task 6.
- Produces: không có mã.

Ràng buộc riêng: `giao-an.md` **không** được thêm vào `GUIDE_FILES`; hai file mới trong `docs/vi/tro-ly/` **không** được dùng liên kết Markdown; `## Câu hỏi bắt buộc` có 1–7 câu đánh số, mỗi câu có `Gợi ý:`; `## Tạo nhanh` có 2–3 câu đánh số.

- [ ] **Step 1: Viết test trước**

Trong `tools/vi/tests/test_vi_layer.py`, sửa hằng số và thêm lớp test.

Sửa `test_agents_vi_keeps_the_three_task_sections_last_in_order` (do gói vi.5 tạo) thành:

```python
    def test_agents_vi_keeps_the_four_task_sections_last_in_order(self):
        headings = h2_headings(read("AGENTS.vi.md"))
        self.assertEqual(headings[-4], AGENTS_VI_ASSISTANT_HEADING)
        self.assertEqual(headings[-3], AGENTS_VI_VIDEO_HEADING)
        self.assertEqual(headings[-2], AGENTS_VI_EXAM_HEADING)
        self.assertEqual(headings[-1], AGENTS_VI_LESSON_HEADING)
```

Sửa dòng cuối của `test_agents_vi_environment_section_points_to_guide` thành:

```python
        self.assertEqual(h2_headings(text)[-4], AGENTS_VI_ASSISTANT_HEADING)
```

Sửa `test_task_type_count_matches_the_table` thành:

```python
    def test_task_type_count_matches_the_table(self):
        agents_vi_body = section(read("AGENTS.vi.md"), AGENTS_VI_ASSISTANT_HEADING)
        self.assertIn("8 loại", agents_vi_body)
        for stale in ("5 loại", "6 loại", "7 loại"):
            self.assertNotIn(stale, agents_vi_body)
        quy_trinh_body = section(read("docs/vi/tro-ly/quy-trinh-hoi.md"), "## Khi nào áp dụng")
        self.assertIn("8 loại", quy_trinh_body)
        for stale in ("5 loại", "6 loại", "7 loại"):
            self.assertNotIn(stale, quy_trinh_body)
```

Thêm vào cuối file, trước khối `if __name__`:

```python
AGENTS_VI_LESSON_HEADING = "## 13. Soạn giáo án tích hợp năng lực số và năng lực AI"
LESSON_GUIDE = "docs/vi/tro-ly/giao-an.md"
FRAMEWORK_GUIDE = "docs/vi/tro-ly/nang-luc-so-va-ai.md"
LESSON_GUIDE_HEADINGS = (
    "## Khi nào dùng",
    "## Câu hỏi bắt buộc",
    "## Câu hỏi tuỳ chọn",
    "## Tạo nhanh",
    "## Cấu trúc giáo án",
    "## Đầu ra",
    "## Ghi vào brief",
)
LESSON_COMMAND = r"python tools\vi\giao_an.py xuat"
ROUTING_QUESTION = "file Word kế hoạch bài dạy"


class LessonGuideTest(unittest.TestCase):
    def test_guide_has_its_own_sections_in_order(self):
        self.assertEqual(h2_headings(read(LESSON_GUIDE)), list(LESSON_GUIDE_HEADINGS))

    def test_guide_is_not_treated_as_a_slide_guide(self):
        self.assertNotIn("giao-an.md", GUIDE_FILES)

    def test_guide_questions_are_limited_and_have_suggestions(self):
        items = numbered_items(section(read(LESSON_GUIDE), "## Câu hỏi bắt buộc"))
        self.assertTrue(1 <= len(items) <= 7, f"{len(items)} câu")
        for item in items:
            self.assertIn("Gợi ý:", item)

    def test_guide_quick_mode_asks_two_or_three_questions(self):
        items = numbered_items(section(read(LESSON_GUIDE), "## Tạo nhanh"))
        self.assertTrue(2 <= len(items) <= 3, f"{len(items)} câu")

    def test_guide_states_the_source_grammar(self):
        body = section(read(LESSON_GUIDE), "## Cấu trúc giáo án")
        for token in ("## MUC TIEU", "## THIET BI", "## TIEN TRINH", "## RUBRIC", "## CAN SOAT",
                      "thoi-luong:", "muc-tieu:", "chuyen-giao:", "ket-luan:", "nls:", "ai:"):
            with self.subTest(token=token):
                self.assertIn(token, body)

    def test_guide_names_the_output_files(self):
        body = section(read(LESSON_GUIDE), "## Đầu ra")
        for name in ("giao-an.docx", "can-soat.md", "projects/_giao-an/"):
            self.assertIn(name, body)

    def test_guide_points_to_the_framework_document(self):
        self.assertIn("nang-luc-so-va-ai.md", read(LESSON_GUIDE))

    def test_guide_forbids_rewriting_the_teacher_content_and_inventing_codes(self):
        text = read(LESSON_GUIDE)
        for phrase in ("giữ nguyên", "không tự đặt mã", "can-soat.md"):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, text)

    def test_tro_ly_files_have_no_markdown_links(self):
        for name in ("giao-an.md", "nang-luc-so-va-ai.md"):
            with self.subTest(file=name):
                self.assertEqual(LINK_RE.findall(read(f"docs/vi/tro-ly/{name}")), [])

    def test_framework_document_credits_its_author_and_sources(self):
        text = read(FRAMEWORK_GUIDE)
        for phrase in ("Lương Hải Anh", "2Anh AI Education", "Bộ GD&ĐT", "Phụ lục III & IV"):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, text)


class LessonWiringTest(unittest.TestCase):
    def test_common_rules_table_lists_the_lesson_plan_task(self):
        body = section(read("docs/vi/tro-ly/quy-trinh-hoi.md"), "## Khi nào áp dụng")
        self.assertIn("giao-an.md", body)
        for keyword in ("kế hoạch bài dạy", "KHBD"):
            self.assertIn(keyword, body)

    def test_common_rules_name_the_ambiguous_lesson_keyword(self):
        body = section(read("docs/vi/tro-ly/quy-trinh-hoi.md"), "## Khi nào áp dụng")
        self.assertIn('"giáo án"', body)
        self.assertIn(ROUTING_QUESTION, body)

    def test_agents_vi_lesson_section_explains_both_flows_and_the_command(self):
        text = read("AGENTS.vi.md")
        self.assertIn(AGENTS_VI_LESSON_HEADING, h2_headings(text))
        body = section(text, AGENTS_VI_LESSON_HEADING)
        self.assertIn(LESSON_COMMAND, body)
        for phrase in (
            "source_to_md.py",
            "trich-sgk",
            "(docs/vi/tro-ly/giao-an.md)",
            "(docs/vi/tro-ly/nang-luc-so-va-ai.md)",
            "projects/_giao-an/",
            "giao-an.md",
            "ảnh",
        ):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, body)

    def test_agents_vi_lesson_section_maps_every_error_step(self):
        body = section(read("AGENTS.vi.md"), AGENTS_VI_LESSON_HEADING)
        for step in ("input", "parse", "framework", "docx", "write"):
            self.assertIn(f"`{step}`", body)
        self.assertIn("requirements-vi.txt", body)

    def test_agents_vi_lesson_section_asks_before_routing(self):
        body = section(read("AGENTS.vi.md"), AGENTS_VI_LESSON_HEADING)
        self.assertIn(ROUTING_QUESTION, body)

    def test_agents_vi_lesson_section_keeps_the_venv_conditional(self):
        body = section(read("AGENTS.vi.md"), AGENTS_VI_LESSON_HEADING)
        self.assertIn("mục 4", body)

    def test_agents_vi_lesson_section_reads_the_curriculum_without_editing_it(self):
        body = section(read("AGENTS.vi.md"), AGENTS_VI_LESSON_HEADING)
        self.assertIn("không sửa", body)
        self.assertIn("phân phối chương trình", body)

    def test_agents_vi_triggers_include_lesson_plan_phrases(self):
        body = section(read("AGENTS.vi.md"), "## 3. Câu lệnh tiếng Việt kích hoạt skill `ppt-master`")
        for phrase in ("kế hoạch bài dạy", "giáo án"):
            self.assertIn(phrase, body)

    def test_brief_template_lists_the_lesson_plan_task(self):
        self.assertIn("Soạn giáo án", read("docs/vi/tro-ly/mau-brief.md"))
```

- [ ] **Step 2: Chạy test để thấy nó fail**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k LessonGuideTest`
Kỳ vọng: FAIL vì chưa có file hướng dẫn.

- [ ] **Step 3: Viết `docs/vi/tro-ly/giao-an.md`**

Tạo file với đúng bảy mục h2 theo thứ tự `LESSON_GUIDE_HEADINGS`. Không dùng liên kết Markdown. Nội dung bắt buộc:

- Dòng 3: `File dành cho AI. Luôn đọc docs/vi/tro-ly/quy-trinh-hoi.md trước file này, và đọc docs/vi/tro-ly/nang-luc-so-va-ai.md trước khi viết mục tiêu năng lực.`
- `## Khi nào dùng`: thầy cô cần file Word kế hoạch bài dạy theo Công văn 5512 có tích hợp năng lực số và năng lực AI, môn nào cũng được. Nêu hai luồng: nâng cấp giáo án có sẵn (luồng A) và soạn mới (luồng B). Hai câu lệnh ví dụ. Ghi rõ: **câu lệnh có chữ "giáo án" thì hỏi một câu định tuyến trước** (xem quy-trinh-hoi.md), và **ảnh chụp giáo án thì không đọc được**, xin PDF hoặc Word.
- `## Câu hỏi bắt buộc`: đúng bảy câu theo §6.6 của spec, mỗi câu kèm `Gợi ý:`.
- `## Câu hỏi tuỳ chọn`: chỉ hỏi khi thầy cô nhắc tới — trường có khung ký duyệt riêng không; lớp có học sinh cần hỗ trợ riêng không.
- `## Tạo nhanh`: hai câu — môn, lớp, tên bài, số tiết; và có giáo án cũ hay soạn mới.
- `## Cấu trúc giáo án`: nêu ngữ pháp `giao-an.md` đầy đủ — khối thông tin đầu, sáu mục h3 của `## MUC TIEU`, hai mục của `## THIET BI`, tám khoá bắt buộc của mỗi hoạt động cộng `nls:`/`ai:`, `## PHIEU HOC TAP`, `## RUBRIC` ba mức, `## CAN SOAT`; ba dấu đánh dấu `~ ~`, `^ ^`, `**`; khoá lặp thành nhiều đoạn. Nêu rõ các phép kiểm sẽ chặn.
- `## Đầu ra`: `projects/_giao-an/<tên_bài>/` chứa `giao-an.md`, `giao-an.docx`, `can-soat.md`. Ghi rõ `can-soat.md` là ghi chú nội bộ, không nằm trong giáo án nộp cho trường.
- `## Ghi vào brief`: loại việc "Soạn giáo án tích hợp năng lực số và năng lực AI", ghi đúng lời thầy cô.

Điều cấm phải ghi nguyên văn: `Luồng A phải giữ nguyên nội dung chuyên môn của thầy cô: không rút gọn, không viết lại cho hay hơn, không đổi bài tập. Chỉ thêm phần năng lực số, năng lực AI và rubric. Không tự đặt mã chỉ báo cho môn chưa có khung. Chỗ nào không đọc được từ file gốc thì ghi vào mục CAN SOAT, không tự viết bù.`

- [ ] **Step 4: Sửa `AGENTS.vi.md`**

1. Mục 3: thêm `"soạn giáo án"`, `"kế hoạch bài dạy"`, `"KHBD"`.
2. Mục 10: đổi "7 loại việc" thành "8 loại việc" và thêm dòng bảng:

```
| Soạn giáo án tích hợp năng lực số và AI | [docs/vi/tro-ly/giao-an.md](docs/vi/tro-ly/giao-an.md) |
```

3. Thêm mục 13 làm mục cuối file, tiêu đề `## 13. Soạn giáo án tích hợp năng lực số và năng lực AI`:

- Câu mở: câu lệnh có chữ "giáo án" thì **hỏi đúng một câu trước**: "Thầy cô cần file Word kế hoạch bài dạy (giáo án 5512), hay slide trình chiếu cho bài này?". Trả lời Word thì theo mục này; trả lời slide thì theo mục 10.
- Đọc `docs/vi/tro-ly/giao-an.md` và `docs/vi/tro-ly/nang-luc-so-va-ai.md`, hỏi một lượt, chờ trả lời. Nhắc như mục 4 về `venv\Scripts\python.exe`.
- Bước 1 (luồng A): đọc giáo án cũ bằng `python skills/ppt-master/scripts/source_to_md.py <file> -o <thư_mục_tạm>`. Thầy cô đưa **ảnh** thì nói rõ không đọc được, xin PDF hoặc Word.
- Bước 2 (tuỳ chọn): có file kế hoạch dạy học hoặc **phân phối chương trình** thì đọc để lấy tuần, tiết thứ, yêu cầu cần đạt và mã năng lực số đã khai. **Không sửa** file đó.
- Bước 3 (tuỳ chọn): cần nội dung SGK thì chuyển quyển SGK sang Markdown **một lần** rồi cắt: `python tools\vi\giao_an.py trich-sgk <sgk.md> --bai "<tên bài>"`. Không nạp cả quyển.
- Bước 4: tạo `projects/_giao-an/<tên_bài>/` và viết `giao-an.md`.
- Bước 5: `python tools\vi\giao_an.py xuat projects\_giao-an\<tên_bài>`; thêm `--plan-only` khi chỉ muốn kiểm.
- Bước 6: đọc dòng JSON. `ready` là `true` thì báo thầy cô đường dẫn hai file, số hoạt động, tổng thời lượng, các mã đã dùng, và đọc nguyên văn `warnings` cùng nội dung `can-soat.md`.
- Bảng xử lý lỗi, mỗi `error.step` một dòng: `input` → chưa có `giao-an.md`; `parse` → sửa đúng dòng `error.message` nêu; `framework` → sửa mã theo `error.fix`, **không tự đặt mã mới**; `docx` → chạy `python -m pip install -r tools/vi/requirements-vi.txt` rồi chạy lại, tối đa một lần; `write` → xin thầy cô đóng file Word rồi chạy lại.
- Điều cấm: không sửa nội dung chuyên môn của thầy cô; không sửa file phân phối chương trình; không tự đặt mã; không in ghi chú nội bộ vào giáo án; không commit gì trong `projects/`.

- [ ] **Step 5: Sửa `docs/vi/tro-ly/quy-trinh-hoi.md`**

1. Trong `## Khi nào áp dụng`, đổi "7 loại việc" thành "8 loại việc" và thêm dòng bảng:

```
| Soạn giáo án tích hợp năng lực số và AI | "kế hoạch bài dạy", "KHBD", "giáo án Word", "giáo án 5512" | [giao-an.md](giao-an.md) |
```

2. Thêm một dòng gạch đầu dòng ngay sau bảng:

```
- Chữ "giáo án" một mình là ca mơ hồ đã biết: nó có thể là file Word kế hoạch bài dạy, cũng có thể là slide. Hỏi đúng một câu trước khi làm gì khác: "Thầy cô cần file Word kế hoạch bài dạy (giáo án 5512), hay slide trình chiếu cho bài này?" Trả lời Word thì dùng giao-an.md; trả lời slide thì dùng bai-giang.md.
```

3. Thêm một dòng: `Loại việc "Soạn giáo án tích hợp năng lực số và AI" không tạo PPTX; nó ghi brief như các loại khác nhưng không đi vào quy trình của upstream.`

- [ ] **Step 6: Sửa `docs/vi/tro-ly/mau-brief.md`**

Thêm `Soạn giáo án tích hợp năng lực số và AI` vào danh sách loại việc.

- [ ] **Step 7: Chạy cả bộ test**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests`
Kỳ vọng: OK. Nếu `test_relative_markdown_links_resolve` fail, kiểm lại liên kết trong `AGENTS.vi.md` và `quy-trinh-hoi.md`.

- [ ] **Step 8: Commit**

```bash
git add docs/vi/tro-ly/giao-an.md docs/vi/tro-ly/quy-trinh-hoi.md docs/vi/tro-ly/mau-brief.md AGENTS.vi.md tools/vi/tests/test_vi_layer.py
git commit -m "docs(vi): add the lesson plan task type for agents"
```

---

### Task 8: Tài liệu cho thầy cô

**Files:**
- Create: `docs/vi/soan-giao-an.md`
- Modify: `docs/vi/xu-ly-loi.md` (thêm `## Xuất giáo án thất bại`)
- Modify: `docs/vi/bat-dau-nhanh.md` (thêm `## Soạn giáo án`)
- Modify: `docs/vi/cau-lenh-mau.md` (thêm `## Soạn giáo án`)
- Modify: `README.md`
- Modify: `tools/vi/tests/test_vi_layer.py` (thêm `LessonUserDocsTest`, thêm file vào `REQUIRED_DOCS`)

**Interfaces:**
- Consumes: các `error.step` ở Task 6.
- Produces: không có mã.

- [ ] **Step 1: Viết test trước**

Thêm `"soan-giao-an.md"` vào `REQUIRED_DOCS`, rồi thêm lớp test:

```python
class LessonUserDocsTest(unittest.TestCase):
    def test_doc_explains_inputs_outputs_and_limits(self):
        text = read("docs/vi/soan-giao-an.md")
        for phrase in ("Word", "PDF", "ảnh", "giao-an.docx", "can-soat.md", "rubric",
                       "năng lực số", "năng lực AI", "phân phối chương trình"):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, text)

    def test_doc_explains_the_routing_question(self):
        self.assertIn("slide", read("docs/vi/soan-giao-an.md"))

    def test_quick_start_mentions_the_lesson_plan_task(self):
        text = read("docs/vi/bat-dau-nhanh.md")
        headings = h2_headings(text)
        self.assertIn("## Soạn giáo án", headings)
        self.assertLess(headings.index("## Soạn giáo án"), headings.index("## Lấy file kết quả"))
        self.assertIn("(soan-giao-an.md)", section(text, "## Soạn giáo án"))

    def test_sample_commands_cover_both_flows(self):
        body = section(read("docs/vi/cau-lenh-mau.md"), "## Soạn giáo án")
        self.assertIn("năng lực số", body)
        self.assertIn("kế hoạch bài dạy", body)

    def test_troubleshooting_has_the_lesson_plan_section(self):
        text = read("docs/vi/xu-ly-loi.md")
        headings = h2_headings(text)
        self.assertIn("## Xuất giáo án thất bại", headings)
        # Đứng NGAY TRƯỚC mục đề thi: gói đề thi khoá mục đề thi ngay trước mục video,
        # và gói video khoá mục video ngay trước "Đường dẫn quá dài".
        self.assertEqual(
            headings.index("## Xuất giáo án thất bại"),
            headings.index("## Xuất đề Word thất bại") - 1,
        )
        body = section(text, "## Xuất giáo án thất bại")
        for phrase in ("requirements-vi.txt", "python-docx", "đang mở trong Word", "Dòng",
                       "giao-an.md", "mã năng lực"):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, body)

    def test_readme_mentions_the_lesson_plan_feature(self):
        self.assertIn("giáo án", section(read("README.md"), "## Làm được gì"))
```

- [ ] **Step 2: Chạy test để thấy nó fail**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests -k LessonUserDocsTest`
Kỳ vọng: FAIL vì chưa có `docs/vi/soan-giao-an.md`.

- [ ] **Step 3: Viết `docs/vi/soan-giao-an.md`**

Viết cho thầy cô. Các mục:

- `# Soạn giáo án tích hợp năng lực số và năng lực AI`
- `## Làm được gì`: từ giáo án cũ hoặc từ tên bài, ra file Word kế hoạch bài dạy 5512 đủ mục tiêu, tiến trình bốn pha, phiếu học tập và rubric đánh giá năng lực số, năng lực AI.
- `## Cách nhắn cho AI`: hai ví dụ, một cho mỗi luồng. Ghi rõ: nói "giáo án" thì AI sẽ hỏi lại Word hay **slide** — trả lời "Word" hoặc "kế hoạch bài dạy".
- `## Thầy cô cần đưa gì`: giáo án cũ dạng Word hoặc PDF; file kế hoạch dạy học hoặc **phân phối chương trình** nếu có (AI chỉ đọc, không sửa). **Ảnh chụp thì AI không đọc được.**
- `## File nhận được`: bảng — `giao-an.docx` (bản nộp) và `can-soat.md` (ghi chú nội bộ, không nộp), lưu ở `projects\_giao-an\<tên_bài>\`.
- `## Máy sẽ chặn những gì`: liệt kê bằng lời cho thầy cô hiểu — hoạt động thiếu bốn thành tố 5512; mục tiêu có năng lực số hoặc AI mà tiến trình không có; mã năng lực không đúng bảng; rubric thiếu mức. Nói rõ đây là để giáo án không bị trả lại.
- `## Môn chưa có khung mã AI`: giải thích vì sao ô mã để trống, và rằng AI không tự đặt mã.
- `## Sửa lại rồi xuất lại`: sửa `giao-an.md` rồi nhờ AI chạy lại.
- `## Giới hạn`: máy chỉ kiểm cấu trúc, không kiểm chất lượng sư phạm; chưa có khung ký duyệt và quốc hiệu; chưa sinh slide từ giáo án.

- [ ] **Step 4: Sửa `docs/vi/xu-ly-loi.md`**

Thêm mục `## Xuất giáo án thất bại` **ngay trước** `## Xuất đề Word thất bại`. Không chèn vào giữa `## Xuất đề Word thất bại`, `## Dựng video thất bại` và `## Đường dẫn quá dài`: test của hai gói trước khoá ba mục đó phải liền nhau theo đúng thứ tự. Viết theo từng mã lỗi: `input`, `parse`, `framework` (mã năng lực sai — làm theo `error.fix`, không tự đặt mã), `docx`, `write`, `internal`.

- [ ] **Step 5: Sửa `docs/vi/bat-dau-nhanh.md` và `docs/vi/cau-lenh-mau.md`**

`bat-dau-nhanh.md`: thêm `## Soạn giáo án` trước `## Lấy file kết quả`, có liên kết `(soan-giao-an.md)`.

`cau-lenh-mau.md`: thêm `## Soạn giáo án` với hai câu lệnh mẫu — "Nâng cấp giáo án Bài 5 Ammonia này thành kế hoạch bài dạy có tích hợp năng lực số và năng lực AI" và "Soạn kế hoạch bài dạy Hoá 11 Bài 5 Ammonia, 2 tiết, có tích hợp năng lực số và năng lực AI".

- [ ] **Step 6: Sửa `README.md`**

1. Mục "Làm được gì": thêm một dòng nêu soạn **giáo án** kế hoạch bài dạy có tích hợp năng lực số và năng lực AI, xuất Word.
2. Bảng tài liệu: thêm dòng `| [Soạn giáo án](docs/vi/soan-giao-an.md) | Kế hoạch bài dạy 5512 tích hợp năng lực số và AI |`.

- [ ] **Step 7: Chạy cả bộ test**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests`
Kỳ vọng: OK.

- [ ] **Step 8: Commit**

```bash
git add docs/vi/soan-giao-an.md docs/vi/xu-ly-loi.md docs/vi/bat-dau-nhanh.md docs/vi/cau-lenh-mau.md README.md tools/vi/tests/test_vi_layer.py
git commit -m "docs(vi): explain lesson plan authoring for teachers"
```

---

### Task 9: Chạy thật một bài

**Files:**
- Create: `docs/vi/phat-trien/2026-09-13-giao-an-kiem-thu.md`
- Không commit gì trong `projects/`.

**Interfaces:**
- Consumes: toàn bộ Task 1–8.
- Produces: báo cáo chạy thật.

Ràng buộc: **không mở Word**; kiểm bằng XML. Không copy file nào của chủ repo vào repo.

- [ ] **Step 1: Dựng giáo án mẫu luồng B**

Tạo `projects/_giao-an/thu-nghiem-ammonia/giao-an.md`. Nội dung: môn `Hoá học`, `grade: 11`, `periods: 2`, bốn hoạt động đủ tám khoá, hai mã NLS, hai mã AI của khung `AI-H11`, một phiếu học tập hai câu có `NH~3~` và `N~2~`, hai tiêu chí rubric ba mức, một dòng `## CAN SOAT`.

- [ ] **Step 2: Chạy và kiểm**

```bash
python tools/vi/giao_an.py xuat projects/_giao-an/thu-nghiem-ammonia
```

Kiểm: mã thoát 0; stdout đúng một dòng JSON; `files` có hai đường dẫn; `activities` là 4; `minutes` là 90; `nls_codes` và `ai_codes` đúng như trong file.

Kiểm nội dung không mở Word:

```bash
python -c "import zipfile; x=zipfile.ZipFile('projects/_giao-an/thu-nghiem-ammonia/giao-an.docx').read('word/document.xml').decode('utf-8'); print(all(s in x for s in ['I. MỤC TIÊU','III. TIẾN TRÌNH DẠY HỌC','Bước 4. Kết luận, nhận định','Hoạt động 1.']), 'subscript' in x, 'Cần thầy cô soát' not in x)"
```

Kỳ vọng: `True True True`.

- [ ] **Step 3: Kiểm phép chặn thiếu thành tố 5512**

Chép `giao-an.md` sang `giao-an.md.bak`, xoá một dòng `bao-cao:` rồi chạy lại. Kỳ vọng: `error.step` là `parse`, `message` nêu đúng số dòng và khoá `bao-cao`. Phục hồi file từ bản `.bak`.

- [ ] **Step 4: Kiểm phép chặn mã sai**

Đổi một mã AI thành `AI-H12` rồi chạy lại. Kỳ vọng: `error.step` là `framework`, `fix` liệt kê các mã của khung `AI-H11`. Phục hồi file.

- [ ] **Step 5: Kiểm luồng môn chưa có khung**

Chép sang `projects/_giao-an/thu-nghiem-ngu-van/`, đổi `subject: Ngữ văn`, đổi mục `### Nang luc AI` thành `- (chưa có khung mã cho môn này)` cộng một dòng hoạt động AI bằng lời, và đổi `ai:` thành `ai: (chưa có mã)`. Chạy và kỳ vọng mã thoát 0, `ai_codes` là danh sách rỗng.

- [ ] **Step 6: Kiểm cắt SGK**

Tạo một file Markdown mẫu ba bài trong `projects/_giao-an/_sgk/mau.md` (tự soạn, **không** dùng SGK thật để tránh chép tài liệu có bản quyền vào chỗ dùng chung), rồi:

```bash
python tools/vi/giao_an.py trich-sgk projects/_giao-an/_sgk/mau.md --bai "Bài 5"
```

Kỳ vọng: `ready` là `true`, `heading` đúng bài 5, file `sgk-trich.md` chỉ chứa phần bài 5.

- [ ] **Step 7: Viết báo cáo**

Tạo `docs/vi/phat-trien/2026-09-13-giao-an-kiem-thu.md`: lệnh đã chạy, dòng JSON thật (rút gọn đường dẫn về dạng tương đối, **không** để lộ tên thư mục riêng của chủ repo), kết quả từng phép chặn, và kết quả kiểm XML. Ghi rõ một dòng: `Chưa mở Word để xem; phần đánh giá trình bày do chủ repo tự kiểm.` Ghi thêm một dòng: `Chưa chạy trên giáo án thật của chủ repo; bản mẫu là giáo án tự soạn.`

- [ ] **Step 8: Chạy cả bộ test lần cuối**

Chạy: `python -m unittest discover -s tools/vi/tests -t tools/vi/tests`
rồi `python skills/ppt-master/scripts/attribution_guard.py; echo $?`
Kỳ vọng: test OK, guard trả 0.

Kiểm `git status --porcelain` không có gì trong `projects/`.

- [ ] **Step 9: Commit**

```bash
git add docs/vi/phat-trien/2026-09-13-giao-an-kiem-thu.md
git commit -m "docs(vi): record the real run of the lesson plan exporter"
```

---

## Tự soát của người viết plan

**Phủ spec:** §1 tiêu chí 1–2 → Task 7 (lượt hỏi) + Task 6 (lệnh) + Task 9 (chạy thật); tiêu chí 3 → Task 5 và test thể thức; tiêu chí 4 → Task 2 (phép chặn cấu trúc) và Task 3 (phép chặn mã); tiêu chí 5 → Task 1 và Task 3; tiêu chí 6 → Task 4 và Task 7 bước 3; tiêu chí 7 → Task 7 bước 4 và 5 cộng test `test_common_rules_name_the_ambiguous_lesson_keyword`; tiêu chí 8 → Global Constraints và Task 9 bước 8; tiêu chí 9 → mọi task kết bằng bước chạy cả bộ test. §4.4 năm phép kiểm đang khoá → Task 7 Step 1 và Task 8 Step 1. §6.1–6.2 ngữ pháp → Task 2. §6.3 mười sáu phép kiểm → Task 2 (cấu trúc, thời lượng, phiếu) và Task 3 (mã) và Task 6 (đường dẫn dài, ghi đè). §6.4 khuôn bảng mã → Task 1. §6.5 lệnh và JSON → Task 6. §6.6 câu hỏi → Task 7 Step 3. §7 mọi ca biên → Task 2, Task 3, Task 4, Task 6. §8 → Task 9. §9 phát hành → ngoài plan, làm sau khi review cuối xanh.

**Không có chỗ trống:** không có "TBD"; mọi bước code có khối mã đầy đủ; các bước viết tài liệu nêu đủ nội dung bắt buộc và đúng những chuỗi mà test sẽ tìm.

**Nhất quán kiểu:** `Lesson`/`Activity`/`Worksheet`/`Rubric` khai ở Task 2 và dùng đúng tên trường ở Task 5; `frameworks.Frameworks` khai ở Task 1, `frameworks.validate` thêm ở Task 3 và gọi ở Task 6; `sgk.extract` khai ở Task 4 và gọi ở Task 6; `docx_build.build`/`write_review` khai ở Task 5 và gọi ở Task 6; `load_docx_build` khai ở Task 6 và được test mock ở cùng task; `word_parts.base` là của gói vi.5, Task 5 chỉ tiêu thụ và có bước dừng nếu chưa có.
