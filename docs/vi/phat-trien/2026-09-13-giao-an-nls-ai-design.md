# Thiết kế: Soạn giáo án tích hợp năng lực số và năng lực AI (v6.3.2-vi.6)

Ngày 2026-09-13. Gói này thêm loại việc thứ 8 cho lớp Việt: soạn Kế hoạch bài dạy (giáo án) theo Công văn 5512, tích hợp Năng lực số và Năng lực trí tuệ nhân tạo, xuất ra file Word đúng thể thức. Đầu ra không phải PPTX.

Nguồn gốc: chủ repo đã có hai skill chạy được ở thư mục riêng (`ke_hoach_bai_day`, `phan_phoi_chuong_trinh`) cùng bốn file khung. Gói này lấy phần khung làm tài liệu cho AI, và đảo lại cách làm: nội dung là dữ liệu, code chỉ là bộ dựng.

## 1. Tiêu chí thành công

1. Thầy cô đưa một giáo án Word hoặc PDF đã có; sau một lượt hỏi, nhận được file Word mới **giữ nguyên nội dung chuyên môn** và đã bổ sung mục tiêu năng lực số, mục tiêu năng lực AI, ít nhất một hoạt động AI có đối chứng, và rubric đánh giá.
2. Thầy cô chỉ có tên bài; AI soạn giáo án mới đủ khung 5512.
3. File Word đúng thể thức đang dùng: A4 dọc, lề 2/2/2,5/2 cm, Times New Roman 14pt, giãn dòng 1,3; phiếu học tập căn lề trái, mỗi câu một đoạn riêng.
4. Công cụ **chặn** giáo án thiếu bất kỳ thành tố nào của Công văn 5512, và chặn giáo án chỉ nhắc năng lực số hoặc AI ở mục tiêu mà không có trong tiến trình.
5. Mã NLS và mã AI luôn thuộc khung chuẩn. Môn chưa có khung AI thì AI **không tự đặt mã**.
6. AI đọc được SGK PDF mà không nạp cả quyển: chuyển một lần thành Markdown, rồi cắt đúng phần của bài.
7. Câu lệnh có chữ "giáo án" luôn được hỏi lại một câu: cần file Word kế hoạch bài dạy, hay slide trình chiếu.
8. Giáo án, đề thi và SGK của thầy cô không bao giờ lên GitHub.
9. Toàn bộ test của lớp Việt vẫn xanh.

## 2. Hiện trạng đã kiểm

- Thư mục tài nguyên của chủ repo có `.agents/skills/ke_hoach_bai_day/` (SKILL.md + 3 file tham khảo + 2 script) và `.agents/skills/phan_phoi_chuong_trinh/` (SKILL.md + 2 file khung + 1 script).
- `generate_khbd.py` là một hàm `build_lesson1_docx()` dài 347 dòng, **toàn bộ nội dung bài học nằm trong code**; `generate_khbd_bai2.py` là bản sao 53 KB cho bài 2. Thêm một bài là thêm một script. Gói này không đi theo cách đó.
- Skill đó dùng dấu `{sub:...}`, `{sup:...}`, `{b:...}`, `{i:...}`. Lớp Việt trong repo dùng một quy ước duy nhất `H~2~SO~4~`, `m/s^2^`, `**in đậm**` (đã chốt ở gói đề thi vi.5).
- Khung NLS có 11 mã: `1.1.NC1a`, `1.1.NC1b`, `1.2.NC1a`, `1.3.NC1a`, `2.1.NC1a`, `2.2.NC1a`, `3.1.NC1a`, `3.2.NC1a`, `4.1.NC1a`, `5.1.NC1a`, `5.3.NC1a`. Khung AI có `AI-H10`, `AI-H11.1`–`AI-H11.4`, `AI-H12`, `AI-H-Chuyên`; **chỉ môn Hoá học có khung**.
- `docs/vi/tro-ly/quy-trinh-hoi.md` hiện xếp `"giáo án"` và `"bài dạy"` là dấu hiệu của loại việc **Bài giảng** (ra PPTX). File đó đã có quy tắc "thuộc hai loại thì hỏi đúng một câu hỏi chọn loại việc" — gói này dùng chính quy tắc đó.
- `skills/ppt-master/scripts/source_to_md.py` nhận `-t {auto,pdf,doc,excel,pptx,web,markdown,text}`; đọc được `.docx`, `.pdf`, `.xlsx`. Không có backend cho ảnh.
- Gói đề thi vi.5 đã có plan nhưng **chưa code**, nên tầng dựng Word tách ra dùng chung được ngay, không phải viết lại.

## 3. Quyết định

| Quyết định | Lý do |
|---|---|
| Chỉ soạn giáo án; PPCT/KHDH chỉ **đọc**, không sửa | Chủ repo chọn. Việc sửa bảng PPCT cả năm là một engine khác (bảng Word ngang + Excel), để gói riêng |
| Dùng được cho **mọi môn** | Chủ repo chọn. Khung NLS vốn không phụ thuộc môn |
| Môn chưa có khung AI: để trống ô mã, ghi đúng chuỗi `(chưa có khung mã cho môn này)` | Đặt mã chính thức là việc của Bộ và của tổ chuyên môn, không phải của AI. Bịa mã `AI-P11` là tạo dữ liệu giả trong hồ sơ chuyên môn |
| Hai luồng: nâng cấp giáo án có sẵn, và soạn mới | Chủ repo chọn. Luồng nâng cấp giữ nội dung chuyên môn của thầy cô |
| Đọc SGK PDF theo cách "chuyển một lần, cắt nhiều lần" | Một quyển SGK 19–22 MB; nạp cả quyển vào ngữ cảnh là không khả thi |
| Câu lệnh có "giáo án" thì **luôn hỏi** Word hay slide | Chủ repo chọn. Không tự đoán, vì đoán sai là làm lại từ đầu |
| Một file nguồn `giao-an.md` sinh ra file Word | Giống gói đề thi: nội dung là dữ liệu; sửa rồi xuất lại, không soạn lại |
| Bảng mã chỉ có một nguồn duy nhất là file tài liệu; module kiểm mã **đọc mã từ file đó** | Khai mã hai chỗ thì chắc chắn vênh |
| Mục "cần soát" ghi ra `can-soat.md` riêng, **không** in vào giáo án | Giáo án là hồ sơ thầy cô nộp cho trường; ghi chú nội bộ không được nằm trong đó |
| Tầng dựng Word dùng chung ở `tools/vi/word_parts/` | Hai gói cùng dựng Word bằng `python-docx`; sửa plan vi.5 để tạo tầng này ngay từ đầu |
| Khung NLS/AI đưa vào repo công khai, có dòng ghi công tác giả | Chủ repo chọn. Không có khung trong repo thì bộ kiểm mã không chạy được ở máy người khác |

## 4. File

### 4.1 Tầng dùng chung (sửa trong plan vi.5)

| File | Trách nhiệm |
|---|---|
| `tools/vi/word_parts/__init__.py` | Gói con |
| `tools/vi/word_parts/inline.py` | Tách `~chỉ số dưới~`, `^chỉ số trên^`, `**in đậm**`; thay cho `de_thi_parts/inline.py` |
| `tools/vi/word_parts/base.py` | Tạo document (khổ giấy, lề, font, giãn dòng), ghi đoạn có kiểu, viền bảng, độ rộng cột, ô bảng, chân trang số trang |

### 4.2 File mới của gói này

| File | Trách nhiệm |
|---|---|
| `tools/vi/giao_an.py` | Điểm vào; hai lệnh con `xuat` và `trich-sgk`; in đúng một dòng JSON |
| `tools/vi/giao_an_parts/__init__.py` | Gói con |
| `tools/vi/giao_an_parts/parse.py` | Đọc `giao-an.md`, chạy các phép kiểm cấu trúc, lỗi kèm số dòng |
| `tools/vi/giao_an_parts/frameworks.py` | Đọc bảng mã từ file tài liệu, chọn khung theo môn/lớp, kiểm mã |
| `tools/vi/giao_an_parts/docx_build.py` | Dựng file Word 5512 |
| `tools/vi/giao_an_parts/sgk.py` | Cắt phần một bài ra khỏi SGK đã chuyển Markdown |
| `docs/vi/tro-ly/giao-an.md` | Loại việc thứ 8 cho AI |
| `docs/vi/tro-ly/nang-luc-so-va-ai.md` | Khung NLS, khung AI theo môn, nguyên tắc tích hợp — **nguồn mã duy nhất** |
| `docs/vi/soan-giao-an.md` | Hướng dẫn cho thầy cô |
| `tools/vi/tests/test_giao_an.py` | Test của gói |

### 4.3 File sửa

| File | Sửa gì |
|---|---|
| `AGENTS.vi.md` | Mục 3 thêm câu lệnh kích hoạt; bảng mục 10 thành 8 loại việc; thêm mục 13 làm mục cuối |
| `docs/vi/tro-ly/quy-trinh-hoi.md` | Thêm dòng thứ 8; ghi rõ cặp mơ hồ "giáo án" kèm câu hỏi mẫu |
| `docs/vi/tro-ly/mau-brief.md` | Thêm loại việc thứ 8 |
| `docs/vi/xu-ly-loi.md` | Thêm mục "Xuất giáo án thất bại" |
| `docs/vi/bat-dau-nhanh.md`, `docs/vi/cau-lenh-mau.md` | Thêm mục soạn giáo án |
| `README.md`, `CHANGELOG-VI.md` | Dòng phiên bản, mục "Làm được gì", mục `6.3.2-vi.6` |
| `tools/vi/tests/test_vi_layer.py` | 8 loại việc; hai file hướng dẫn mới; mục 13 là mục cuối |

### 4.4 Test đang khoá sẽ phải đổi

1. `giao-an.md` **không** vào danh sách sáu file hướng dẫn có `## Khổ slide`; nó dùng bộ mục riêng của loại việc ra văn bản, giống `de-khtn-tieng-anh.md`.
2. Mọi chỗ đếm loại việc đổi thành 8.
3. Phép kiểm thứ tự mục cuối của `AGENTS.vi.md` đổi thành mục 10 → 11 → 12 → 13.
4. Bảng nhận diện trong `quy-trinh-hoi.md` có đúng 8 dòng.
5. Quy tắc "không dùng liên kết Markdown trong `docs/vi/tro-ly/`" áp cho cả hai file mới.

## 5. Luồng

### 5.1 Định tuyến chữ "giáo án"

Câu lệnh có chữ "giáo án" là ca mơ hồ đã biết. AI hỏi đúng một câu trước khi làm gì khác:

> Thầy cô cần file Word kế hoạch bài dạy (giáo án 5512), hay slide trình chiếu cho bài này?

"kế hoạch bài dạy", "KHBD", "giáo án Word" là rõ ràng, đi thẳng vào loại việc này. "bài giảng", "slide", "trình chiếu" đi vào loại việc Bài giảng như hiện tại.

### 5.2 Luồng A — nâng cấp giáo án có sẵn

1. Đọc `docs/vi/tro-ly/quy-trinh-hoi.md`, `docs/vi/tro-ly/giao-an.md`, `docs/vi/tro-ly/nang-luc-so-va-ai.md`.
2. Đọc giáo án thầy cô đưa: `source_to_md.py <file> -o <thư_mục_tạm>`. Thầy cô đưa ảnh: nói rõ không đọc được ảnh, xin PDF hoặc Word.
3. Có file PPCT/KHDH thì đọc luôn để lấy tuần, tiết thứ, số tiết, yêu cầu cần đạt và **mã NLS đã khai** cho bài đó. Dùng đúng mã đó, không chọn mã mới.
4. Hỏi một lượt (mục 6.5), chờ trả lời.
5. Viết `projects/_giao-an/<tên_bài>/giao-an.md`: **giữ nguyên** kiến thức, nội dung hoạt động và bài tập của thầy cô; chỉ thêm mục tiêu NLS, mục tiêu AI, dòng `nls:`/`ai:` cho các hoạt động phù hợp, một hoạt động AI có đối chứng, và rubric.
6. Chạy `giao_an.py xuat`, đọc JSON, báo thầy cô đường dẫn file và nội dung `can-soat.md`.

### 5.3 Luồng B — soạn mới

Như luồng A, bỏ bước 2. Nếu cần nội dung SGK: chuyển quyển SGK một lần rồi cắt phần của bài (mục 6.4).

### 5.4 Điều AI không được làm

- Không sửa, rút gọn hay "viết lại cho hay hơn" nội dung chuyên môn trong giáo án gốc của thầy cô.
- Không tự đặt mã AI cho môn chưa có khung.
- Không dùng mã NLS ngoài bảng của Bộ.
- Không sửa file PPCT/KHDH của thầy cô.
- Không in ghi chú nội bộ vào file giáo án.
- Không commit bất cứ gì trong `projects/`.
- Không chạy `project_manager.py init`, không tạo SVG, không chạm `skills/`.

## 6. Hợp đồng

### 6.1 Khối thông tin đầu `giao-an.md`

Giữa hai dòng `---`, mỗi dòng một cặp `khoá: giá trị`:

| Khoá | Bắt buộc | Ý nghĩa |
|---|---|---|
| `school` | có | Tên trường |
| `subject` | có | Tên môn, ví dụ `Hoá học` |
| `grade` | có | Lớp, chỉ chữ số |
| `lesson` | có | Tên bài, ví dụ `Bài 5. Ammonia và một số hợp chất ammonium` |
| `periods` | có | Số tiết, chỉ chữ số |
| `department` | không | Tổ chuyên môn |
| `teacher` | không | Họ tên giáo viên |
| `track` | không | `chuyen` khi dạy hệ chuyên |
| `week` | không | Tuần dạy |
| `period_numbers` | không | Tiết thứ trong phân phối chương trình |
| `school_year` | không | Năm học |

### 6.2 Các mục của `giao-an.md`

Mục h2 viết không dấu, theo đúng thứ tự: `## MUC TIEU`, `## THIET BI`, `## TIEN TRINH`, `## PHIEU HOC TAP` (tuỳ chọn), `## RUBRIC`, `## CAN SOAT` (tuỳ chọn, phải là mục cuối).

`## MUC TIEU` có sáu mục h3 không dấu, theo thứ tự: `### Kien thuc`, `### Nang luc chung`, `### Nang luc dac thu`, `### Nang luc so`, `### Nang luc AI`, `### Pham chat`. Thân mỗi mục là các dòng `- `.

Dòng trong `### Nang luc so` và `### Nang luc AI` phải mở đầu bằng mã rồi dấu `— `:

```
### Nang luc so
- 1.1.NC1a — Tìm kiếm, chọn lọc dữ liệu về ứng dụng của ammonia.
- 5.1.NC1a — Dùng thí nghiệm ảo PhET kiểm chứng chuyển dịch cân bằng.

### Nang luc AI
- AI-H11.2 — Viết prompt để AI dự đoán chiều chuyển dịch cân bằng khi tăng áp suất.
- AI-H11.4 — Đối chiếu kết quả AI với thí nghiệm, chỉ ra chỗ AI nói sai.
```

Môn chưa có khung AI: mục `### Nang luc AI` có đúng một dòng đầu là `- (chưa có khung mã cho môn này)`, các dòng sau là hoạt động AI viết bằng lời, không có mã.

`## THIET BI` có hai mục h3: `### Giao vien`, `### Hoc sinh`; thân là các dòng `- `.

`## TIEN TRINH` có một mục h3 cho mỗi hoạt động; tên hoạt động là chữ tự do (được dùng dấu). Số hoạt động do thứ tự trong file quyết định. Mỗi hoạt động gồm các dòng `khoá: giá trị`:

| Khoá | Bắt buộc | Ý nghĩa |
|---|---|---|
| `thoi-luong` | có | Số phút, chỉ chữ số |
| `muc-tieu` | có | Mục tiêu của hoạt động |
| `noi-dung` | có | Nhiệm vụ giao cho học sinh |
| `san-pham` | có | Kết quả cụ thể học sinh hoàn thành |
| `chuyen-giao` | có | Bước 1: chuyển giao nhiệm vụ |
| `thuc-hien` | có | Bước 2: học sinh thực hiện |
| `bao-cao` | có | Bước 3: báo cáo, thảo luận |
| `ket-luan` | có | Bước 4: kết luận, nhận định |
| `nls` | không | Một hoặc nhiều mã NLS, cách nhau bằng dấu phẩy |
| `ai` | không | Một hoặc nhiều mã AI, cách nhau bằng dấu phẩy |

Khoá lặp lại thành nhiều đoạn liên tiếp: hai dòng `noi-dung:` cho ra hai đoạn. Đây là điểm khác gói đề thi, nơi khoá lặp là lỗi.

Môn chưa có khung AI: hoạt động AI vẫn phải có dòng `ai:`, và giá trị hợp lệ duy nhất lúc đó là đúng chuỗi `(chưa có mã)`. Nhờ vậy phép kiểm "phải có hoạt động AI trong tiến trình" chạy được cho mọi môn.

`## PHIEU HOC TAP` có một mục h3 cho mỗi phiếu; thân là các dòng `- `, mỗi dòng thành **một đoạn căn lề trái** trong file Word.

`## RUBRIC` có một mục h3 cho mỗi tiêu chí; thân là đúng ba dòng mở đầu bằng `- Mức 1:`, `- Mức 2:`, `- Mức 3:`.

`## CAN SOAT` là các dòng `- `; nội dung đi vào `can-soat.md`, **không** vào file giáo án.

Trong mọi dòng văn bản chỉ dùng ba dấu đánh dấu `~chỉ số dưới~`, `^chỉ số trên^`, `**in đậm**`.

### 6.3 Các phép kiểm

| Phép kiểm | Kết quả |
|---|---|
| Hoạt động thiếu một trong tám khoá bắt buộc | Lỗi `parse`, nêu số dòng và tên khoá |
| `### Nang luc so` không có dòng nào | Lỗi `parse` |
| `### Nang luc AI` không có dòng nào | Lỗi `parse` |
| Không hoạt động nào có `nls:` | Lỗi `parse` |
| Không hoạt động nào có `ai:` | Lỗi `parse` |
| Mã NLS không thuộc bảng của Bộ | Lỗi `framework`, liệt kê mã hợp lệ |
| Mã AI không thuộc khung của môn và lớp đó | Lỗi `framework` |
| Môn không có khung AI nhưng mục AI vẫn ghi mã | Lỗi `framework`, nhắc dùng chuỗi `(chưa có khung mã cho môn này)` |
| Hoạt động dùng mã chưa khai trong mục tiêu | Lỗi `framework`, nêu mã và tên hoạt động |
| `## RUBRIC` có ít hơn 2 tiêu chí | Lỗi `parse` |
| Tiêu chí rubric không đủ ba mức, hoặc sai thứ tự | Lỗi `parse` |
| Tổng `thoi-luong` khác `periods × 45` | Cảnh báo, vẫn xuất file |
| Hoạt động nhắc một phiếu học tập không có trong `## PHIEU HOC TAP` | Cảnh báo |
| Thiếu `## PHIEU HOC TAP` | Cảnh báo |
| Đường dẫn thư mục dài hơn 200 ký tự | Cảnh báo |
| File Word đã có | Cảnh báo, nêu tên file bị ghi đè |

### 6.4 Bảng mã trong file tài liệu

`docs/vi/tro-ly/nang-luc-so-va-ai.md` là nguồn mã duy nhất. `frameworks.py` đọc theo hai khuôn cố định:

- Mục `## Khung năng lực số (mọi môn)`: mỗi mã là một dòng `` - `<mã>` — <mô tả> ``.
- Mục `## Khung năng lực AI theo môn`: mỗi khung là một mục h3 theo khuôn `### <môn> — <phạm vi> — khung \`<tên khung>\``, trong đó `<phạm vi>` là `lớp 10`, `lớp 11`, `lớp 12` hoặc `hệ chuyên`; các dòng `` - `<mã>` — <mô tả> `` sau đó thuộc khung ấy.

Chọn khung: so `subject` (không phân biệt chữ hoa) và phạm vi (`hệ chuyên` khi `track: chuyen`, còn lại `lớp <grade>`). Không tìm được khung nào thì môn đó coi như chưa có khung AI.

File có một dòng ghi công: khung do Lương Hải Anh — 2Anh AI Education biên soạn, dựa trên Bảng mã Năng lực số của Bộ GD&ĐT và Phụ lục III & IV về giáo dục AI.

### 6.5 Lệnh và JSON

```
python tools/vi/giao_an.py xuat <thư_mục_giáo_án> [--plan-only]
python tools/vi/giao_an.py trich-sgk <sgk.md> --bai "<tên bài>" [--ra <file.md>]
```

`xuat` in một dòng JSON:

```json
{
  "ready": true,
  "files": ["...giao-an.docx", "...can-soat.md"],
  "activities": 4,
  "periods": 2,
  "minutes": 90,
  "nls_codes": ["1.1.NC1a", "5.1.NC1a"],
  "ai_codes": ["AI-H11.2", "AI-H11.4"],
  "warnings": [],
  "error": null
}
```

`trich-sgk` in một dòng JSON:

```json
{
  "ready": true,
  "files": ["...sgk-trich.md"],
  "heading": "Bài 5. Ammonia và một số hợp chất ammonium",
  "size_kb": 18.4,
  "warnings": [],
  "error": null
}
```

Khi lỗi, `ready` là `false` và `error` là `{"step": ..., "message": ..., "fix": ...}` với `step` thuộc `{input, parse, framework, docx, write, internal}`:

| `step` | Nghĩa |
|---|---|
| `input` | Không có thư mục, không có `giao-an.md`, không có file SGK, hoặc tham số sai |
| `parse` | `giao-an.md` sai cấu trúc; `message` nêu số dòng |
| `framework` | Mã NLS hoặc mã AI sai, hoặc dùng mã chưa khai trong mục tiêu |
| `docx` | Thiếu `python-docx`; `fix` là lệnh cài |
| `write` | Không ghi được file (đang mở trong Word, hết đĩa, đường dẫn quá dài) |
| `internal` | Lỗi ngoài dự kiến. Có bậc này để stdout không bao giờ trống |

### 6.6 Câu hỏi cho thầy cô

Tối đa 7 câu, theo cách hỏi của `quy-trinh-hoi.md`. Câu hỏi định tuyến ở mục 5.1 **không tính** vào 7 câu này.

1. Môn, lớp, tên bài, và bài dạy trong bao nhiêu tiết? *Gợi ý: lấy từ câu lệnh nếu đã có; chỉ hỏi phần còn thiếu.*
2. Thầy cô có giáo án cũ của bài này không, hay em soạn mới? *Gợi ý: có thì gửi file Word hoặc PDF; ảnh chụp thì em không đọc được.*
3. Có file kế hoạch dạy học hoặc phân phối chương trình để em lấy tuần, tiết thứ và mã năng lực số đã khai không? *Gợi ý: không có thì em tự chọn mã theo khung của Bộ.*
4. Lớp có thiết bị gì để làm hoạt động số (máy tính, điện thoại, mạng, phòng máy)? *Gợi ý: học sinh dùng điện thoại theo nhóm, có wifi.*
5. Thầy cô muốn dùng công cụ AI nào với học sinh? *Gợi ý: một trợ lý chat để dự đoán, kèm PhET hoặc thí nghiệm thật để kiểm chứng.*
6. Bài này có thí nghiệm thật không, và có phiếu học tập không? *Gợi ý: có một thí nghiệm đối chứng và một phiếu học tập.*
7. Trường mình yêu cầu thể thức riêng nào không (ký duyệt, quốc hiệu, năm học)? *Gợi ý: giữ thể thức A4 dọc, Times New Roman 14pt, giãn dòng 1,3.*

Mục "Tạo nhanh": câu 1 và câu 2.

## 7. Ca biên

| Tình huống | Xử lý |
|---|---|
| `giao-an.md` không có hoạt động nào | Lỗi `parse` |
| Hoạt động có `thoi-luong: 0` | Lỗi `parse` |
| `grade` không phải chữ số | Lỗi `parse` |
| Mục h2 sai thứ tự hoặc lặp | Lỗi `parse` |
| `## CAN SOAT` không ở cuối | Lỗi `parse` |
| Môn Ngữ văn (chưa có khung AI) mà mục AI ghi `AI-H11.2` | Lỗi `framework` |
| Môn Ngữ văn, mục AI ghi `(chưa có khung mã cho môn này)`, hoạt động không có `ai:` | Lỗi `parse` — vẫn phải có hoạt động AI, chỉ là không có mã; hoạt động dùng `ai: (chưa có mã)` |
| Thiếu `python-docx` | Lỗi `docx` kèm lệnh cài |
| File Word đang mở trong Word | Lỗi `write` |
| `trich-sgk` không tìm thấy tiêu đề bài | Lỗi `parse`, kèm ba tiêu đề gần nhất tìm được |
| `trich-sgk` cắt ra hơn 200 KB | Cảnh báo: có thể đã cắt sang bài sau |
| Đường dẫn có dấu tiếng Việt và khoảng trắng | Phải chạy được; mọi đường dẫn `resolve()` trước khi dùng |
| `--plan-only` | Không ghi file nào, `files` là danh sách rỗng |
| Giáo án 8 hoạt động, 4 tiết | Vẫn xuất xong; không giới hạn số hoạt động |

Ghi chú về ca môn chưa có khung: hoạt động vẫn bắt buộc có dòng `ai:`, và giá trị hợp lệ khi môn không có khung là đúng chuỗi `(chưa có mã)`. Nhờ vậy phép kiểm "phải có hoạt động AI" vẫn chạy cho mọi môn.

## 8. Kiểm thử

### 8.1 Test tự động

- `frameworks.py`: đọc đúng 11 mã NLS và bốn khung AI từ file tài liệu; `framework_for("Hoá học", "11", "")` trả `AI-H11`; `framework_for("Hoá học", "11", "chuyen")` trả `AI-H-Chuyên`; `framework_for("Ngữ văn", "11", "")` trả `None`.
- `parse.py`: một giáo án hợp lệ hai tiết bốn hoạt động; và từng ca biên ở mục 7, mỗi ca assert đúng số dòng trong `message`.
- Các phép kiểm mục 6.3: mỗi phép một test riêng.
- `sgk.py`: cắt đúng phần giữa hai tiêu đề bài; không tìm thấy thì báo kèm tiêu đề gần nhất; cắt quá lớn thì cảnh báo.
- `docx_build.py`: mở lại file đã xuất và soi XML — A4 dọc, lề 2/2/2,5/2 cm, Times New Roman 14pt, giãn dòng 1,3; **mỗi câu trong phiếu học tập là một đoạn riêng và căn lề trái**; rubric là bảng 4 cột; chỉ số dưới có `vertAlign="subscript"`; file giáo án **không** chứa chuỗi của `## CAN SOAT`.
- `giao_an.py`: hợp đồng JSON một dòng cho cả hai lệnh con; từng `error.step`; `--plan-only` không ghi file; đường dẫn có dấu tiếng Việt.

### 8.2 Chạy thật

Một bài thật: lấy giáo án **Bài 5 Ammonia và một số hợp chất ammonium** mà chủ repo đã có, chạy luồng A, rồi so mục lục và số hoạt động với bản đang dùng. Kiểm file bằng XML. **Không mở Word** — phần đánh giá trình bày do chủ repo tự kiểm.

Không commit bất kỳ file nào của thầy cô: giáo án, đề thi, SGK PDF đều nằm ngoài repo hoặc trong `projects/` đã gitignore.

## 9. Phát hành

Nhánh `feat/vi-giao-an`, phiên bản `v6.3.2-vi.6`. Làm **sau** gói đề thi `v6.3.2-vi.5`, vì gói đó tạo ra tầng `tools/vi/word_parts/`.

## 10. Ngoài phạm vi

- Sửa file PPCT/KHDH và chèn cột năng lực AI vào bảng cả năm: gói riêng sau. Gói này chỉ **đọc** PPCT.
- Sinh slide từ giáo án.
- Xuất PDF.
- Đặt mã chỉ báo AI cho môn chưa có khung.
- Đọc ảnh chụp giáo án (OCR).
- Đánh giá, chấm điểm, sổ điểm.

## 11. Rủi ro

- Khung AI chỉ có cho môn Hoá. Môn khác dùng được gói này nhưng ô mã để trống, nên hồ sơ có thể bị tổ chuyên môn hỏi lại. Đây là giới hạn của khung hiện hành, không phải của công cụ.
- Luồng A phụ thuộc chất lượng chuyển đổi `.docx`/`.pdf` sang Markdown. Giáo án có bảng lồng nhau hoặc công thức dạng ảnh sẽ mất định dạng; AI phải nói rõ chỗ nào không đọc được thay vì tự viết bù.
- Công cụ kiểm được **cấu trúc**, không kiểm được chất lượng sư phạm. Một giáo án đủ tám khoá vẫn có thể nhạt; phần đó vẫn là việc của thầy cô.
- Thể thức trường có thể khác: bản này chưa có khung ký duyệt và chưa chèn quốc hiệu.
- Chưa chạy trên một bài thật của chủ repo tại thời điểm viết spec.
