# Chạy thật một bài (Task 9)

Kiểm thử Task 9 của kế hoạch `2026-09-13-giao-an-plan`: dựng giáo án mẫu luồng B,
chạy `tools/vi/giao_an.py xuat` và `trich-sgk` trên máy thật, kiểm cả năm phép
chặn (thiếu thành tố 5512, mã sai, môn chưa có khung, cắt SGK), rồi chạy luồng A
trên một giáo án Bài 5 thật của chủ repo. Mọi file nguồn và `.docx` sinh ra nằm
dưới `projects/_giao-an/` (nằm trong `.gitignore`, không commit). Chưa mở Word
để xem; phần đánh giá trình bày do chủ repo tự kiểm.

Python dùng: không có `venv\Scripts\python.exe` ở gốc repo → dùng `python`.

## Bước 1–2 — Giáo án mẫu luồng B và kiểm

Tạo `projects/_giao-an/thu-nghiem-ammonia/giao-an.md`: Hoá học, lớp 11, 2 tiết,
4 hoạt động (đủ 8 khoá mỗi hoạt động), 2 mã năng lực số (`1.1.NC1a`, `5.1.NC1a`),
2 mã AI của khung `AI-H11` (`AI-H11.2`, `AI-H11.4`), 1 phiếu học tập có câu chứa
`NH~3~` và `N~2~`, rubric 2 tiêu chí 3 mức, 1 dòng `## CAN SOAT`.

```
python tools/vi/giao_an.py xuat projects/_giao-an/thu-nghiem-ammonia
```
```json
{"ready": true, "files": ["projects/_giao-an/thu-nghiem-ammonia/giao-an.docx", "projects/_giao-an/thu-nghiem-ammonia/can-soat.md"], "warnings": ["Ghi đè file có sẵn: giao-an.docx", "Ghi đè file có sẵn: can-soat.md"], "error": null, "activities": 4, "periods": 2, "minutes": 90, "nls_codes": ["1.1.NC1a", "5.1.NC1a"], "ai_codes": ["AI-H11.2", "AI-H11.4"]}
```
Mã thoát 0, stdout đúng một dòng JSON (đã tách stdout/stderr để xác nhận dòng
"Đã xuất giáo án và file ghi chú." nằm ở stderr). `files` có 2 đường dẫn,
`activities` = 4, `minutes` = 90, `nls_codes`/`ai_codes` đúng như khai. Hai
cảnh báo "Ghi đè file có sẵn" xuất hiện vì `giao-an.docx` và `can-soat.md` đã
tồn tại từ lần chạy trước đó trong cùng phiên kiểm thử — đây là hành vi cố ý
khi ghi đè file cũ, không phải lỗi.

Kiểm nội dung không mở Word:
```
python -c "import zipfile; x=zipfile.ZipFile('projects/_giao-an/thu-nghiem-ammonia/giao-an.docx').read('word/document.xml').decode('utf-8'); print(all(s in x for s in ['I. MỤC TIÊU','III. TIẾN TRÌNH DẠY HỌC','Bước 4. Kết luận, nhận định','Hoạt động 1.']), 'subscript' in x, 'Cần thầy cô soát' not in x)"
```
→ `True True True`, đúng kỳ vọng.

## Bước 3 — Phép chặn thiếu thành tố 5512

Sao lưu `giao-an.md` → `.bak`, xoá dòng `bao-cao:` của Hoạt động 1, chạy lại:
```json
{"error": {"step": "parse", "message": "Dòng 46: hoạt động 'Hoạt động 1: Mở đầu' thiếu khoá: bao-cao", "fix": "Sửa đúng dòng đó trong giao-an.md theo docs/vi/tro-ly/giao-an.md rồi chạy lại."}}
```
Mã thoát 1. `error.step` là `parse`, thông báo nêu đúng số dòng (dòng tiêu đề
hoạt động) và đúng tên khoá còn thiếu (`bao-cao`). Đã phục hồi từ `.bak`, chạy
lại xác nhận `ready: true`.

## Bước 4 — Phép chặn mã sai

Đổi mã ở cả dòng khai mục tiêu (`### Nang luc AI`) lẫn dòng hoạt động
(`AI-H11.2` → `AI-H12`), chạy lại:
```json
{"error": {"step": "framework", "message": "Mã AI không thuộc khung AI-H11 của môn Hoá học lớp 11: AI-H12", "fix": "Dùng một trong các mã: AI-H11.1, AI-H11.2, AI-H11.3, AI-H11.4"}}
```
Mã thoát 1. `error.step` là `framework`, `fix` liệt kê đủ 4 mã của khung
`AI-H11`. Đã phục hồi từ `.bak`, chạy lại xác nhận `ready: true`.

Ghi nhận thêm (giới hạn của công cụ, không phải chỉ là điểm tài liệu): nếu
chỉ đổi mã ở dòng `ai:` của hoạt động mà không khai mã đó ở mục tiêu
(`### Nang luc AI`), `frameworks.validate` báo lỗi khác cũng ở bước
`framework` — "Hoạt động '...' dùng mã AI-H12 chưa khai ở mục
'### Nang luc AI'" — nhưng `fix` của lỗi này ("Thêm mã đó vào mục tiêu, hoặc
bỏ khỏi hoạt động.") không liệt kê các mã hợp lệ của khung `AI-H11`, khác với
lỗi "mã không thuộc khung" ở trên vốn liệt kê đủ. Export vẫn bị chặn đúng
(mã thoát 1) trong cả hai trường hợp, chỉ khác nhau ở độ đầy đủ của `fix`.

## Bước 5 — Luồng môn chưa có khung (Ngữ văn)

Sao chép sang `projects/_giao-an/thu-nghiem-ngu-van/giao-an.md`: đổi
`subject: Ngữ văn`, đổi mục `### Nang luc AI` thành đúng chuỗi
`- (chưa có khung mã cho môn này)` cộng một dòng hoạt động AI bằng lời, đổi
các dòng `ai:` thành `ai: (chưa có mã)`.

```
python tools/vi/giao_an.py xuat projects/_giao-an/thu-nghiem-ngu-van
```
```json
{"ready": true, "files": ["projects/_giao-an/thu-nghiem-ngu-van/giao-an.docx", "projects/_giao-an/thu-nghiem-ngu-van/can-soat.md"], "warnings": ["Ghi đè file có sẵn: giao-an.docx", "Ghi đè file có sẵn: can-soat.md"], "error": null, "activities": 4, "periods": 2, "minutes": 90, "nls_codes": ["1.1.NC1a", "5.1.NC1a"], "ai_codes": []}
```
Mã thoát 0, `ai_codes` là danh sách rỗng — đúng kỳ vọng. Hai cảnh báo "Ghi đè
file có sẵn" xuất hiện cùng lý do như ở Bước 1–2 (file đã tồn tại từ lần
chạy trước).

## Bước 6 — Kiểm cắt SGK

Tạo file mẫu tự soạn 3 "bài" (không dùng SGK thật):
`projects/_giao-an/_sgk/mau.md`.
```
python tools/vi/giao_an.py trich-sgk projects/_giao-an/_sgk/mau.md --bai "Bài 5"
```
```json
{"ready": true, "files": ["projects/_giao-an/_sgk/sgk-trich.md"], "warnings": [], "error": null, "heading": "Bài 5. Hợp chất mẫu chứa nitơ", "size_kb": 0.3}
```
Mã thoát 0, `heading` đúng Bài 5. Đọc lại `sgk-trich.md`: chỉ chứa đúng phần
Bài 5, dừng đúng trước tiêu đề Bài 6.

## Bước 6b — Luồng A trên giáo án Bài 5 thật (spec §8.2)

Nguồn: giáo án Bài 5 thật (PDF) của chủ repo. Chỉ đọc, không sửa, không chép
file gốc vào repo; chỉ giữ bản Markdown chuyển đổi và `giao-an.md` mới soạn
trong `projects/_giao-an/bai-5-ammonia-that/` (đã gitignore).

Chuyển đổi:
```
python skills/ppt-master/scripts/source_to_md.py "<đường dẫn PDF của chủ repo>" -o projects/_giao-an/bai-5-ammonia-that/nguon --json
```
Mã thoát 0. Vì chỉ có một input và `-o` không kết thúc bằng `/`, tool ghi ra
đúng một file `projects/_giao-an/bai-5-ammonia-that/nguon.md` (đúng theo
`--help` của script). Chuyển đổi đọc được, phát hiện 3 bảng, không lỗi; công
thức hoá học bị OCR tách rời chỉ số dưới khỏi kí hiệu nguyên tố ở nhiều chỗ
nhưng nội dung chữ vẫn đọc được rõ, không phải chuyển đổi hỏng.

**Mục lục và số hoạt động của giáo án gốc:**
- I. MỤC TIÊU (Kiến thức; Năng lực chung, Năng lực đặc thù; Phẩm chất)
- II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU (Giáo viên; Học sinh)
- III. TIẾN TRÌNH DẠY HỌC — **5 hoạt động**: Khởi động (5 phút), Hình thành
  kiến thức mới (70 phút, gồm 2 nội dung con trình bày dạng bảng GV/HS),
  Luyện tập (10 phút), Vận dụng (3 phút), Củng cố dặn dò (2 phút) — tổng
  90 phút = 2 tiết.
- PHỤ LỤC: Phiếu học tập số 1, Phiếu học tập số 2.

Giáo án gốc không khai mã năng lực số/AI nào và không ghi tên trường trong
nội dung đọc được.

**`giao-an.md` (luồng A):** giữ nguyên nội dung chuyên môn của giáo án gốc
(kiến thức, PTHH, hai phiếu học tập, các bước tổ chức thực hiện), không rút
gọn; giữ nguyên 5 hoạt động và đúng số phút gốc. Vì đây là lớp chuyên, dùng
`track: chuyen` và khung `AI-H-Chuyên` (khung chỉ có đúng 1 mã). Giáo án gốc
không có mã NLS/AI sẵn nên chọn theo bảng: `1.1.NC1a`, `5.1.NC1a` cho năng
lực số; `AI-H-Chuyên` cho năng lực AI, gắn vào một hoạt động AI có đối
chứng (một nhóm dùng AI dự đoán phương trình/điều kiện phản ứng khử của
NH~3~ với O~2~ rồi đối chiếu với PTHH và thí nghiệm/mô phỏng thật trong
SGK). Thêm rubric 2 tiêu chí 3 mức. Những chỗ không đọc được rõ (tên trường,
công thức bị OCR xáo trộn, cấu trúc bảng GV/HS phải diễn giải lại theo khuôn
4 bước của công cụ, đáp án trắc nghiệm không có sẵn trong PDF) đã ghi vào
`## CAN SOAT`, không tự viết bù nội dung chuyên môn.

Ghi nhận: khung `AI-H-Chuyên` mô tả thiên về hữu cơ/phổ nghiệm (NMR, MS, IR,
cơ chế SN1/SN2/E1/E2, Python/machine learning cho động hoá học), không khớp
hẳn chủ đề Ammonia (vô cơ, cân bằng Le Chatelier). Tài liệu hướng dẫn không
nói phải làm gì khi khung duy nhất của track "chuyên" không khớp chủ đề bài
học cụ thể; đã dùng đúng mã `AI-H-Chuyên` (mã duy nhất của khung) và diễn đạt
nhiệm vụ AI theo hướng gần nhất có thể với nội dung bài, nhưng đây là một
giới hạn thực sự của khung mã hiện tại.

Chạy và kiểm:
```
python tools/vi/giao_an.py xuat projects/_giao-an/bai-5-ammonia-that
```
```json
{"ready": true, "files": ["projects/_giao-an/bai-5-ammonia-that/giao-an.docx", "projects/_giao-an/bai-5-ammonia-that/can-soat.md"], "warnings": ["Ghi đè file có sẵn: giao-an.docx", "Ghi đè file có sẵn: can-soat.md"], "error": null, "activities": 5, "periods": 2, "minutes": 90, "nls_codes": ["1.1.NC1a", "5.1.NC1a"], "ai_codes": ["AI-H-Chuyên"]}
```
Mã thoát 0. Hai cảnh báo "Ghi đè file có sẵn" xuất hiện cùng lý do như ở Bước
1–2 (file đã tồn tại từ lần chạy trước trong cùng phiên kiểm thử), không ảnh
hưởng tới các trường còn lại. **`activities` = 5, khớp đúng 5 hoạt động của
giáo án gốc**; `minutes` = 90, khớp tổng thời lượng gốc.

Kiểm XML:
```
python -c "import zipfile; x=zipfile.ZipFile('projects/_giao-an/bai-5-ammonia-that/giao-an.docx').read('word/document.xml').decode('utf-8'); print(all(s in x for s in ['I. MỤC TIÊU','III. TIẾN TRÌNH DẠY HỌC','Bước 4. Kết luận, nhận định','Hoạt động 1.']), 'subscript' in x, 'Cần thầy cô soát' not in x)"
```
→ `True True True`, đúng kỳ vọng.

## Kiểm cuối

```
python -m unittest discover -s tools/vi/tests -t tools/vi/tests
```
→ `Ran 494 tests` — `OK`.

```
python skills/ppt-master/scripts/attribution_guard.py
```
→ mã thoát 0.

`git status --porcelain` không có dòng nào dưới `projects/`.

## Kết luận

Không có lỗi nào chặn việc xuất giáo án; mọi ca thử (kể cả các ca cố ý sai)
đều bị chặn đúng (mã thoát khác 0) hoặc chạy đúng (mã thoát 0) như kỳ vọng.

Một giới hạn thật của công cụ, không chỉ là điểm tài liệu: khi một mã AI sai
chỉ xuất hiện ở dòng `ai:` của hoạt động và chưa được khai ở mục tiêu
(`### Nang luc AI`), `frameworks.validate` báo "chưa khai ở mục
'### Nang luc AI'" và `fix` của lỗi đó không liệt kê các mã hợp lệ của khung
— khác với ca mã sai được khai luôn ở mục tiêu, nơi `fix` liệt kê đủ. Export
vẫn bị chặn đúng trong cả hai trường hợp (mã thoát 1); chỉ phần gợi ý sửa
(`fix`) là chưa đầy đủ như nhau. Ghi nhận thêm ở bước 6b: khung `AI-H-Chuyên`
(hữu cơ/phổ nghiệm) không khớp hẳn chủ đề Ammonia (vô cơ) — đây là giới hạn
nội dung của khung mã hiện có, tài liệu hướng dẫn chưa nói rõ cách xử lý khi
gặp tình huống này.
