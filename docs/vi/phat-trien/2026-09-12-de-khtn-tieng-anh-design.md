# Thiết kế: Soạn đề KHTN bằng tiếng Anh (v6.3.2-vi.5)

Ngày 2026-09-12. Gói này thêm một họ đầu ra mới cho lớp Việt: đề kiểm tra môn khoa học tự nhiên viết bằng tiếng Anh, xuất ra file Word in được. Đầu ra không phải PPTX và không đi qua lõi PPT Master.

## 1. Tiêu chí thành công

1. Thầy cô đưa một đề tiếng Việt (Word hoặc PDF) hoặc dán đề vào chat; sau một lượt hỏi tối đa 7 câu, nhận được ba file Word mở được trong Microsoft Word và in ra đúng thể thức đề thi.
2. Thầy cô không có sẵn gì; sau lượt hỏi, AI soạn đề mới hoàn toàn bằng tiếng Anh kèm ma trận đặc tả.
3. Tiếng Anh đúng thuật ngữ của môn, đúng động từ lệnh hỏi, đúng đơn vị và ký hiệu; không dịch từng chữ.
4. Khi chuyển từ đề tiếng Việt, số liệu, đáp án đúng và thứ tự câu không đổi.
5. Mọi chỗ AI chưa chắc chắn đều hiện ra trong mục "Cần thầy cô soát" ở cuối file đáp án, không im lặng bỏ qua.
6. Thầy cô không phải cài Pandoc. Bộ cài chỉ thêm đúng một thư viện Python.
7. Không sửa bất cứ gì trong `skills/`, `LICENSE`, `SPONSORS*.md`. Đề của thầy cô không bao giờ lên GitHub.
8. Toàn bộ test của lớp Việt vẫn xanh sau khi thêm gói này.

## 2. Hiện trạng đã kiểm

- `skills/ppt-master/scripts/source_to_md.py` nhận `-t {auto,pdf,doc,excel,pptx,web,markdown,text}`. **Không có backend cho ảnh**, nên đề chụp bằng điện thoại không đọc được; phải là PDF hoặc thầy cô gõ lại.
- `.gitignore` có `projects/*` (chỉ trừ `projects/README.md`), nên `projects/_de-thi/` được bỏ qua sẵn, cùng chỗ với `projects/_ho-so-don-vi.md` và `projects/_brief-*.md` của các gói trước.
- `skills/ppt-master/requirements.txt` **không** khai `python-docx` (đã đọc toàn bộ file: PyYAML, python-pptx, XlsxWriter, skia-pathops, uharfbuzz, edge-tts, PyMuPDF, mammoth, markdownify, ebooklib, nbconvert, openpyxl, Pillow, numpy, requests, beautifulsoup4, curl_cffi, google-genai, flask). File này thuộc upstream nên không được sửa; thư viện mới phải khai ở chỗ khác.
- `docs/vi/tro-ly/` hiện có 6 loại việc. `quy-trinh-hoi.md` mở đầu bằng "trước khi tạo PPTX", tuy loại việc thứ 6 (video bài giảng) đã không ra PPTX.
- `tools/vi/tests/test_vi_layer.py` khoá danh sách sáu file hướng dẫn, các chỗ đếm "6 loại", và khoá việc mục 11 của `AGENTS.vi.md` là mục cuối.
- Bài học từ gói video còn giá trị ở đây: đường dẫn phải `resolve()` trước khi dùng, và `.ps1` phải giữ BOM UTF-8.

## 3. Quyết định

| Quyết định | Lý do |
|---|---|
| Bốn khung môn: KHTN THCS 6–9 tích hợp, và Vật lí / Hoá học / Sinh học THPT 10–12 | Chủ repo chọn. Trường THPT dùng ba môn tách; THCS dùng môn tích hợp |
| Giữ cấu trúc đề Việt Nam (Phần I trắc nghiệm, Phần II đúng/sai, Phần III trả lời ngắn), thêm bản song ngữ | Chủ repo chọn. Tổ chuyên môn cần bản đối chiếu để soát bản dịch |
| Không làm từ điển thuật ngữ, không làm script kiểm thuật ngữ | Chủ repo chọn. Bù lại: quy ước và ví dụ mẫu nằm trong file hướng dẫn cho AI, và mục "Cần thầy cô soát" bắt AI nói ra chỗ chưa chắc |
| Dựng Word bằng `python-docx`, bỏ Pandoc khỏi gói này | Pandoc không xếp được đáp án thành cột, không đặt được số trang, không làm chỉ số dưới. Tự viết XML thì một lỗi nhỏ là Word từ chối mở file — rủi ro không đáng nhận |
| Một file nguồn `de.md` sinh ra cả ba file Word | Đề và đáp án không thể lệch nhau. Ma trận đặc tả được tính từ trường `level:` nên không thể mâu thuẫn với đề |
| Chính tả IUPAC / Anh-Anh cho tên hợp chất | Sách khoa học phổ thông quốc tế dùng hệ này (*sulfuric acid*, *sulfate*, *aluminium*) |
| Thang điểm in trong file đáp án, và cảnh báo khi tổng khác `points` | Phép tính, không phải đánh giá ngôn ngữ, nên máy kiểm được mà không trái quyết định bỏ script kiểm |
| Đề để ở `projects/_de-thi/<tên_đề>/` | `projects/` đã gitignore; cùng quy ước tiền tố `_` với các file của gói trước |

## 4. File

### 4.1 File mới

| File | Trách nhiệm |
|---|---|
| `docs/vi/tro-ly/de-khtn-tieng-anh.md` | Loại việc thứ 7: khi nào dùng, câu hỏi cho thầy cô, hai luồng A/B, ghi brief |
| `docs/vi/tro-ly/tieng-anh-khoa-hoc.md` | Chín nguyên tắc viết tiếng Anh khoa học, ví dụ mẫu theo môn |
| `tools/vi/de_thi.py` | Điểm vào duy nhất; in đúng một dòng JSON |
| `tools/vi/de_thi_parts/__init__.py` | Gói con |
| `tools/vi/de_thi_parts/parse.py` | Đọc `de.md` thành cấu trúc; lỗi kèm số dòng |
| `tools/vi/de_thi_parts/docx_build.py` | Dựng ba file Word bằng `python-docx` |
| `tools/vi/requirements-vi.txt` | `python-docx>=1.1.0` |
| `docs/vi/soan-de-tieng-anh.md` | Hướng dẫn cho thầy cô |
| `tools/vi/tests/test_de_thi.py` | Test của gói |

### 4.2 File sửa

| File | Sửa gì |
|---|---|
| `AGENTS.vi.md` | Mục 3 thêm câu lệnh kích hoạt; mục 10 thành 7 loại việc và nới cách diễn đạt; thêm mục 12 "Soạn đề KHTN bằng tiếng Anh" làm mục cuối |
| `docs/vi/tro-ly/quy-trinh-hoi.md` | Thêm dòng thứ 7 vào bảng nhận diện; nới câu mở đầu khỏi "trước khi tạo PPTX" |
| `docs/vi/xu-ly-loi.md` | Thêm mục "Xuất đề Word thất bại" |
| `docs/vi/bat-dau-nhanh.md` | Thêm mục "Soạn đề tiếng Anh" |
| `tools/vi/pptmaster.ps1` | Cài `tools/vi/requirements-vi.txt` trong bước setup |
| `tools/vi/setup.sh` | Cài `tools/vi/requirements-vi.txt` |
| `tools/vi/doctor.py` | Kiểm `python-docx`, báo thiếu kèm lệnh sửa |
| `README.md` | Dòng phiên bản, mục "Làm được gì" |
| `CHANGELOG-VI.md` | Mục `6.3.2-vi.5` |
| `tools/vi/tests/test_vi_layer.py` | 7 loại việc; hai file hướng dẫn mới; mục 12 là mục cuối |

### 4.3 Test đang khoá sẽ phải đổi

Phải sửa đúng những chỗ này, không nới lỏng phép kiểm:

1. Danh sách file hướng dẫn: thêm `de-khtn-tieng-anh.md` và `tieng-anh-khoa-hoc.md`.
2. Mọi chỗ đếm "6 loại việc" trong `AGENTS.vi.md`, `quy-trinh-hoi.md`, `mau-brief.md`, `bat-dau-nhanh.md` thành 7.
3. Phép kiểm "mục 11 là mục cuối của `AGENTS.vi.md`" đổi thành mục 12; mục 11 vẫn phải nằm ngay trước mục 12.
4. Bảng nhận diện trong `quy-trinh-hoi.md` có đúng 7 dòng.

## 5. Luồng

### 5.1 Luồng A — có đề tiếng Việt sẵn

1. Nhận diện loại việc theo `quy-trinh-hoi.md`, đọc `de-khtn-tieng-anh.md` và `tieng-anh-khoa-hoc.md`.
2. Thầy cô đưa file: đọc bằng `source_to_md.py <file> -o <thư_mục_tạm>`. Thầy cô đưa ảnh: nói rõ không đọc được ảnh, xin bản PDF hoặc Word.
3. Hỏi một lượt (mục 6.4), chờ trả lời.
4. Chuyển sang tiếng Anh theo chín nguyên tắc; viết `projects/_de-thi/<tên_đề>/de.md`.
5. Chạy `de_thi.py`, đọc JSON, báo thầy cô đường dẫn ba file và mục "Cần thầy cô soát".

### 5.2 Luồng B — tạo đề mới hoàn toàn

1. Nhận diện và đọc hai file hướng dẫn như trên.
2. Hỏi một lượt, chờ trả lời.
3. Soạn ma trận trước: chủ đề × mức độ × số câu, theo câu trả lời của thầy cô.
4. Soạn câu hỏi **trực tiếp bằng tiếng Anh**, không soạn tiếng Việt rồi dịch — câu văn tự nhiên hơn. Viết bản tiếng Việt vào trường `vi:` cho thầy cô soát.
5. Chạy `de_thi.py` và báo như luồng A.

### 5.3 Điều AI không được làm

- Không tự sửa câu hỏi gốc của thầy cô cho "hay hơn"; sai hoặc mơ hồ thì ghi vào mục cần soát.
- Không tự đổi số liệu để số tròn hơn.
- Không tự đổi đáp án đúng, kể cả khi tin là đáp án gốc sai — ghi vào mục cần soát.
- Không đổi ngữ cảnh Việt Nam thành ngữ cảnh nước ngoài.
- Không chạy `project_manager.py init`, không tạo SVG, không chạm `skills/`.

## 6. Hợp đồng

### 6.1 Định dạng `de.md`

Khối meta ở đầu file, giữa hai dòng `---`, mỗi dòng một cặp `khoá: giá trị`:

| Khoá | Bắt buộc | Ý nghĩa |
|---|---|---|
| `school` | có | Tên trường, in hoa |
| `department` | không | Tổ chuyên môn |
| `title` | có | Tên kỳ kiểm tra |
| `subject` | có | Môn và lớp bằng tiếng Anh, ví dụ `PHYSICS — Grade 10` |
| `time` | có | Số phút làm bài, chỉ chữ số |
| `code` | không | Mã đề |
| `points` | không | Tổng điểm, mặc định `10` |

Mỗi phần mở bằng một dòng `## PART I`, `## PART II` hoặc `## PART III`, theo đúng thứ tự đó; phần không dùng thì bỏ hẳn dòng đó. Mỗi câu mở bằng `### <số>`, đếm lại từ 1 trong mỗi phần.

Các dòng trong một câu:

| Khoá | Ở phần | Bắt buộc | Ý nghĩa |
|---|---|---|---|
| `en` | I, II, III | có | Đề bài tiếng Anh |
| `vi` | I, II, III | không | Bản tiếng Việt; thiếu thì bản song ngữ ghi "(chưa có bản tiếng Việt)" và thêm cảnh báo |
| `A`–`D` | I | có, đủ bốn | Bốn lựa chọn |
| `a`–`d` | II | có, đủ bốn | Bốn ý, mỗi dòng kết bằng ` \| T` hoặc ` \| F` |
| `key` | I, III | có | Phần I là một chữ trong `A`–`D`; phần III là đáp số |
| `unit` | III | không | Đơn vị của đáp số |
| `why` | I, II, III | không | Giải thích ngắn, chỉ in trong file đáp án |
| `level` | I, II, III | có | `biet`, `hieu` hoặc `vandung` |
| `topic` | I, II, III | không | Tên chương hoặc chủ đề, dùng để dựng ma trận |

Trong `en:` và `vi:` chỉ dùng ba dấu đánh dấu sau, ngoài ra không dùng Markdown nào khác:

- `H~2~O` → chỉ số dưới
- `m/s^2^` → chỉ số trên
- `**not**` → in đậm

### 6.2 Thang điểm

In trong file đáp án, theo thang đang dùng cho đề trắc nghiệm hiện hành:

- Phần I: mỗi câu 0,25 điểm.
- Phần II: mỗi câu tối đa 1,0 điểm — đúng 1 ý 0,1; 2 ý 0,25; 3 ý 0,5; 4 ý 1,0.
- Phần III: mỗi câu 0,25 điểm.

Tổng khác `points` thì thêm một cảnh báo, không phải lỗi: tổ chuyên môn có thể dùng thang khác.

### 6.3 Lệnh và JSON

```
python tools/vi/de_thi.py <thư_mục_đề> [--phan de|song-ngu|dap-an|tat-ca] [--plan-only]
```

`--phan` nhận một giá trị hoặc nhiều giá trị cách nhau bằng dấu phẩy (ví dụ `--phan de,dap-an`), mặc định `tat-ca`. `--plan-only` chỉ đọc và kiểm `de.md`, không ghi file nào.

stdout đúng một dòng JSON; tiến trình đi ra stderr:

```json
{
  "ready": true,
  "files": ["...de-en.docx", "...de-song-ngu.docx", "...dap-an.docx"],
  "questions": {"part1": 18, "part2": 4, "part3": 6},
  "points": 10.0,
  "warnings": [],
  "error": null
}
```

Khi lỗi, `ready` là `false` và `error` là `{"step": ..., "message": ..., "fix": ...}` với `step` thuộc `{input, parse, docx, write}`:

| `step` | Nghĩa |
|---|---|
| `input` | Không có thư mục đề, hoặc không có `de.md` |
| `parse` | `de.md` sai cú pháp; `message` nêu số dòng |
| `docx` | Thiếu `python-docx`; `fix` là lệnh cài |
| `write` | Không ghi được file (đang mở trong Word, hết đĩa, đường dẫn quá dài) |

### 6.4 Câu hỏi cho thầy cô

Tối đa 7 câu, theo đúng cách hỏi của `quy-trinh-hoi.md` (xưng "em", mỗi câu kèm "Gợi ý:"):

1. Môn và lớp (KHTN 6–9, hay Vật lí / Hoá học / Sinh học 10–12)?
2. Đề cho kỳ nào, làm trong bao nhiêu phút, tổng bao nhiêu điểm?
3. Nội dung thuộc chương hay chủ đề nào?
4. Số câu mỗi phần (Phần I trắc nghiệm, Phần II đúng/sai, Phần III trả lời ngắn)? Gợi ý THPT: 18 – 4 – 6 theo đề tham khảo 2025. THCS không có định dạng chung nên **phải hỏi**, không được lấy gợi ý làm mặc định.
5. Tỉ lệ mức độ biết – hiểu – vận dụng? Gợi ý 4 – 4 – 2.
6. Đề dùng để làm gì (kiểm tra lớp song ngữ, đề luyện thêm, đề tham khảo cho tổ)? Câu này quyết định mức từ vựng tiếng Anh.
7. Có cần bản song ngữ để tổ soát không? Trả lời "không" thì chạy với `--phan de,dap-an`, không xuất `de-song-ngu.docx`; trường `vi:` trong `de.md` vẫn ghi để sau này xuất lại được.

Luồng A đã có đề sẵn thì câu 3, 4 và 5 đọc được từ chính đề, chỉ hỏi phần còn thiếu.

Mục "Tạo nhanh": câu 1 và câu 4.

## 7. Ca biên

| Tình huống | Xử lý |
|---|---|
| `de.md` không có câu nào | `error.step = parse` |
| Phần I có câu thiếu lựa chọn `C` | `error.step = parse`, nêu số dòng và khoá còn thiếu |
| `key: E` ở phần I | `error.step = parse` |
| `level` sai chính tả | `error.step = parse`, liệt kê ba giá trị hợp lệ |
| Thiếu `vi:` | Cảnh báo; bản song ngữ ghi "(chưa có bản tiếng Việt)" |
| Tổng điểm khác `points` | Cảnh báo, vẫn xuất file |
| Chưa cài `python-docx` | `error.step = docx`, `fix` nêu lệnh cài |
| File docx đang mở trong Word | `error.step = write`, `fix` bảo đóng file rồi chạy lại |
| Thư mục đích đã có ba file docx | Ghi đè, kèm cảnh báo nêu tên file bị ghi đè |
| Đường dẫn có dấu tiếng Việt và khoảng trắng | Phải chạy được; mọi đường dẫn `resolve()` trước khi dùng |
| Đường dẫn dài hơn 200 ký tự | Cảnh báo về giới hạn đường dẫn của Windows |
| Đề 40 câu | Vẫn xuất xong; không có giới hạn số câu |
| Thầy cô đưa ảnh chụp đề | AI nói rõ không đọc được ảnh, xin PDF hoặc Word; không tự đoán nội dung đề |
| `--plan-only` | Không ghi file nào, `files` là danh sách rỗng |

## 8. Kiểm thử

### 8.1 Test tự động (`tools/vi/tests/test_de_thi.py`)

- `parse.py`: đề hợp lệ ba phần; từng ca biên sai cú pháp ở bảng trên, mỗi ca assert đúng số dòng trong `message`; đánh dấu `~ ~` và `^ ^` lồng trong một dòng; `topic` thiếu thì ma trận gom vào một dòng "Không ghi chủ đề".
- Thang điểm: tính đúng 4,5 + 4,0 + 1,5 = 10 với đề 18–4–6; cảnh báo khi lệch.
- `docx_build.py`: mở lại file đã xuất bằng `zipfile` và soi XML — khổ A4 và bốn lề đúng, font Times New Roman trong `styles.xml`, có trường `PAGE` ở chân trang, đáp án ngắn xếp 4 cột và đáp án dài xếp 1 cột, chỉ số dưới có `vertAlign="subscript"`, file đề **không** chứa chuỗi của `key:` hay `why:`.
- `de_thi.py`: thiếu `de.md` → `input`; `python-docx` không import được → `docx`; `--plan-only` không tạo file; stdout đúng một dòng JSON; đường dẫn có dấu tiếng Việt chạy được.

### 8.2 Chạy thật

Hai đề mẫu, chạy trên máy thật:

1. Một đề Vật lí 10 ngắn (5 câu) chuyển từ một đề tiếng Việt mẫu.
2. Một đề KHTN 8 tạo mới hoàn toàn (5 câu).

Xuất đủ ba file mỗi đề, kiểm bằng XML. **Không mở Word trong lúc kiểm** — chủ repo không muốn cửa sổ bật lên bất ngờ; phần đánh giá "đẹp" do chủ repo tự mở xem.

## 9. Phát hành

Nhánh `feat/vi-de-thi`, phiên bản `v6.3.2-vi.5`. Tag này gộp luôn commit `7534f614` (sửa mã hoá console của `video.py`) hiện đang nằm ngoài tag `v6.3.2-vi.4`.

`CHANGELOG-VI.md` ghi cả phần rủi ro, trong đó nêu rõ phần nào chủ repo chưa nghiệm thu.

## 10. Ngoài phạm vi

- **Trộn mã đề** (đảo thứ tự câu và đáp án thành nhiều mã 101, 102, 103…): chủ repo chốt làm thành một gói riêng tích hợp sau. `de.md` là nguồn máy đọc được nên gói đó dùng lại được ngay, không phải đổi định dạng nguồn.
- Đọc ảnh chụp đề (OCR).
- Xuất PDF: Word đã in ra PDF được; thêm bước này cần LibreOffice hoặc bộ LaTeX trên máy.
- Chấm bài, nhập điểm, ngân hàng câu hỏi tái sử dụng.
- Các môn ngoài KHTN (Toán, Tin, Ngữ văn…).
- Từ điển thuật ngữ và script kiểm thuật ngữ: chủ repo đã quyết định không làm trong gói này.

## 11. Rủi ro

- Thêm `python-docx` vào bộ cài: máy đã cài xong bản vi.4 phải chạy lại `CAI-DAT.bat` hoặc lệnh cài thư viện. `doctor.py` sẽ báo thiếu, kèm lệnh sửa.
- Không có bước máy nào kiểm thuật ngữ, nên hai lần chạy có thể dùng hai thuật ngữ khác nhau cho cùng khái niệm. Chỗ chặn duy nhất là chín nguyên tắc trong file hướng dẫn và mục "Cần thầy cô soát".
- Đọc đề PDF dựa vào PyMuPDF của upstream (giấy phép AGPL-3.0, chỉ dùng khi chuyển PDF).
- "Đẹp" là đánh giá của người. Test chỉ kiểm được cấu trúc XML, không kiểm được cảm nhận khi in ra giấy.
- Đề của trường có thể dùng thể thức đầu đề riêng (tên sở, logo). Bản này lấy `school` và `title` từ meta, chưa nạp mẫu đầu đề riêng và chưa chèn logo.
