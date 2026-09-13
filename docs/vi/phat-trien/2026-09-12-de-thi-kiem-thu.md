# Chạy thật hai đề mẫu (Task 8)

Kiểm thử Task 8 của kế hoạch `2026-09-12-de-thi-plan`: dựng hai đề kiểm tra KHTN tiếng Anh thật (luồng A — chuyển từ đề tiếng Việt, và luồng B — soạn mới trực tiếp bằng tiếng Anh), chạy `tools/vi/de_thi.py` trên máy thật và kiểm nội dung file `.docx` bằng cách đọc XML, không mở Word.

Hai đề mẫu và các file `.docx` sinh ra nằm dưới `projects/_de-thi/` (thư mục này nằm trong `.gitignore`, không commit). Chủ repo tự mở các file Word để đánh giá trình bày.

## Đề mẫu 1 — `projects/_de-thi/thu-nghiem-ly-10/` (luồng A)

Chuyển một đề Vật lí 10 chương "Động lực học" (định luật II Newton, lực ma sát trượt) từ bản tiếng Việt tự soạn sang tiếng Anh: 3 câu Phần I, 1 câu Phần II, 1 câu Phần III; mọi câu có `vi:`, `why:`, `topic:`; có mục `## CAN SOAT` với một dòng flag thuật ngữ "buoyant force" (lực đẩy Ác-si-mét); `points: 10`.

Lệnh chạy: `python tools/vi/de_thi.py projects/_de-thi/thu-nghiem-ly-10`

Kết quả (đường dẫn đã rút gọn về dạng tương đối):

```json
{"ready": true, "files": ["projects/_de-thi/thu-nghiem-ly-10/de-en.docx", "projects/_de-thi/thu-nghiem-ly-10/de-song-ngu.docx", "projects/_de-thi/thu-nghiem-ly-10/dap-an.docx"], "questions": {"part1": 3, "part2": 1, "part3": 1}, "points": 2.0, "warnings": ["Thang điểm mặc định cho ra 2 điểm nhưng đề ghi points: 10. Kiểm tra lại số câu mỗi phần, hoặc sửa dòng points trong de.md."], "error": null}
```

Mã thoát 0, đúng một dòng JSON, đủ ba đường dẫn, cả ba file tồn tại. Cảnh báo lệch điểm xuất hiện đúng như dự tính — đề mẫu chỉ 5 câu (2,0 điểm theo thang mặc định) trong khi `points: 10`, đây là hành vi mong đợi cho một đề rút gọn, không phải lỗi.

### Kiểm XML — `de-en.docx`

```
python -c "import zipfile,sys; p='projects/_de-thi/thu-nghiem-ly-10/de-en.docx'; x=zipfile.ZipFile(p).read('word/document.xml').decode('utf-8'); print('subscript' in x or 'superscript' in x, 'Full name' in x, 'THE END' in x)"
```
→ `True True True`, đúng kỳ vọng của kế hoạch.

Kiểm thêm (không trong danh sách gốc của kế hoạch, làm thêm để chắc đáp án không lọt sang đề học sinh):

- `de-en.docx` **không** chứa nội dung bất kỳ dòng `why:` nào đã viết (ví dụ giải thích "F = μmg = 0,2 × 2 × 10 = 4 N"), **không** chứa chữ "Đáp án" và **không** chứa "Hướng dẫn giải". Đạt.
- `de-en.docx` **không** chứa câu `vi:` nào (thử với câu tiếng Việt của câu 1 Phần I). Đạt.
- `de-song-ngu.docx` **có** chứa câu `vi:` đó và có nhãn "BẢN SONG NGỮ". Đạt.
- `dap-an.docx` có tiêu đề "Cần thầy cô soát" và có đúng nội dung dòng trong `## CAN SOAT` của `de.md` (thuật ngữ "Ác-si-mét"). Cũng có "Ma trận đặc tả", "Thang điểm", "Hướng dẫn giải". Đạt.

## Đề mẫu 2 — `projects/_de-thi/thu-nghiem-khtn-8/` (luồng B)

Soạn mới trực tiếp bằng tiếng Anh, môn `NATURAL SCIENCES — Grade 8`: 3 câu Phần I (trong đó có công thức hoá học `H~2~SO~4~` dùng dấu `~ ~`), 1 câu Phần II (hô hấp tế bào), 1 câu Phần III (tính áp suất, dùng đơn vị `N/m^2^` với dấu `^ ^`). Ba chủ đề khác nhau (Acids and their formulas, Cellular respiration, Pressure) để ma trận có nhiều hơn hai dòng. Không có mục `## CAN SOAT`. `points: 10`.

Lệnh chạy (chỉ xuất đáp án): `python tools/vi/de_thi.py projects/_de-thi/thu-nghiem-khtn-8 --phan dap-an`

```json
{"ready": true, "files": ["projects/_de-thi/thu-nghiem-khtn-8/dap-an.docx"], "questions": {"part1": 3, "part2": 1, "part3": 1}, "points": 2.0, "warnings": ["Thang điểm mặc định cho ra 2 điểm nhưng đề ghi points: 10. Kiểm tra lại số câu mỗi phần, hoặc sửa dòng points trong de.md."], "error": null}
```

### Kiểm XML — `dap-an.docx`

Có "Ma trận đặc tả", có "Thang điểm", có đúng ba tên chủ đề đã đặt trong `de.md` ("Acids and their formulas", "Pressure", "Cellular respiration"). Có dấu subscript/superscript trong tài liệu.

**Điểm khác với kỳ vọng ban đầu của kế hoạch:** kế hoạch dự kiến `dap-an.docx` có dòng "Không có mục nào cần soát" vì `de.md` không có mục `## CAN SOAT`. Thực tế file không có dòng đó — thay vào đó mục "Cần thầy cô soát" in ra chính cảnh báo lệch điểm ở trên ("Thang điểm mặc định cho ra 2 điểm nhưng đề ghi points: 10…"). Đọc lại `tools/vi/de_thi.py` (hàm `run`) và `docx_build._add_review`: hàm này in `exam.review_notes + warnings` (danh sách cảnh báo được truyền vào, gồm cả cảnh báo lệch điểm), và chỉ in "Không có mục nào cần soát" khi danh sách đó rỗng. Vì đề mẫu 5 câu luôn lệch điểm so với `points: 10`, danh sách không bao giờ rỗng nên dòng "Không có mục nào cần soát" không xuất hiện ở đề mẫu ngắn này. Đây đúng là hành vi được lường trước trong hồ sơ nhiệm vụ (mục "Additional checks" #3) cho trường hợp có cảnh báo lệch điểm, không phải lỗi công cụ.

### Kiểm `--plan-only` không ghi file

Xoá `dap-an.docx` trong `projects/_de-thi/thu-nghiem-khtn-8/`, chạy:

```
python tools/vi/de_thi.py projects/_de-thi/thu-nghiem-khtn-8 --plan-only
```

```json
{"ready": true, "files": [], "questions": {"part1": 3, "part2": 1, "part3": 1}, "points": 2.0, "warnings": ["Thang điểm mặc định cho ra 2 điểm nhưng đề ghi points: 10. Kiểm tra lại số câu mỗi phần, hoặc sửa dòng points trong de.md."], "error": null}
```

`files` rỗng, và sau lệnh không có file `.docx` nào trong thư mục đề. Đạt đúng kỳ vọng.

Sau đó chạy lại đầy đủ (`python tools/vi/de_thi.py projects/_de-thi/thu-nghiem-khtn-8`, không cờ) để thư mục đề mẫu này cũng có đủ ba file `.docx` cho chủ repo mở, cùng cách với đề mẫu 1.

### Kiểm ca thiếu thư viện

```
python -c "import sys; sys.argv=['de_thi.py','projects/_de-thi/thu-nghiem-khtn-8']; sys.path.insert(0,'tools/vi'); import de_thi; de_thi.load_docx_build=lambda: (_ for _ in ()).throw(ImportError('No module named docx')); sys.exit(de_thi.main())"
```

Mã thoát 1, stdout đúng một dòng JSON:

```json
{"ready": false, "files": [], "questions": {"part1": 3, "part2": 1, "part3": 1}, "points": 2.0, "warnings": ["Thang điểm mặc định cho ra 2 điểm nhưng đề ghi points: 10. Kiểm tra lại số câu mỗi phần, hoặc sửa dòng points trong de.md."], "error": {"step": "docx", "message": "Chưa cài thư viện python-docx (No module named docx)", "fix": "Cài thư viện bằng: python -m pip install -r tools/vi/requirements-vi.txt (hoặc chạy lại CAI-DAT.bat)"}}
```

`error.step` là `"docx"`, `fix` nêu đúng `tools/vi/requirements-vi.txt`. Đạt.

## Số câu và tổng điểm

| Đề | Phần I | Phần II | Phần III | Tổng điểm (thang mặc định) | `points` khai báo |
|---|---|---|---|---|---|
| `thu-nghiem-ly-10` | 3 | 1 | 1 | 2,0 | 10 |
| `thu-nghiem-khtn-8` | 3 | 1 | 1 | 2,0 | 10 |

Cả hai đều nhận cảnh báo lệch điểm — đúng dự tính cho đề mẫu 5 câu, không phải lỗi.

## Chạy cả bộ test và guard toàn vẹn

- `python -m unittest discover -s tools/vi/tests -t tools/vi/tests` → `Ran 355 tests in 19.711s` — `OK`.
- `python skills/ppt-master/scripts/attribution_guard.py; echo $?` → thoát mã `0`.

## Ghi chú bắt buộc

Chưa mở Word để xem; phần đánh giá trình bày do chủ repo tự kiểm.

Hai đề mẫu này chưa từng chạy trên một đề thi thật của trường; nội dung khoa học và đáp án số do người viết báo cáo tự soạn và tự tính lại, chưa qua tổ chuyên môn nào soát.
