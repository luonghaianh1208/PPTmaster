# Chạy thật hai đề mẫu (Task 8)

Kiểm thử Task 8 của kế hoạch `2026-09-12-de-thi-plan`: dựng hai đề kiểm tra KHTN tiếng Anh thật (luồng A — chuyển từ đề tiếng Việt, và luồng B — soạn mới trực tiếp bằng tiếng Anh), chạy `tools/vi/de_thi.py` trên máy thật và kiểm nội dung file `.docx` bằng cách đọc XML, không mở Word.

Hai đề mẫu và các file `.docx` sinh ra nằm dưới `projects/_de-thi/` (thư mục này nằm trong `.gitignore`, không commit). Chủ repo tự mở các file Word để đánh giá trình bày.

Báo cáo này đã qua một vòng sửa (Fix round 1) sau khi review phát hiện một câu Vật lí có hai đáp án đúng; nội dung dưới đây phản ánh đề mẫu **sau khi sửa**. Chi tiết vòng sửa nằm ở cuối file.

## Đề mẫu 1 — `projects/_de-thi/thu-nghiem-ly-10/` (luồng A)

Chuyển một đề Vật lí 10 chương "Động lực học" (định luật II Newton, lực ma sát trượt) từ bản tiếng Việt tự soạn sang tiếng Anh: 3 câu Phần I, 1 câu Phần II, 1 câu Phần III; mọi câu có `vi:`, `why:`, `topic:`; `points: 10`.

Lệnh chạy: `python tools/vi/de_thi.py projects/_de-thi/thu-nghiem-ly-10`

Kết quả (đường dẫn đã rút gọn về dạng tương đối):

```json
{"ready": true, "files": ["projects/_de-thi/thu-nghiem-ly-10/de-en.docx", "projects/_de-thi/thu-nghiem-ly-10/de-song-ngu.docx", "projects/_de-thi/thu-nghiem-ly-10/dap-an.docx"], "questions": {"part1": 3, "part2": 1, "part3": 1}, "points": 2.0, "warnings": ["Thang điểm mặc định cho ra 2 điểm nhưng đề ghi points: 10. Kiểm tra lại số câu mỗi phần, hoặc sửa dòng points trong de.md."], "error": null}
```

Mã thoát 0, đúng một dòng JSON, đủ ba đường dẫn, cả ba file tồn tại. Cảnh báo lệch điểm xuất hiện đúng như dự tính — đề mẫu chỉ 5 câu (2,0 điểm theo thang mặc định) trong khi `points: 10`; một đề 5 câu không thể đạt đúng 10 điểm theo thang mặc định (0,25 / 1,0 / 0,25 mỗi phần), nên công cụ luôn in cảnh báo lệch điểm — đây là hành vi được thiết kế đúng như vậy, không phải lỗi.

### Kiểm XML — `de-en.docx`

```
python -c "import zipfile,sys; p='projects/_de-thi/thu-nghiem-ly-10/de-en.docx'; x=zipfile.ZipFile(p).read('word/document.xml').decode('utf-8'); print('subscript' in x or 'superscript' in x, 'Full name' in x, 'THE END' in x)"
```
→ `True True True`, đúng kỳ vọng của kế hoạch.

Kiểm thêm:

- `de-en.docx` chứa đúng câu dẫn của câu 2 Phần I sau khi sửa: "A book is pushed across a horizontal table". Đạt.
- `de-en.docx` **không** chứa "Static friction" và **không** chứa "The buoyant force" (hai lựa chọn của câu cũ có lỗi khoa học, đã bị thay hẳn). Đạt.
- `de-en.docx` **không** chứa nội dung bất kỳ dòng `why:` nào đã viết, **không** chứa chữ "Đáp án" và **không** chứa "Hướng dẫn giải". Đạt.
- `de-en.docx` **không** chứa câu `vi:` nào (thử với câu tiếng Việt của câu 1 Phần I). Đạt.
- `de-song-ngu.docx` **có** chứa câu `vi:` đó và có nhãn "BẢN SONG NGỮ". Đạt.
- `dap-an.docx` **không** chứa cụm "ba lực còn lại" (lời giải sai khoa học của câu cũ). Đạt.
- `dap-an.docx` có "Ma trận đặc tả", "Thang điểm", "Hướng dẫn giải". Đạt.

## Đề mẫu 2 — `projects/_de-thi/thu-nghiem-khtn-8/` (luồng B)

Soạn mới trực tiếp bằng tiếng Anh, môn `NATURAL SCIENCES — Grade 8`: 3 câu Phần I (trong đó có công thức hoá học `H~2~SO~4~` dùng dấu `~ ~`), 1 câu Phần II (hệ tuần hoàn ở người), 1 câu Phần III (tính áp suất, dùng đơn vị `N/m^2^` với dấu `^ ^`). Không có mục `## CAN SOAT`. `points: 10`.

**Lệch so với kế hoạch gốc:** kế hoạch yêu cầu "hai `topic:` khác nhau để ma trận có hai dòng". Đề mẫu này có **ba** chủ đề khác nhau ("Acids and their formulas", "Pressure", "Hệ tuần hoàn ở người"), nhiều hơn mức tối thiểu kế hoạch yêu cầu. Đây là lựa chọn có chủ đích để bài kiểm thử luyện tới việc gom nhóm ma trận với nhiều hơn hai dòng, không phải thiếu sót.

Lệnh chạy: `python tools/vi/de_thi.py projects/_de-thi/thu-nghiem-khtn-8`

```json
{"ready": true, "files": ["projects/_de-thi/thu-nghiem-khtn-8/de-en.docx", "projects/_de-thi/thu-nghiem-khtn-8/de-song-ngu.docx", "projects/_de-thi/thu-nghiem-khtn-8/dap-an.docx"], "questions": {"part1": 3, "part2": 1, "part3": 1}, "points": 2.0, "warnings": ["Thang điểm mặc định cho ra 2 điểm nhưng đề ghi points: 10. Kiểm tra lại số câu mỗi phần, hoặc sửa dòng points trong de.md."], "error": null}
```

### Kiểm XML

- `de-en.docx` chứa "human circulatory system" (câu Phần II sau khi sửa) và **không** chứa "Cellular respiration" (chủ đề cũ, sai cấp lớp — hô hấp tế bào thuộc chương trình KHTN 7, không phải KHTN 8). Đạt.
- `de-en.docx` **không** chứa nội dung `why:`, **không** chứa "Đáp án", **không** chứa "Hướng dẫn giải". Đạt.
- `de-song-ngu.docx` chứa "sulfuric acid" (tên hoá chất viết theo cách sách giáo khoa KHTN 8 hiện dùng, thay cho "Axit sulfuric" Việt hoá). Đạt.
- `dap-an.docx` có "Ma trận đặc tả", "Thang điểm", có đủ ba tên chủ đề ("Acids and their formulas", "Pressure", "Hệ tuần hoàn ở người"), có dấu subscript/superscript.
- `dap-an.docx` **không** có dòng "Không có mục nào cần soát". Một đề 5 câu không đạt đúng 10 điểm theo thang mặc định, nên công cụ luôn in cảnh báo lệch điểm vào mục "Cần thầy cô soát" thay vì in "Không có mục nào cần soát" — đây là hành vi được thiết kế đúng như vậy cho mọi đề mẫu ngắn kiểu này (`de.md` không có `## CAN SOAT` không có nghĩa là mục "Cần thầy cô soát" sẽ trống, một khi vẫn có cảnh báo khác cần đọc).

### Kiểm `--plan-only` (cả hai đề)

```
python tools/vi/de_thi.py projects/_de-thi/thu-nghiem-ly-10 --plan-only
python tools/vi/de_thi.py projects/_de-thi/thu-nghiem-khtn-8 --plan-only
```

Cả hai đều `ready: true`, `files: []`, không ghi thêm hay xoá file `.docx` nào trong thư mục đề (ba file `.docx` sinh từ lần chạy đầy đủ trước đó vẫn còn nguyên). Đạt đúng kỳ vọng.

### Kiểm ca thiếu thư viện (đã chạy ở lần kiểm thử đầu, không lặp lại ở vòng sửa)

```
python -c "import sys; sys.argv=['de_thi.py','projects/_de-thi/thu-nghiem-khtn-8']; sys.path.insert(0,'tools/vi'); import de_thi; de_thi.load_docx_build=lambda: (_ for _ in ()).throw(ImportError('No module named docx')); sys.exit(de_thi.main())"
```

Mã thoát 1, stdout đúng một dòng JSON với `error.step == "docx"` và `fix` nêu đúng `tools/vi/requirements-vi.txt`. Không liên quan tới nội dung `de.md` nên không cần chạy lại sau vòng sửa.

## Số câu và tổng điểm

| Đề | Phần I | Phần II | Phần III | Tổng điểm (thang mặc định) | `points` khai báo |
|---|---|---|---|---|---|
| `thu-nghiem-ly-10` | 3 | 1 | 1 | 2,0 | 10 |
| `thu-nghiem-khtn-8` | 3 | 1 | 1 | 2,0 | 10 |

Cả hai đều nhận cảnh báo lệch điểm — đúng dự tính cho đề mẫu 5 câu, không phải lỗi.

## Chạy cả bộ test và guard toàn vẹn (sau vòng sửa)

- `python -m unittest discover -s tools/vi/tests -t tools/vi/tests` → `Ran 355 tests in 20.086s` — `OK`.
- `python skills/ppt-master/scripts/attribution_guard.py; echo $?` → thoát mã `0`.

## Ghi chú bắt buộc

Chưa mở Word để xem; phần đánh giá trình bày do chủ repo tự kiểm.

Hai đề mẫu này chưa từng chạy trên một đề thi thật của trường; nội dung khoa học và đáp án số do người viết báo cáo tự soạn và tự tính lại, chưa qua tổ chuyên môn nào soát.

## Fix round 1 — sửa lỗi khoa học và tinh chỉnh chất lượng

Review phát hiện câu 2 Phần I của đề Vật lí 10 có **hai đáp án đúng**: đề gốc hỏi lực nào KHÔNG tác dụng lên một vật đứng yên trên bàn, không chịu lực nào khác — với vật đứng yên và không có gì đẩy/kéo, lực ma sát nghỉ bằng 0 và không tác dụng, nên "Static friction" (C) đúng ngang với "The buoyant force" (D); lời giải `why:` khẳng định "ba lực còn lại đều tác dụng lên vật" là sai khoa học (lực ma sát nghỉ không tác dụng). Đã thay toàn bộ câu này bằng một câu mới, không còn mơ hồ: sách trượt trên bàn với vận tốc không đổi (vật đang chuyển động thẳng đều nên chỉ có lực ma sát **trượt** — không phải ma sát nghỉ — tác dụng ngược chiều chuyển động; trọng lực và phản lực vuông góc với chuyển động, lực đẩy cùng chiều chuyển động), đáp án đúng duy nhất là "Kinetic friction" (C).

Ngoài ra, gộp bảy chỉnh sửa chất lượng nhỏ vì dù sao cũng phải regenerate lại hai đề, và đây là những gì chủ repo mở ra đầu tiên:

1. **`thu-nghiem-ly-10/de.md`, câu 1 Phần II ý c** — câu gốc "If the coefficient... is 0.25, this value is consistent with g = 10 m/s²" đọc hơi gượng. Viết lại: "The coefficient of kinetic friction between the tyres and the road is 0.25 (take g = 10 m/s^2^)." Tính lại: μ = F/(mg) = 2500/(1000×10) = 0,25, khớp với dữ kiện đề — vẫn đúng (T). Câu dẫn tổng quát (`vi:` của cả câu) không nhắc riêng ý c nên không cần sửa thêm.
2. **`thu-nghiem-khtn-8/de.md`, câu 1 Phần II** — hô hấp tế bào là chủ đề của KHTN 7, không phải KHTN 8. Đã thay bằng câu Phần II mới cùng dạng, chủ đề "Hệ tuần hoàn ở người": tim bơm máu, động mạch dẫn máu đi từ tim, tĩnh mạch dẫn máu về tim, hồng cầu vận chuyển oxygen. Xem mục kiểm tính đúng/sai bên dưới.
3. **`thu-nghiem-khtn-8/de.md`** — trong `vi:` (câu 1, câu 2 Phần I) và `why:` (câu 2 Phần I) đổi "Axit sulfuric" → "sulfuric acid", "natri hiđroxit" → "sodium hydroxide", theo đúng cách sách giáo khoa KHTN 8 hiện dùng viết tên hoá chất (giữ nguyên tên tiếng Anh trong câu tiếng Việt), phần còn lại của câu giữ nguyên tự nhiên.
4. **`thu-nghiem-ly-10/de.md`, câu 3 Phần I và câu 1 Phần III** — đổi ký hiệu đơn vị trong `vi:` từ chữ "m/s²" (Unicode dựng sẵn) sang đúng dấu đánh dấu `m/s^2^` như bên `en:`, để nhất quán cách dựng chỉ số trên giữa hai ngôn ngữ.
5. **`thu-nghiem-ly-10/de.md`, câu 3 Phần I** — hạ mức độ từ `vandung` xuống `hieu`: đây chỉ là một bước thế số trực tiếp vào công thức F = μmg, không cần biến đổi hay suy luận nhiều bước.
6. Câu Phần I mới ("A book is pushed…") không thay `topic:` cũ ("Newton's laws and contact forces") vì vẫn đúng chủ đề của câu; giữ `level: biet` như chỉ đạo vì đây vẫn là nhận biết lực đơn giản.
7. Đã xoá mục `## CAN SOAT` (một dòng) của `thu-nghiem-ly-10/de.md`: dòng đó chỉ nhắc riêng thuật ngữ "the buoyant force", vốn thuộc câu 2 Phần I đã bị thay hẳn — giữ lại sẽ là một ghi chú "cần soát" trỏ tới một cụm từ không còn xuất hiện trong đề. Đây là dọn dẹp phần thừa do chính việc sửa câu 2 gây ra, không phải một trong bảy điểm được liệt kê ở trên.

### Kiểm câu Phần II KHTN 8 mới ("human circulatory system") không mơ hồ

- a) "The heart pumps blood through the blood vessels." — đúng, đây là chức năng cơ bản của tim.
- b) "Arteries carry blood away from the heart." — đúng theo đúng định nghĩa của động mạch (dẫn máu đi ra khỏi tim), không phụ thuộc máu giàu hay nghèo oxygen (tránh nhầm lẫn thường gặp rằng "động mạch luôn chứa máu giàu oxygen" — động mạch phổi dẫn máu nghèo oxygen nhưng vẫn đi ra từ tim nên vẫn là động mạch).
- c) "Veins carry blood away from the heart." — sai, tĩnh mạch dẫn máu **về** tim, không phải đi ra; đánh dấu F đúng.
- d) "Red blood cells transport oxygen around the body." — đúng, hồng cầu chứa hemoglobin vận chuyển oxygen.

Bốn ý này dùng đúng định nghĩa giải phẫu (động mạch = dẫn máu ra khỏi tim, tĩnh mạch = dẫn máu về tim), không phụ thuộc vào tình huống đặc biệt (tuần hoàn phổi) nên không có ý nào mơ hồ hay có hai cách hiểu.

### Regenerate sau khi sửa

```
$ python tools/vi/de_thi.py projects/_de-thi/thu-nghiem-ly-10
EXIT:0
{"ready": true, "files": ["projects/_de-thi/thu-nghiem-ly-10/de-en.docx", "projects/_de-thi/thu-nghiem-ly-10/de-song-ngu.docx", "projects/_de-thi/thu-nghiem-ly-10/dap-an.docx"], "questions": {"part1": 3, "part2": 1, "part3": 1}, "points": 2.0, "warnings": ["Thang điểm mặc định cho ra 2 điểm nhưng đề ghi points: 10. Kiểm tra lại số câu mỗi phần, hoặc sửa dòng points trong de.md.", "Ghi đè file có sẵn: de-en.docx", "Ghi đè file có sẵn: de-song-ngu.docx", "Ghi đè file có sẵn: dap-an.docx"], "error": null}
```

```
$ python tools/vi/de_thi.py projects/_de-thi/thu-nghiem-khtn-8
EXIT:0
{"ready": true, "files": ["projects/_de-thi/thu-nghiem-khtn-8/de-en.docx", "projects/_de-thi/thu-nghiem-khtn-8/de-song-ngu.docx", "projects/_de-thi/thu-nghiem-khtn-8/dap-an.docx"], "questions": {"part1": 3, "part2": 1, "part3": 1}, "points": 2.0, "warnings": ["Thang điểm mặc định cho ra 2 điểm nhưng đề ghi points: 10. Kiểm tra lại số câu mỗi phần, hoặc sửa dòng points trong de.md.", "Ghi đè file có sẵn: de-en.docx", "Ghi đè file có sẵn: de-song-ngu.docx", "Ghi đè file có sẵn: dap-an.docx"], "error": null}
```

Cảnh báo "Ghi đè file có sẵn" xuất hiện vì đây là lần chạy lại trên các file đã sinh trước đó — đúng như dự kiến khi regenerate.

Tất cả các phép kiểm XML, `--plan-only`, bộ test đầy đủ (355 test — `OK`) và guard toàn vẹn (thoát mã 0) đã chạy lại sau khi sửa và đều đạt (xem các mục ở trên, đã cập nhật theo kết quả sau sửa).
