# Chạy thật hai đề mẫu (Task 8)

Kiểm thử Task 8 của kế hoạch `2026-09-12-de-thi-plan`: dựng hai đề kiểm tra KHTN tiếng Anh thật (luồng A — chuyển từ đề tiếng Việt, và luồng B — soạn mới trực tiếp bằng tiếng Anh), chạy `tools/vi/de_thi.py` trên máy thật và kiểm nội dung file `.docx` bằng cách đọc XML, không mở Word.

Hai đề mẫu và các file `.docx` sinh ra nằm dưới `projects/_de-thi/` (thư mục này nằm trong `.gitignore`, không commit). Chủ repo tự mở các file Word để đánh giá trình bày.

Báo cáo này đã qua hai vòng sửa sau review: Fix round 1 (sửa một câu Vật lí có hai đáp án đúng, cộng bảy chỉnh sửa chất lượng) và Fix round 2 (khôi phục mục "cần soát" cho đề Vật lí, cộng bốn chỉnh sửa nhất quán). Nội dung chính dưới đây phản ánh đề mẫu **sau cả hai vòng sửa**; chi tiết từng vòng nằm ở cuối file.

## Đề mẫu 1 — `projects/_de-thi/thu-nghiem-ly-10/` (luồng A)

Chuyển một đề Vật lí 10 chương "Động lực học" (định luật II Newton, lực ma sát trượt) từ bản tiếng Việt tự soạn sang tiếng Anh: 3 câu Phần I, 1 câu Phần II, 1 câu Phần III; mọi câu có `vi:`, `why:`, `topic:` (tên chủ đề tiếng Việt); có mục `## CAN SOAT` với một dòng; `points: 10`.

Lệnh chạy: `python tools/vi/de_thi.py projects/_de-thi/thu-nghiem-ly-10`

Kết quả (đường dẫn đã rút gọn về dạng tương đối):

```json
{"ready": true, "files": ["projects/_de-thi/thu-nghiem-ly-10/de-en.docx", "projects/_de-thi/thu-nghiem-ly-10/de-song-ngu.docx", "projects/_de-thi/thu-nghiem-ly-10/dap-an.docx"], "questions": {"part1": 3, "part2": 1, "part3": 1}, "points": 2.0, "warnings": ["Thang điểm mặc định cho ra 2 điểm nhưng đề ghi points: 10. Kiểm tra lại số câu mỗi phần, hoặc sửa dòng points trong de.md."], "error": null}
```

Mã thoát 0, đúng một dòng JSON, đủ ba đường dẫn, cả ba file tồn tại. Cảnh báo lệch điểm xuất hiện đúng như dự tính: một đề 5 câu không thể đạt đúng 10 điểm theo thang mặc định (0,25 / 1,0 / 0,25 mỗi phần), nên công cụ luôn in cảnh báo lệch điểm — đây là hành vi được thiết kế đúng như vậy, không phải lỗi.

Bốn chủ đề dùng trong ma trận (đặt tên tiếng Việt theo đúng phong cách chương/chủ đề của ví dụ trong `docs/vi/tro-ly/de-khtn-tieng-anh.md`): "Định luật II Newton" (câu 1 Phần I), "Định luật Newton và lực tiếp xúc" (câu 2 Phần I), "Lực ma sát" (câu 3 Phần I), "Định luật II Newton và lực ma sát" (dùng chung cho câu Phần II và câu Phần III).

### Kiểm XML — `de-en.docx`

```
python -c "import zipfile,sys; p='projects/_de-thi/thu-nghiem-ly-10/de-en.docx'; x=zipfile.ZipFile(p).read('word/document.xml').decode('utf-8'); print('subscript' in x or 'superscript' in x, 'Full name' in x, 'THE END' in x)"
```
→ `True True True`, đúng kỳ vọng của kế hoạch.

Kiểm thêm:

- `de-en.docx` chứa đúng câu dẫn của câu 2 Phần I: "A book is pushed across a horizontal table". Đạt.
- `de-en.docx` **không** chứa "Static friction" và **không** chứa "The buoyant force" (hai lựa chọn của câu cũ có lỗi khoa học, đã bị thay hẳn). Đạt.
- `de-en.docx` **không** chứa nội dung bất kỳ dòng `why:` nào đã viết, **không** chứa chữ "Đáp án" và **không** chứa "Hướng dẫn giải". Đạt.
- `de-en.docx` **không** chứa câu `vi:` nào (thử với câu tiếng Việt của câu 1 Phần I). Đạt.
- `de-en.docx` **không** chứa nội dung ghi chú "cần soát" ("thầy cô xác nhận giúp em") — ghi chú dành cho tổ chuyên môn, không phải cho học sinh. Đạt.
- `de-song-ngu.docx` **có** chứa câu `vi:` của câu 1 Phần I và có nhãn "BẢN SONG NGỮ". Đạt.
- `dap-an.docx` **không** chứa cụm "ba lực còn lại" (lời giải sai khoa học của câu cũ). Đạt.
- `dap-an.docx` chứa "Cần thầy cô soát", chứa đúng nội dung mục `## CAN SOAT` ("thầy cô xác nhận giúp em đúng cách gọi..."), và chứa "0,25" (giá trị của `why:` câu Phần II sau khi viết lại theo chiều tính μ từ g). Mục "Cần thầy cô soát" của đề này gồm hai phần: ghi chú thuật ngữ vừa nêu, cộng cảnh báo lệch điểm ở trên (vì cảnh báo được gộp vào cùng danh sách "cần soát" khi in ra). Đạt.
- `dap-an.docx` có "Ma trận đặc tả", "Thang điểm", "Hướng dẫn giải", và có đủ bốn tên chủ đề tiếng Việt liệt kê ở trên trong bảng ma trận; không còn tên chủ đề tiếng Anh nào trong ma trận (cụm "Kinetic friction" duy nhất còn lại trong file nằm trong câu chữ của chính ghi chú "cần soát" ở trên, vốn cố ý trích dẫn thuật ngữ tiếng Anh để tổ chuyên môn đối chiếu — không phải tên chủ đề). Đạt.

## Đề mẫu 2 — `projects/_de-thi/thu-nghiem-khtn-8/` (luồng B)

Soạn mới trực tiếp bằng tiếng Anh, môn `NATURAL SCIENCES — Grade 8`: 3 câu Phần I (trong đó có công thức hoá học `H~2~SO~4~` dùng dấu `~ ~`), 1 câu Phần II (hệ tuần hoàn ở người), 1 câu Phần III (tính áp suất, dùng đơn vị `N/m^2^` với dấu `^ ^`). Không có mục `## CAN SOAT`. `points: 10`.

**Lệch so với kế hoạch gốc:** kế hoạch yêu cầu "hai `topic:` khác nhau để ma trận có hai dòng". Đề mẫu này có **ba** chủ đề khác nhau, tất cả đặt tên tiếng Việt: "Acid" (câu 1 và câu 2 Phần I), "Áp suất" (câu 3 Phần I và câu Phần III), "Hệ tuần hoàn ở người" (câu Phần II) — nhiều hơn mức tối thiểu kế hoạch yêu cầu. Đây là lựa chọn có chủ đích để bài kiểm thử luyện tới việc gom nhóm ma trận với nhiều hơn hai dòng, không phải thiếu sót.

Lệnh chạy: `python tools/vi/de_thi.py projects/_de-thi/thu-nghiem-khtn-8`

```json
{"ready": true, "files": ["projects/_de-thi/thu-nghiem-khtn-8/de-en.docx", "projects/_de-thi/thu-nghiem-khtn-8/de-song-ngu.docx", "projects/_de-thi/thu-nghiem-khtn-8/dap-an.docx"], "questions": {"part1": 3, "part2": 1, "part3": 1}, "points": 2.0, "warnings": ["Thang điểm mặc định cho ra 2 điểm nhưng đề ghi points: 10. Kiểm tra lại số câu mỗi phần, hoặc sửa dòng points trong de.md."], "error": null}
```

### Kiểm XML

- `de-en.docx` chứa "human circulatory system" (câu Phần II) và **không** chứa "Cellular respiration" (chủ đề cũ, sai cấp lớp — hô hấp tế bào thuộc chương trình KHTN 7, không phải KHTN 8). Đạt.
- `de-en.docx` **không** chứa nội dung `why:`, **không** chứa "Đáp án", **không** chứa "Hướng dẫn giải". Đạt.
- `de-song-ngu.docx` chứa "sulfuric acid" (tên hoá chất viết theo cách sách giáo khoa KHTN 8 hiện dùng, thay cho "Axit sulfuric" Việt hoá). Đạt.
- `dap-an.docx` chứa "Sulfuric acid" (viết hoa đầu câu trong `why:` câu 1 Phần I, khớp cách viết ở câu 2) và **không** còn chứa "Axit sulfuric" ở đâu trong file. Đạt.
- `dap-an.docx` có "Ma trận đặc tả", "Thang điểm", có đủ ba tên chủ đề tiếng Việt ("Acid", "Áp suất", "Hệ tuần hoàn ở người") trong bảng ma trận, không còn tên chủ đề tiếng Anh nào ("Acids and their formulas", "Pressure" không còn xuất hiện); có dấu subscript/superscript.
- `dap-an.docx` **không** có dòng "Không có mục nào cần soát". Một đề 5 câu không đạt đúng 10 điểm theo thang mặc định, nên công cụ luôn in cảnh báo lệch điểm vào mục "Cần thầy cô soát" thay vì in "Không có mục nào cần soát" — đây là hành vi thiết kế đúng như vậy: đề `de.md` không có `## CAN SOAT` không có nghĩa mục "Cần thầy cô soát" sẽ trống, một khi vẫn còn cảnh báo khác (như lệch điểm) cần thầy cô đọc.

### Kiểm `--plan-only` (cả hai đề)

```
python tools/vi/de_thi.py projects/_de-thi/thu-nghiem-ly-10 --plan-only
python tools/vi/de_thi.py projects/_de-thi/thu-nghiem-khtn-8 --plan-only
```

Cả hai đều `ready: true`, `files: []`, không ghi thêm hay xoá file `.docx` nào trong thư mục đề (ba file `.docx` sinh từ lần chạy đầy đủ trước đó vẫn còn nguyên). Đạt đúng kỳ vọng.

### Kiểm ca thiếu thư viện (đã chạy ở lần kiểm thử đầu, không lặp lại ở các vòng sửa)

```
python -c "import sys; sys.argv=['de_thi.py','projects/_de-thi/thu-nghiem-khtn-8']; sys.path.insert(0,'tools/vi'); import de_thi; de_thi.load_docx_build=lambda: (_ for _ in ()).throw(ImportError('No module named docx')); sys.exit(de_thi.main())"
```

Mã thoát 1, stdout đúng một dòng JSON với `error.step == "docx"` và `fix` nêu đúng `tools/vi/requirements-vi.txt`. Không liên quan tới nội dung `de.md` nên không cần chạy lại sau các vòng sửa.

## Số câu và tổng điểm

| Đề | Phần I | Phần II | Phần III | Tổng điểm (thang mặc định) | `points` khai báo |
|---|---|---|---|---|---|
| `thu-nghiem-ly-10` | 3 | 1 | 1 | 2,0 | 10 |
| `thu-nghiem-khtn-8` | 3 | 1 | 1 | 2,0 | 10 |

Cả hai đều nhận cảnh báo lệch điểm — đúng dự tính cho đề mẫu 5 câu, không phải lỗi.

## Chạy cả bộ test và guard toàn vẹn (sau vòng sửa cuối)

- `python -m unittest discover -s tools/vi/tests -t tools/vi/tests` → `Ran 355 tests in 22.522s` — `OK`.
- `python skills/ppt-master/scripts/attribution_guard.py; echo $?` → thoát mã `0`.

## Ghi chú bắt buộc

Chưa mở Word để xem; phần đánh giá trình bày do chủ repo tự kiểm.

Hai đề mẫu này chưa từng chạy trên một đề thi thật của trường; nội dung khoa học và đáp án số do người viết báo cáo tự soạn và tự tính lại, chưa qua tổ chuyên môn nào soát.

## Fix round 1 — sửa lỗi khoa học và tinh chỉnh chất lượng

Review phát hiện câu 2 Phần I của đề Vật lí 10 có **hai đáp án đúng**: đề gốc hỏi lực nào KHÔNG tác dụng lên một vật đứng yên trên bàn, không chịu lực nào khác — với vật đứng yên và không có gì đẩy/kéo, lực ma sát nghỉ bằng 0 và không tác dụng, nên "Static friction" (C) đúng ngang với "The buoyant force" (D); lời giải `why:` khẳng định "ba lực còn lại đều tác dụng lên vật" là sai khoa học (lực ma sát nghỉ không tác dụng). Đã thay toàn bộ câu này bằng một câu mới, không còn mơ hồ: sách trượt trên bàn với vận tốc không đổi (vật đang chuyển động thẳng đều nên chỉ có lực ma sát **trượt** — không phải ma sát nghỉ — tác dụng ngược chiều chuyển động; trọng lực và phản lực vuông góc với chuyển động, lực đẩy cùng chiều chuyển động), đáp án đúng duy nhất là "Kinetic friction" (C).

Bảy chỉnh sửa chất lượng nhỏ đã gộp cùng lúc (vì dù sao cũng phải regenerate lại hai đề, và đây là những gì chủ repo mở ra đầu tiên):

1. `thu-nghiem-ly-10/de.md`, câu 1 Phần II ý c — câu gốc "If the coefficient... is 0.25, this value is consistent with g = 10 m/s²" đọc hơi gượng. Viết lại: "The coefficient of kinetic friction between the tyres and the road is 0.25 (take g = 10 m/s^2^)." Tính lại: μ = F/(mg) = 2500/(1000×10) = 0,25, khớp với dữ kiện đề — vẫn đúng (T).
2. `thu-nghiem-khtn-8/de.md`, câu 1 Phần II — hô hấp tế bào là chủ đề của KHTN 7, không phải KHTN 8. Thay bằng câu Phần II mới cùng dạng, chủ đề "hệ tuần hoàn ở người": tim bơm máu, động mạch dẫn máu đi từ tim, tĩnh mạch dẫn máu về tim, hồng cầu vận chuyển oxygen (kiểm tính đúng/sai xem mục riêng bên dưới).
3. `thu-nghiem-khtn-8/de.md` — trong `vi:` (câu 1, câu 2 Phần I) và `why:` (câu 2 Phần I) đổi "Axit sulfuric" → "sulfuric acid", "natri hiđroxit" → "sodium hydroxide", theo đúng cách sách giáo khoa KHTN 8 hiện dùng viết tên hoá chất.
4. `thu-nghiem-ly-10/de.md`, câu 3 Phần I và câu 1 Phần III — đổi ký hiệu đơn vị trong `vi:` từ chữ "m/s²" (Unicode dựng sẵn) sang đúng dấu đánh dấu `m/s^2^` như bên `en:`.
5. `thu-nghiem-ly-10/de.md`, câu 3 Phần I — hạ mức độ từ `vandung` xuống `hieu`: đây chỉ là một bước thế số trực tiếp vào công thức F = μmg, không cần biến đổi hay suy luận nhiều bước.
6. Báo cáo — bỏ cách diễn đạt trỏ tới hồ sơ nhiệm vụ không commit; thay bằng giải thích trực tiếp về cảnh báo lệch điểm (xem mục "Kiểm XML" của đề KHTN 8 ở trên).
7. Báo cáo — đánh dấu rõ số lượng chủ đề KHTN 8 (ba, không phải hai như kế hoạch yêu cầu) là một điểm lệch có chủ đích so với kế hoạch gốc (xem mục "Lệch so với kế hoạch gốc" của đề mẫu 2 ở trên).

**Quyết định giữ nguyên (không phải chỉnh sửa):** câu 2 Phần I mới giữ nguyên `topic:` cũ ("Newton's laws and contact forces" tại thời điểm đó, trước khi đổi sang tiếng Việt ở Fix round 2) vì vẫn đúng chủ đề của câu, và giữ `level: biet` vì đây vẫn là nhận biết lực đơn giản.

**Dọn dẹp phát sinh (không nằm trong bảy điểm trên):** đã xoá mục `## CAN SOAT` (một dòng) của `thu-nghiem-ly-10/de.md` ở vòng sửa này, vì dòng đó chỉ nhắc riêng thuật ngữ "the buoyant force" — vốn thuộc câu 2 Phần I đã bị thay hẳn — giữ lại sẽ là một ghi chú "cần soát" trỏ tới một cụm từ không còn xuất hiện trong đề. *Re-review ở Fix round 2 chỉ ra việc xoá này đúng nhưng để lại đề Vật lí không còn minh hoạ đường "cần soát" nào nữa; mục `## CAN SOAT` đã được khôi phục với nội dung mới ở Fix round 2, xem bên dưới.*

### Kiểm câu Phần II KHTN 8 mới ("human circulatory system") không mơ hồ

- a) "The heart pumps blood through the blood vessels." — đúng, đây là chức năng cơ bản của tim.
- b) "Arteries carry blood away from the heart." — đúng theo đúng định nghĩa của động mạch (dẫn máu đi ra khỏi tim), không phụ thuộc máu giàu hay nghèo oxygen (tránh nhầm lẫn thường gặp rằng "động mạch luôn chứa máu giàu oxygen" — động mạch phổi dẫn máu nghèo oxygen nhưng vẫn đi ra từ tim nên vẫn là động mạch).
- c) "Veins carry blood away from the heart." — sai, tĩnh mạch dẫn máu **về** tim, không phải đi ra; đánh dấu F đúng.
- d) "Red blood cells transport oxygen around the body." — đúng, hồng cầu chứa hemoglobin vận chuyển oxygen.

Bốn ý này dùng đúng định nghĩa giải phẫu (động mạch = dẫn máu ra khỏi tim, tĩnh mạch = dẫn máu về tim), không phụ thuộc vào tình huống đặc biệt (tuần hoàn phổi) nên không có ý nào mơ hồ hay có hai cách hiểu.

## Fix round 2 — khôi phục mục "cần soát" và thống nhất tên chủ đề

Re-review chỉ ra: việc xoá mục `## CAN SOAT` ở Fix round 1 là đúng (dòng cũ trỏ tới một lựa chọn đã bị xoá), nhưng khiến không còn đề mẫu nào minh hoạ đường "cần soát" trong `dap-an.docx`, và báo cáo chưa nêu đây là một điểm lệch so với yêu cầu gốc của Task 8 (yêu cầu đề Vật lí phải có `## CAN SOAT` với một dòng).

**Khôi phục mục cần soát.** Thêm vào cuối `thu-nghiem-ly-10/de.md` đúng nội dung sau:

```
## CAN SOAT
- Thuật ngữ "Kinetic friction" dùng cho "lực ma sát trượt" và "normal force" dùng cho "phản lực"; thầy cô xác nhận giúp em đúng cách gọi trong tài liệu tổ đang dùng.
```

**Bốn chỉnh sửa nhất quán đã gộp cùng lúc:**

1. `thu-nghiem-ly-10/de.md`, `why:` của câu Phần II — dòng cũ tính ngược từ μ ra g ("μ = F/(mg) cho g = 2500/(0,25×1000) = 10 m/s², khớp với dữ kiện"), trong khi ý c của câu đã đổi hướng ở Fix round 1 thành "cho g, hỏi μ". Sửa lại đúng chiều: "μ = F/(mg) = 2500/(1000 × 10) = 0,25", giữ nguyên phần giải thích còn lại của `why:` vẫn đúng (a = F/m cho ý a; lực ma sát cản trở chuyển động cho ý b; thời gian dừng xe phụ thuộc v0 cho ý d).
2. `thu-nghiem-khtn-8/de.md`, `why:` câu 1 Phần I — đổi "Axit sulfuric" thành "Sulfuric acid" (viết hoa đầu câu), khớp cách viết đã dùng ở `why:` câu 2 Phần I.
3. Đặt lại toàn bộ `topic:` của cả hai đề thành tên chủ đề/chương tiếng Việt có dấu đúng, theo đúng phong cách ví dụ trong `docs/vi/tro-ly/de-khtn-tieng-anh.md` (ma trận trước đó lẫn cả tên tiếng Anh và tiếng Việt — không nhất quán):
   - Vật lí 10: "Newton's second law" → "Định luật II Newton"; "Newton's laws and contact forces" → "Định luật Newton và lực tiếp xúc"; "Kinetic friction" → "Lực ma sát"; "Newton's second law and kinetic friction" (dùng chung cho câu Phần II và câu Phần III) → "Định luật II Newton và lực ma sát".
   - KHTN 8: "Acids and their formulas" (dùng chung cho câu 1 và câu 2 Phần I) → "Acid"; "Pressure" (dùng chung cho câu 3 Phần I và câu Phần III) → "Áp suất"; "Hệ tuần hoàn ở người" giữ nguyên (đã là tiếng Việt).
   - Giữ đúng nhóm chủ đề cũ khi đổi tên: câu nào trước đây dùng chung một `topic:` thì sau khi đổi tên vẫn dùng chung đúng tên mới đó, nên số dòng của ma trận đặc tả không đổi (Vật lí 10 vẫn 4 dòng, KHTN 8 vẫn 3 dòng).
4. Báo cáo (mục "Fix round 1" ở trên) — viết lại phần liệt kê "bảy chỉnh sửa" cho nhất quán: bỏ mục từng đánh số 6 (thật ra là một quyết định giữ nguyên, không phải chỉnh sửa) và mục từng đánh số 7 (việc xoá `## CAN SOAT`, tự mâu thuẫn vì viết "không phải một trong bảy điểm" ngay trong chính danh sách bảy điểm đó) ra khỏi danh sách đánh số; tách hai nội dung này thành hai mục riêng có tiêu đề rõ ("Quyết định giữ nguyên" và "Dọn dẹp phát sinh") ngay sau danh sách bảy điểm.

### Regenerate và kiểm lại sau Fix round 2

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

Toàn bộ phép kiểm XML được liệt kê lại ở phần "Đề mẫu 1" và "Đề mẫu 2" phía trên (bao gồm phép kiểm được khôi phục: `dap-an.docx` của Vật lí chứa "Cần thầy cô soát" và đúng nội dung ghi chú mới, `de-en.docx` không chứa ghi chú đó, và bốn tên chủ đề tiếng Việt xuất hiện đúng trong từng ma trận, không còn tên tiếng Anh) đã chạy lại sau Fix round 2 và đều đạt. `--plan-only` trên cả hai thư mục vẫn `ready: true`, `files: []`. Bộ test đầy đủ (355 test — `OK`) và guard toàn vẹn (thoát mã 0) đã chạy lại lần cuối và đều đạt.
