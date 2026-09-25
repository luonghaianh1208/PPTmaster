# Biên bản kiểm thử: hình ảnh, bàn tay cầm bút và máy quay cho video giải thích (v6.3.2-vi.10)

Ngày 2026-09-26. Máy chủ repo: Windows 11, 6 lõi, Chromium headless shell của Playwright, FFmpeg bản Gyan. Nhánh `feat/vi-video-hinh`. Spec: `docs/vi/phat-trien/2026-09-25-video-ma-hinh-anh-design.md`.

## 1. Lỗi font tiếng Việt

### 1.1 Nguyên nhân

Video vi.9 viết chữ bằng Segoe Print. Đọc bảng `cmap` của font bằng bộ đọc trong `tools/vi/video_ma_parts/phong.py` (chạy lại khi viết biên bản này):

| Font | Chữ có dấu có trong font (trên 134) | Thiếu |
|---|---|---|
| Segoe Print (`segoepr.ttf`) | 42 | 92, ví dụ ạ ả ậ ầ ấ ẩ ẫ ặ ằ ắ ẳ ẵ |
| Ink Free (`Inkfree.ttf`) | 54 | 80 |
| Comic Sans MS (`comic.ttf`) | 42 | 92 |
| Itim (`Itim-Regular.ttf`, 371.384 byte) | 134 | 0 |

Chromium lấy từng chữ thiếu từ một font khác, nên một từ bị trộn hai kiểu nét ("Chiều" có "ều" khác nét). Phép thử font của vi.9 dùng `document.fonts.check`; hàm này chỉ cho biết font đã nạp, không kiểm font có đủ chữ, nên biên bản vi.9 (mục 1.1) đã kết luận sai rằng Segoe Print "hiện đủ dấu".

### 1.2 Sửa và bằng chứng

- Đóng gói `Itim-Regular.ttf` và `OFL.txt` (SIL Open Font License 1.1) trong `tools/vi/video_ma_parts/runtime/fonts/`, nhúng vào trang cảnh bằng `@font-face` dạng `data:`. Phụ đề in lên hình cũng dùng Itim qua `fontsdir=.khung/fonts`.
- Test cmap (`test_video_ma_phong.py`): Itim chứa đủ 134 chữ và "ĐƯƠ"; `segoepr.ttf` thiếu "ề" (chứng minh bộ đọc phân biệt được font thiếu chữ). Hai font tổng hợp bằng `struct` kiểm nhánh định dạng 4 có `idRangeOffset` và định dạng 12.
- Test đo độ rộng trong Chromium: đo `measureText` từng chữ trong 134 chữ với `'Itim', monospace` và `'Itim', serif`; mọi chữ lệch dưới 0,01 px, tức không chữ nào rơi sang font dự phòng.
- Lần đầu test đo độ rộng trượt khi chạy riêng lẻ ("ạ" lệch 4,24 px) vì font nạp lười. Đã sửa cả test và đường chụp thật: `chup.mo_trang` chờ `document.fonts.load` cho toàn bộ 134 chữ trước khi chụp khung đầu.
- Kiểm bằng mắt: chữ THỬ, sẵn, kỷ, kỹ, mỹ trong tiêu đề, nhãn hình và phụ đề cùng một nét Itim (xem mục 4). Chữ trên canvas cảnh thí nghiệm ban đầu vẫn là Segoe UI; đã sửa sang Itim (`daed06d1`), có test Chromium cho hai mẫu thí nghiệm.

## 2. Ba lần chạy test cuối

Chạy sau commit tài liệu, trên máy chủ repo:

- `venv\Scripts\python.exe -m unittest discover -s tools/vi/tests`: **850 test, OK, bỏ qua 31** (các test cần Chromium/FFmpeg tự bỏ qua vì venv chính không có playwright).
- `C:/Users/ADMIN/vmt/v/Scripts/python.exe -m unittest discover -s tools/vi/tests -p "test_video_ma_*.py"`: **199 test, OK, không bỏ qua test nào, 146,6 giây** (Chromium và FFmpeg chạy thật, không mở cửa sổ).
- `node --test tools/vi/tests/js/test_canh.js tools/vi/tests/js/test_chuyen_dong.js`: **39 test, pass 39, fail 0**.

Khi chạy bộ venv, test `test_changes_limited_to_vietnamese_layer` trượt vì git in tên file có dấu của fixture mới (`tools/vi/fixtures/video-hinh/anh/quả nặng dọc.png`) dưới dạng mã bát phân trong ngoặc kép, nên không khớp mẫu `tools/vi/*`. Đã sửa test cho git in tên thô (`-c core.quotepath=off`); danh sách đường dẫn được phép không đổi.

## 3. Tốc độ dựng

Kịch bản `tools/vi/fixtures/video-mau/video.md` (8 cảnh, có con lắc), tiếng giả 6 giây mỗi cảnh, ra video 61,3 giây, 1280×720, 30/1, 1.839 khung:

| Số tiến trình chụp | Thời gian dựng | Suy ra video 5 phút |
|---|---|---|
| 3 (mặc định trên máy 6 lõi) | 87 giây (lần hai 93 giây): chụp khoảng 58 giây, FFmpeg khoảng 31 giây | khoảng 7,1–7,6 phút |
| 1 | 130 giây | khoảng 10,6 phút |

- Tiêu chí "video 5 phút dựng không quá 8 phút trên máy chủ repo" đạt, nhưng sát ngưỡng.
- Spec mục 10 ghi "máy 2 lõi chạy một tiến trình, video 5 phút mất khoảng 7 phút". Số đo một tiến trình trên máy này cho khoảng 10,6 phút, nên câu đó sai. Tài liệu cho AI và thầy cô đã ghi khoảng 11 phút cho máy 2–3 lõi. Chưa đo trên máy 2 lõi thật.
- Chạy 1 và 2 tiến trình không cho khung giống hệt từng byte: 3 khung lau bảng của cảnh mở đầu dải thứ hai lệch tối đa 6/255 ở 139 điểm ảnh vùng gạch chân, nằm dưới giẻ lau, mắt không thấy. Có test khoá mức lệch này (không quá 8/255, chỉ ở các khung đó).

## 4. Xem bằng mắt

### 4.1 Ảnh xem trước và khung giữa cảnh

- Fixture `tools/vi/fixtures/video-hinh/` (6 cảnh: tiêu đề + đồng hồ, khái niệm + đồng hồ cát, ý từng ý + ảnh dọc tên có dấu cách và dấu, minh hoạ 3 hình, ảnh thật + `nguon:`, thí nghiệm con lắc): đã xem cả 6 ảnh `--xem-truoc`.
- Lỗi gặp khi xem và đã sửa:
  - Dòng nguồn của ảnh dọc ở cột phải bị ngắt 2 dòng và đè lên ảnh. Sửa: dòng nguồn nằm dưới khung ảnh, một dòng (`f0af8074`).
  - Ở giây 20,5 máy quay phóng vào đồng hồ bấm giây đang vẽ, nhưng bàn tay lại vẽ gạch chân tiêu đề ngoài khung hình. Sửa: tay ưu tiên mục máy quay đang nhìn (`eb5fbb6d`). Có test Node khoá: ở mọi thời điểm, mục máy quay nhắm (nếu đang vẽ) chính là mục tay đang vẽ.
  - Chữ trên canvas thí nghiệm là Segoe UI (mục 1.2).
- Khung trích từ video dựng thật đã xem: ngòi bút đúng cuối chữ đang viết; lau bảng giữa chừng (nửa trái sạch, nửa phải còn khung cuối cảnh trước, giẻ ở mép quét); máy quay phóng vào định nghĩa và vào hình đang vẽ; Ken Burns ảnh thật; thí nghiệm "Số dao động" bằng Itim.
- Lần xem ở Task 4 cũng đã sửa: tiêu đề cách hình quá xa, và các ý của cùng một cảnh có cỡ chữ khác nhau.

### 4.2 Demo giọng thật

Thư mục `projects/_video/con-lac-don-demo3/` (không commit):

- 10 cảnh, đủ 10 loại cảnh; hình ở tiêu đề, khái niệm, công thức; minh hoạ 3 hình; cảnh ảnh thật con lắc Foucault (Wikimedia, Syced, CC0); ảnh đồng hồ quả lắc ở cột phải cảnh ý từng ý (Openverse, Christoph Braun, CC0).
- Giọng máy edge-tts thật, giọng nữ, tốc độ vừa; 3 tiến trình Chromium.
- Kết quả: `ready: true`, **104,3 giây** (3.129 khung), **8,55 MB**, dựng trong **186 giây** kể cả tạo giọng; h264 1280×720, 30/1, aac. Hai cảnh báo "mốc câu ước lượng" (cảnh 3 và 6).
- Ảnh đồng hồ quả lắc gốc nặng 8,7 MB. Lần dựng đầu báo đúng lỗi `canh` ("nặng 8708982 byte, tối đa 8388608 (dòng 60)"). Dùng bản thu nhỏ 1024 px mà `image_search.py` ghi sẵn ở `anh/.review/` thì dựng được. Giới hạn 8 MB giữ nguyên; tài liệu cho AI đã ghi cách dùng bản thu nhỏ.
- Đã xem 10 ảnh xem trước và 9 khung trích (giây 2, 9,85, 13, 30, 50,5, 60,5, 64, 86, 99): không tràn chữ, tay ở ngòi, phụ đề Itim đủ dấu. Video còn chờ chủ repo xem toàn bộ và nghiệm thu bằng mắt (spec mục 8).

## 5. Sửa sau review

Mỗi mục có test trượt trước khi sửa và đạt sau khi sửa.

- **Đường dẫn thoát khỏi thư mục** (`hinh.py`, `anh.py`): giá trị `hinh:` hoặc `anh:` có `..`, `/`, `\` hay ổ đĩa từng đọc được file ngoài thư mục biểu tượng hoặc `anh/`. Sửa: chặn trước khi ghép đường dẫn, rồi kiểm lại đường dẫn đã phân giải phải nằm trong thư mục gốc (`67820427`).
- **Thuộc tính SVG lạ** (`hinh.js`): khoá có namespace (ví dụ `{urn:x-thu}nhan`) từng được Chromium gán thẳng lên phần tử; khoá `on…` (kể cả `ONCLICK`) giờ cũng bị bỏ. Test Chromium xác nhận không thuộc tính nào như vậy được gán và không hàm xử lý nào chạy (`2e4a9d0e`).
- Tên biểu tượng có dấu tiếng Việt ("bình thí nghiệm") vẫn có gợi ý tên gần đúng; ảnh WEBP có test đọc kích thước.
- Chụp song song: nền cảnh trước được giải phóng sau khi chụp xong; chỉ lỗi ghi khung là `write`, lỗi dựng trang vẫn là `dung` kèm số cảnh (`3b38c34d`, `95d0a40e`).

## 6. Chưa kiểm

Các mục sau **chưa** được kiểm, không suy diễn là đạt:

- **Cài Chromium từ đầu trên máy sạch**: máy kiểm đã có sẵn Playwright và Chromium.
- **Phiên Antigravity thật với luật mới**: file `.agents/rules/ppt-master-vi.md` có thêm dòng tải ảnh (6.704 ký tự, dưới ngưỡng 12.000), nhưng chưa chạy thử trong Antigravity.
- **Thời gian dựng trên máy 2 lõi**: con số khoảng 11 phút suy ra từ một tiến trình trên máy 6 lõi.
- Tải ảnh bằng `image_search.py` khi máy không có Pillow: khi đó lệnh không ghi bản `.review`; chưa thử.

Các điểm còn để lại, không chặn phát hành:

- Tiêu đề cảnh vẫn tự hiện dần cùng lúc tay viết mục đầu tiên (lịch có từ vi.9).
- Máy quay khá năng động: gần như mỗi dòng chữ ngắn đều được phóng tới 1,35 lần.
- Cảnh thí nghiệm đẩy lên 1,06 lần quanh tâm, khung mô hình sát mép trái khoảng 4 px và mép dưới lấn nhẹ vào vùng phụ đề.
- Trang cảnh có ảnh lớn nặng hơn (khoảng 3,3 MB với ảnh 2,4 MB) vì ảnh nhúng dạng `data:`.
