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

### 2.1 Hai lần chạy test toàn bộ (Task 11)

- `venv\Scripts\python.exe -m unittest discover -s tools/vi/tests`: **745 test, OK, bỏ qua 12** (các test cần Chromium/FFmpeg mà venv chính không có, đúng như dự kiến).
- `C:/Users/ADMIN/vmt/v/Scripts/python.exe -m unittest discover -s tools/vi/tests -p "test_video_ma_*.py"` (máy này có sẵn playwright và edge-tts): **97 test, OK, không bỏ qua test nào** — bao gồm cả lớp Chromium (`test_video_ma_chup.py`) và lớp dựng video thật bằng FFmpeg (`test_video_ma_tich_hop.py`, `test_video_ma_ghep.py`).
- `node --test tools/vi/tests/js/test_canh.js`: **10 test, pass 10, fail 0**.

Ba lần chạy đều xanh, không có test nào bị nới lỏng hay bỏ qua ngoài các skip đã biết trước (thiếu Chromium/FFmpeg ở venv chính).

### 2.2 Ảnh xem trước tám loại cảnh (Task 9, sau khi sửa)

| Cảnh | Kết quả nhìn ảnh xem trước |
|---|---|
| tieu-de | Dấu tiếng Việt đủ ("Con lắc đơn và chu kì dao động", "Vật lí 11"). Đường gạch chân lệch 300 px khỏi tiêu đề lúc đầu — đã sửa, canh ngay dưới tiêu đề. |
| khai-niem | Khung, thuật ngữ, định nghĩa đúng vị trí, định nghĩa xuống 2 dòng gọn. Đường gạch chân từng sát định nghĩa — đã sửa cùng đợt với tieu-de. |
| cong-thuc | Công thức "T = 2π√(l/g)" nằm gọn trong khung, ba dòng giải thích cách đều, "m/s²" hiện đúng. Không lỗi ở lần xem đầu; sau đó khung `bieu-thuc` được tăng cao 110→120px khi kiểm với chữ dài thật (Task 9b). |
| y-tung-y | Các gạch đầu dòng và ba ý đúng. Lỗi gạch chân tiêu đề chung (xem tieu-de) — đã sửa. Giới hạn ký tự của mục `y` sau đó hạ 80→60 vì chữ 80 ký tự tràn khung (Task 9b). |
| quy-trinh | Ba ô kèm mũi tên, chữ xuống dòng trong ô không bị cắt. Ô trông trống khi chữ ngắn — coi là vấn đề thẩm mỹ, không phải lỗi, còn để lại. |
| so-sanh | Hai cột, đường chia không chồng lấn. Bốn dòng hai cột mỗi bên vẫn vừa khung nhưng khá sát nhau (~10 px giữa các dòng) — còn để lại. |
| do-thi | Đủ 4 điểm khoanh tròn và nối, nhãn trục "Chu kì T (s)", "Chiều dài l (m)" và các mốc 0,25/2, 1/2,84 đều hiện đúng. Trục ngang (`truc-ngang`) sau đó được nới rộng 440→680px và cỡ chữ 24→20 vì chữ dài thật xuống 2 dòng trong khung 40px (Task 9b). |
| thi-nghiem | Con lắc vẽ trong khung, cột phải hiện "Chiều dài dây l" và "Chu kì T". Ảnh xem trước ban đầu dừng ở trạng thái đầu cảnh (khoảng giây 1,4, con lắc còn ở giá trị nhỏ) — đã sửa để dừng ở trạng thái cuối cảnh (Task 9b, `thoiDiemCuoi()` trả về gần cuối cảnh thay vì ngay sau mốc câu cuối). |

Toàn bộ 8 cảnh sau khi sửa đều giữ dấu tiếng Việt nguyên vẹn, không chồng chữ lên phụ đề (thấy rõ nhất ở do-thi và thi-nghiem, nơi phụ đề từng đè lên nhãn trục và mép khung).

### 2.3 Video mẫu dựng thật (Task 9)

- Lệnh: `vmt python tools/vi/video_ma.py projects/_video/_thu_con_lac`, thuyết minh là 8 file mp3 sóng sin tổng hợp (8 giây/cảnh, 8 cảnh).
- Kết quả JSON: `ready: true`, `so_canh: 8`, `thoi_luong_giay: 74.67`, `giong: "co-san"`, 8 cảnh báo "mốc câu ước lượng theo số ký tự" (đúng như dự kiến vì dùng mp3 có sẵn, không có `notes/*.md` để tách mốc thật).
- Thời gian dựng: **99 giây** (lần chạy thành công đầu tiên là 91 giây).
- Dung lượng: **1.967.025 byte (1,97 MB)**, bit rate tổng 210,8 kb/s.
- Thời lượng: **74,667 giây** = 8 × 9,333 giây/cảnh, đúng công thức `thoi_luong_canh(8.0)`.
- `ffprobe`: luồng hình h264, **1280×720**, yuv420p, `r_frame_rate 30/1`, 2240 khung hình; luồng tiếng aac, 44100 Hz, mono. Đúng yêu cầu "chụp 15 khung/giây → dựng ra 30 khung/giây, có tiếng".
- File mp3 do "thầy cô" cung cấp (`giong/canh-1.mp3`) giữ nguyên mã băm trước và sau khi dựng — không bị ghi đè.

### 2.4 Các khung giữa cảnh đã xem (Task 9, `ffmpeg -ss`)

- **Giây 23,3 (cảnh 3 – cong-thuc):** công thức đã hiện đủ, ghi chú 1 hiện đủ, ghi chú 2 đang viết dở ("l là chiều dài dây, đ") đúng vị trí. Đây là bằng chứng chữ được "viết tay dở" theo mốc câu, không hiện hết ngay từ đầu cảnh.
- **Giây 60,7 (cảnh 7 – do-thi):** trục và nhãn đã vẽ, 2 trong 4 điểm đã khoanh tròn và nối, các điểm còn lại chưa vẽ — đúng kiểu "vẽ dần theo thời gian".
- **Giây 70,0 và 70,5 (cảnh 8 – thi-nghiem):** hai khung khác nhau rõ (lệnh `cmp` báo DIFFERENT). Con lắc chuyển động: ở giây 70,0 dây gần thẳng đứng (l = 1,33 m, T = 2,318 s); ở giây 70,5 con lắc đã lệch sang phải, dây dài hơn rõ (l = 1,43 m, T = 2,397 s) — xác nhận con lắc chuyển động thật trong video, không phải ảnh tĩnh.

### 2.5 Lỗi thật đã gặp khi dựng và cách sửa

1. **Đường gạch chân tiêu đề lệch khỏi chữ** (mọi cảnh dùng `tieuDe` dùng chung, tức cảnh 1, 4–8): tiêu đề canh trên, đường gạch chân đặt cố định phía dưới nên khi tiêu đề ngắn/1 dòng, khoảng cách lên tới ~300 px (cảnh 1) hoặc ~77–80 px (các cảnh còn lại). Sửa bằng cách canh chữ xuống đáy khung (`day: true`, flex column `margin-top:auto`) và đặt lại toạ độ đường gạch chân sát chữ. (commit `5f628a99`)
2. **Phụ đề in đè lên nội dung cảnh:** ở cảnh do-thi phụ đề đè lên nhãn trục "Chiều dài l (m)" và các mốc chia; ở cảnh thi-nghiem phụ đề cắt qua mép dưới khung thí nghiệm. Sửa bằng cách hạ `MarginV` (24→10) và nâng nội dung các cảnh lên trên đường an toàn y=630 (do-thi, thi-nghiem, so-sanh, khai-niem, y-tung-y, cong-thuc đều chỉnh toạ độ). (commit `67e4a2c2`)
3. **Đường dẫn thư mục tương đối làm FFmpeg không tìm thấy file tiếng:** lệnh dựng dùng đường dẫn tương đối `projects/_video/_thu_con_lac`, trong khi `ghep_video` chạy FFmpeg với `cwd=thu_muc` và cần đường dẫn tuyệt đối cho file mp3 trong lệnh `-i`; kết quả `ready:false`, lỗi "Error opening input". Sửa bằng `Path(args.thu_muc).resolve()` trong `video_ma.py`. (commit `f9833ae5`)
4. **Giới hạn ký tự không khớp kích thước khung** (phát hiện ở Task 9b khi kiểm bằng chữ tiếng Việt thật, không phải placeholder ngắn): mục `y` (y-tung-y) và `giai-thich` (cong-thuc) cho phép 80 ký tự nhưng ở cỡ chữ 30/28 px thì 80 ký tự xuống 2 dòng và tràn khung — hạ giới hạn 80→60. Ba khung khác tuy giới hạn ký tự không đổi (90/40/90) vẫn tràn khi thử chữ tiếng Việt dài thật: `bieu-thuc` (cong-thuc, cần 113px cho khung 110px), `truc-ngang` (do-thi, xuống 2 dòng trong khung 40px), `phu` (tieu-de, cần 95px cho khung 90px) — sửa bằng cách nới khung, không nới giới hạn chữ. (commit `c31df7f7`)
5. **Ảnh xem trước cảnh thí nghiệm dừng ở trạng thái đầu, không phải cuối cảnh:** `thoiDiemCuoi()` trả về ngay sau mốc câu tĩnh cuối cùng (khoảng giây 1,4/9,3), nên xem trước hiện tham số ban đầu (l = 0,69 m) thay vì giá trị cuối cùng mà con lắc đạt tới. Sửa để trả về gần cuối thời lượng cảnh (`thoiLuong - 0.2`). (commit `797064cf`)

### 2.5b Dựng demo với giọng máy thật

Chạy `video_ma.py` đầy đủ trên kịch bản mẫu 8 cảnh với giọng `vi-VN-HoaiMyNeural` lấy qua mạng (edge-tts 7.2.8), thư mục `projects/_video/con-lac-don-demo/`: `ready: true`, `giong: may`, 74,0 giây, 1,77 MB, 8 file `giong/canh-N.mp3` kèm sổ `canh-N.json` có mốc câu thật. Một cảnh báo: cảnh 3 phải ước lượng mốc câu vì edge-tts trả ít mốc hơn số câu tách được. Xem ba khung giữa video (giây 12, 41, 62): chữ đang viết dở, đồ thị đang nối điểm, phụ đề đúng câu đang đọc. Lưu ý môi trường: playwright và edge-tts phải nằm trong cùng một trình thông dịch; lệnh cài Chromium của `pptmaster.ps1` cài playwright vào `venv` của repo, nơi đã có edge-tts.

### 2.6 Chưa kiểm

Trung thực ghi nhận các mục sau **chưa** được kiểm trong các lần chạy trên, không suy diễn là đạt:

- **Cài Chromium từ đầu trên máy sạch** — máy dùng để kiểm (`C:/Users/ADMIN/vmt/v/Scripts/python.exe`) đã có sẵn playwright/Chromium từ trước; chưa thử cài mới hoàn toàn bằng `pptmaster.ps1 -Action tool -Name chromium` trên máy chưa có gì.
- **Chạy trong Antigravity thật** — chủ repo cần tự thử bằng câu "Làm video giải thích bài Con lắc đơn" để xác nhận file luật (`.agents/rules/ppt-master-vi.md`) và luồng hỏi/tạo có hoạt động đúng trong môi trường đó.

Các việc thẩm mỹ còn để lại, không chặn phát hành: ô `quy-trinh` trông trống khi chữ ngắn; cảnh `cong-thuc` không có tiêu đề nên đỉnh khung trống; phụ đề hơi sít chữ; con lắc nhỏ khi dây ngắn (nằm ở `thi_nghiem_parts/`, không thuộc phạm vi task này).
