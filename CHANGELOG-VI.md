# Nhật ký thay đổi — Bản Việt

## 6.3.2-vi.4 — 2026-09-12

Làm video bài giảng: bài giảng đã có thành video MP4 có lời giảng tiếng Việt và phụ đề, chạy được cả khi máy không có PowerPoint.

### Thêm
- `tools/vi/video.py`: một lệnh để dựng video, tự chọn đường PowerPoint (giữ hiệu ứng chuyển cảnh) hoặc đường FFmpeg ghép ảnh slide (không cần PowerPoint). Kết quả trả về dạng JSON cho AI đọc.
- Phụ đề `.srt` dựng từ phụ đề từng slide: đường FFmpeg cộng dồn thời lượng tiếng, đường PowerPoint đọc mốc thời gian trong bản PPTX đã gắn tiếng. Thầy cô chọn để file rời hoặc in lên hình.
- `-Action tool -Name chromium` cài Chromium khi cần chụp ảnh slide, và chỉ khi thật sự cần chụp.
- Loại việc thứ 6 cho AI: `docs/vi/tro-ly/video-bai-giang.md` (4 câu hỏi: giọng đọc, tốc độ, phụ đề, độ phân giải); `AGENTS.vi.md` mục 11 nêu thứ tự làm và cách chuyển câu trả lời thành tham số.
- Tài liệu `docs/vi/lam-video.md` và mục "Dựng video thất bại" trong Xử lý lỗi, viết theo từng mã lỗi.

### Không thay đổi
- Lõi PPT Master v6.3.2 của Hugo He giữ nguyên. Tiếng đọc vẫn do `edge-tts` của dự án gốc tạo, giọng `vi-VN-HoaiMyNeural` và `vi-VN-NamMinhNeural`.

### Rủi ro
- Đường FFmpeg xuất video đúng bằng kích thước ảnh slide (1280×720 với khổ 16:9), không kéo giãn lên 1080. Cần 1080 thật thì chờ bản sau.
- Đường PowerPoint mở cửa sổ PowerPoint và chiếm máy vài phút. Bài dài khoảng 15 slide trở lên sẽ hiện cảnh báo phụ đề lệch chừng 1 giây, do PowerPoint tự thêm khoảng đệm mỗi slide; cần phụ đề chính xác thì dùng đường FFmpeg.
- Chủ repo chưa nghiệm thu trên một bài giảng thật của mình. Bằng chứng hiện có là một lần dựng thật trên dự án mẫu 3 slide, cả hai đường đều ra video xem được.
- Video 10 phút nặng khoảng 100–300 MB; ổ đĩa nên còn trống ít nhất 2 GB. Tiếng máy đọc có thể sai tên riêng nước ngoài và công thức, nên nghe lại trước khi giao cho học sinh.

## 6.3.2-vi.3 — 2026-09-12

AI tự cài đặt: thầy cô chỉ cần dán link repo vào Antigravity (hoặc mở thư mục đã tải), AI tự kiểm tra máy, cài những gì còn thiếu rồi báo khi sẵn sàng tạo slide.

### Thêm
- `tools/vi/pptmaster.ps1 -Action setup -Auto`: cài Python 3.12 cho riêng tài khoản (không cần quyền quản trị), tạo môi trường Python riêng `venv\`, cài thư viện, tạo `.env`, kiểm tra có xuất thử PPTX, trả kết quả JSON cho AI.
- `-Action tool -Name ffmpeg|pandoc`: chỉ cài FFmpeg hoặc Pandoc khi cần.
- `tools/vi/doctor.py --json`.
- Hướng dẫn cài cho AI `docs/vi/cai-dat-bang-ai.md`; `AGENTS.vi.md` mục 9 cho AI tự kiểm tra môi trường trước lệnh Python đầu tiên của repo trong mỗi cuộc trò chuyện.
- Tài liệu: mục "Để AI tự cài" trong Bắt đầu nhanh, mục "Máy trường chặn cài đặt" kèm đoạn gửi bộ phận IT.

### Không thay đổi
- `CAI-DAT.bat`, `KIEM-TRA.bat`, `CAP-NHAT.bat` vẫn dùng như cũ (cả ba ưu tiên `venv` nếu có).
- Lõi PPT Master v6.3.2 của Hugo He giữ nguyên.

### Rủi ro
- Thư mục bộ công cụ ở đường dẫn dài (khoảng 150 ký tự) có thể làm cài thư viện lỗi vì giới hạn đường dẫn của Windows. Bộ cài cảnh báo khi đường dẫn dài hơn 80 ký tự; nên đặt bộ công cụ ở `D:\PPTmaster`.
- Bộ cài dùng Python 3.10 trở lên có sẵn trên máy, kể cả bản mới hơn 3.12 có thể chưa có gói build sẵn cho vài thư viện; khi đó lỗi hiện ở bước cài thư viện. Máy có Python 3.12 cài cho mọi người dùng nhưng không có trong PATH hay `py` có thể không được nhận ra.
- Bước AI tự cài Python (winget hoặc bộ cài python.org) chưa được chạy thật trên một máy chưa có Python. Nếu bước này lỗi, cài Python theo [Cài đặt trên Windows](docs/vi/cai-dat-windows.md) rồi chạy lại.

## 6.3.2-vi.2 — 2026-09-11

Thêm trợ lý hỏi đáp cho thầy cô: AI hỏi một lượt ngắn trước khi tạo slide.

### Thêm
- Bộ hướng dẫn cho AI trong `docs/vi/tro-ly/`: quy trình hỏi chung và 5 loại việc (bài giảng, báo cáo – tổng kết, hoạt động Đoàn – sự kiện, poster Zalo – Facebook, tập huấn/workshop).
- Hồ sơ đơn vị lưu một lần ở `projects/_ho-so-don-vi.md` (không đưa lên GitHub); câu trả lời của thầy cô được ghi thành brief trong thư mục nguồn của dự án.
- "tạo nhanh" chỉ hỏi 2–3 câu thật cần thiết; "không cần hỏi lại" thì không hỏi câu nào.
- Bước xác nhận có thể làm ngay trong khung chat, không cần mở trang web.
- `AGENTS.vi.md` mục 10 trỏ AI tới bộ hướng dẫn; tài liệu người dùng giải thích lượt hỏi.

### Không thay đổi
- Lõi PPT Master v6.3.2 của Hugo He giữ nguyên; bước xác nhận của dự án gốc vẫn áp dụng.

## 6.3.2-vi.1 — 2026-09-11

Chuyển sang nền PPT Master **v6.3.2** của Hugo He và đóng gói lại cho người dùng Việt Nam.

### Thêm
- Ghi công dự án gốc trong `NOTICE` và `README.md`.
- Bộ cài một lần bấm cho Windows: `CAI-DAT.bat`, `KIEM-TRA.bat`, `CAP-NHAT.bat`; `tools/vi/setup.sh` cho macOS/Linux.
- Kiểm tra môi trường `tools/vi/doctor.py`, có xuất thử một file PPTX tiếng Việt.
- Quy tắc tiếng Việt cho AI editor: `AGENTS.vi.md`, tự nạp trong Claude Code, Cursor, Antigravity.
- Tài liệu tiếng Việt trong `docs/vi/`.

### Thay đổi so với bản cũ (v2.x)
- Toàn bộ lõi cập nhật lên v6.3.2: PPTX chỉnh sửa trực tiếp, hiệu ứng động, thuyết minh, làm đẹp PPTX có sẵn.
- Khổ Zalo/Facebook dùng tên key của dự án gốc (`moments`, `xiaohongshu`, `wechat`); bảng quy đổi nằm trong `AGENTS.vi.md`.

### Gỡ bỏ
- Thư mục `examples/`, `viewer.html`, `index.html` cũ. Bộ ví dụ xem tại https://github.com/hugohe3/ppt-master-examples.
- Trạng thái bản cũ vẫn giữ ở tag `v2-vi-legacy`.
