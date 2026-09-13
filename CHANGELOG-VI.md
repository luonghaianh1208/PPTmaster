# Nhật ký thay đổi — Bản Việt

## 6.3.2-vi.6 — 2026-09-13

Soạn giáo án tích hợp năng lực số và năng lực AI: từ giáo án cũ (Word hoặc PDF) hoặc chỉ từ tên bài, ra file Word Kế hoạch bài dạy theo Công văn 5512, cho mọi môn.

### Thêm
- `tools/vi/giao_an.py xuat`: đọc `giao-an.md`, kiểm cấu trúc 5512 và mã năng lực, dựng `giao-an.docx` (A4 dọc, lề 2/2/2,5/2 cm, Times New Roman 14pt, giãn dòng 1,3; phiếu học tập căn trái; rubric dạng bảng) và `can-soat.md` (ghi chú nội bộ, không nằm trong giáo án). `--plan-only` chỉ kiểm. Kết quả trả về dạng JSON cho AI đọc.
- Công cụ chặn giáo án thiếu mục tiêu, nội dung, sản phẩm hoặc bốn bước tổ chức của một hoạt động; chặn giáo án chỉ nhắc năng lực số hay AI ở mục tiêu mà tiến trình không có; chặn mã năng lực không đúng bảng; chặn rubric thiếu mức. Lỗi cấu trúc nêu đúng dòng cần sửa.
- `tools/vi/giao_an.py trich-sgk`: cắt đúng phần một bài ra khỏi SGK đã chuyển sang Markdown, để AI không phải nạp cả quyển; không bao giờ ghi đè file SGK nguồn.
- Khung năng lực số (11 mã của Bộ GD&ĐT) và khung năng lực AI môn Hoá học (lớp 10, 11, 12, hệ chuyên) trong `docs/vi/tro-ly/nang-luc-so-va-ai.md`, do Lương Hải Anh — 2Anh AI Education biên soạn; đây là nguồn mã duy nhất của công cụ. Môn chưa có khung AI thì ô mã ghi "(chưa có khung mã cho môn này)", AI không tự đặt mã.
- Loại việc thứ 8 cho AI: `docs/vi/tro-ly/giao-an.md`; `AGENTS.vi.md` mục 13. Câu lệnh chỉ nói "giáo án" thì AI hỏi lại cần file Word hay slide; "kế hoạch bài dạy", "KHBD" đi thẳng vào loại việc này. Giáo án cũ được đọc trước khi hỏi, nội dung chuyên môn của thầy cô được giữ nguyên; file phân phối chương trình chỉ được đọc, không sửa.
- Tài liệu `docs/vi/soan-giao-an.md` và mục "Xuất giáo án thất bại" trong Xử lý lỗi.

### Không thay đổi
- Lõi PPT Master v6.3.2 của Hugo He giữ nguyên. Soạn giáo án không tạo dự án PPTX và không dùng bước xác nhận của dự án gốc. Dùng lại thư viện `python-docx` và tầng dựng Word của bản 6.3.2-vi.5.

### Rủi ro
- `trich-sgk` với SGK chuyển ra toàn dòng thường (không có tiêu đề `#`) và có mục lục: cắt bài **cuối cùng** của sách sẽ lấy cả phần từ mục lục tới hết sách (chỉ có cảnh báo lớn hơn 200 KB). Các bài khác và sách có tiêu đề `#` không bị. Sẽ sửa ở bản sau.
- Khung mã AI mới chỉ có môn Hoá học; khung hệ chuyên có một mã, thiên về hữu cơ và phổ, nên có thể không hợp một số bài vô cơ. Môn khác dùng được nhưng ô mã AI để trống, hồ sơ có thể bị tổ chuyên môn hỏi lại.
- Công cụ kiểm cấu trúc và mã, không kiểm chất lượng sư phạm. Chưa có khung ký duyệt và quốc hiệu.
- Chủ repo chưa mở file Word để nghiệm thu bằng mắt. Bằng chứng hiện có: chạy thật trên một giáo án Bài 5 thật (giữ đủ 5 hoạt động, 90 phút), một giáo án tự soạn và một giáo án Ngữ văn, kiểm tra cấu trúc file bằng test.
- Giáo án có bảng lồng nhau hoặc công thức dạng ảnh có thể mất định dạng khi chuyển sang Markdown; AI phải ghi chỗ không đọc được vào "Cần thầy cô soát".

## 6.3.2-vi.5 — 2026-09-13

Soạn đề kiểm tra KHTN bằng tiếng Anh: chuyển đề tiếng Việt có sẵn hoặc soạn đề mới, dùng thuật ngữ khoa học đúng môn thay vì dịch từng chữ, xuất file Word đúng thể thức đề thi.

### Thêm
- `tools/vi/de_thi.py`: đọc `de.md` rồi dựng ba file Word `de-en.docx` (đề tiếng Anh), `de-song-ngu.docx` (song ngữ Anh – Việt) và `dap-an.docx` (đáp án, thang điểm, ma trận đặc tả, mục "Cần thầy cô soát"). `--phan de,dap-an` bỏ bản song ngữ; `--plan-only` chỉ kiểm cú pháp. Kết quả trả về dạng JSON cho AI đọc.
- Cấu trúc đề theo định dạng hiện hành: Phần I trắc nghiệm nhiều lựa chọn (0,25 điểm/câu), Phần II đúng/sai (1 điểm/câu), Phần III trả lời ngắn (0,25 điểm/câu); áp dụng cho KHTN THCS 6–9 và Vật lí, Hoá học, Sinh học THPT 10–12. Khổ A4, Times New Roman, đáp án ngắn tự xếp 4 cột, có số trang.
- Loại việc thứ 7 cho AI: `docs/vi/tro-ly/de-khtn-tieng-anh.md` (ngữ pháp `de.md` và đề mẫu) và `docs/vi/tro-ly/tieng-anh-khoa-hoc.md` (quy ước thuật ngữ, đơn vị, ký hiệu, cách gọi tên chất); `AGENTS.vi.md` mục 12 nêu thứ tự làm cho hai luồng.
- Thư viện `python-docx` (`tools/vi/requirements-vi.txt`) được cài cùng bộ công cụ ở mức khuyến nghị: cài hỏng thì chỉ cảnh báo, máy vẫn làm PPTX bình thường. `doctor.py` có thêm mục "Thư viện lớp Việt".
- Tài liệu `docs/vi/soan-de-tieng-anh.md` và mục "Xuất đề Word thất bại" trong Xử lý lỗi.

### Sửa
- `tools/vi/video.py` ghi JSON bằng UTF-8 khi console Windows dùng bảng mã cũ (commit sau tag 6.3.2-vi.4, nay mới vào bản phát hành).

### Không thay đổi
- Lõi PPT Master v6.3.2 của Hugo He giữ nguyên. Soạn đề không tạo dự án PPTX và không dùng bước xác nhận của dự án gốc.

### Rủi ro
- Chủ repo chưa mở các file Word mẫu để nghiệm thu bằng mắt (viền bảng, số trang). Bằng chứng hiện có là hai lần dựng thật (Vật lí 10, KHTN 8) và kiểm tra cấu trúc file bằng test.
- Đáp án Phần III chỉ nhận số (ví dụ `36`, `12,5`, `-0.25`); câu cần đáp án bằng chữ phải chuyển sang Phần I hoặc II.
- Chất lượng tiếng Anh phụ thuộc AI; công cụ không có từ điển hay bộ kiểm thuật ngữ. Thầy cô cần đọc lại đề, nhất là các mục trong "Cần thầy cô soát".
- Máy cài hỏng `python-docx` sẽ báo sẵn sàng kèm cảnh báo; lần soạn đề đầu tiên báo lỗi `docx` kèm lệnh cài. Chế độ `-Auto` của bộ cài sẽ thử cài lại thư viện này mỗi lần chạy.
- `video.py` vẫn có thể in hai dòng JSON trong trường hợp hiếm stdout bị lỗi khi đang ghi.

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
