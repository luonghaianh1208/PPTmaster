# AGENTS.vi.md — Quy tắc bổ sung cho bản Việt

File này bổ sung ngữ cảnh Việt Nam cho [AGENTS.md](AGENTS.md). Nó không thay thế quy tắc nào của dự án gốc.

## 1. Thứ tự ưu tiên

- `skills/ppt-master/SKILL.md` và `AGENTS.md` luôn được ưu tiên khi có mâu thuẫn với file này.
- Không bao giờ sửa, bỏ qua hay tìm cách "sửa chữa" `skills/ppt-master/scripts/attribution_guard.py`. Nếu kiểm tra toàn vẹn thất bại, dừng lại và hướng dẫn người dùng tải lại bản đầy đủ.
- `tools/vi/` và `docs/vi/` là lớp Việt hoá của bản phân phối này; `tools/vi/tests/` là test riêng của lớp đó.

## 2. Ngôn ngữ

- Giữ quy tắc ngôn ngữ trong `SKILL.md`: trả lời theo ngôn ngữ người dùng đang dùng.
- Khi người dùng viết tiếng Việt và không yêu cầu ngôn ngữ khác, đặt ngôn ngữ nội dung của dự án (`primary_language`) là `vi`.

## 3. Câu lệnh tiếng Việt kích hoạt skill `ppt-master`

"tạo PPT", "làm slide", "làm bài giảng", "tạo bài thuyết trình", "làm poster", "làm báo cáo", "thêm thuyết minh", "làm đẹp slide".

Các cụm "tạo nhanh", "làm nhanh", "không cần hỏi lại" là yêu cầu Quick tường minh (với bài tạo mới thông thường là hồ sơ `workflows/profiles/quick-generate.md` của upstream). Việc chọn hồ sơ và các bước thực hiện vẫn theo đúng `SKILL.md`.

## 4. Chạy lệnh trên Windows

- Có `venv\Scripts\python.exe` ở thư mục gốc repo thì dùng nó cho mọi lệnh `python3 …` hoặc `python …` của repo, ví dụ `venv\Scripts\python.exe skills/ppt-master/scripts/project_manager.py init <tên_dự_án>`.
- Nếu lệnh `python3 ...` báo không tìm thấy, trả mã 49 hoặc mở Microsoft Store, chạy lại đúng lệnh đó với `python`.
- Không dùng `cp`, `mkdir -p`, `/tmp` hay heredoc của bash trong PowerShell. Dùng lệnh tương đương (`Copy-Item`, `New-Item -ItemType Directory -Force`, `$env:TEMP`) hoặc Python.
- Khi ghi file có chữ tiếng Việt từ PowerShell 5.1, không dùng `Set-Content` hay `Out-File` mặc định. Dùng Python hoặc `[IO.File]::WriteAllText(path, text, (New-Object Text.UTF8Encoding $false))`.

## 5. Khổ canvas theo cách gọi của người Việt

<!-- format-map:start -->
| Người dùng nói | Key canvas |
|---|---|
| Slide 16:9, trình chiếu | `ppt169` |
| Slide 4:3, máy chiếu cũ | `ppt43` |
| Bài đăng Facebook/TikTok dọc 3:4 | `xiaohongshu` |
| Ảnh vuông Zalo/Facebook/Instagram | `moments` |
| Story, Reels, TikTok 9:16 | `story` |
| Ảnh bìa bài viết Zalo OA (2.35:1) | `wechat` |
<!-- format-map:end -->

## 6. Font

Với chữ tiếng Việt, chỉ dùng font có sẵn trên Windows hiển thị đủ dấu: Segoe UI, Arial, Calibri, Times New Roman (văn bản hành chính). Không dùng PingFang SC hay Microsoft YaHei cho chữ tiếng Việt.

## 7. Giọng thuyết minh (edge-tts)

- Giọng nữ: `vi-VN-HoaiMyNeural`
- Giọng nam: `vi-VN-NamMinhNeural`

## 8. Kết quả

Sau khi xuất, cho người dùng biết đường dẫn file PPTX trong thư mục dự án dưới `projects/`.

## 9. Môi trường: tự kiểm tra, tự cài và xử lý lỗi

- Mỗi cuộc trò chuyện, trước lệnh Python đầu tiên của repo (ví dụ `project_manager.py init`, `source_to_md.py`): chạy `tools/vi/doctor.py --no-smoke --json` (bằng Python của `venv` nếu có). Kết quả có `"ready": true` thì làm tiếp; chưa sẵn sàng, hoặc không chạy được Python, thì làm theo [docs/vi/cai-dat-bang-ai.md](docs/vi/cai-dat-bang-ai.md) rồi mới chạy lệnh đó.
- Yêu cầu thuộc mục 10: chạy kiểm tra trên (chỉ đọc, vài giây) trước khi gửi tin nhắn hỏi thầy cô. Chưa sẵn sàng thì thêm đúng một dòng ở cuối tin nhắn hỏi (ngay trước dòng chốt cách xác nhận, nếu có): "Máy chưa cài xong bộ công cụ; sau khi thầy cô trả lời, em sẽ cài trước (khoảng 5–10 phút) rồi làm bài." Thầy cô trả lời xong thì cài ngay theo [docs/vi/cai-dat-bang-ai.md](docs/vi/cai-dat-bang-ai.md), rồi mới làm tiếp.
- Lệnh Python nào của repo báo không tìm thấy Python (kể cả sau khi chạy lại với `python` theo mục 4) hoặc báo `ModuleNotFoundError`: dừng, làm theo [docs/vi/cai-dat-bang-ai.md](docs/vi/cai-dat-bang-ai.md); không tự cài Python hay thư viện theo cách khác.
- Người dùng nhờ cài đặt, kiểm tra máy, hoặc dán link repo để cài: làm theo [docs/vi/cai-dat-bang-ai.md](docs/vi/cai-dat-bang-ai.md).
- Cần FFmpeg (thuyết minh, video) hoặc Pandoc (tài liệu định dạng cũ như `.doc`, `.odt`, `.rtf`) mà máy chưa có: làm theo mục "Công cụ tuỳ chọn" của file đó.
- Người dùng gặp lỗi môi trường: đề nghị chạy `KIEM-TRA.bat` (Windows) hoặc `python tools/vi/doctor.py`, rồi đối chiếu với [docs/vi/xu-ly-loi.md](docs/vi/xu-ly-loi.md).

## 10. Hỗ trợ thầy cô trước khi tạo PPTX

Khi người dùng viết tiếng Việt và yêu cầu thuộc một trong 5 loại việc dưới đây, đọc [docs/vi/tro-ly/quy-trinh-hoi.md](docs/vi/tro-ly/quy-trinh-hoi.md) trước, rồi đọc file của loại việc đó. Hỏi thầy cô một lượt và chờ trả lời trước khi khởi tạo dự án.

| Loại việc | File hướng dẫn |
|---|---|
| Bài giảng | [docs/vi/tro-ly/bai-giang.md](docs/vi/tro-ly/bai-giang.md) |
| Báo cáo – tổng kết | [docs/vi/tro-ly/bao-cao-tong-ket.md](docs/vi/tro-ly/bao-cao-tong-ket.md) |
| Hoạt động Đoàn – sự kiện | [docs/vi/tro-ly/hoat-dong-doan.md](docs/vi/tro-ly/hoat-dong-doan.md) |
| Poster/ấn phẩm Zalo – Facebook | [docs/vi/tro-ly/poster-mang-xa-hoi.md](docs/vi/tro-ly/poster-mang-xa-hoi.md) |
| Tập huấn/workshop | [docs/vi/tro-ly/tap-huan-workshop.md](docs/vi/tro-ly/tap-huan-workshop.md) |

- `SKILL.md` vẫn được ưu tiên. Lượt hỏi này chỉ tạo thêm tài liệu nguồn; bước xác nhận của upstream vẫn bắt buộc, trừ khi người dùng yêu cầu tạo nhanh (xem mục 3). Khi tạo nhanh, kể cả với "không cần hỏi lại", vẫn đọc `docs/vi/tro-ly/quy-trinh-hoi.md` và làm theo mục "Tạo nhanh" của file đó: có thể không hỏi câu nào, nhưng vẫn ghi brief.
- Yêu cầu không thuộc 5 loại (bối cảnh trường học hay Đoàn một mình không đủ để xếp loại), hoặc người dùng không viết tiếng Việt: làm theo `SKILL.md` như bình thường, không tìm hồ sơ đơn vị và không dùng bộ câu hỏi Việt.
