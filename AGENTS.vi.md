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

## 4. Chạy lệnh trên Windows

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

## 9. Khi người dùng gặp lỗi môi trường

Đề nghị chạy `KIEM-TRA.bat` (Windows) hoặc `python tools/vi/doctor.py`, rồi đối chiếu với [docs/vi/xu-ly-loi.md](docs/vi/xu-ly-loi.md).
