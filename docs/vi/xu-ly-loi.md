# Xử lý lỗi

Gặp lỗi, hãy chạy `KIEM-TRA.bat` trước, rồi đối chiếu dòng ❌ với các mục bên dưới.

## Gõ python mà mở Microsoft Store

Lệnh `python` đang bị lối tắt của Microsoft Store chiếm chỗ, chưa phải Python thật.

- Vào **Settings → Apps → Advanced app settings → App execution aliases**, tắt `python.exe` và `python3.exe`.
- Hoặc cài Python từ https://www.python.org/downloads/, nhớ tick **"Add python.exe to PATH"**.
- Sau khi sửa, đóng cửa sổ dòng lệnh đang mở rồi mở lại (hoặc bấm lại `CAI-DAT.bat`).

## Cài thư viện thất bại

Vài nguyên nhân thường gặp:

- **Mạng trường học/công ty có proxy chặn:** thử đổi sang mạng khác (ví dụ 4G), hoặc chạy:
  ```
  python -m pip install --proxy http://<máy chủ>:<cổng> -r requirements.txt
  ```
- **Thiếu quyền ghi:**
  ```
  python -m pip install --user -r requirements.txt
  ```
- **Python quá mới, chưa có bản build cho vài gói:** cài thêm Python 3.12 và dùng bản đó.
- **Phần mềm diệt virus chặn quá trình cài:** tạm tắt diệt virus trong lúc cài, bật lại sau khi xong.

## KIEM-TRA báo thiếu thư viện

- Bấm lại `CAI-DAT.bat` để cài lại thư viện.
- Nếu máy có nhiều bản Python, lệnh `python` trong PATH có thể không phải bản đã cài thư viện. Kiểm tra bằng:
  ```
  python -c "import sys; print(sys.executable)"
  ```
  và đảm bảo đây là bản bạn đã dùng để cài thư viện.

## Lỗi "Tính toàn vẹn skill"

Bộ công cụ bị sửa đổi hoặc thiếu file bản quyền (`LICENSE`, `SKILL.md`, `SPONSORS*.md`).

- Tải lại bản đầy đủ từ repo.
- Không tự chỉnh sửa các file trên.

## Xuất thử PPTX thất bại

- Chạy lại `KIEM-TRA.bat` để xem lỗi có lặp lại không.
- Kiểm tra xem phần mềm diệt virus có đang quét thư mục Temp của Windows không — việc quét có thể làm chậm hoặc chặn quá trình xuất file.
- Nếu vẫn lỗi, khi hỏi hỗ trợ hãy chụp toàn bộ màn hình kết quả `KIEM-TRA.bat` để gửi kèm.

## Cập nhật thất bại

- Thông báo **"not a Git checkout"**: thư mục này là bản tải bằng ZIP, không tự cập nhật được. Tải bản mới và chép thư mục `projects\` của bạn sang.
- Thông báo **"Tracked local changes"**: bạn đã sửa file thuộc bộ công cụ. Chạy `git status` để xem file nào đã đổi, chuyển các file cá nhân của bạn vào thư mục `projects\`, rồi khôi phục file bộ công cụ bằng:
  ```
  git restore <tên file>
  ```
  **Cảnh báo:** lệnh này sẽ xoá mọi thay đổi bạn đã làm trong file đó.
- Thiếu Git: cài Git (bấm `CAI-DAT.bat`, bộ cài sẽ đề nghị cài Git).

## Chữ tiếng Việt lỗi dấu trong PowerPoint

Thường do máy không có sẵn font đang dùng. Nhắn cho AI, ví dụ "đổi toàn bộ font sang Segoe UI" (hoặc Arial, Times New Roman), rồi xuất lại file.

## Hỏi hỗ trợ

Tạo issue tại https://github.com/luonghaianh1208/PPTmaster/issues, kèm theo kết quả chạy `KIEM-TRA.bat`. **Xoá mọi API key** trong nội dung trước khi gửi.
