# Xử lý lỗi

Gặp lỗi, hãy chạy `KIEM-TRA.bat` trước, rồi đối chiếu dòng ❌ [LỖI] (hoặc ⚠️ [CẢNH BÁO]) với các mục bên dưới.

## Gõ python mà mở Microsoft Store

Lệnh `python` đang bị lối tắt của Microsoft Store chiếm chỗ, chưa phải Python thật.

- Vào **Settings → Apps → Advanced app settings → App execution aliases**, tắt `python.exe` và `python3.exe`.
- Hoặc cài Python từ https://www.python.org/downloads/, nhớ tick **"Add python.exe to PATH"**.
- Sau khi sửa, đóng cửa sổ dòng lệnh đang mở rồi mở lại (hoặc bấm lại `CAI-DAT.bat`).

## Đã cài Python nhưng bộ cài báo không tìm thấy

`CAI-DAT.bat` báo **"Đã cài Python ... nhưng bản này chưa có trong PATH"**, hoặc gõ `python` báo không tìm thấy dù máy đã cài Python. Nguyên nhân thường gặp: lúc cài chưa tick **"Add python.exe to PATH"** (bộ cài của python.org mặc định không tick ô này). Cài lại bằng winget không giúp được vì Python đã có sẵn trên máy.

1. Mở **Settings → Apps → Installed apps** (Windows 10: **Apps & features**), tìm **Python 3.x** rồi chọn **Modify**. Cách khác: chạy lại bộ cài đã tải từ python.org và chọn **Modify**.
2. Bấm **Next** tới trang **Advanced Options**, tick **"Add Python to environment variables"**, rồi bấm **Install**.
3. Đóng cửa sổ dòng lệnh đang mở, mở lại rồi bấm lại `CAI-DAT.bat`.

## PowerShell bị chặn trên máy trường hoặc công ty

Bấm `CAI-DAT.bat` hoặc `KIEM-TRA.bat` mà thấy thông báo tiếng Anh có cụm **"running scripts is disabled on this system"** hoặc nhắc tới **execution policy**, rồi dừng: máy do nhà trường hoặc công ty quản lý đang chặn chạy script PowerShell. Chính sách này mạnh hơn tuỳ chọn mà các file `.bat` dùng nên bộ cài không tự vượt qua được.

- Nhờ quản trị máy (bộ phận IT) cho phép chạy script PowerShell.
- Hoặc tự chạy từng lệnh trong cửa sổ dòng lệnh mở tại thư mục bộ công cụ:
  ```
  python -m pip install -r requirements.txt
  python tools\vi\doctor.py
  ```
  Nếu chưa có file `.env`, tạo bằng lệnh `copy .env.example .env`.

## Máy trường chặn cài đặt

Dấu hiệu: AI báo không tải được bộ công cụ hoặc Python, không chạy được bộ cài Python, không cài được thư viện vì mạng, hoặc PowerShell bị chặn chạy script. Máy do nhà trường quản lý có thể chặn các việc này; AI không tìm cách vượt qua.

Gửi bộ phận IT đoạn sau:

> Nhờ anh/chị hỗ trợ để tôi dùng bộ công cụ PPT Master trên máy này:
> 1. Cho tài khoản Windows của tôi truy cập: python.org, pypi.org, files.pythonhosted.org, github.com, codeload.github.com; thêm cdn.winget.microsoft.com, objects.githubusercontent.com (khi cần FFmpeg/Pandoc).
> 2. Cho phép cài Python 3.12 cho riêng tài khoản của tôi (không cần quyền quản trị).
> 3. Cho phép chạy PowerShell với tuỳ chọn `-ExecutionPolicy Bypass` cho từng lệnh.

Nếu chỉ bị chặn chạy script PowerShell mà máy đã có Python, xem mục **PowerShell bị chặn trên máy trường hoặc công ty** ở trên.

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

- Thư mục bộ công cụ có thư mục `venv` (bộ công cụ do AI cài): nhờ AI chạy lại lệnh cài, hoặc chạy lệnh sau trong cửa sổ dòng lệnh mở tại thư mục bộ công cụ:
  ```
  venv\Scripts\python.exe -m pip install -r requirements.txt
  ```
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
- Nếu thông báo có `FileNotFoundError` kèm một đường dẫn rất dài, xem mục **Đường dẫn quá dài** ngay bên dưới.
- Nếu vẫn lỗi, khi hỏi hỗ trợ hãy chụp toàn bộ màn hình kết quả `KIEM-TRA.bat` để gửi kèm.

## Đường dẫn quá dài

Dấu hiệu: `KIEM-TRA.bat` hoặc lúc xuất bài báo lỗi có chữ `FileNotFoundError`, kèm một đường dẫn rất dài (thường chứa `.pptx-build-`). Windows mặc định chỉ cho phép đường dẫn dài khoảng 260 ký tự, trong khi quá trình xuất PPTX tạo thêm nhiều thư mục con bên trong thư mục dự án.

- Dời cả thư mục bộ công cụ sang đường dẫn ngắn, ví dụ `D:\PPTmaster`.
- Tránh giải nén ZIP thành thư mục lồng nhau như `PPTmaster-main\PPTmaster-main`, và tránh đặt trong thư mục OneDrive nhiều cấp.
- Hoặc nhờ quản trị máy bật **Long Paths** của Windows (cần quyền quản trị).

## Cập nhật thất bại

- Thông báo **"not a Git checkout"**: thư mục này là bản tải bằng ZIP, không tự cập nhật được. Tải bản mới rồi chép thư mục `projects\` và file `.env` của bạn sang.
- Thông báo **"Tracked local changes"**: bạn đã sửa file thuộc bộ công cụ. Chạy `git status` để xem file nào đã đổi, chuyển các file cá nhân của bạn vào thư mục `projects\`, rồi khôi phục file bộ công cụ bằng:
  ```
  git restore <tên file>
  ```
  **Cảnh báo:** lệnh này sẽ xoá mọi thay đổi bạn đã làm trong file đó.
- Thiếu Git: cài Git (bấm `CAI-DAT.bat`, bộ cài sẽ đề nghị cài Git).

## Dựng video thất bại

Xem dòng kết quả AI đọc được, phần `error`:

- `chromium`: máy chưa có Chromium để chụp ảnh slide. Cho AI chạy `powershell -NoProfile -ExecutionPolicy Bypass -File tools\vi\pptmaster.ps1 -Action tool -Name chromium` (tải 150–300 MB).
- `ffmpeg`: máy chưa có FFmpeg. Cho AI chạy lệnh trên với `-Name ffmpeg`.
- `audio`: bài giảng chưa có tiếng đọc. Nhờ AI tạo lời giảng và tiếng đọc trước.
- `powerpoint`: PowerPoint không xuất được video. Thử lại bằng cách ghép ảnh: `venv\Scripts\python.exe tools\vi\video.py <đường_dẫn_dự_án> --cach ffmpeg`.
- `render`: thường do hết dung lượng ổ đĩa hoặc đường dẫn quá dài. Dọn ổ đĩa, hoặc chuyển bộ công cụ sang `D:\PPTmaster` rồi làm lại.

## Chữ tiếng Việt lỗi dấu trong PowerPoint

Thường do máy không có sẵn font đang dùng. Nhắn cho AI, ví dụ "đổi toàn bộ font sang Segoe UI" (hoặc Arial, Times New Roman), rồi xuất lại file.

## Hỏi hỗ trợ

Tạo issue tại https://github.com/luonghaianh1208/PPTmaster/issues, kèm theo kết quả chạy `KIEM-TRA.bat`. **Xoá mọi API key** trong nội dung trước khi gửi.
