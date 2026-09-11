# Cài đặt bộ công cụ bằng AI

File dành cho AI agent (Antigravity, Claude Code, Cursor) trên Windows. Làm đúng các bước dưới đây, không tự nghĩ cách cài khác. Xưng "em", gọi người dùng là "thầy cô".

## Khi nào dùng

- Người dùng nhờ cài PPT Master, dán link `https://github.com/luonghaianh1208/PPTmaster`, hoặc nhờ "cài đặt", "kiểm tra máy".
- `AGENTS.vi.md` mục 9 báo môi trường chưa sẵn sàng trước khi tạo slide.
- Máy không phải Windows: không dùng file này; chạy `sh tools/vi/setup.sh` như README.

## Tải bộ công cụ

Bỏ qua mục này nếu thư mục đang mở đã có `tools/vi/pptmaster.ps1`.

1. Chọn thư mục đích: thư mục đang mở còn trống thì dùng chính nó; đã có file khác thì dùng thư mục con `PPTmaster` bên trong (tạo bằng `New-Item -ItemType Directory -Force <thư_mục_đích>`).
2. Chọn nhánh: mặc định là `main`. Người dùng nêu rõ một nhánh khác (ví dụ "nhánh feat/vi-tu-cai") thì dùng nhánh đó.
3. Có Git (`git --version` chạy được):

   ```
   git clone -b <nhánh> https://github.com/luonghaianh1208/PPTmaster.git <thư_mục_đích>
   ```

   Với nhánh `main` có thể bỏ `-b <nhánh>`: `git clone https://github.com/luonghaianh1208/PPTmaster.git <thư_mục_đích>`.

4. Không có Git: tải ZIP của nhánh rồi giải nén bằng PowerShell:

   ```
   $zip = Join-Path $env:TEMP 'PPTmaster.zip'
   $unzip = Join-Path $env:TEMP ('PPTmaster-zip-' + [guid]::NewGuid())
   Invoke-WebRequest -Uri 'https://github.com/luonghaianh1208/PPTmaster/archive/refs/heads/<nhánh>.zip' -OutFile $zip -UseBasicParsing
   Expand-Archive -Path $zip -DestinationPath $unzip -Force
   $src = Get-ChildItem -Path $unzip -Directory | Select-Object -First 1
   Get-ChildItem -Path $src.FullName -Force | Move-Item -Destination '<thư_mục_đích>'
   ```

5. Đọc `AGENTS.md` và `AGENTS.vi.md` trong thư mục đích và áp dụng từ đây. Mọi lệnh sau chạy từ thư mục đích.

## Cài đặt

1. Gửi thầy cô đúng một tin nhắn, không hỏi lại: "Em sẽ cài Python và thư viện cho PPT Master, mất khoảng 5–10 phút. Nếu Antigravity hỏi cho phép chạy lệnh, thầy cô bấm đồng ý giúp em."
2. Chạy từ thư mục đích:

   ```
   powershell -NoProfile -ExecutionPolicy Bypass -File tools\vi\pptmaster.ps1 -Action setup -Auto
   ```

3. Tiến trình cài hiện ở stderr. Kết quả là đúng một dòng JSON ở stdout; đọc riêng stdout, không gộp stderr vào (không dùng `2>&1`).

## Đọc kết quả

- `ready` là `true`: sang mục "Báo thầy cô".
- `error` khác `null`, hoặc `ready` là `false`: đọc `error.fix` và trường `fix` của các mục bắt buộc chưa đạt trong `checks`. Việc nào AI làm được ngay trong thư mục bộ công cụ thì làm, rồi chạy lại đúng lệnh ở mục "Cài đặt", tối đa một lần. Vẫn chưa sẵn sàng thì báo bị chặn.
- Lệnh cài không chạy được vì PowerShell bị chặn chạy script (thông báo có "running scripts is disabled" hoặc "execution policy"): nếu máy đã có Python 3.10 trở lên, chạy các lệnh tay trong mục "PowerShell bị chặn trên máy trường hoặc công ty" của `docs/vi/xu-ly-loi.md`, rồi kiểm tra bằng `python tools/vi/doctor.py --json`. Chưa có Python thì báo bị chặn.
- `warnings` có nội dung: chép lại cho thầy cô trong tin nhắn cuối.

## Báo thầy cô

- Sẵn sàng: nói ngắn gọn đã cài gì theo `installed` (`python` = Python 3.12, `venv` = môi trường Python riêng, `packages` = thư viện, `env` = file cấu hình; danh sách rỗng thì nói máy đã cài sẵn từ trước), kết quả mục "Xuất thử PPTX", các cảnh báo nếu có, và một câu lệnh mẫu như `Tạo bài giảng Toán 6 bài Phân số`. Nếu vừa tải bộ công cụ về, thêm: "Lần sau thầy cô mở thư mục <thư_mục_đích> trong Antigravity là dùng được ngay."
- Bị chặn: nêu bước lỗi và `error.message` bằng lời dễ hiểu, cách tự xử lý nếu có, rồi chép nguyên đoạn gửi bộ phận IT trong mục "Máy trường chặn cài đặt" của `docs/vi/xu-ly-loi.md`.
- Thầy cô đã gửi yêu cầu tạo slide trước đó: khi đã sẵn sàng thì làm tiếp yêu cầu đó.

## Công cụ tuỳ chọn

- Chỉ cài khi cần: thuyết minh và video cần FFmpeg; tài liệu định dạng cũ (`.doc`, `.odt`, `.rtf`…) cần Pandoc.
- Chạy (thay `ffmpeg` bằng `pandoc` khi cần Pandoc):

   ```
   powershell -NoProfile -ExecutionPolicy Bypass -File tools\vi\pptmaster.ps1 -Action tool -Name ffmpeg
   ```

- JSON có `found` là `true`: lấy `dir`, rồi chạy lệnh cần công cụ đó trong cùng một lệnh PowerShell, thêm thư mục vào PATH trước: `$env:Path = "<dir>;$env:Path"; <lệnh>`.
- `found` là `false`: báo thầy cô `error.message` và `error.fix`; làm tiếp những phần không cần công cụ đó.

## Không được làm

- Không chạy script tải thẳng từ mạng (`iex`, `Invoke-Expression`, `irm … | iex`).
- Không xin quyền quản trị, không chạy lệnh ở chế độ "Run as administrator".
- Không tắt phần mềm diệt virus hay tường lửa.
- Không đổi ExecutionPolicy vĩnh viễn (`Set-ExecutionPolicy`); chỉ dùng `-ExecutionPolicy Bypass` cho từng lệnh như trên.
- Không cài phần mềm nào khác ngoài Python 3.12, thư viện trong `requirements.txt`, Pandoc và FFmpeg; chỉ cài Git khi người dùng yêu cầu.
- Không tự nghĩ cách cài khác thay cho lệnh ở mục "Cài đặt" (ví dụ tự gõ `pip install`, tự tải Python từ nơi khác), trừ các lệnh tay được nêu ở mục "Đọc kết quả".
