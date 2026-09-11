# Bảo trì bản Việt

## Nguyên tắc

Kiến trúc của bản này là **upstream nguyên vẹn + lớp Việt hoá cộng thêm**: lớp Việt hoá chỉ được thêm file mới hoặc sửa một danh sách đóng các file upstream, không được đụng vào phần còn lại.

Danh sách file upstream được phép sửa và cách giữ lại khi đồng bộ upstream mới nằm trong [spec §4](2026-09-10-dong-goi-ban-viet-design.md#4-nguyên-tắc-kiến-trúc).

Không sửa file danh tính của skill: `skills/ppt-master/SKILL.md`, `LICENSE`, `SPONSORS*.md`.

Không đổi nội dung `CAP-NHAT.bat` sau khi đã phát hành. cmd.exe đọc file `.bat` theo vị trí byte trong lúc chạy, mà `git pull` bên trong có thể ghi đè chính file đó, nên phần còn lại có thể bị đọc lệch và chạy sai. Nếu buộc phải đổi, ghi rõ rủi ro này trong `CHANGELOG-VI.md` và dặn người dùng tải lại bộ công cụ thay vì bấm `CAP-NHAT.bat`.

## Chuẩn bị một lần

```
git remote add upstream https://github.com/hugohe3/ppt-master.git
git config merge.ours.driver true
```

Ngoài ra cần một virtualenv (venv) có đủ thư viện trong `requirements.txt` để chạy test và doctor.

## Đồng bộ phiên bản upstream mới

```
powershell -NoProfile -ExecutionPolicy Bypass -File tools\vi\sync_upstream.ps1 -Tag <tag> -Python <python>
```

Với các file index template (`decks_index.json`, `brands_index.json`...), script tự lấy bản upstream rồi dựng lại index nên không cần sửa tay.

Nếu script dừng và in `[DỪNG]`, đọc thông báo để biết đang ở loại nào:

1. **Dừng trước khi merge** — thiếu git hoặc Python, cây làm việc còn thay đổi chưa commit, không fetch được upstream hoặc không có tag, `git config` thất bại: chưa có gì thay đổi. Sửa nguyên nhân rồi chạy lại script.
2. **Dừng giữa merge** — còn xung đột ở file không tự giải được (script liệt kê các file cần xử lý tay), hoặc không commit được merge: sửa từng file, `git add` các file đã sửa, rồi `git commit` để hoàn tất merge. Muốn bỏ lần đồng bộ này thì chạy `git merge --abort`.
3. **Dừng sau khi đã commit merge** — `attribution_guard.py`, test lớp Việt hoá hoặc `doctor.py` thất bại: merge đã được commit nhưng **chưa push**. Cập nhật lớp Việt hoá cho khớp upstream mới rồi commit tiếp, hoặc hoàn tác merge bằng:
   ```
   git reset --keep ORIG_HEAD
   ```

## Chạy test

```
python -m unittest discover -s tools/vi/tests -v
python skills/ppt-master/scripts/attribution_guard.py
python tools/vi/doctor.py
```

## Bộ cài Python cố định

Chế độ `-Auto` của `tools/vi/pptmaster.ps1` cài Python bằng winget; khi máy không có winget hoặc winget lỗi, script dùng bộ cài python.org với phiên bản `3.12.10` (bản 3.12 cuối cùng có bộ cài cho Windows), đường dẫn và mã SHA256 ghi cứng trong biến `$PythonVersion` và `$PythonInstallers` cho `amd64` và `arm64`.

Khi đổi phiên bản:

1. Tải hai file cài từ `https://www.python.org/ftp/python/<phiên bản>/` và tính mã bằng `Get-FileHash -Algorithm SHA256`.
2. Tính thêm `Get-FileHash -Algorithm MD5` và đối chiếu với trang phát hành trên python.org trước khi ghi SHA256 vào script.
3. Cập nhật `$PythonVersion`, `$PythonInstallers`, mã gói winget `Python.Python.3.12` và thư mục `Python312` trong `Get-UserPythonPath` nếu đổi nhánh phiên bản (ví dụ lên 3.13).
4. Chạy lại test, rồi thử `-Action setup -Auto` trên một máy chưa có Python.

## Đánh số phiên bản

Số phiên bản có dạng `<tag upstream>-vi.<n>`, ví dụ `6.3.2-vi.1`.

- Chỉ đổi lớp Việt hoá (không đổi nền upstream): tăng `n` lên 1.
- Đồng bộ theo một tag upstream mới: đưa `n` về `1`.

## Phát hành

1. Cập nhật `CHANGELOG-VI.md` và dòng phiên bản trong `README.md`.
2. Đảm bảo toàn bộ test ở mục "Chạy test" đều đạt.
3. Gắn tag:
   ```
   git tag -a v<phiên bản> -m "..."
   ```
4. Đẩy lên remote:
   ```
   git push origin main
   git push origin v<phiên bản>
   ```
5. Tạo bản phát hành trên GitHub:
   ```
   gh release create v<phiên bản> --title "PPT Master bản Việt <phiên bản>" --notes "<ghi chú>"
   ```
