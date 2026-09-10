# Bảo trì bản Việt

## Nguyên tắc

Kiến trúc của bản này là **upstream nguyên vẹn + lớp Việt hoá cộng thêm**: lớp Việt hoá chỉ được thêm file mới hoặc sửa một danh sách đóng các file upstream, không được đụng vào phần còn lại.

Danh sách file upstream được phép sửa và cách giữ lại khi đồng bộ upstream mới nằm trong [spec §4](2026-09-10-dong-goi-ban-viet-design.md#4-nguyên-tắc-kiến-trúc).

Không sửa file danh tính của skill: `skills/ppt-master/SKILL.md`, `LICENSE`, `SPONSORS*.md`.

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

Nếu script dừng và in `[DỪNG]`, đó là một xung đột cần xử lý tay: sửa file xung đột, `git add` các file đã sửa, rồi `git commit` để hoàn tất merge. Với các file index template (`decks_index.json`, `brands_index.json`...), script tự lấy bản upstream rồi dựng lại index nên không cần sửa tay.

## Chạy test

```
python -m unittest discover -s tools/vi/tests -v
python skills/ppt-master/scripts/attribution_guard.py
python tools/vi/doctor.py
```

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
