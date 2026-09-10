# Nhật ký thay đổi — Bản Việt

## 6.3.2-vi.1 — 2026-09-10

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
