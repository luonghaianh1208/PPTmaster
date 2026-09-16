---
trigger: always_on
---

# PPT Master bản Việt — luật luôn áp dụng

Khi làm việc trong repo này, luôn đọc và áp dụng @../../AGENTS.md và @../../AGENTS.vi.md.

Antigravity không chép nội dung các file nhắc bằng `@` vào luật, nên các điểm bắt buộc dưới đây được ghi thẳng ở đây. Chi tiết nằm ở `AGENTS.vi.md` và `docs/vi/tro-ly/`.

## Hỏi thầy cô trước khi làm

Người dùng viết tiếng Việt, bối cảnh là trường học hoặc Đoàn, và yêu cầu thuộc một trong 8 loại việc ở bảng dưới: trước mọi lệnh của quy trình tạo bài, làm đúng ba bước.

1. Đọc `docs/vi/tro-ly/quy-trinh-hoi.md`, rồi đọc file hướng dẫn của loại việc.
2. Gửi thầy cô một tin nhắn hỏi theo hai file đó.
3. Dừng và chờ thầy cô trả lời. Trong lúc chờ không chạy `project_manager.py init`, không tra cứu, không tìm hay tạo ảnh, không viết SVG hay file Word.

- Lượt hỏi này không trái `SKILL.md`: nó chỉ tạo thêm tài liệu nguồn (brief) trước khi quy trình của upstream bắt đầu. Bước xác nhận của upstream vẫn giữ nguyên.
- Turbo Mode, Always Proceed hay việc người dùng cho phép chạy lệnh không cần duyệt không phải là yêu cầu tạo nhanh, không bỏ được lượt hỏi.
- Câu lệnh có "không cần hỏi lại" hoặc "không hỏi gì": không hỏi câu nào. Câu lệnh chỉ có "tạo nhanh" hoặc "làm nhanh": vẫn hỏi các câu còn thiếu trong mục "Tạo nhanh" của file loại việc. Cả hai trường hợp đều đọc mục "Tạo nhanh" của `docs/vi/tro-ly/quy-trinh-hoi.md` và ghi brief.
- Trước khi gửi tin nhắn hỏi, chạy kiểm tra máy theo `AGENTS.vi.md` mục 9 (`tools/vi/doctor.py --no-smoke --json`) và thêm dòng báo máy chưa cài xong nếu cần.
- Chữ "giáo án" một mình: hỏi đúng một câu "Thầy cô cần file Word kế hoạch bài dạy (giáo án 5512), hay slide trình chiếu cho bài này?" trước khi làm gì khác.
- Soạn đề KHTN tiếng Anh và soạn giáo án có tài liệu thầy cô gửi: đọc tài liệu đó trước, rồi mới hỏi (xem `AGENTS.vi.md` mục 12 và 13).

| Loại việc | Dấu hiệu nhận biết | File hướng dẫn |
|---|---|---|
| Bài giảng | "bài giảng", "giáo án", "tiết học", "bài dạy", tên môn kèm lớp | `docs/vi/tro-ly/bai-giang.md` |
| Báo cáo – tổng kết | "báo cáo", "sơ kết", "tổng kết", "thi đua", "hội nghị viên chức" | `docs/vi/tro-ly/bao-cao-tong-ket.md` |
| Hoạt động Đoàn – sự kiện | "Đoàn", "chi đoàn", "cuộc thi", "sự kiện", "lễ kỷ niệm", "trao giải" | `docs/vi/tro-ly/hoat-dong-doan.md` |
| Poster/ấn phẩm Zalo – Facebook | "poster", "ảnh đăng Zalo", "bài đăng Facebook", "story", "TikTok" | `docs/vi/tro-ly/poster-mang-xa-hoi.md` |
| Tập huấn/workshop | "tập huấn", "bồi dưỡng", "sinh hoạt chuyên môn", "workshop", "chia sẻ chuyên đề" | `docs/vi/tro-ly/tap-huan-workshop.md` |
| Video bài giảng | "làm video", "xuất video", "lồng tiếng", "video bài giảng" | `docs/vi/tro-ly/video-bai-giang.md` |
| Soạn đề KHTN tiếng Anh | "soạn đề", "đề kiểm tra", "đề tiếng Anh", "đề KHTN", "chuyển đề sang tiếng Anh" | `docs/vi/tro-ly/de-khtn-tieng-anh.md` |
| Soạn giáo án tích hợp năng lực số và AI | "kế hoạch bài dạy", "KHBD", "giáo án Word", "giáo án 5512" | `docs/vi/tro-ly/giao-an.md` |

## Ảnh minh hoạ

Với bài mới thuộc 5 loại việc tạo PPTX ở bảng trên (không áp dụng làm đẹp hay sửa PPTX có sẵn), kể cả khi tạo nhanh: không làm bài toàn chữ, nhưng cũng không chèn ảnh trang trí cho đủ số.

- Sự vật, chất, hiện tượng, dụng cụ thí nghiệm, địa danh, nhân vật có thật: lên kế hoạch nguồn `web` ngay từ đầu, tìm bằng `skills/ppt-master/scripts/image_search.py` (Openverse, Wikimedia, không cần khoá). Quá trình, cấu tạo, sơ đồ thí nghiệm: vẽ sơ đồ trên slide.
- Thiếu khoá API tạo ảnh không phải lý do bỏ ảnh `web`; không mở file cấu hình để dò khoá.
- Chi tiết: mục "Ảnh minh hoạ" của `docs/vi/tro-ly/quy-trinh-hoi.md`.

## Các việc khác

Chạy lệnh trên Windows, kiểm tra môi trường trước lệnh Python đầu tiên, làm video, soạn đề, soạn giáo án: đọc `AGENTS.vi.md` trước khi làm.
