# Bắt đầu nhanh

## Mở thư mục trong AI editor

- **Claude Code:** nếu dùng extension trong VS Code, chọn **File → Open Folder** rồi mở khung Claude. Nếu dùng bản dòng lệnh, mở terminal tại thư mục này và gõ `claude`.
- **Cursor:** **File → Open Folder** rồi mở khung chat.
- **Antigravity:** **Open Folder** rồi mở khung Agent.

Bản Việt đã có sẵn quy tắc cho cả 3 công cụ này, không cần cấu hình thêm gì.

## Câu lệnh đầu tiên

Nhắn cho AI một câu như:

```
Tạo bài thuyết trình 8 slide giới thiệu câu lạc bộ tin học của trường, phong cách trẻ trung
```

Nếu có sẵn tài liệu, kéo thả file PDF/Word vào khung chat, hoặc ghi thẳng đường dẫn file trong câu lệnh.

## AI sẽ hỏi gì

Với bài giảng, báo cáo – tổng kết, hoạt động Đoàn – sự kiện, poster Zalo – Facebook và tập huấn/workshop, AI hỏi một lượt ngắn trước khi làm: tối đa 7 câu, mỗi câu có sẵn gợi ý. Thầy cô chỉ cần trả lời "đồng ý" hoặc sửa những ý chưa đúng.

- **Lần đầu**, AI hỏi thêm hồ sơ đơn vị: tên trường, cấp học, người trình bày, logo, màu chủ đạo. AI không điền sẵn tên trường hay họ tên, thầy cô tự ghi. Hồ sơ lưu ở `projects\_ho-so-don-vi.md` trên máy của thầy cô, lần sau AI dùng lại.
- **Sau đó**, AI tóm tắt đề xuất trong khung chat để thầy cô duyệt lần cuối, rồi mới dựng slide.
- **Việc khác** (ví dụ giới thiệu sản phẩm), AI xác nhận lại đối tượng người xem, số trang, phong cách và mẫu thiết kế như trước.

Với 5 loại việc trên, muốn AI làm luôn, thêm chữ "tạo nhanh" vào câu lệnh: AI chỉ hỏi 2–3 câu thật cần thiết. Không muốn trả lời câu nào, thêm "không cần hỏi lại": AI không hỏi gì, tự đề xuất các phần còn thiếu rồi làm luôn.

## Lấy file kết quả

File PPTX nằm trong `projects\<tên dự án>\exports\` — AI sẽ báo đường dẫn cụ thể sau khi xuất xong. Mở file bằng PowerPoint để xem và chỉnh sửa tiếp.

## Mẹo

- Muốn sửa một slide cụ thể, cứ nhắn thẳng, ví dụ "sửa slide 3 cho ít chữ hơn".
- Xem thêm câu lệnh mẫu tại [Câu lệnh mẫu](cau-lenh-mau.md).
- Gặp lỗi trong quá trình dùng, xem [Xử lý lỗi](xu-ly-loi.md).
