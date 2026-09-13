# Bắt đầu nhanh

## Để AI tự cài

Dùng Antigravity và máy chưa cài gì: không cần tải ZIP hay bấm `CAI-DAT.bat`.

1. Mở Antigravity → **Open Folder** → chọn một thư mục trống, ví dụ `D:\PPTmaster` (tránh thư mục đang đồng bộ OneDrive).
2. Mở khung **Agent**, dán câu lệnh:

   ```
   Cài PPT Master từ https://github.com/luonghaianh1208/PPTmaster vào thư mục này rồi báo khi sẵn sàng tạo slide
   ```

3. Chờ khoảng 5–10 phút. Khi Antigravity hỏi cho phép chạy lệnh, bấm đồng ý. Nếu AI xin chạy với quyền quản trị hoặc tắt phần mềm diệt virus, bấm từ chối.
4. AI báo "sẵn sàng" kèm một câu lệnh mẫu là dùng được. Lần sau mở thư mục AI báo trong tin nhắn sẵn sàng.

Đã tải ZIP và mở thư mục rồi: nhắn "cài đặt giúp em", hoặc gửi luôn yêu cầu tạo slide. AI tự kiểm tra và cài những gì còn thiếu trước khi làm.

Máy trường chặn cài đặt: AI sẽ báo lý do và đưa một đoạn để gửi bộ phận IT (xem [Xử lý lỗi](xu-ly-loi.md#máy-trường-chặn-cài-đặt)).

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

Với 8 loại việc trên, muốn AI làm luôn, thêm chữ "tạo nhanh" vào câu lệnh: AI chỉ hỏi 2–3 câu thật cần thiết. Không muốn trả lời câu nào, thêm "không cần hỏi lại": AI không hỏi gì, tự đề xuất các phần còn thiếu rồi làm luôn.

## Làm video bài giảng

Có bài giảng rồi, nhắn `Làm video bài giảng này` là AI đọc lời giảng, ghép thành video kèm phụ đề. Chi tiết trong [Làm video bài giảng](lam-video.md).

## Soạn đề tiếng Anh

Cần đề kiểm tra KHTN, Vật lí, Hoá học hay Sinh học bằng tiếng Anh, nhắn `Soạn đề Vật lí 10 tiếng Anh, 45 phút, chương động lực học` (chưa có đề) hoặc `Chuyển đề giữa kì Hoá 11 ở file này sang tiếng Anh` (đã có đề tiếng Việt). Chi tiết trong [Soạn đề tiếng Anh](soan-de-tieng-anh.md).

## Soạn giáo án

Cần kế hoạch bài dạy (giáo án 5512) có tích hợp năng lực số và năng lực AI, nhắn `Soạn kế hoạch bài dạy Hoá 11 Bài 5 Ammonia, 2 tiết, có tích hợp năng lực số và năng lực AI` (soạn mới) hoặc `Nâng cấp giáo án Bài 5 Ammonia này thành kế hoạch bài dạy có tích hợp năng lực số và năng lực AI` (đã có giáo án cũ). Nói "giáo án" thì AI hỏi lại Word hay slide trước khi làm. Chi tiết trong [Soạn giáo án](soan-giao-an.md).

## Lấy file kết quả

File PPTX nằm trong `projects\<tên dự án>\exports\` — AI sẽ báo đường dẫn cụ thể sau khi xuất xong. Mở file bằng PowerPoint để xem và chỉnh sửa tiếp.

## Mẹo

- Muốn sửa một slide cụ thể, cứ nhắn thẳng, ví dụ "sửa slide 3 cho ít chữ hơn".
- Xem thêm câu lệnh mẫu tại [Câu lệnh mẫu](cau-lenh-mau.md).
- Gặp lỗi trong quá trình dùng, xem [Xử lý lỗi](xu-ly-loi.md).
