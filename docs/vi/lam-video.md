# Làm video bài giảng

Bộ công cụ biến một bài giảng đã làm thành video có lời giảng tiếng Việt và phụ đề.

## Cần gì trước

- Một bài giảng đã tạo trong `projects\`.
- Lời giảng cho từng slide. Chưa có thì cứ nhắn AI viết giúp, rồi thầy cô sửa lại cho đúng ý.

## Cách làm

Nhắn cho AI: `Làm video bài giảng này`. AI hỏi bốn câu ngắn (giọng đọc, tốc độ, phụ đề, độ phân giải) rồi làm.

Có hai cách dựng. AI dùng PowerPoint khi máy có PowerPoint **và** bài giảng đã có bản PPTX gắn tiếng (AI xuất bản này trước khi dựng video); cách đó giữ được hiệu ứng chuyển cảnh, nhưng **cửa sổ PowerPoint sẽ hiện lên và chiếm máy vài phút**. Còn lại AI ghép bằng FFmpeg, lần đầu phải tải Chromium khoảng 150–300 MB để chụp ảnh từng slide.

## Độ phân giải

Đường FFmpeg ghép từ ảnh chụp slide, nên độ phân giải cao nhất bằng khổ slide: 1280×720 với slide 16:9. Chọn 1080 ở đường này không làm chữ nét hơn, chỉ nặng thêm — bộ công cụ giữ đúng 720 dòng. Đường PowerPoint xuất đúng 1080 dòng.

## Thời gian và dung lượng

- Tạo tiếng đọc: khoảng một phút cho mỗi 10 slide.
- Dựng video: vài phút, tuỳ độ dài bài và tốc độ máy.
- Video 10 phút nặng khoảng 100–300 MB. Ổ đĩa nên còn trống ít nhất 2 GB.

## Phụ đề

- **File riêng** (nên dùng): được file `.srt` cạnh video. YouTube nhận file này, học sinh bật tắt được.
- **In lên hình**: chữ nằm sẵn trong video, gửi Zalo hay chiếu ngoại tuyến đều thấy, nhưng không tắt được.

## Trước khi giao cho học sinh

Nghe lại vài đoạn. Máy đọc tên riêng nước ngoài và công thức đôi khi chưa đúng; sửa lại lời giảng trong ghi chú rồi nhắn AI làm lại là được.

Gặp lỗi, xem [Xử lý lỗi](xu-ly-loi.md).
