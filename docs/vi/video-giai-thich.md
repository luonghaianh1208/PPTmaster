# Làm video giải thích

Video giải thích là một video ngắn kiểu viết tay: tiêu đề, khái niệm, công thức, các ý, sơ đồ, đồ thị và cả thí nghiệm ảo được viết dần ra trên nền giấy, khớp với giọng đọc tiếng Việt, có phụ đề. Khác "Làm video bài giảng" (dựng từ file slide có sẵn), video giải thích dựng thẳng từ **nội dung bài**, không cần làm slide trước.

| File | Dùng để |
|---|---|
| `video.mp4` | Video 1280×720, chiếu trên lớp hoặc đưa lên YouTube, Zalo |
| `phu-de.srt` | Phụ đề để riêng, chỉ có khi thầy cô chọn không in phụ đề lên hình |
| `video.md` | Kịch bản từng cảnh; sửa file này rồi dựng lại là video đổi theo |

## Cách yêu cầu

Nhắn cho AI, ví dụ: `Làm video giải thích bài Con lắc đơn bằng kiểu viết tay`, rồi dán nội dung hoặc dàn ý bài. AI hỏi một lượt ngắn (bài nào, học sinh cần hiểu gì, dài bao lâu, giọng nam hay nữ, có cảnh thí nghiệm ảo không, phụ đề in lên hình hay để riêng, có giọng thu sẵn không), chia bài thành các cảnh để thầy cô duyệt, rồi dựng video trong `projects\_video\<tên_video>\`.

Nói "làm video" mà không rõ loại, AI sẽ hỏi lại: làm video từ bài giảng slide đã có, hay dựng video giải thích mới.

## Máy cần gì

- **Chromium** (trình duyệt dùng để vẽ cảnh): tải một lần, khoảng 150–300 MB. AI hỏi thầy cô trước khi tải.
- **FFmpeg** để ghép video: AI cài theo hướng dẫn nếu máy chưa có.
- **Mạng** khi dùng giọng máy (giọng nữ HoaiMy hoặc giọng nam NamMinh). Không có mạng, hoặc muốn dùng giọng của chính mình: thu từng cảnh thành file `giong\canh-1.mp3`, `giong\canh-2.mp3`… đặt trong thư mục video; cảnh có file thu sẵn thì không cần mạng.

## Thời gian dựng

Trước khi dựng thật, AI dựng thử mỗi cảnh một ảnh (thư mục `xem-truoc`) để soát chữ có tràn khung không. Dựng thật mất vài phút: video dài 5 phút cần khoảng 4–5 phút, cộng thời gian tạo giọng. Trong lúc đó thầy cô vẫn dùng máy bình thường.

## Sửa một cảnh

Mở `video.md`, sửa chữ hoặc lời của cảnh đó (hoặc nhờ AI sửa), rồi nhờ AI dựng lại. Chỉ cảnh bị sửa lời mới phải tạo giọng lại; các cảnh khác dùng lại giọng cũ nên dựng lại nhanh hơn. File giọng thầy cô thu sẵn không bao giờ bị ghi đè.

## Những điều cần biết

- Bản này có một phong cách viết tay, khổ ngang 16:9. Chưa có khổ dọc cho TikTok, không chèn ảnh hay video tư liệu, không có nhạc nền.
- Cảnh thí nghiệm ảo dùng 8 mô hình đã kiểm bằng số của phần [Làm thí nghiệm ảo](thi-nghiem-ao.md); số hiện trên hình luôn khớp công thức.
- AI không tự sửa số liệu, công thức hay lời giảng của thầy cô; khi chữ quá dài cho khung, AI chỉ rút gọn chữ trên hình hoặc tách cảnh.
- Kịch bản và video nằm trong `projects\` trên máy thầy cô, không được đưa lên GitHub.
- Gặp lỗi khi dựng: xem mục **Dựng video giải thích thất bại** trong [Xử lý lỗi](xu-ly-loi.md).
