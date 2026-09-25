# Làm video giải thích

Video giải thích là một video ngắn kiểu viết tay: tiêu đề, khái niệm, công thức, các ý, sơ đồ, đồ thị và cả thí nghiệm ảo được viết dần ra trên nền giấy, khớp với giọng đọc tiếng Việt, có phụ đề. Video có hình minh hoạ vẽ dần từng nét, có thể có ảnh chụp thật; một bàn tay cầm bút đi theo nét đang vẽ, lau bảng khi sang cảnh mới, và máy quay phóng vào phần đang nói rồi thu về toàn cảnh. Khác "Làm video bài giảng" (dựng từ file slide có sẵn), video giải thích dựng thẳng từ **nội dung bài**, không cần làm slide trước.

| File | Dùng để |
|---|---|
| `video.mp4` | Video 1280×720, chiếu trên lớp hoặc đưa lên YouTube, Zalo |
| `phu-de.srt` | Phụ đề để riêng, chỉ có khi thầy cô chọn không in phụ đề lên hình |
| `video.md` | Kịch bản từng cảnh; sửa file này rồi dựng lại là video đổi theo |

## Cách yêu cầu

Nhắn cho AI, ví dụ: `Làm video giải thích bài Con lắc đơn bằng kiểu viết tay`, rồi dán nội dung hoặc dàn ý bài. AI hỏi một lượt ngắn (bài nào, học sinh cần hiểu gì, dài bao lâu, giọng nam hay nữ, có cảnh thí nghiệm ảo hay ảnh chụp thật không, phụ đề in lên hình hay để riêng, có giọng thu sẵn không), chia bài thành các cảnh để thầy cô duyệt, rồi dựng video trong `projects\_video\<tên_video>\`.

Nói "làm video" mà không rõ loại, AI sẽ hỏi lại: làm video từ bài giảng slide đã có, hay dựng video giải thích mới.

## Máy cần gì

- **Chromium** (trình duyệt dùng để vẽ cảnh): tải một lần, khoảng 150–300 MB. AI hỏi thầy cô trước khi tải.
- **FFmpeg** để ghép video: AI cài theo hướng dẫn nếu máy chưa có.
- **Mạng** khi dùng giọng máy (giọng nữ HoaiMy hoặc giọng nam NamMinh). Không có mạng, hoặc muốn dùng giọng của chính mình: thu từng cảnh thành file `giong\canh-1.mp3`, `giong\canh-2.mp3`… đặt trong thư mục video; cảnh có file thu sẵn thì không cần mạng.

## Thời gian dựng

Trước khi dựng thật, AI dựng thử mỗi cảnh một ảnh (thư mục `xem-truoc`) để soát chữ có tràn khung không. Video có ảnh chụp thật thì AI gửi thầy cô xem các ảnh đó trước (lệnh `--xem-truoc`) và chỉ dựng khi thầy cô đồng ý.

Dựng thật mất khoảng 1,5 lần thời lượng video trên máy 6 lõi: video 5 phút mất khoảng 7–8 phút; máy 2–3 lõi mất khoảng 11 phút; cộng thời gian tạo giọng. Máy càng nhiều lõi thì càng nhanh vì công cụ chụp khung bằng nhiều trình duyệt chạy song song. Trong lúc đó máy chạy nặng hơn bình thường.

## Hình và ảnh

- **Hình vẽ nét** lấy từ bộ biểu tượng có sẵn trong bộ công cụ, không cần mạng. AI tự chọn hình theo nội dung bài.
- **Ảnh chụp thật** do AI tải từ kho ảnh mở (Openverse, Wikimedia), không cần khoá. Mỗi ảnh luôn có dòng ghi tác giả và giấy phép ở góc. Thầy cô muốn dùng ảnh tự chụp thì gửi file cho AI và cho biết ai chụp.
- Muốn video bớt chuyển động: nhờ AI tắt bàn tay, tắt máy quay hoặc tắt lau bảng; mỗi thứ tắt riêng được.
- Chữ trên hình và phụ đề dùng font Itim (giấy phép mở SIL OFL 1.1) đi kèm bộ công cụ, đủ mọi chữ có dấu, máy nào cũng hiện giống nhau; không cần cài font.

## Sửa một cảnh

Mở `video.md`, sửa chữ hoặc lời của cảnh đó (hoặc nhờ AI sửa), rồi nhờ AI dựng lại. Chỉ cảnh bị sửa lời mới phải tạo giọng lại; các cảnh khác dùng lại giọng cũ nên dựng lại nhanh hơn. File giọng thầy cô thu sẵn không bao giờ bị ghi đè, kể cả khi thầy cô chép nó đè lên giọng máy cũ mà file `giong\canh-N.json` vẫn còn (xoá file `.json` đó cũng không sao).

## Những điều cần biết

- Bản này có một phong cách viết tay, khổ ngang 16:9. Chưa có khổ dọc cho TikTok, không chèn video tư liệu, không có nhạc nền.
- Ảnh tải về có thể sai nội dung: thầy cô xem kỹ ảnh AI gửi trước khi đồng ý dựng.
- Cảnh thí nghiệm ảo dùng 8 mô hình đã kiểm bằng số của phần [Làm thí nghiệm ảo](thi-nghiem-ao.md); số hiện trên hình luôn khớp công thức.
- AI không tự sửa số liệu, công thức hay lời giảng của thầy cô; khi chữ quá dài cho khung, AI chỉ rút gọn chữ trên hình hoặc tách cảnh.
- Kịch bản và video nằm trong `projects\` trên máy thầy cô, không được đưa lên GitHub.
- Gặp lỗi khi dựng: xem mục **Dựng video giải thích thất bại** trong [Xử lý lỗi](xu-ly-loi.md).
