# Loại việc: Video bài giảng

File dành cho AI. Luôn đọc `docs/vi/tro-ly/quy-trinh-hoi.md` trước file này.

## Khi nào dùng

Thầy cô muốn biến một bài giảng đã có trong bộ công cụ thành video có lời giảng, để đăng LMS, YouTube hoặc gửi học sinh tự học.

Ví dụ câu lệnh:
- "Làm video bài giảng này"
- "Lồng tiếng rồi xuất video bài Phân số"

## Câu hỏi bắt buộc

1. Dùng giọng đọc nào: giọng nữ hay giọng nam?
   Gợi ý: giọng nữ `vi-VN-HoaiMyNeural`.
2. Tốc độ đọc thế nào: chậm cho lớp nhỏ, vừa, hay nhanh?
   Gợi ý: vừa, giữ tốc độ mặc định.
3. Phụ đề để thành file riêng hay in thẳng lên hình?
   Gợi ý: file riêng, vì YouTube nhận file phụ đề và học sinh bật tắt được.
4. Độ phân giải video: 1080 cho máy chiếu và YouTube, hay 720 cho file nhẹ?
   Gợi ý: 1080; cách ghép ảnh slide chỉ cho tối đa 1280×720 nên lúc đó video giữ 720 dòng.

## Câu hỏi tuỳ chọn

- Thầy cô muốn giữ hiệu ứng chuyển cảnh của slide không? Chỉ hỏi khi máy có PowerPoint và bài giảng có hiệu ứng.

## Tạo nhanh

1. Giọng đọc (câu hỏi bắt buộc 1).
2. Phụ đề (câu hỏi bắt buộc 3).

## Cấu trúc gợi ý

- Giữ nguyên thứ tự slide của bài giảng; mỗi slide một đoạn lời giảng trong `notes/`.
- Chưa có ghi chú thì viết ghi chú trước, mỗi slide 3–6 câu, đọc khoảng 30–60 giây.

## Phong cách gợi ý

- Lời giảng viết như nói: câu ngắn, xưng hô thống nhất, đọc số và công thức thành lời.
- Tên riêng tiếng Anh nên viết lại theo cách đọc để máy đọc đúng.

## Khổ slide

- Mặc định `ppt169`. Video xuất ra 1080 hoặc 720 dòng, 30 hình mỗi giây. Cách ghép ảnh slide cao nhất chỉ 1280×720 vì ảnh chụp đúng khổ slide.

## Ghi vào brief

- Loại việc: Video bài giảng.
- Thầy cô yêu cầu: giọng đọc, tốc độ, phụ đề, độ phân giải, cách dựng nếu thầy cô có ý riêng.
- AI đề xuất (thầy cô đã đồng ý): các gợi ý thầy cô chấp nhận.
- Viết theo mẫu `docs/vi/tro-ly/mau-brief.md`.
