# Soạn đề KHTN bằng tiếng Anh

Bộ công cụ soạn đề kiểm tra môn khoa học tự nhiên (KHTN, Vật lí, Hoá học, Sinh học) bằng tiếng Anh, từ đề tiếng Việt có sẵn hoặc từ con số không.

## Làm được gì

- **Có đề tiếng Việt rồi:** chuyển sang tiếng Anh, giữ nguyên số liệu, đáp án và thứ tự câu.
- **Chưa có gì:** AI soạn đề mới hoàn toàn bằng tiếng Anh theo môn, lớp, chương và số câu thầy cô yêu cầu.
- Cả hai trường hợp đều ra ba file Word: đề tiếng Anh để in cho học sinh, đề song ngữ Anh–Việt để tổ chuyên môn soát, và đáp án kèm ma trận đặc tả.

## Cách nhắn cho AI

Có đề sẵn:

```
Chuyển đề giữa kì Hoá 11 ở file này sang tiếng Anh
```

Chưa có đề:

```
Soạn đề Vật lí 10 tiếng Anh, 45 phút, chương động lực học
```

## AI sẽ hỏi gì

AI hỏi tối đa 7 câu ngắn trước khi soạn, mỗi câu có sẵn gợi ý để thầy cô chọn nhanh. Ba câu quan trọng nhất:

- Môn và lớp.
- Số câu mỗi phần (trắc nghiệm, đúng/sai, trả lời ngắn).
- Đề dùng để làm gì (kiểm tra lớp song ngữ, đề luyện thêm, hay đề tham khảo cho tổ) — câu này quyết định mức từ vựng tiếng Anh.

Các câu còn lại hỏi về kỳ kiểm tra, thời gian, tổng điểm, chương/chủ đề, tỉ lệ mức độ biết – hiểu – vận dụng, và có cần bản song ngữ để tổ soát không.

## Thầy cô cần đưa gì

Một file Word hoặc PDF của đề đã có (nếu có). **Một tấm ảnh chụp đề thì AI không đọc được** — thầy cô chụp lại thành file PDF, hoặc gõ lại đề thành file Word/PDF rồi gửi.

## File nhận được

AI lưu kết quả tại `projects\_de-thi\<tên_đề>\`:

| File | Nội dung |
|---|---|
| `de.md` | Nguồn của đề — sửa file này rồi xuất lại được, không cần soạn lại từ đầu |
| `de-en.docx` | Đề tiếng Anh, in cho học sinh |
| `de-song-ngu.docx` | Đề song ngữ Anh–Việt, để tổ chuyên môn soát |
| `dap-an.docx` | Đáp án, giải thích, thang điểm và ma trận đặc tả |

Trả lời không cần bản song ngữ thì AI chỉ xuất `de-en.docx` và `dap-an.docx`.

## Sửa lại rồi xuất lại

Muốn sửa một câu hỏi, mở `de.md`, sửa trực tiếp rồi nhờ AI chạy lại lệnh xuất. Không phải soạn lại đề từ đầu.

## Cần đọc lại trước khi in

Trước khi in hoặc gửi đề, đọc mục **"Cần thầy cô soát"** ở cuối file `dap-an.docx`. Mục này liệt kê những chỗ AI chưa chắc: thuật ngữ tiếng Anh, câu hỏi gốc mơ hồ hoặc có thể sai, và những đoạn AI phải thêm giải thích. Máy chỉ kiểm được cú pháp và số liệu, không kiểm được thuật ngữ có đúng chuyên môn hay không, nên thầy cô cần đọc kỹ phần này và kiểm lại thang điểm trước khi in.

## Giới hạn

- Chưa nạp được mẫu đầu đề riêng của trường (quốc hiệu, tên trường, khung điểm theo mẫu riêng) và chưa chèn được logo.
- Chưa trộn được nhiều mã đề (đảo thứ tự câu/đáp án) trong một lần chạy; tính năng này dự kiến làm ở gói riêng.
- Chưa có phần tự luận; đề dùng Phần III trả lời ngắn thay cho câu tự luận.
- Chưa xuất được PDF trực tiếp; mở file Word rồi lưu thành PDF nếu cần.
