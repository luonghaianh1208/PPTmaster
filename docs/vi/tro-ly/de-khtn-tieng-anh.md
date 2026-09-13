# Loại việc: Soạn đề KHTN tiếng Anh

File dành cho AI. Luôn đọc docs/vi/tro-ly/quy-trinh-hoi.md trước file này, và đọc docs/vi/tro-ly/tieng-anh-khoa-hoc.md trước khi viết câu hỏi tiếng Anh.

## Khi nào dùng

Thầy cô cần đề kiểm tra môn khoa học tự nhiên viết bằng tiếng Anh, cho môn KHTN lớp 6–9 (chương trình tích hợp), hoặc Vật lí / Hoá học / Sinh học lớp 10–12 (chương trình tách môn).

Có hai trường hợp:
- Luồng A — thầy cô đã có đề tiếng Việt, cần chuyển sang tiếng Anh kèm bản song ngữ để tổ chuyên môn soát.
- Luồng B — thầy cô chưa có gì, cần AI soạn đề mới hoàn toàn bằng tiếng Anh.

Ví dụ câu lệnh:
- "Chuyển đề giữa kì Hoá 11 này sang tiếng Anh"
- "Soạn đề Vật lí 10 tiếng Anh 45 phút chương động lực học"

**Thầy cô đưa ảnh chụp đề thì không đọc được.** Xin bản PDF hoặc Word, không tự đoán nội dung đề.

Chuyển từ đề tiếng Việt thì không đổi số liệu, không đổi đáp án đúng, không đổi thứ tự câu. Câu gốc sai hoặc mơ hồ thì không tự sửa, ghi vào mục "Cần thầy cô soát".

## Câu hỏi bắt buộc

1. Môn và lớp (KHTN 6–9, hay Vật lí / Hoá học / Sinh học 10–12)?
   Gợi ý: lấy từ câu lệnh nếu đã có; chỉ hỏi phần còn thiếu.
2. Đề cho kỳ nào, làm trong bao nhiêu phút, tổng bao nhiêu điểm?
   Gợi ý: giữa kì hoặc cuối kì theo lịch nhà trường; 45 phút, thang điểm 10.
3. Nội dung thuộc chương hay chủ đề nào?
   Gợi ý: có đề tiếng Việt sẵn thì lấy theo đúng đề đó, không hỏi lại.
4. Số câu mỗi phần (Phần I trắc nghiệm, Phần II đúng/sai, Phần III trả lời ngắn)?
   Gợi ý: THPT dùng 18 – 4 – 6 theo đề tham khảo 2025. Môn KHTN ở THCS không có định dạng chung nên phải hỏi, không được lấy gợi ý làm mặc định.
5. Tỉ lệ mức độ biết – hiểu – vận dụng?
   Gợi ý: 4 – 4 – 2.
6. Đề dùng để làm gì (kiểm tra lớp song ngữ, đề luyện thêm, đề tham khảo cho tổ)?
   Gợi ý: kiểm tra lớp song ngữ thì dùng từ vựng đơn giản hơn đề tham khảo cho tổ chuyên môn. Câu này quyết định mức từ vựng tiếng Anh.
7. Có cần bản song ngữ để tổ soát không?
   Gợi ý: có, để tổ chuyên môn đối chiếu. Trả lời "không" thì chạy với `--phan de,dap-an`, không xuất `de-song-ngu.docx`; trường `vi:` trong `de.md` vẫn ghi để sau này xuất lại được.

Luồng A đã có đề sẵn thì câu 3, 4 và 5 đọc được từ chính đề, chỉ hỏi phần còn thiếu.

## Câu hỏi tuỳ chọn

- Có cần mã đề không? Chỉ hỏi khi thầy cô nhắc tới mã đề.
- Có cần in chỗ trống cho học sinh trình bày bài tự luận không? Chỉ hỏi khi thầy cô nhắc tới việc này.

## Tạo nhanh

1. Môn và lớp (câu hỏi bắt buộc 1).
2. Số câu mỗi phần (câu hỏi bắt buộc 4).

## Cấu trúc đề

Giữ cấu trúc đề Việt Nam đang dùng:
- Phần I — trắc nghiệm bốn lựa chọn (`A`–`D`).
- Phần II — đúng/sai bốn ý (`a`–`d`).
- Phần III — trả lời ngắn.

Thang điểm in trong file đáp án: Phần I mỗi câu 0,25 điểm; Phần II mỗi câu tối đa 1,0 điểm (chia theo số ý đúng); Phần III mỗi câu 0,25 điểm.

Ngữ pháp `de.md`: khối meta ở đầu file giữa hai dòng `---`; mỗi câu bắt đầu bằng `### <số>` và dùng các khoá `en:` (đề bài tiếng Anh, bắt buộc), `vi:` (bản tiếng Việt, không bắt buộc — thiếu thì chỉ là cảnh báo), `A:`–`D:` (bốn lựa chọn ở Phần I), `a:`–`d:` (bốn ý ở Phần II, mỗi dòng kết bằng ` | T` hoặc ` | F`), `key:` (đáp án đúng), `unit:` (đơn vị đáp số ở Phần III), `why:` (giải thích ngắn, chỉ in trong file đáp án), `level:` (`biet`, `hieu` hoặc `vandung`), `topic:` (chủ đề, dùng dựng ma trận).

Trong `en:` và `vi:` chỉ dùng ba dấu đánh dấu: `~ ~` (chỉ số dưới, ví dụ `H~2~O`), `^ ^` (chỉ số trên, ví dụ `m/s^2^`), `**` (in đậm, ví dụ `**not**`). Không dùng Markdown nào khác trong hai trường này.

Sau phần cuối, file có thể có mục `## CAN SOAT`: mỗi dòng bắt đầu bằng `- ` là một việc AI muốn thầy cô soát lại. Mục này in ở cuối `dap-an.docx` thành "Cần thầy cô soát"; không có mục này thì file đáp án ghi "Không có mục nào cần soát".

Đọc chi tiết đầy đủ và các ca biên ở docs/vi/phat-trien/2026-09-12-de-khtn-tieng-anh-design.md trước khi viết `de.md` cho đề dài hoặc đề có tình huống lạ.

## Đầu ra

`projects/_de-thi/<tên_đề>/` chứa:
- `de.md` — nguồn duy nhất, sinh ra cả ba file dưới.
- `de-en.docx` — đề tiếng Anh để in.
- `de-song-ngu.docx` — bản song ngữ Anh–Việt để tổ soát.
- `dap-an.docx` — đáp án, thang điểm, ma trận đặc tả và mục "Cần thầy cô soát".

## Ghi vào brief

- Loại việc: Soạn đề KHTN tiếng Anh.
- Thầy cô yêu cầu: ghi đúng lời thầy cô cho từng câu hỏi bắt buộc và tuỳ chọn.
- AI đề xuất (thầy cô đã đồng ý): các gợi ý thầy cô chấp nhận, theo đúng quy ước của docs/vi/tro-ly/mau-brief.md.
- Viết theo mẫu docs/vi/tro-ly/mau-brief.md.
