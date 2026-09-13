# Soạn giáo án tích hợp năng lực số và năng lực AI

Bộ công cụ soạn kế hoạch bài dạy (giáo án) theo Công văn 5512, có tích hợp năng lực số và năng lực AI, từ giáo án cũ có sẵn hoặc từ tên bài mới hoàn toàn.

## Làm được gì

- **Có giáo án cũ rồi:** nâng cấp thêm phần năng lực số, năng lực AI và rubric đánh giá, giữ nguyên nội dung chuyên môn thầy cô đã viết.
- **Chưa có gì:** AI soạn giáo án mới hoàn toàn theo môn, lớp, tên bài và số tiết thầy cô yêu cầu.
- Cả hai trường hợp đều ra file Word kế hoạch bài dạy đủ mục tiêu, tiến trình dạy học, phiếu học tập và rubric đánh giá năng lực số, năng lực AI.

## Cách nhắn cho AI

Nâng cấp giáo án có sẵn:

```
Nâng cấp giáo án Bài 5 Ammonia này thành kế hoạch bài dạy có tích hợp năng lực số và năng lực AI
```

Soạn giáo án mới hoàn toàn:

```
Soạn kế hoạch bài dạy Hoá 11 Bài 5 Ammonia, 2 tiết, có tích hợp năng lực số và năng lực AI
```

Câu lệnh có chữ "giáo án" là ca mơ hồ: AI sẽ hỏi lại thầy cô cần file Word kế hoạch bài dạy hay **slide** trình chiếu cho bài này. Trả lời "Word" hoặc "kế hoạch bài dạy" để đi theo hướng dẫn này.

## Thầy cô cần đưa gì

- Giáo án cũ của bài này, dạng Word hoặc PDF, nếu muốn nâng cấp thay vì soạn mới.
- File kế hoạch dạy học hoặc phân phối chương trình, nếu có — AI chỉ đọc để lấy tuần, tiết thứ và mã năng lực số đã khai, **không sửa** file này.
- **Ảnh chụp giáo án thì AI không đọc được.** Gửi bản PDF hoặc Word, không chụp ảnh màn hình hay chụp giấy.

## File nhận được

AI lưu kết quả tại `projects\_giao-an\<tên_bài>\`:

| File | Nội dung |
|---|---|
| `giao-an.md` | Nguồn của giáo án — sửa file này rồi xuất lại được, không cần soạn lại từ đầu |
| `giao-an.docx` | Kế hoạch bài dạy để nộp cho trường |
| `can-soat.md` | Ghi chú nội bộ (những chỗ cần thầy cô kiểm lại), không nộp cho trường |

## Máy sẽ chặn những gì

Trước khi xuất file Word, máy tự kiểm cấu trúc giáo án để tránh bị trả lại:

- Hoạt động trong tiến trình thiếu một trong bốn thành tố bắt buộc (mục tiêu, nội dung, sản phẩm, các bước tổ chức chuyển giao – thực hiện – báo cáo – kết luận).
- Mục tiêu có nêu năng lực số hoặc năng lực AI nhưng tiến trình không có hoạt động nào dùng tới.
- Mã năng lực số hoặc mã năng lực AI không đúng với bảng mã của Bộ hoặc của tổ chuyên môn.
- Rubric thiếu tiêu chí, hoặc một tiêu chí không đủ ba mức đánh giá.

Gặp lỗi này, AI báo lại đúng theo lỗi: lỗi cấu trúc thì nêu đúng dòng cần sửa trong `giao-an.md`, lỗi mã năng lực thì nêu mã sai và các mã dùng được — không tự đoán rồi xuất liều.

## Môn chưa có khung mã AI

Một số môn chưa có bảng mã năng lực AI chính thức. Với môn đó, ô mã năng lực AI trong giáo án để trống — ghi rõ là chưa có khung mã cho môn này — thay vì một mã do AI tự đặt ra. AI không bao giờ tự bịa mã năng lực; việc ban hành mã là của Bộ và của tổ chuyên môn.

## Sửa lại rồi xuất lại

Muốn sửa một mục tiêu, một hoạt động hay một tiêu chí rubric, mở `giao-an.md`, sửa trực tiếp rồi nhờ AI chạy lại lệnh xuất. Không phải soạn lại giáo án từ đầu.

## Giới hạn

- Máy chỉ kiểm cấu trúc giáo án (đủ mục, đúng mã, đủ mức rubric), không kiểm được chất lượng sư phạm hay nội dung chuyên môn có chính xác hay không.
- Chưa nạp được khung ký duyệt và quốc hiệu theo mẫu riêng của từng trường.
- Chưa sinh được slide trình chiếu từ giáo án; cần slide cho bài này thì nhắn riêng cho AI.
