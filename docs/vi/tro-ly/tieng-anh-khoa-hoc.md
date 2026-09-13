# Viết tiếng Anh cho đề khoa học tự nhiên

File dành cho AI. Đọc cùng docs/vi/tro-ly/de-khtn-tieng-anh.md.

Chín nguyên tắc dưới đây áp dụng cho mọi câu hỏi tiếng Anh trong `de.md`, dù là chuyển từ đề tiếng Việt hay soạn mới trực tiếp bằng tiếng Anh.

## 1. Dịch theo khái niệm, không dịch từng chữ

Đừng ghép từng từ tiếng Việt sang tiếng Anh. Dịch đúng khái niệm khoa học, kể cả khi câu tiếng Anh dùng cấu trúc khác hẳn câu tiếng Việt.

- Sai: `uniformly variable rectilinear motion` (dịch từng chữ "chuyển động thẳng biến đổi đều").
- Đúng: `uniformly accelerated motion`.

## 2. Thuật ngữ theo môn

Dùng đúng thuật ngữ chuẩn của từng môn, không dùng một từ tiếng Anh cho nhiều khái niệm khác nhau. Bảng dưới nêu các cặp bắt buộc dùng đúng:

| Môn | Khái niệm tiếng Việt | Thuật ngữ tiếng Anh |
|---|---|---|
| Vật lí | lực ma sát trượt | `kinetic friction` |
| Vật lí | suất điện động | `electromotive force (emf)` |
| Vật lí | công của lực | `work done by a force` |
| Vật lí | tốc độ | `speed` |
| Hoá học | khối lượng mol | `molar mass` |
| Hoá học | hiệu suất phản ứng | `percentage yield` |
| Hoá học | nồng độ mol | `molar concentration` |
| Hoá học | khối lượng riêng | `density` |
| Sinh học | hô hấp tế bào | `cellular respiration` |
| Sinh học | trao đổi chất | `metabolism` |
| Sinh học | cơ thể | `organism` |
| KHTN 6–9 | tốc độ | `speed` |
| KHTN 6–9 | khối lượng riêng | `density` |
| KHTN 6–9 | hô hấp tế bào | `cellular respiration` |

Chỉ dùng thuật ngữ có trong bảng này hoặc thuật ngữ tiếng Anh chuẩn của sách giáo khoa quốc tế mà AI chắc chắn đúng. Không chắc chắn thì ghi vào mục `## CAN SOAT` thay vì đoán.

## 3. Động từ lệnh hỏi và khuôn câu

Dùng đúng động từ lệnh hỏi theo mức độ: `State`, `Describe`, `Explain`, `Calculate`, `Determine`, `Deduce`, `Compare`, `Predict`.

Khuôn trắc nghiệm chuẩn: `Which of the following statements is correct?`

Dạng phủ định phải in đậm chữ `not` bằng `**not**` để học sinh không bỏ sót, ví dụ: `Which of the following is **not** a property of...?`

## 4. Chính tả IUPAC / Anh-Anh

Dùng chính tả IUPAC và tiếng Anh-Anh cho tên hợp chất: `sulfuric acid`, `sulfate`, `aluminium`. Ký hiệu nguyên tố và phương trình hoá học giữ nguyên, không dịch hay đổi định dạng.

## 5. Số và đơn vị

- Dấu thập phân: `25,5` (tiếng Việt) viết thành `25.5` (tiếng Anh).
- Dấu phân cách hàng nghìn: `1.000.000` viết thành `1 000 000`.
- Luôn có khoảng trắng giữa số và đơn vị: `5 kg`, `25 °C`.
- "Ở đktc" viết rõ thành `at 0 °C and 1 atm`, **không** dịch thành `at STP`.

## 6. Ngữ cảnh Việt Nam

Giữ nguyên ngữ cảnh Việt Nam trong đề, kèm một cụm giải thích ngắn cho học sinh nước ngoài đọc cũng hiểu: `terraced fields`, `fish sauce (a traditional Vietnamese condiment)`. Không thay ngữ cảnh Việt Nam bằng ngữ cảnh nước ngoài để "cho dễ".

## 7. Độ khó nằm ở khoa học, không ở tiếng Anh

Câu dẫn ngắn, chỉ một mệnh đề chính. Bốn đáp án trắc nghiệm phải cùng dạng ngữ pháp và xấp xỉ bằng nhau về độ dài, để học sinh không đoán được đáp án đúng nhờ câu chữ dài ngắn khác thường.

## 8. Chuyển từ đề tiếng Việt

Chuyển từ đề tiếng Việt thì không đổi nội dung khoa học của câu gốc. Câu gốc có chơi chữ tiếng Việt (đồng âm, nói lái) mà tiếng Anh không dịch được thì nói rõ với thầy cô là không chuyển được, chờ thầy cô quyết thay vì tự bỏ hay tự viết lại.

## 9. Tự soát trước khi xuất

Trước khi ghi `de.md`, tự soát lại toàn bộ thuật ngữ và câu chữ vừa viết. Chỗ nào chưa chắc chắn (thuật ngữ hiếm gặp, câu gốc mơ hồ, ngữ cảnh vừa thêm giải thích) thì ghi thành một dòng `- ` trong mục `## CAN SOAT` của `de.md`, không im lặng bỏ qua. Mục đó in ra thành "Cần thầy cô soát" ở cuối file đáp án.
