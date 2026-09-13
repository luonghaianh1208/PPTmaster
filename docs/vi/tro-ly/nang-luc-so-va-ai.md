# Khung năng lực số và năng lực AI

File dành cho AI. Đọc cùng docs/vi/tro-ly/giao-an.md.

Khung do Lương Hải Anh — 2Anh AI Education biên soạn, dựa trên Bảng mã Năng lực số của Bộ GD&ĐT (mức nâng cao NC1 cho THPT) và Phụ lục III & IV về triển khai giáo dục AI. Đây là nguồn mã duy nhất của bộ công cụ: `tools/vi/giao_an.py` đọc mã trực tiếp từ file này, nên sửa mã ở đây là sửa cho cả AI và cho bộ kiểm tra.

## Nguyên tắc tích hợp

- AI là công cụ hỗ trợ tư duy, không thay việc học. Học sinh phải là người viết câu lệnh và là người đánh giá kết quả.
- Mỗi hoạt động dùng AI để dự đoán đều phải có bước đối chứng: thí nghiệm thật, thí nghiệm ảo, hoặc tra cứu nguồn tin cậy.
- Phải có ít nhất một hoạt động yêu cầu học sinh chỉ ra chỗ AI nói sai hoặc nói thiếu.
- Năng lực số và năng lực AI phải nằm trong tiến trình dạy học, không chỉ nằm ở mục tiêu.
- Không đặt mã chỉ báo mới. Môn chưa có khung thì ghi `(chưa có khung mã cho môn này)` ở mục tiêu và `(chưa có mã)` ở hoạt động.

## Khung năng lực số (mọi môn)

- `1.1.NC1a` — Tìm kiếm và chọn lọc dữ liệu, thông tin, nội dung số phục vụ học tập.
- `1.1.NC1b` — Đánh giá độ tin cậy và độ chính xác của nguồn dữ liệu số.
- `1.2.NC1a` — Phân tích, so sánh, đối chiếu các nguồn dữ liệu khoa học số.
- `1.3.NC1a` — Tổ chức, phân loại, lưu trữ và truy xuất thông tin, dữ liệu trong môi trường số.
- `2.1.NC1a` — Tương tác, trao đổi ý kiến qua nền tảng học tập trực tuyến.
- `2.2.NC1a` — Chia sẻ học liệu số và làm việc nhóm trực tuyến an toàn.
- `3.1.NC1a` — Dùng phần mềm chuyên dụng tạo sơ đồ tư duy, báo cáo, infographic, hình ảnh 3D.
- `3.2.NC1a` — Chỉnh sửa, tích hợp nội dung số đa phương tiện.
- `4.1.NC1a` — Bảo vệ thiết bị, dữ liệu cá nhân, tuân thủ bản quyền và an toàn số.
- `5.1.NC1a` — Dùng thiết bị số, phần mềm mô phỏng, thí nghiệm ảo để giải quyết vấn đề học tập và thực tiễn.
- `5.3.NC1a` — Đổi mới cách học bằng cách áp dụng giải pháp công nghệ số.

## Khung năng lực AI theo môn

### Hoá học — lớp 10 — khung `AI-H10`

- `AI-H10` — Dùng AI trực quan hoá phân tử 3D theo VSEPR và lai hoá AO; mô phỏng tốc độ phản ứng theo Arrhenius; giải thích vai trò của dữ liệu huấn luyện; nhận diện orbital p, d, f và dự đoán góc liên kết.

### Hoá học — lớp 11 — khung `AI-H11`

- `AI-H11.1` — Khai thác và nhận diện: tra cứu PubChem, NIST; phân tích phổ IR để nhận diện nhóm chức, phổ MS để xác định phân tử khối.
- `AI-H11.2` — Mô hình hoá và dự đoán: dự đoán chiều chuyển dịch cân bằng theo Le Chatelier, hướng cộng theo Markovnikov, phản ứng tách theo Zaitsev, vị trí thế trên nhân thơm.
- `AI-H11.3` — Thiết kế và tối ưu hoá: dùng prompt engineering gợi ý sơ đồ tổng hợp hữu cơ, tối ưu thông số Haber-Bosch và Contact process, tối ưu quy trình STEM.
- `AI-H11.4` — Đánh giá và phản biện: kiểm chứng kết quả AI với lý thuyết và thực nghiệm, nhận diện rủi ro môi trường và tuân thủ đạo đức AI.

### Hoá học — lớp 12 — khung `AI-H12`

- `AI-H12` — Dùng AI tính suất điện động pin Galvani, dự đoán sản phẩm điện phân, mô tả cấu trúc phức chất; dự đoán tính chất polymer và giải pháp tái chế; viết prompt tra cứu hằng số.

### Hoá học — hệ chuyên — khung `AI-H-Chuyên`

- `AI-H-Chuyên` — Phân tích phổ NMR, MS, IR phức tạp; mô phỏng cơ chế SN1, SN2, E1, E2, cộng ái nhân, cộng ái điện tử; dùng Python và machine learning xử lý dữ liệu động hoá học, xác định bậc phản ứng.

## Môn chưa có khung mã AI

Môn nào không có mục khung ở trên thì coi như chưa có khung. Khi đó:

- Mục `### Nang luc AI` trong giáo án có dòng đầu là đúng chuỗi `- (chưa có khung mã cho môn này)`, các dòng sau viết hoạt động AI bằng lời, không có mã.
- Hoạt động AI trong tiến trình ghi `ai: (chưa có mã)`.
- Tuyệt đối không tự đặt mã mới kiểu `AI-P11` hay `AI-S12`. Đặt mã chỉ báo là việc của Bộ và của tổ chuyên môn.
- Tổ nào đã có khung riêng thì thầy cô gửi file khung; lúc đó thêm một mục khung mới vào file này theo đúng khuôn ở trên.

## Gợi ý hoạt động AI cho môn chưa có khung

- Toán: dùng AI sinh phản ví dụ rồi học sinh kiểm chứng bằng lập luận; dùng AI giải rồi tìm chỗ sai trong lời giải.
- Ngữ văn: dùng AI viết một đoạn theo yêu cầu rồi học sinh nhận xét giọng điệu, dẫn chứng, chỗ bịa; so bản AI với bản của mình.
- Lịch sử, Địa lí: dùng AI tóm tắt một nguồn rồi học sinh đối chiếu với tư liệu gốc, chỉ ra chỗ thiếu bối cảnh.
- Vật lí, Sinh học: dùng AI dự đoán kết quả thí nghiệm rồi làm thí nghiệm thật hoặc mô phỏng để kiểm chứng.
- Tin học: dùng AI sinh mã rồi học sinh chạy, tìm lỗi và giải thích vì sao sai.
- Ngoại ngữ: dùng AI sửa bài viết rồi học sinh giải thích từng chỗ sửa, giữ lại chỗ không đồng ý.
