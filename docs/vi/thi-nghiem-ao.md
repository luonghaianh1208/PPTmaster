# Làm thí nghiệm ảo

Thí nghiệm ảo là một trang web nhỏ chạy ngay trên máy của thầy cô: học sinh đổi tham số, đo, ghi số liệu, vẽ đồ thị rồi rút ra kết luận. Bộ công cụ tạo cho mỗi thí nghiệm ba file:

| File | Dùng để |
|---|---|
| `thi-nghiem.html` | Mở bằng trình duyệt (Chrome, Edge) là chạy, **không cần mạng** |
| `phieu-hoc-tap.docx` | Phiếu in cho học sinh; trang cuối là đáp án dành cho giáo viên |
| `can-soat.md` | Công thức của mô hình, điều kiện lí tưởng hoá và những điểm thầy cô cần soát |

## Cách yêu cầu

Nhắn cho AI, ví dụ: `Làm thí nghiệm ảo con lắc đơn cho Vật lí 11, học sinh làm theo nhóm, có phiếu học tập`. AI hỏi một lượt ngắn (thí nghiệm nào, ai thao tác, học sinh cần rút ra kết luận gì, đo mấy lần, có bật sai số đo không), rồi tạo file trong `projects\_thi-nghiem\<tên_thí_nghiệm>\`.

## Các thí nghiệm có sẵn

| Môn | Thí nghiệm | Học sinh thay đổi | Học sinh đo |
|---|---|---|---|
| Vật lí | Ném xiên | vận tốc đầu, góc ném, độ cao, g | tầm xa, độ cao cực đại, thời gian bay |
| Vật lí | Con lắc đơn | chiều dài dây, g, khối lượng | chu kì, thời gian 10 dao động |
| Vật lí | Đoạn mạch nối tiếp và song song | hiệu điện thế, hai điện trở, kiểu mắc | cường độ dòng điện, hiệu điện thế |
| Hoá học | Chuẩn độ acid – base | acid mạnh hay yếu, nồng độ, thể tích NaOH nhỏ vào, chất chỉ thị | pH |
| Hoá học | Cân bằng N₂O₄ ⇌ 2NO₂ | nhiệt độ, áp suất | phần mol NO₂, độ đậm màu nâu |
| Hoá học | Tốc độ phản ứng | nồng độ, nhiệt độ, chất xúc tác | thời gian phản ứng |
| Toán | Khảo sát hàm số | các hệ số, tiếp điểm | giá trị, hệ số góc tiếp tuyến, cực trị |
| Toán | Xác suất thực nghiệm | phép thử, số lần gieo | tần số, tần suất |

Tám thí nghiệm này đã được kiểm bằng số: mỗi mô hình được tính lại độc lập bằng một chương trình khác và so trên hàng trăm bộ số, cộng với kiểm định luật (bảo toàn cơ năng, pH = 7 tại điểm tương đương, tần suất tiến về xác suất…).

Thầy cô cần thí nghiệm khác thì AI vẫn viết được mô hình mới. Khi đó `can-soat.md` ghi rõ **mô hình do AI viết, chưa có người duyệt**: thầy cô soát công thức và tính lại ít nhất hai dòng của bảng số kiểm bằng máy tính cầm tay trước khi dùng trên lớp.

## Dùng trên lớp

- **Máy chiếu:** chép thư mục thí nghiệm sang USB, bấm đúp `thi-nghiem.html`. Không cần mạng, không cần cài gì.
- **Học sinh dùng điện thoại:** điện thoại không mở được file từ USB. Thầy cô đưa `thi-nghiem.html` lên một trang web tĩnh (ví dụ kéo thả vào Netlify) rồi gửi địa chỉ cho học sinh.
- **Ba bước trên trang:** *Dự đoán* → *Quan sát* → *Giải thích*. Học sinh phải chốt dự đoán thì mới làm được thí nghiệm; ghi đủ số lần đo thì mục Giải thích mới mở và lúc đó trang mới đối chiếu dự đoán với kết quả. Nút **Chế độ giáo viên** bỏ mọi khoá để thầy cô trình diễn tự do.
- **Ghi lần đo** đưa số liệu đang hiện vào bảng; đồ thị vẽ từ chính bảng đó, kèm đường thẳng khớp và hệ số góc. **Chép số liệu** rồi dán thẳng vào Excel.
- **Sai số đo:** bật thì mỗi lần đo lệch ngẫu nhiên một chút như đo thật, để học sinh tập lấy trung bình và xử lí số liệu.
- Trang không lưu điểm và không gửi dữ liệu đi đâu.

## Những điều cần biết

- Thí nghiệm ảo là **mô hình lí tưởng**, không thay thí nghiệm thật. Cuối trang luôn ghi điều kiện lí tưởng hoá (bỏ qua sức cản, góc lệch nhỏ, dung dịch loãng…); ngoài các điều kiện đó kết quả thật sẽ khác.
- Mỗi lần mở, trang **Tự kiểm** mô hình của mình và ghi kết quả ở cuối trang. Nếu hiện dải đỏ "Mô hình không qua tự kiểm" thì không dùng file đó để dạy; nhờ AI tạo lại.
- Máy có cài Node thì công cụ kiểm số ngay lúc tạo file; máy không có Node thì bước này bỏ qua và trang tự kiểm khi mở. Không cần cài Node để dùng.
- Muốn gắn thí nghiệm vào bài giảng PowerPoint, nhờ AI thêm một trang có liên kết mở `thi-nghiem.html`; nhớ chép file HTML đi cùng file PPTX.
