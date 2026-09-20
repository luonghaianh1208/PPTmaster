# Loại việc: Thí nghiệm ảo

File dành cho AI. Luôn đọc `docs/vi/tro-ly/quy-trinh-hoi.md` trước file này. Đầu ra của loại việc này là một file HTML thí nghiệm ảo chạy không cần mạng và một phiếu học tập Word, không phải PPTX.

## Khi nào dùng

Thầy cô cần một thí nghiệm ảo, mô phỏng tương tác cho Toán, Vật lí hoặc Hoá học để học sinh thay đổi tham số, đo, ghi số liệu và rút ra kết luận.

Ví dụ câu lệnh:
- "Làm thí nghiệm ảo con lắc đơn cho Vật lí 11"
- "Tạo mô phỏng chuẩn độ acid – base cho Hoá 11"
- "Làm thí nghiệm ảo gieo xúc xắc cho bài xác suất lớp 10"

Thầy cô chỉ cần hình minh hoạ một thí nghiệm trên slide thì không dùng loại việc này; đó là việc của bài giảng.

## Câu hỏi bắt buộc

1. Môn, lớp, tên bài, và thí nghiệm nào?
   Gợi ý: mẫu gần nhất trong `docs/vi/tro-ly/mo-hinh-thi-nghiem.md`. Thí nghiệm ngoài danh mục thì nói rõ: em sẽ viết mô hình mới, thầy cô cần soát công thức trước khi dùng.
2. Thí nghiệm dùng ở hoạt động nào: mở đầu, hình thành kiến thức, hay luyện tập?
   Gợi ý: hình thành kiến thức.
3. Ai thao tác: giáo viên trên máy chiếu, hay học sinh theo nhóm trên máy tính hoặc điện thoại?
   Gợi ý: giáo viên trình diễn, các nhóm ghi số liệu vào phiếu.
4. Học sinh cần rút ra kết luận gì từ thí nghiệm?
   Gợi ý: AI đề xuất một kết luận theo yêu cầu cần đạt của bài để thầy cô sửa.
5. Tham số nào được thay đổi, trong khoảng nào, và đo bao nhiêu lần?
   Gợi ý: một tham số chính theo khoảng mặc định của mẫu, 5 lần đo.
6. Có bật sai số đo để học sinh tập xử lí số liệu không?
   Gợi ý: tắt với THCS, bật với THPT.
7. Có cần phiếu học tập Word để in không?
   Gợi ý: có.

## Câu hỏi tuỳ chọn

- Thầy cô có địa chỉ web đã đưa file lên (ví dụ Netlify) để gắn vào slide không? Chỉ hỏi khi thầy cô nhắc tới việc cho học sinh dùng điện thoại.

## Tạo nhanh

1. Môn, lớp, tên bài và thí nghiệm (câu hỏi bắt buộc 1).
2. Ai thao tác (câu hỏi bắt buộc 3).

## Cấu trúc thi-nghiem.md

Viết một file `thi-nghiem.md` trong `projects/_thi-nghiem/<tên_thí_nghiệm>/`. File mở đầu bằng khối thông tin giữa hai dòng `---`, rồi đúng năm mục theo thứ tự: `## Tham số`, `## Dự đoán`, `## Quan sát`, `## Giải thích`, `## Kết luận`.

Khối thông tin:

| Khoá | Bắt buộc | Giá trị |
|---|---|---|
| `tieu-de` | có | tên thí nghiệm hiện trên trang |
| `mon`, `lop` | có | môn và lớp |
| `mau` | có | mã mẫu trong `mo-hinh-thi-nghiem.md`, hoặc `moi` khi AI viết mô hình mới |
| `nguoi-thao-tac` | không | `giao-vien` (mặc định, không khoá) hoặc `nhom` (khoá cho tới khi học sinh chốt dự đoán) |
| `sai-so` | không | `tat` (mặc định) hoặc `bat` |

Mục `## Tham số`, mỗi dòng một tham số của mẫu; tham số không ghi thì giữ cố định ở mặc định của mẫu:

- Thanh trượt: `<mã>: <nhỏ nhất>..<lớn nhất> buoc <bước> mac-dinh <giá trị>`
- Giữ cố định: `<mã>: co-dinh <giá trị>`
- Tham số lựa chọn: `<mã>: chon <mã_1>, <mã_2>` (mục đầu là mặc định) hoặc `<mã>: co-dinh <mã_lựa_chọn>`
- Khoảng không được vượt khoảng của mẫu, vì ngoài khoảng đó công thức của mẫu không còn đúng. Phải có ít nhất một tham số thay đổi được.

Mục `## Dự đoán`: dòng `cau:`; nếu là trắc nghiệm thì thêm `A:` đến `D:` (ít nhất hai) và `dap-an:`. Không có lựa chọn thì là câu mở, học sinh tự viết.

Mục `## Quan sát`:

- `so-lan-do:` số nguyên từ 3 đến 20.
- `do:` các cột của bảng số liệu, là mã tham số hoặc mã đại lượng đo, cách nhau bằng dấu phẩy, tối đa 6 cột, có ít nhất một đại lượng đo.
- `do-thi:` (tuỳ chọn) `<trục tung> theo <trục hoành>`. Mỗi trục là một mã có trong `do:`, tuỳ chọn kèm một phép biến đổi: `<mã>^2`, `1/<mã>`, `ln(<mã>)`, `sqrt(<mã>)`. Chọn phép biến đổi để đồ thị thành đường thẳng thì học sinh đọc được quy luật từ hệ số góc. Đây là cú pháp tính toán, khác quy ước hiển thị `T^2^` dùng trong các dòng chữ.

Mục `## Giải thích`: dòng `cau:` và dòng `goi-y-dap-an:`. Mục `## Kết luận`: một đoạn văn.

Quy ước chữ: `H~2~SO~4~` cho chỉ số dưới, `m/s^2^` cho chỉ số trên, `**in đậm**`. Không chèn địa chỉ web vào file: thí nghiệm chạy không cần mạng.

Ví dụ Vật lí (học sinh thao tác theo nhóm, có sai số đo):

```
---
tieu-de: Chu kì con lắc đơn
mon: Vật lí
lop: 11
mau: li-con-lac-don
nguoi-thao-tac: nhom
sai-so: bat
---

## Tham số
chieu-dai: 0.4..1.6 buoc 0.2 mac-dinh 1.0
g: co-dinh 9.8

## Dự đoán
cau: Khi tăng chiều dài dây gấp 4 lần, chu kì thay đổi thế nào?
A: Tăng 4 lần
B: Tăng 2 lần
C: Không đổi
dap-an: B

## Quan sát
so-lan-do: 5
do: chieu-dai, chu-ki
do-thi: chu-ki^2 theo chieu-dai

## Giải thích
cau: Từ đồ thị, T^2^ và l liên hệ thế nào? Suy ra công thức tính chu kì.
goi-y-dap-an: T^2^ tỉ lệ thuận với l, hệ số góc bằng 4π^2^/g; suy ra T = 2π√(l/g).

## Kết luận
Chu kì con lắc đơn chỉ phụ thuộc chiều dài dây và gia tốc trọng trường, không phụ thuộc khối lượng quả nặng.
```

Ví dụ Hoá học (giáo viên trình diễn, dự đoán câu mở):

```
---
tieu-de: Chuyển dịch cân bằng N~2~O~4~ ⇌ 2NO~2~
mon: Hoá học
lop: 11
mau: hoa-can-bang-no2
---

## Tham số
nhiet-do: 0..100 buoc 10 mac-dinh 20
ap-suat: co-dinh 1

## Dự đoán
cau: Ngâm bình khí vào nước nóng thì màu nâu đỏ của bình thay đổi thế nào? Vì sao?

## Quan sát
so-lan-do: 6
do: nhiet-do, phan-mol-no2, nong-do-no2

## Giải thích
cau: Phản ứng thuận thu nhiệt hay toả nhiệt? Dùng nguyên lí Le Chatelier để giải thích số liệu.
goi-y-dap-an: Tăng nhiệt độ thì phần mol NO~2~ tăng, màu đậm hơn; cân bằng chuyển dịch theo chiều thu nhiệt, vậy phản ứng thuận thu nhiệt.

## Kết luận
Khi tăng nhiệt độ, cân bằng chuyển dịch theo chiều phản ứng thu nhiệt; khi giảm nhiệt độ, cân bằng chuyển dịch theo chiều toả nhiệt.
```

Ví dụ Toán (tham số lựa chọn giữ cố định):

```
---
tieu-de: Tần suất và xác suất khi gieo hai xúc xắc
mon: Toán
lop: 10
mau: toan-xac-suat
nguoi-thao-tac: nhom
---

## Tham số
phep-thu: co-dinh hai-xuc-xac-tong-7
so-lan: 100..10000 buoc 100 mac-dinh 100
hat-giong: 1..999 buoc 1 mac-dinh 1

## Dự đoán
cau: Gieo hai xúc xắc càng nhiều lần thì tần suất xuất hiện tổng bằng 7 thay đổi thế nào?
A: Tăng dần tới 1
B: Dao động rồi ổn định quanh một số
C: Giảm dần về 0
dap-an: B

## Quan sát
so-lan-do: 6
do: so-lan, tan-so, tan-suat

## Giải thích
cau: Tần suất ổn định quanh giá trị nào? So sánh với xác suất tính bằng cách đếm số kết quả thuận lợi.
goi-y-dap-an: Có 6 trên 36 kết quả cho tổng bằng 7, nên xác suất là 1/6 ≈ 0,1667; tần suất tiến gần giá trị này khi số lần gieo lớn.

## Kết luận
Khi số lần thử đủ lớn, tần suất của biến cố xấp xỉ xác suất của biến cố đó.
```

## Đầu ra

Chạy `python tools\vi\thi_nghiem.py projects\_thi-nghiem\<tên_thí_nghiệm>`. Công cụ ghi vào chính thư mục đó:

- `thi-nghiem.html`: mở bằng trình duyệt là chạy, không cần mạng. Có ba bước Dự đoán – Quan sát – Giải thích, bảng số liệu, đồ thị, nút chép số liệu sang Excel, và tự kiểm mô hình mỗi lần mở.
- `phieu-hoc-tap.docx`: trang học sinh và trang đáp án dành cho giáo viên.
- `can-soat.md`: công thức, điều kiện lí tưởng hoá, kết quả kiểm số, và các điểm thầy cô cần soát. Đọc nguyên văn file này cho thầy cô.

Thí nghiệm ngoài danh mục: đặt `mau: moi`, rồi viết `mo-hinh.json` và `mo-hinh.js` trong cùng thư mục theo `docs/vi/tro-ly/mo-hinh-thi-nghiem.md`. Không viết file HTML bằng tay, không chèn thư viện từ Internet, không bỏ công thức, điều kiện áp dụng hay bảng số kiểm.

## Nối vào bài giảng

Thầy cô làm slide cho cùng bài thì thêm một trang "Thí nghiệm ảo": sơ đồ thí nghiệm vẽ trên slide, một câu nhiệm vụ, và liên kết mở `thi-nghiem.html`. Chép file HTML vào cạnh file PPTX và dùng đường dẫn tương đối; có địa chỉ web thầy cô cung cấp thì dùng địa chỉ đó và có thể thêm mã QR. Không làm mã QR tới file trong máy.

## Ghi vào brief

- Loại việc: Thí nghiệm ảo.
- Thầy cô yêu cầu: môn, lớp, bài, thí nghiệm, hoạt động, người thao tác, kết luận cần rút ra, tham số và số lần đo, sai số đo, phiếu học tập — ghi đúng lời thầy cô.
- AI đề xuất (thầy cô đã đồng ý): mẫu chọn dùng, khoảng tham số, câu dự đoán và các gợi ý thầy cô chấp nhận.
- Loại việc này không tạo PPTX: ghi brief vào `projects/_thi-nghiem/<tên_thí_nghiệm>/brief.md`; không có dòng chốt cách xác nhận, không chạy `import-sources`, không có bước xác nhận của upstream.
- Viết theo mẫu `docs/vi/tro-ly/mau-brief.md`.
