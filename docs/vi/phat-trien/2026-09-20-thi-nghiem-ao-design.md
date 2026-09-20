# Thiết kế: Thí nghiệm ảo cho Toán, Vật lí, Hoá học (v6.3.2-vi.8)

Ngày 2026-09-20. Gói này thêm loại việc thứ 9 cho lớp Việt: từ một bài dạy, tạo ra một file HTML thí nghiệm ảo tương tác chạy không cần mạng, kèm phiếu học tập Word. Đầu ra không phải PPTX; bài giảng PPTX chỉ có thêm một trang nối sang.

## 1. Tiêu chí thành công

1. Thầy cô nêu môn, lớp, bài và thí nghiệm; sau một lượt hỏi, nhận được `thi-nghiem.html` mở bằng trình duyệt là chạy, và `phieu-hoc-tap.docx` để in.
2. File HTML **không tải gì từ Internet**: không có một địa chỉ `http` nào trong file. Chép USB sang máy chiếu là dùng được; kéo lên Netlify là học sinh dùng được bằng điện thoại.
3. Tám mẫu trong thư viện cho kết quả số đúng, được kiểm bằng bản tính lại độc lập và bằng định luật, không bằng niềm tin.
4. Thí nghiệm ngoài thư viện vẫn làm được, nhưng **không bao giờ lặng lẽ**: công thức, điều kiện áp dụng và bảng số kiểm là bắt buộc, và thầy cô luôn được nhắc soát.
5. Mọi thí nghiệm có cùng giao diện và cùng khung Dự đoán – Quan sát – Giải thích, nên thầy cô học một lần.
6. Học sinh ghi được số liệu, thấy đồ thị vẽ từ chính số liệu đó, chép ra Excel được; có thể bật sai số đo để tập xử lí số liệu.
7. File tự kiểm mỗi lần mở; mô hình sai thì hiện dải đỏ thay vì chạy sai.
8. AI trong Antigravity đi đúng luồng này: cổng hỏi và dòng loại việc thứ 9 nằm thẳng trong file luật Antigravity.
9. Thí nghiệm, phiếu và số liệu của thầy cô không bao giờ lên GitHub.
10. Toàn bộ test của lớp Việt vẫn xanh; không sửa `skills/`, `requirements.txt` gốc hay `attribution_guard.py`.

## 2. Hiện trạng đã kiểm

- Repo chưa có gì tạo thí nghiệm ảo. Giáo án vi.6 chỉ nhắc tới PhET như công cụ ngoài.
- Chủ repo có skill `tro-choi-giao-duc` ngoài repo: một file HTML tự chứa, nhưng tải font Google và KaTeX từ CDN, và AI viết nguyên file mỗi lần. Gói này lấy ý "một file HTML" và bỏ hai điểm còn lại.
- PowerPoint không chạy JavaScript, nên mô phỏng có tham số chỉ chạy được trên trình duyệt.
- `tools/vi/word_parts/` (base.py, inline.py) và `python-docx` đã có từ vi.5, dùng lại cho phiếu học tập. Quy ước chữ: `H~2~SO~4~`, `m/s^2^`, `**in đậm**`.
- Máy chủ repo có Node 22; máy thầy cô không chắc có. `tools/vi/doctor.py` không kiểm Node.
- Bài học vi.7: Antigravity chép nguyên `AGENTS.md` vào luật nhưng không chép nội dung file nhắc bằng `@`; quy tắc nào cần tới Antigravity phải nằm thẳng trong `.agents/rules/ppt-master-vi.md`, mỗi file luật tối đa 12.000 ký tự (hiện 4.693).
- Test `test_antigravity_rule_task_table_matches_common_rules` khoá bảng loại việc ở file luật phải khớp `quy-trinh-hoi.md` và đúng 8 dòng; `test_task_type_count_matches_the_table` khoá chữ "8 loại".

## 3. Quyết định

| # | Quyết định | Lý do |
|---|---|---|
| Q1 | Nằm trong PPT Master bản Việt, loại việc thứ 9 | Thầy cô đã cài repo; phát hành qua `CAP-NHAT.bat`; có test và công cụ kiểm |
| Q2 | Đầu ra là file HTML tương tác; slide chỉ có trang nối sang | PowerPoint không chạy mô phỏng |
| Q3 | Thư viện mẫu + AI lắp ghép; ngoài thư viện AI viết mô hình mới theo khuôn | Khoa học đúng được kiểm; vẫn không phải từ chối thầy cô |
| Q4 | Bản đầu 8 mẫu: Lí 3, Hoá 3, Toán 2 | Kiểm kỹ từng mẫu, mở rộng theo phản hồi |
| Q5 | Chạy không cần mạng | Wifi trường không ổn định |
| Q6 | Kèm Dự đoán – Quan sát – Giải thích, bảng số liệu + đồ thị, sai số đo, phiếu Word | Chủ repo chọn cả bốn, làm trọn trong vi.8 |
| Q7 | Máy dựng từ linh kiện: khung chạy chung + file mô hình + `thi-nghiem.md` | AI yếu viết 500 dòng mô phỏng dễ sai ngầm; ghép linh kiện đã kiểm thì không |
| Q8 | Kiểm số ba lớp: test repo, Node trên máy thầy cô nếu có, tự kiểm trong HTML | Lớp thứ ba không cần cài gì nên máy nào cũng có |
| Q9 | Không lưu điểm, không gửi dữ liệu | File không có mạng; đây không phải hệ thống chấm điểm |
| Q10 | Công thức hiển thị bằng HTML thuần và ký hiệu Unicode, font hệ thống | Không CDN; đủ cho chỉ số, số mũ, căn, mũi tên cân bằng |
| Q11 | Mỗi mô hình là hai file: `<mã>.json` (khai báo) và `<mã>.js` (hàm `tinh`, `ve`) | Python kiểm được khai báo, khoảng tham số và bảng số kiểm mà không cần Node |
| Q12 | Đại lượng đo không phụ thuộc thời gian; thời gian chỉ dùng cho hình vẽ | Bảng số kiểm, phiếu và đồ thị chỉ cần `tinh(thamSo)`; đơn giản và kiểm được |
| Q13 | Bản tính lại bằng Python nằm ở `thi_nghiem_parts/tham_chieu.py` | Dùng cho cả test lẫn bảng số liệu lí tưởng của phiếu, nên không thể nằm trong thư mục test |
| Q14 | `thi-nghiem.md` không được chứa địa chỉ web | Giữ nguyên tắc "không có `http` trong file HTML" mà không phải phân biệt chữ của thầy cô với mã |

## 4. Kiến trúc và file

```
tools/vi/thi_nghiem.py                 kiểm → ghép HTML → phiếu Word → đúng một dòng JSON
tools/vi/thi_nghiem_parts/
    parse.py                           đọc thi-nghiem.md, lỗi nêu đúng dòng
    build_html.py                      ghép khung chạy + mô hình + cấu hình thành một file
    thu_vien.py                        nạp mô hình (thư viện hoặc mới), kiểm khai báo và mã
    kiem_so.py                         chạy bangKiem qua Node nếu máy có
    tham_chieu.py                      bản tính lại bằng Python của 8 mô hình
    phieu.py                           phiếu học tập .docx (dùng word_parts/)
    runtime/khung.js, runtime/khung.css, runtime/chay_node.js
    mo_hinh/<mã>.json + <mã>.js        li-nem-xien, li-con-lac-don, li-mach-ohm, hoa-chuan-do,
                                       hoa-can-bang-no2, hoa-toc-do, toan-ham-so, toan-xac-suat
tools/vi/tests/test_thi_nghiem_*.py, tools/vi/tests/thi_nghiem_mau.py, tools/vi/tests/js/test_khung.js
docs/vi/tro-ly/thi-nghiem-ao.md        câu hỏi cho thầy cô, ngữ pháp thi-nghiem.md, thứ tự làm
docs/vi/tro-ly/mo-hinh-thi-nghiem.md   khuôn viết mô hình mới, danh mục 8 mẫu và tham số
docs/vi/thi-nghiem-ao.md               tài liệu cho thầy cô
```

Sửa: `AGENTS.vi.md` (mục 3 câu lệnh kích hoạt, mục 10 bảng 9 loại, mục 14 mới), `.agents/rules/ppt-master-vi.md`, `docs/vi/tro-ly/quy-trinh-hoi.md`, `docs/vi/xu-ly-loi.md`, `docs/vi/cau-lenh-mau.md`, `README.md`, `CHANGELOG-VI.md`, `tools/vi/tests/test_vi_layer.py`.

Đầu ra ở `projects/_thi-nghiem/<tên>/`: `thi-nghiem.md`, `thi-nghiem.html`, `phieu-hoc-tap.docx`, `can-soat.md`, và `mo-hinh.json` cùng `mo-hinh.js` khi là mô hình mới. Không chạy `project_manager.py init`, không tạo SVG.

Ràng buộc: công cụ Python chỉ dùng thư viện chuẩn cộng `python-docx` (đã có); khung chạy và mô hình là JavaScript thuần, không thư viện ngoài, chạy được cả trong trình duyệt lẫn Node (không dùng DOM trong phần tính).

## 5. File HTML

- Một file tự chứa: CSS, khung chạy, mô hình, cấu hình (JSON) nằm trong file. `build_html.py` từ chối ghi nếu kết quả chứa chuỗi `http://` hoặc `https://`.
- Font: `"Segoe UI", Arial, sans-serif`. Công thức: `<sub>`, `<sup>` và ký hiệu Unicode (√, π, ⇌, Δ) viết trên một dòng.
- Bố cục co giãn: màn rộng thì khung mô phỏng bên trái, điều khiển và nhiệm vụ bên phải; màn hẹp (khoảng 400px) thì xếp dọc. Chữ đủ lớn cho máy chiếu.
- Thành phần khung chạy: thanh trượt hoặc ô chọn cho tham số, danh sách tham số giữ cố định, nút Chạy/Dừng/Đặt lại, khung vẽ `canvas`, ô số đo, bảng số liệu, đồ thị, ba bước nhiệm vụ, nút Chế độ giáo viên, dòng chú thích điều kiện lí tưởng hoá, dòng tự kiểm. Sai số đo bật hay tắt do `thi-nghiem.md` quyết định; trang chỉ ghi chú khi đang bật.
- Dòng chú thích lấy từ `congThuc.dieuKien` của mô hình, luôn hiện.

## 6. Khuôn mô hình

Mỗi mô hình khai đúng sáu thứ, chia hai file. File `.json` giữ bốn khai báo dữ liệu, file `.js` giữ hai hàm:

| Khai báo | Ở file | Nội dung |
|---|---|---|
| `thamSo` | json | mỗi tham số: `ma`, `ten`, `kieu` (`so` hoặc `chon`); loại `so` có `donVi`, `min`, `max`, `buoc`, `macDinh`; loại `chon` có `luaChon` và `macDinh` |
| `daiLuongDo` | json | mỗi đại lượng đo được: `ma`, `ten`, `donVi`, `saiSo` (độ lệch chuẩn của nhiễu đo), `chuSo` (số chữ số thập phân) |
| `congThuc` | json | `bieuThuc` (chữ, theo quy ước `~`, `^`), `dieuKien` (điều kiện áp dụng), `nguon` (nguồn hằng số nếu có) |
| `bangKiem` | json | ít nhất 5 dòng `{vao, ra, saiSo}`: tham số khác mặc định → giá trị đúng của đại lượng đo, sai số tuyệt đối cho phép |
| `tinh(thamSo)` | js | trả về giá trị mọi đại lượng đo; hàm thuần, không phụ thuộc thời gian, không đụng DOM |
| `ve(ctx, thamSo, t, kichThuoc, ketQua)` | js | vẽ lên `canvas` tại thời điểm `t`; mô hình `hoatHinh: mot-lan` có thêm `thoiLuong(thamSo, ketQua)` |

File `.json` còn có `ten`, `mon` và `hoatHinh` (`khong`, `mot-lan` hoặc `lap`).

Khoảng cho phép của `thamSo` là ranh giới của điều kiện áp dụng: `thi-nghiem.md` không được vượt (ví dụ góc lệch con lắc tối đa 15°).

### Tám mẫu bản đầu

| Mã | Tham số chính | Đại lượng đo | Điều kiện lí tưởng hoá |
|---|---|---|---|
| `li-nem-xien` | vận tốc đầu, góc, độ cao đầu, g | tầm xa, độ cao cực đại, thời gian bay | bỏ qua sức cản không khí |
| `li-con-lac-don` | chiều dài, g, góc lệch, khối lượng | chu kì | góc lệch nhỏ, dây không giãn |
| `li-mach-ohm` | suất điện động, hai điện trở, kiểu mắc | cường độ, hiệu điện thế từng điện trở | nguồn và dây dẫn lí tưởng |
| `hoa-chuan-do` | acid (HCl hoặc CH~3~COOH), nồng độ và thể tích acid, nồng độ NaOH, thể tích NaOH đã nhỏ, chất chỉ thị | pH | dung dịch loãng, 25 °C |
| `hoa-can-bang-no2` | nhiệt độ, áp suất chung (bar) | K~p~, phần mol NO~2~, độ phân li, nồng độ NO~2~ (độ đậm màu) | khí lí tưởng; Δ~r~H°, Δ~r~S° không đổi trong 0–100 °C |
| `hoa-toc-do` | nồng độ, nhiệt độ, xúc tác | thời gian tới khi vẩn đục | phản ứng giả định bậc 1, E~a~ khai trong mô hình |
| `toan-ham-so` | hệ số a, b, c, d (bậc ba; a = 0 thành bậc hai), hoành độ tiếp điểm | giá trị, hệ số góc tiếp tuyến, hoành độ cực đại và cực tiểu | — |
| `toan-xac-suat` | loại phép thử (xúc xắc, đồng xu, hai xúc xắc), số lần, biến cố | tần số, tần suất | bộ sinh số ngẫu nhiên có hạt giống |

## 7. Giữ cho khoa học đúng

1. **Test trong repo.** Mỗi mẫu có bản tính lại độc lập bằng Python trong `thi_nghiem_parts/tham_chieu.py`. Test chạy `tinh()` qua Node trên lưới vài trăm điểm và so với Python. Thêm test định luật: ném xiên bảo toàn cơ năng; con lắc T² tỉ lệ với l; mạch song song tổng dòng nhánh bằng dòng chính; chuẩn độ acid mạnh – base mạnh có pH 7 tại điểm tương đương và acid yếu có pH = pK~a~ tại nửa điểm tương đương; N~2~O~4~ ⇌ 2NO~2~ tăng nhiệt độ thì phần mol NO~2~ tăng, tăng áp suất thì giảm; tốc độ tăng theo nhiệt độ đúng hệ thức Arrhenius; hàm bậc hai có đỉnh tại −b/2a; tần suất gieo 100.000 lần lệch xác suất dưới 1%. Máy không có Node thì các test này bỏ qua có ghi lý do; máy phát hành phải có Node.
2. **Trên máy thầy cô.** `kiem_so.py` chạy `bangKiem` qua Node nếu có. Không có Node: `kiem_so.chay` là `false` và có cảnh báo "chưa chạy kiểm số trên máy này; file sẽ tự kiểm khi mở".
3. **Trong file HTML.** Mỗi lần mở, file chạy `bangKiem`. Đạt: dòng nhỏ "Tự kiểm: 5/5 đạt". Trượt: dải đỏ "Mô hình không qua tự kiểm — không dùng để dạy", nêu dòng trượt.

Mô hình mới do AI viết (`mau: moi`): công cụ chặn nếu thiếu một trong sáu khai báo, `congThuc.dieuKien` rỗng, `bangKiem` dưới 5 dòng, hoặc file có `http`, `fetch`, `XMLHttpRequest`, `import`, `eval`. `can-soat.md` luôn có đoạn: mô hình do AI viết, chưa có người duyệt; công thức là …; bảng số kiểm do AI tự tính, thầy cô kiểm lại ít nhất hai dòng bằng máy tính cầm tay. Bảng kiểm do AI viết chỉ bắt được lỗi lập trình, không bắt được lỗi hiểu sai kiến thức; vì vậy đoạn nhắc này không tắt được. Khi `check` trượt, AI sửa mô hình; sửa bảng kiểm thì phải nêu rõ trong `can-soat.md`.

Sai số đo: mặc định tắt. Bật thì mỗi lần ghi lần đo cộng nhiễu chuẩn có độ lệch `daiLuongDo.saiSo`; hình vẽ dùng giá trị thật; tự kiểm dùng giá trị không nhiễu. Nhiễu dùng bộ sinh số có hạt giống để test được.

## 8. Lượt hỏi và `thi-nghiem.md`

### Câu hỏi bắt buộc (7 câu, mỗi câu có gợi ý)

1. Môn, lớp, tên bài, và thí nghiệm nào? Gợi ý: mẫu gần nhất trong thư viện; ngoài thư viện thì nói rõ sẽ viết mô hình mới và thầy cô cần soát công thức.
2. Thí nghiệm dùng ở hoạt động nào: mở đầu, hình thành kiến thức, hay luyện tập? Gợi ý: hình thành kiến thức.
3. Ai thao tác: giáo viên trên máy chiếu, hay học sinh theo nhóm? Gợi ý: giáo viên trình diễn, nhóm học sinh ghi phiếu.
4. Học sinh cần rút ra kết luận gì? Gợi ý: một kết luận theo yêu cầu cần đạt.
5. Khoảng giá trị và số lần đo? Gợi ý: mặc định của mẫu, 5 lần đo.
6. Có bật sai số đo không? Gợi ý: tắt với THCS, bật với THPT.
7. Có cần phiếu học tập Word không? Gợi ý: có.

Tạo nhanh: câu 1 và câu 3. Câu lệnh kích hoạt: "thí nghiệm ảo", "mô phỏng thí nghiệm", "làm thí nghiệm ảo". Loại việc này ghi brief như các loại khác nhưng không đi vào quy trình PPTX của upstream (cùng cách với đề thi và giáo án).

### Ngữ pháp

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
cau: Từ đồ thị, T^2^ và l liên hệ thế nào? Suy ra công thức.
goi-y-dap-an: T^2^ tỉ lệ thuận với l; T = 2π√(l/g).

## Kết luận
Chu kì con lắc đơn chỉ phụ thuộc chiều dài dây và g, không phụ thuộc khối lượng.
```

- `nguoi-thao-tac`: `giao-vien` (mở sẵn Chế độ giáo viên, là mặc định) hoặc `nhom`. `sai-so`: `bat` hoặc `tat` (mặc định).
- Tham số loại lựa chọn: `<mã>: chon <mã_1>, <mã_2>` (mục đầu là mặc định) hoặc `<mã>: co-dinh <mã_lựa_chọn>`.
- Dòng tham số có hai dạng: `<mã>: <nhỏ nhất>..<lớn nhất> buoc <bước> mac-dinh <giá trị>` (có thanh trượt), hoặc `<mã>: co-dinh <giá trị>` (ẩn thanh trượt). Tham số không khai trong `## Tham số` thì cố định ở mặc định của mẫu.
- `do-thi`: `<biểu thức> theo <biểu thức>`; biểu thức là một mã tham số hoặc đại lượng đo, tuỳ chọn kèm đúng một phép biến đổi: `<mã>^2`, `1/<mã>`, `ln(<mã>)`, `sqrt(<mã>)`. Đây là cú pháp tính toán, khác quy ước hiển thị `^2^` dùng trong các dòng chữ (`cau:`, `goi-y-dap-an:`, Kết luận).
- `## Dự đoán` có thể là trắc nghiệm (`A:`…`dap-an:`) hoặc câu mở (chỉ `cau:`).
- Lỗi nêu đúng dòng: mẫu không tồn tại; tham số hay đại lượng không có trong mẫu; khoảng vượt khoảng cho phép; bước không dương; mặc định ngoài khoảng; thiếu một trong các mục Tham số, Dự đoán, Quan sát, Giải thích, Kết luận; `dap-an` không nằm trong lựa chọn; `so-lan-do` ngoài 3–20; `mau: moi` mà thiếu `mo-hinh.json` hoặc `mo-hinh.js`; trong file có địa chỉ web.
- File hướng dẫn cho AI phải có một `thi-nghiem.md` mẫu cho mỗi môn, và test chạy các mẫu đó qua parser thật (bài học vi.5).

## 9. Khung Dự đoán – Quan sát – Giải thích

- Ba bước khoá tuần tự: chưa dự đoán thì nút Chạy bị khoá; dự đoán xong không sửa được và chưa bị chấm; ghi đủ `so-lan-do` thì mở Giải thích, lúc đó mới đối chiếu dự đoán với kết quả và hiện gợi ý đáp án khi bấm.
- Chế độ giáo viên bỏ mọi khoá.
- Bảng số liệu: Ghi lần đo, xoá từng dòng, cột dẫn xuất theo `do-thi`, đồ thị điểm kèm đường khớp tuyến tính với hệ số góc và hệ số tương quan, nút Chép số liệu (dạng tab, dán được vào Excel).
- Không lưu gì xuống máy, không gửi gì đi.

## 10. Phiếu học tập Word

A4 dọc, thể thức như đề thi vi.5. Trang học sinh: tiêu đề, họ tên – nhóm – lớp, câu dự đoán, bảng trống đúng cột và số lần đo, khung kẻ ô vẽ đồ thị có tên trục và đơn vị, câu giải thích, kết luận để trống. Trang giáo viên (trang riêng cuối): đáp án, bảng số liệu lí tưởng tính từ mô hình tại các giá trị tham số đều trong khoảng, công thức và điều kiện. Bảng lí tưởng được tính bằng `tham_chieu.py` cho 8 mẫu; mô hình mới thì tính qua Node nếu có, không thì bỏ bảng và ghi cảnh báo.

## 11. Công cụ

```
python tools/vi/thi_nghiem.py <thư_mục> [--phan tat-ca|html|phieu] [--plan-only]
```

stdout đúng một dòng JSON: `ready`, `files`, `mau`, `tham_so`, `so_lan_do`, `kiem_so` (`{"chay", "dat", "tong"}`), `warnings`, `error`. Mã thoát 0 hoặc 1.

| `error.step` | Nghĩa | Xử lý |
|---|---|---|
| `input` | thiếu thư mục hoặc `thi-nghiem.md`, sai tham số lệnh | viết file rồi chạy lại |
| `parse` | sai ngữ pháp, nêu đúng dòng | sửa dòng đó |
| `model` | mẫu không có; mô hình mới thiếu khai báo, thiếu điều kiện, bảng kiểm dưới 5 dòng, có lệnh mạng | sửa theo `error.fix`, không bỏ bảng kiểm |
| `check` | bảng số kiểm trượt khi chạy qua Node | sửa mô hình; sửa bảng kiểm thì phải ghi vào `can-soat.md` |
| `docx` | chưa có `python-docx` | cài `tools/vi/requirements-vi.txt`, chạy lại một lần |
| `write` | không ghi được file | đóng file đang mở rồi chạy lại |
| `internal` | lỗi ngoài dự kiến | dán nguyên `error.message` cho người bảo trì |

## 12. Nối vào bài giảng

Khi thầy cô làm slide cho cùng bài và đã có thí nghiệm ảo, AI thêm một trang "Thí nghiệm ảo": ảnh sơ đồ thí nghiệm (vẽ SVG), một câu nhiệm vụ, và liên kết mở file HTML (đường dẫn tương đối tới file chép cạnh PPTX, hoặc địa chỉ Netlify thầy cô cung cấp). Mã QR chỉ làm khi có địa chỉ web. Phần này chỉ là hướng dẫn trong `thi-nghiem-ao.md`, dùng cơ chế liên kết sẵn có của upstream (`native-hyperlinks.md`), không có mã mới.

## 13. Hướng dẫn cho AI và luật

- `AGENTS.vi.md` mục 14 "Làm thí nghiệm ảo": thứ tự làm (hỏi → tạo thư mục → viết `thi-nghiem.md`, và `mo-hinh.js` nếu mới → chạy công cụ → đọc JSON → báo thầy cô đường dẫn file, kết quả kiểm số, nguyên văn `warnings` và `can-soat.md`), bảng `error.step`, điều cấm (không tự bỏ bảng kiểm, không chèn thư viện từ Internet, không commit `projects/`, không chạm `skills/`).
- Bảng loại việc thành 9 dòng ở `AGENTS.vi.md`, `quy-trinh-hoi.md` và file luật Antigravity; chữ "8 loại" thành "9 loại"; test tương ứng cập nhật. Thứ tự các mục cuối của `AGENTS.vi.md` mà test đang khoá được mở rộng thêm mục 14.
- File luật Antigravity thêm dòng bảng và một đoạn ngắn "Thí nghiệm ảo": đọc `thi-nghiem-ao.md`, chạy `tools\vi\thi_nghiem.py`, không viết HTML tay. Vẫn dưới 12.000 ký tự.

## 14. Test

- Khung chạy (qua Node): khoá ba bước, ghi và xoá lần đo, cột dẫn xuất, khớp tuyến tính, nhiễu có hạt giống, tự kiểm báo trượt khi bảng kiểm sai.
- Tám mô hình: so với Python tham chiếu và test định luật (mục 7).
- Parser: mọi lỗi ở mục 8 nêu đúng dòng; ba file mẫu trong hướng dẫn chạy qua được.
- Công cụ: mọi nhánh ra đúng một dòng JSON; HTML không chứa `http`; mô hình cố tình sai bị `check` chặn; mô hình thiếu `congThuc` hoặc có `fetch` bị `model` chặn; không có Node thì vẫn `ready` kèm cảnh báo.
- Phiếu Word: mở lại bằng `python-docx`, đúng số cột và số dòng, có trang giáo viên.
- Lớp hướng dẫn: `test_vi_layer.py` cho loại việc thứ 9, mục 14, file luật, tài liệu thầy cô, mục xử lý lỗi.
- Kiểm bằng mắt trước khi phát hành: mở 8 file HTML bằng trình duyệt (có mở cửa sổ), chụp màn hình rộng và hẹp, gửi chủ repo. Chủ repo soát riêng ba mẫu Hoá.

## 15. Ngoài phạm vi

Lưu điểm và tài khoản học sinh; gửi số liệu về giáo viên; mô phỏng 3D; Sinh học và Khoa học Trái Đất; mô phỏng trong slide bằng Morph (thuộc gói sơ đồ thí nghiệm); tự đưa lên Netlify; nhúng HTML vào PowerPoint.

## 16. Rủi ro

- Mô hình Hoá cần hằng số thật (K~a~, ΔH, K~p~ theo nhiệt độ, năng lượng hoạt hoá); sai nguồn là sai cả mẫu. Ghi nguồn trong `congThuc.nguon`, chủ repo soát.
- Mô phỏng là mô hình lí tưởng, không thay thí nghiệm thật; dòng chú thích điều kiện luôn hiện.
- Mô hình do AI viết mới chỉ được bảo vệ bằng bảng kiểm do chính AI viết và lời nhắc thầy cô soát.
- Điện thoại không mở được file từ USB; dùng điện thoại phải đưa lên web. Tài liệu thầy cô ghi rõ.
- Máy thầy cô không có Node thì lớp kiểm thứ hai không chạy; còn lớp tự kiểm trong HTML.
- Bản phát hành lớn vì gồm cả phiếu Word; thời gian dài hơn các gói trước.
