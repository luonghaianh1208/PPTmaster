# Kiểm thử chấp nhận: Thí nghiệm ảo (v6.3.2-vi.8)

Ngày 2026-09-20. Chạy thật `tools/vi/thi_nghiem.py` trên cả 8 mẫu trong thư viện, kiểm bằng mắt qua ảnh chụp trình duyệt và kiểm cấu trúc phiếu Word. Thực hiện theo Task 13 của kế hoạch, do controller làm trực tiếp (không giao subagent).

## Ghi chú môi trường

Máy chạy phiên này chặn `file://` với trình duyệt headless (`chrome-headless-shell` trả về `<body></body>` rỗng với mọi file cục bộ, kể cả file HTML một dòng, trong khi `data:` URL vẫn chạy được). Đây là khác biệt so với lúc dựng thử ở giai đoạn thiết kế trên một môi trường khác. Đã dùng `python -m http.server` chạy cục bộ tại `127.0.0.1` để phục vụ thư mục `projects/_thi-nghiem/`, và chạy `chrome-headless-shell` nhắm vào địa chỉ `http://127.0.0.1:<cổng>/...` thay cho `file://`. Bản thân file HTML khi thầy cô mở bằng trình duyệt thật (bấm đúp trong Windows) không đi qua giới hạn này — giới hạn chỉ riêng của công cụ headless dùng để kiểm thử tự động trong phiên này.

Vì `file://` bị chặn, bước "mở tay bằng Edge" của thầy cô không kiểm thử được trong phiên này; hạng mục còn treo, ghi ở mục Rủi ro.

## Step 1 — Chạy công cụ trên 8 mẫu

| Mẫu | `ready` | `kiem_so.chay` | `dat`/`tong` | `so_lan_do` | `tham_so` |
|---|---|---|---|---|---|
| `li-con-lac-don` | true | true | 6/6 | 5 | chieu-dai |
| `hoa-can-bang-no2` | true | true | 6/6 | 6 | nhiet-do |
| `toan-xac-suat` | true | true | 6/6 | 6 | so-lan, hat-giong |
| `li-nem-xien` | true | true | 6/6 | 4 | goc |
| `li-mach-ohm` | true | true | 5/5 | 4 | dien-tro-2, kieu-mac |
| `hoa-chuan-do` | true | true | 8/8 | 4 | the-tich-base |
| `hoa-toc-do` | true | true | 6/6 | 4 | nhiet-do, xuc-tac |
| `toan-ham-so` | true | true | 6/6 | 4 | a, x0 |

Cả 8 mẫu đều `ready: true`, không có `error`, không có `warnings`. Máy này có Node nên `kiem_so.chay` là `true` ở mọi mẫu; `dat` bằng `tong` ở mọi mẫu (không mẫu nào trượt bảng số kiểm).

## Step 2 — Ảnh chụp và kiểm DOM

Với mỗi mẫu: chụp ảnh khổ rộng (1366×1500) và khổ hẹp/điện thoại (400×1900), dump DOM.

| Mẫu | `id="tu-kiem"` | Số khối `class="buoc"` | Dải đỏ thật trong thân trang |
|---|---|---|---|
| `li-con-lac-don` | Tự kiểm: 6/6 đạt. | 3 | không |
| `hoa-can-bang-no2` | Tự kiểm: 6/6 đạt. | 3 | không |
| `toan-xac-suat` | Tự kiểm: 6/6 đạt. | 3 | không |
| `li-nem-xien` | Tự kiểm: 6/6 đạt. | 3 | không |
| `li-mach-ohm` | Tự kiểm: 5/5 đạt. | 3 | không |
| `hoa-chuan-do` | Tự kiểm: 8/8 đạt. | 3 | không |
| `hoa-toc-do` | Tự kiểm: 6/6 đạt. | 3 | không |
| `toan-ham-so` | Tự kiểm: 6/6 đạt. | 3 | không |

Ghi chú kiểm tra "dải đỏ": chuỗi `dai-do` luôn xuất hiện 2 lần trong file HTML (một lần trong CSS, một lần trong mã JavaScript tạo phần tử khi tự kiểm trượt) — đây là mã nguồn nhúng, không phải phần tử thật. Đã kiểm đúng bằng cách loại bỏ nội dung bên trong thẻ `<script>` trước khi tìm `class="dai-do"` trong thân trang: không mẫu nào có dải đỏ thật, khớp với việc cả 8 mẫu đều tự kiểm đạt 100%.

Xem bằng mắt 4/8 ảnh (con lắc đơn, mạch điện, chuẩn độ, tốc độ phản ứng) ở cả khổ rộng và khổ hẹp: khung vẽ không tràn, chữ không méo, bố cục hẹp xếp dọc đúng như thiết kế, số liệu tính đúng (ví dụ mạch song song R₁=10Ω, R₂=20Ω, U=12V ra I=1,8A, Rtđ=6,67Ω — khớp công thức song song).

## Step 3 — Luồng bấm có khoá (mẫu con lắc đơn, `nguoi-thao-tac: nhom`)

Kịch bản: chọn dự đoán A, bấm "Chốt dự đoán", đổi thanh trượt 5 lần (0,4/0,6/1,0/1,4/1,6 m) và bấm "Ghi lần đo" sau mỗi lần, bấm "Xem gợi ý đáp án và kết luận".

Kết quả: `XONG 5 Đường thẳng khớp: hệ số góc = 4,2252; tung độ gốc = -0,1931; hệ số tương quan r = 0,9987`.

- 5 lần đo được ghi đúng, khớp số lần đã đổi thanh trượt.
- Hệ số góc lí thuyết là 4π²/g ≈ 4,03; kết quả đo được 4,2252 lệch khoảng 5%, hợp lý vì mẫu này bật sai số đo (`sai-so: bat`) nên mỗi lần chạy cho một bộ nhiễu ngẫu nhiên khác nhau (hạt giống lấy theo thời điểm mở trang, không cố định).
- Hệ số tương quan r = 0,9987, rất gần 1, cho thấy quan hệ tuyến tính T² theo l vẫn rõ ràng dù có nhiễu.
- Khoá tuần tự hoạt động đúng: nút "Ghi lần đo" chỉ bấm được sau khi đã chốt dự đoán (kịch bản chốt dự đoán trước mới bấm được, đúng thiết kế).

## Step 4 — Phiếu học tập Word

Không mở được bằng giao diện Word thật trong phiên này (môi trường không có màn hình tương tác cho GUI). Thay bằng kiểm cấu trúc trực tiếp bằng `python-docx` trên 2 mẫu (`li-con-lac-don`, `hoa-chuan-do`):

| Mẫu | Số bảng | Kích thước bảng | Có "PHIẾU HỌC TẬP" | Có "DÀNH CHO GIÁO VIÊN" | Trang giáo viên sau trang học sinh | Có ngắt trang | Chỉ số trên | Chỉ số dưới |
|---|---|---|---|---|---|---|---|---|
| `li-con-lac-don` | 3 | (6,3) bảng đo · (14,20) lưới đồ thị · (6,2) bảng đáp án | có | có | có | có | có (T²) | không (công thức không có chỉ số dưới) |
| `hoa-chuan-do` | 3 | (5,3) bảng đo · (14,20) lưới đồ thị · (5,2) bảng đáp án | có | có | có | có | có | có (CH₃COOH, Kₐ) |

Cả hai đều đúng cấu trúc: bảng số liệu học sinh đúng số cột theo `do:`, lưới đồ thị 14×20 ô vuông, trang giáo viên tách riêng sau dấu ngắt trang, chữ chỉ số trên/dưới dựng đúng run riêng (không lẫn vào chữ thường).

## Kết quả chung

Không phát hiện lỗi nào cần sửa code trong phiên kiểm thử này. Không có commit sửa nào phát sinh từ Task 13.

## Cần thầy cô soát 3 mẫu Hoá

Ba mẫu dùng hằng số khoa học thật, xin thầy cô Hoá học kiểm lại trước khi phát hành:

- **`hoa-chuan-do`** (Chuẩn độ acid – base): Kw = 1,0·10⁻¹⁴; Kₐ(CH₃COOH) = 1,75·10⁻⁵ ở 25 °C, nguồn CRC Handbook of Chemistry and Physics.
- **`hoa-can-bang-no2`** (Cân bằng N₂O₄ ⇌ 2NO₂): Δᵣ = 57,20 kJ/mol, ΔᵣS° = 175,83 J/(mol·K), tính từ ΔfH° và S° ở 298 K, nguồn Atkins Physical Chemistry (bảng dữ liệu nhiệt động).
- **`hoa-toc-do`** (Tốc độ phản ứng): mô hình giả định bậc 1, k(25 °C) = 1,25·10⁻³ s⁻¹, Eₐ = 50 kJ/mol (không xúc tác) và 45 kJ/mol (có xúc tác) — đây là số liệu minh hoạ quy luật Arrhenius, không phải của một phản ứng cụ thể; trang đã ghi rõ điều này trong điều kiện lí tưởng hoá.

File `can-soat.md` đầy đủ của ba mẫu này (công thức, điều kiện, kết quả kiểm số) nằm trong `projects/_thi-nghiem/_thu-hoa-chuan-do/`, `_thu-hoa-can-bang-no2/`, `_thu-hoa-toc-do/` (không commit lên GitHub theo quy tắc).

## Rủi ro còn treo

- Chưa kiểm được bằng mắt qua Word thật (giao diện) do môi trường phiên này không có màn hình tương tác; đã thay bằng kiểm cấu trúc `python-docx`, không thay thế hoàn toàn việc thầy cô tự mở file bằng mắt.
- Chưa mở thử bằng bấm đúp trực tiếp trong Windows Explorer (đường `file://` thật của thầy cô); phiên kiểm thử này chỉ xác nhận được qua HTTP cục bộ do môi trường chặn `file://` cho trình duyệt headless. Cách thầy cô thực sự dùng (bấm đúp mở bằng Chrome/Edge) là đường `file://` bình thường của hệ điều hành, không đi qua giới hạn sandbox của phiên này, nhưng chưa có xác nhận trực tiếp.
- Chủ repo chưa xác nhận ba mẫu Hoá (xem mục trên).
