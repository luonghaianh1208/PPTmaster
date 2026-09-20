# Mô hình thí nghiệm ảo: danh mục mẫu và khuôn viết mô hình mới

File dành cho AI. Đọc cùng `docs/vi/tro-ly/thi-nghiem-ao.md`. Phần đầu là danh mục 8 mẫu đã kiểm trong thư viện; phần sau là khuôn bắt buộc khi thầy cô cần một thí nghiệm ngoài danh mục.

## Danh mục mẫu

Dùng đúng các mã dưới đây trong `thi-nghiem.md`. Khoảng ở cột cuối là giới hạn áp dụng của công thức; `thi-nghiem.md` chỉ được chọn khoảng nằm trong đó.

### `hoa-can-bang-no2` — Cân bằng N₂O₄ ⇌ 2NO₂ (Hoá học)

| Tham số | Mã | Khoảng, bước, mặc định |
|---|---|---|
| Nhiệt độ | `nhiet-do` | 0 đến 100 °C, bước 5, mặc định 25 |
| Áp suất chung | `ap-suat` | 0,5 đến 5 bar, bước 0,1, mặc định 1 |

| Đại lượng đo | Mã | Sai số đo |
|---|---|---|
| Phần mol NO~2~ | `phan-mol-no2` | 0,005 |
| Độ phân li của N~2~O~4~ | `do-phan-li` | 0,005 |
| Nồng độ NO~2~ (độ đậm màu nâu) (mol/L) | `nong-do-no2` | 0,0005 |
| Hằng số cân bằng K~p~ | `kp` | 0 |

- Mô hình: N~2~O~4~(g) ⇌ 2NO~2~(g), Δ~r~H° = +57,2 kJ; K~p~ = exp(−(Δ~r~H° − TΔ~r~S°)/RT); K~p~ = x^2^P/(1 − x) với x là phần mol NO~2~.
- Điều kiện lí tưởng hoá: Hỗn hợp khí lí tưởng đã đạt cân bằng; coi Δ~r~H° và Δ~r~S° không đổi trong khoảng 0–100 °C; áp suất tính bằng bar.
- Nguồn số liệu: Δ~r~H° = 57,20 kJ/mol và Δ~r~S° = 175,83 J/(mol·K), tính từ Δ~f~H° và S° ở 298 K (Atkins, Physical Chemistry, bảng dữ liệu nhiệt động).

### `hoa-chuan-do` — Chuẩn độ acid – base bằng dung dịch NaOH (Hoá học)

| Tham số | Mã | Khoảng, bước, mặc định |
|---|---|---|
| Acid cần chuẩn độ | `loai-acid` | lựa chọn: `hcl` (HCl (acid mạnh)), `ch3cooh` (CH~3~COOH (acid yếu)); mặc định `hcl` |
| Nồng độ acid | `nong-do-acid` | 0,01 đến 0,5 mol/L, bước 0,01, mặc định 0,1 |
| Thể tích acid | `the-tich-acid` | 10 đến 50 mL, bước 5, mặc định 20 |
| Nồng độ NaOH | `nong-do-base` | 0,01 đến 0,5 mol/L, bước 0,01, mặc định 0,1 |
| Thể tích NaOH đã nhỏ | `the-tich-base` | 0 đến 50 mL, bước 0,1, mặc định 0 |
| Chất chỉ thị | `chi-thi` | lựa chọn: `phenolphtalein` (Phenolphtalein), `metyl-da-cam` (Methyl da cam), `bromothymol` (Bromothymol xanh); mặc định `phenolphtalein` |

| Đại lượng đo | Mã | Sai số đo |
|---|---|---|
| pH của dung dịch | `ph` | 0,05 |

- Mô hình: Bảo toàn điện tích: [H^+^] + [Na^+^] = [OH^−^] + [A^−^]; K~w~ = [H^+^][OH^−^]; với acid yếu [A^−^] = C·K~a~/(K~a~ + [H^+^]).
- Điều kiện lí tưởng hoá: Dung dịch loãng ở 25 °C (K~w~ = 1,0·10^−14^); coi hoạt độ bằng nồng độ; thể tích cộng tính.
- Nguồn số liệu: K~a~(CH~3~COOH) = 1,75·10^−5^ ở 25 °C (CRC Handbook of Chemistry and Physics).

### `hoa-toc-do` — Các yếu tố ảnh hưởng đến tốc độ phản ứng (Hoá học)

| Tham số | Mã | Khoảng, bước, mặc định |
|---|---|---|
| Nồng độ chất tham gia | `nong-do` | 0,02 đến 0,5 mol/L, bước 0,02, mặc định 0,1 |
| Nhiệt độ | `nhiet-do` | 10 đến 60 °C, bước 5, mặc định 25 |
| Chất xúc tác | `xuc-tac` | lựa chọn: `khong` (Không có), `co` (Có xúc tác); mặc định `khong` |

| Đại lượng đo | Mã | Sai số đo |
|---|---|---|
| Thời gian đến khi vẩn đục che dấu X (s) | `thoi-gian` | 0,5 |

- Mô hình: v = k·C; k = A·exp(−E~a~/RT); thời gian t = ΔC/v với ΔC = 0,005 mol/L.
- Điều kiện lí tưởng hoá: Phản ứng giả định bậc 1 theo chất tham gia, dùng tốc độ đầu; E~a~ = 50 kJ/mol khi không có xúc tác và 45 kJ/mol khi có; k = 1,25·10^−3^ s^−1^ ở 25 °C. Số liệu minh hoạ quy luật, không phải của một phản ứng cụ thể.

### `li-con-lac-don` — Con lắc đơn (Vật lí)

| Tham số | Mã | Khoảng, bước, mặc định |
|---|---|---|
| Chiều dài dây l | `chieu-dai` | 0,2 đến 2 m, bước 0,05, mặc định 1 |
| Gia tốc trọng trường g | `g` | 1,6 đến 24,8 m/s^2^, bước 0,1, mặc định 9,8 |
| Góc lệch ban đầu | `goc-lech` | 2 đến 15 °, bước 1, mặc định 8 |
| Khối lượng quả nặng m | `khoi-luong` | 0,05 đến 1 kg, bước 0,05, mặc định 0,2 |

| Đại lượng đo | Mã | Sai số đo |
|---|---|---|
| Chu kì T (s) | `chu-ki` | 0,02 |
| Thời gian 10 dao động (s) | `thoi-gian-10-dao-dong` | 0,1 |

- Mô hình: T = 2π√(l/g)
- Điều kiện lí tưởng hoá: Góc lệch nhỏ (không quá 15°); dây không giãn, khối lượng dây không đáng kể; bỏ qua sức cản.

### `li-mach-ohm` — Đoạn mạch nối tiếp và song song (Vật lí)

| Tham số | Mã | Khoảng, bước, mặc định |
|---|---|---|
| Hiệu điện thế nguồn U | `suat-dien-dong` | 1,5 đến 24 V, bước 0,5, mặc định 12 |
| Điện trở R~1~ | `dien-tro-1` | 1 đến 100 Ω, bước 1, mặc định 10 |
| Điện trở R~2~ | `dien-tro-2` | 1 đến 100 Ω, bước 1, mặc định 20 |
| Kiểu mắc | `kieu-mac` | lựa chọn: `noi-tiep` (Nối tiếp), `song-song` (Song song); mặc định `noi-tiep` |

| Đại lượng đo | Mã | Sai số đo |
|---|---|---|
| Cường độ mạch chính I (A) | `cuong-do-mach-chinh` | 0,01 |
| Cường độ qua R~1~ (A) | `cuong-do-1` | 0,01 |
| Cường độ qua R~2~ (A) | `cuong-do-2` | 0,01 |
| Hiệu điện thế hai đầu R~1~ (V) | `hieu-dien-the-1` | 0,05 |
| Hiệu điện thế hai đầu R~2~ (V) | `hieu-dien-the-2` | 0,05 |
| Điện trở tương đương (Ω) | `dien-tro-tuong-duong` | 0 |

- Mô hình: Nối tiếp: R = R~1~ + R~2~, I = U/R. Song song: 1/R = 1/R~1~ + 1/R~2~, I = I~1~ + I~2~.
- Điều kiện lí tưởng hoá: Nguồn có điện trở trong bằng 0; dây nối và ampe kế có điện trở không đáng kể; vôn kế có điện trở rất lớn.

### `li-nem-xien` — Chuyển động ném xiên (Vật lí)

| Tham số | Mã | Khoảng, bước, mặc định |
|---|---|---|
| Vận tốc đầu v~0~ | `van-toc-dau` | 5 đến 50 m/s, bước 1, mặc định 20 |
| Góc ném α | `goc` | 0 đến 85 °, bước 1, mặc định 45 |
| Độ cao ban đầu h | `do-cao-dau` | 0 đến 50 m, bước 1, mặc định 0 |
| Gia tốc trọng trường g | `g` | 1,6 đến 24,8 m/s^2^, bước 0,1, mặc định 9,8 |

| Đại lượng đo | Mã | Sai số đo |
|---|---|---|
| Tầm xa L (m) | `tam-xa` | 0,2 |
| Độ cao cực đại H (m) | `do-cao-cuc-dai` | 0,1 |
| Thời gian bay t (s) | `thoi-gian-bay` | 0,02 |

- Mô hình: t = (v~0~sinα + √(v~0~^2^sin^2^α + 2gh))/g; L = v~0~cosα·t; H = h + v~0~^2^sin^2^α/(2g)
- Điều kiện lí tưởng hoá: Bỏ qua sức cản không khí; g không đổi; vật coi là chất điểm.

### `toan-ham-so` — Khảo sát hàm số y = ax³ + bx² + cx + d (Toán)

| Tham số | Mã | Khoảng, bước, mặc định |
|---|---|---|
| Hệ số a | `a` | -5 đến 5, bước 0,5, mặc định 1 |
| Hệ số b | `b` | -5 đến 5, bước 0,5, mặc định 0 |
| Hệ số c | `c` | -5 đến 5, bước 0,5, mặc định -3 |
| Hệ số d | `d` | -5 đến 5, bước 0,5, mặc định 0 |
| Hoành độ tiếp điểm x~0~ | `x0` | -5 đến 5, bước 0,1, mặc định 1 |

| Đại lượng đo | Mã | Sai số đo |
|---|---|---|
| Giá trị y(x~0~) | `gia-tri` | 0 |
| Hệ số góc tiếp tuyến y'(x~0~) | `he-so-goc` | 0 |
| Hoành độ điểm cực đại | `hoanh-do-cuc-dai` | 0 |
| Hoành độ điểm cực tiểu | `hoanh-do-cuc-tieu` | 0 |

- Mô hình: y' = 3ax^2^ + 2bx + c; cực trị tại nghiệm của y' = 0 nơi y' đổi dấu; a = 0 thì là hàm bậc hai, đỉnh tại x = −c/(2b).
- Điều kiện lí tưởng hoá: Hệ số thực; a = 0 và b = 0 thì hàm bậc nhất hoặc hằng, không có cực trị.

### `toan-xac-suat` — Xác suất thực nghiệm (Toán)

| Tham số | Mã | Khoảng, bước, mặc định |
|---|---|---|
| Phép thử và biến cố | `phep-thu` | lựa chọn: `dong-xu-ngua` (Tung đồng xu: mặt ngửa), `xuc-xac-mat-6` (Gieo xúc xắc: mặt 6 chấm), `xuc-xac-chan` (Gieo xúc xắc: số chấm chẵn), `hai-xuc-xac-tong-7` (Gieo hai xúc xắc: tổng bằng 7), `hai-xuc-xac-tong-12` (Gieo hai xúc xắc: tổng bằng 12); mặc định `dong-xu-ngua` |
| Số lần thử n | `so-lan` | 10 đến 10000 lần, bước 10, mặc định 100 |
| Lượt gieo số | `hat-giong` | 1 đến 999, bước 1, mặc định 1 |

| Đại lượng đo | Mã | Sai số đo |
|---|---|---|
| Số lần biến cố xảy ra (lần) | `tan-so` | 0 |
| Tần suất | `tan-suat` | 0 |
| Xác suất lí thuyết | `xac-suat-li-thuyet` | 0 |

- Mô hình: Tần suất f = (số lần biến cố xảy ra)/n; khi n lớn, f tiến gần xác suất P.
- Điều kiện lí tưởng hoá: Đồng xu và xúc xắc cân đối, các lần thử độc lập; số ngẫu nhiên sinh bằng thuật toán có hạt giống nên cùng một lượt gieo cho cùng kết quả.

## Khuôn mô hình mới

Chỉ viết mô hình mới khi không mẫu nào trong danh mục dùng được. Đặt `mau: moi` trong `thi-nghiem.md` và viết hai file trong cùng thư mục thí nghiệm.

`mo-hinh.json` khai bốn thứ, cùng `ten`, `mon` và `hoatHinh`:

| Khoá | Nội dung |
|---|---|
| `hoatHinh` | `khong` (hình tĩnh, đổi theo tham số), `mot-lan` (chạy một lần rồi mới có số đo) hoặc `lap` (chạy lặp lại) |
| `thamSo` | mỗi tham số có `ma`, `ten`, `kieu`. `kieu: so` cần `donVi`, `min`, `max`, `buoc`, `macDinh`. `kieu: chon` cần `luaChon` (ít nhất hai mục có `ma`, `ten`) và `macDinh`. `min` và `max` là giới hạn áp dụng của công thức. |
| `daiLuongDo` | mỗi đại lượng có `ma`, `ten`, `donVi`, `saiSo` (độ lệch chuẩn của nhiễu khi bật sai số đo; 0 nếu không có nhiễu), `chuSo` (số chữ số thập phân hiển thị) |
| `congThuc` | `bieuThuc` (công thức bằng chữ), `dieuKien` (điều kiện áp dụng, bắt buộc), `nguon` (nguồn của hằng số; để chuỗi rỗng nếu không dùng hằng số tra cứu) |
| `bangKiem` | ít nhất 5 dòng `{"vao": ..., "ra": ..., "saiSo": ...}`. `vao` chỉ ghi tham số khác mặc định; `ra` là giá trị đúng của đại lượng đo, tính tay từ công thức; `saiSo` là sai số tuyệt đối cho phép. Chọn các dòng ở biên và ở giữa khoảng, không chỉ ở mặc định. |

`mo-hinh.js` gán `THI_NGHIEM_MO_HINH` với hai hàm, thêm `thoiLuong` khi `hoatHinh` là `mot-lan`:

- `tinh(p)`: nhận đối tượng tham số theo mã, trả về đối tượng có mọi mã trong `daiLuongDo`. Hàm thuần: không phụ thuộc thời gian, không dùng số ngẫu nhiên ngoài `THI_NGHIEM_KHUNG.taoNgauNhien(hạt_giống)`, trả `null` cho đại lượng không xác định.
- `ve(ctx, p, t, kt, d)`: vẽ lên `canvas` 2D; `t` là thời gian tính bằng giây, `kt.rong` và `kt.cao` là kích thước khung, `d` là kết quả của `tinh(p)`. Dùng `THI_NGHIEM_KHUNG.MAU` cho màu, `THI_NGHIEM_KHUNG.PHONG` cho font, `THI_NGHIEM_KHUNG.dinhDang(số, chữ_số)` để in số.
- `thoiLuong(p, d)`: số giây của một lần chạy.
- Không dùng `document`, `window`, `fetch`, `import`, `require`, `eval`, địa chỉ web hay thư viện ngoài: công cụ chặn các file như vậy.

Ví dụ đủ khuôn, định luật Hooke. File `mo-hinh.json`:

```json
{
  "ma": "moi",
  "ten": "Định luật Hooke",
  "mon": "Vật lí",
  "hoatHinh": "khong",
  "thamSo": [
    {
      "ma": "do-cung",
      "ten": "Độ cứng k",
      "kieu": "so",
      "donVi": "N/m",
      "min": 10,
      "max": 100,
      "buoc": 10,
      "macDinh": 50
    },
    {
      "ma": "do-gian",
      "ten": "Độ giãn x",
      "kieu": "so",
      "donVi": "m",
      "min": 0,
      "max": 0.2,
      "buoc": 0.01,
      "macDinh": 0.1
    }
  ],
  "daiLuongDo": [
    {
      "ma": "luc",
      "ten": "Lực đàn hồi F",
      "donVi": "N",
      "saiSo": 0.05,
      "chuSo": 2
    }
  ],
  "congThuc": {
    "bieuThuc": "F = k·x",
    "dieuKien": "Lò xo còn trong giới hạn đàn hồi.",
    "nguon": ""
  },
  "bangKiem": [
    {
      "vao": {},
      "ra": {
        "luc": 5.0
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "do-gian": 0
      },
      "ra": {
        "luc": 0.0
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "do-gian": 0.2
      },
      "ra": {
        "luc": 10.0
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "do-cung": 100
      },
      "ra": {
        "luc": 10.0
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "do-cung": 10,
        "do-gian": 0.05
      },
      "ra": {
        "luc": 0.5
      },
      "saiSo": 0.001
    }
  ]
}
```

File `mo-hinh.js`:

```js
(function (root) {
  function tinh(p) { return { 'luc': p['do-cung'] * p['do-gian'] }; }
  function ve(ctx, p, t, kt, d) { ctx.fillRect(20, 20, 40 + 600 * p['do-gian'], 20); }
  root.THI_NGHIEM_MO_HINH = { tinh: tinh, ve: ve };
})(typeof globalThis !== 'undefined' ? globalThis : this);
```

## Khi mô hình do AI viết

- Công cụ chặn mô hình thiếu `congThuc.dieuKien`, thiếu `bangKiem` đủ 5 dòng, hoặc có lệnh mạng (`error.step` là `model`).
- Máy có Node thì công cụ chạy `bangKiem`; trượt là `error.step` `check`. Khi đó sửa hàm `tinh` cho khớp bảng. Chỉ sửa bảng số kiểm khi chính bảng sai, và ghi rõ dòng đã sửa cùng lý do vào tin nhắn báo thầy cô.
- Bảng số kiểm do AI tự tính chỉ bắt được lỗi lập trình, không bắt được lỗi hiểu sai kiến thức. Vì vậy `can-soat.md` luôn nhắc thầy cô soát công thức và tính lại ít nhất hai dòng bằng máy tính cầm tay; AI đọc nguyên văn lời nhắc đó cho thầy cô, không lược bỏ.
- Không viết mô hình cho hiện tượng mà AI không nêu được công thức và điều kiện áp dụng; khi đó nói rõ với thầy cô và đề xuất mẫu gần nhất trong danh mục.
