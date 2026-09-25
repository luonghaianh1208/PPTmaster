# Cảnh video giải thích: danh mục tám loại cảnh

File dành cho AI. Đọc cùng `docs/vi/tro-ly/video-giai-thich.md`. Mỗi cảnh trong `video.md` có `loai:`, `loi:` và các trường của loại cảnh dưới đây; trường không có trong danh sách của loại cảnh là lỗi `parse`.

Quy ước chung:

- Giới hạn tính theo số ký tự hiện ra, không tính dấu quy ước `~`, `^`, `**`. Vượt giới hạn là lỗi `canh`.
- Trường lặp viết nhiều dòng cùng khoá, theo đúng thứ tự muốn hiện.
- Cảnh có danh sách: ý thứ k hiện khi câu thứ k của `loi` bắt đầu; lời ít câu hơn số ý thì các ý hiện đều nhau theo thời lượng giọng.
- Không chèn địa chỉ web vào bất kỳ trường nào.

## Tiêu đề

Mã loại: `tieu-de`.

| Trường | Bắt buộc | Giới hạn |
|---|---|---|
| `chu` | có | 90 ký tự |
| `phu` | không | 90 ký tự |

Cách hiện: chữ lớn được viết ra giữa khung, dòng phụ hiện bên dưới. Dùng cho cảnh mở đầu hoặc mở một phần mới.

```
## Cảnh 1
loai: tieu-de
chu: Con lắc đơn và chu kì dao động
phu: Vật lí 11
loi: Hôm nay chúng ta tìm hiểu chu kì của con lắc đơn phụ thuộc vào yếu tố nào.
```

## Khái niệm

Mã loại: `khai-niem`.

| Trường | Bắt buộc | Giới hạn |
|---|---|---|
| `thuat-ngu` | có | 60 ký tự |
| `dinh-nghia` | có | 220 ký tự |

Cách hiện: khung được vẽ nét, thuật ngữ rồi định nghĩa được viết vào trong khung.

```
## Cảnh 2
loai: khai-niem
thuat-ngu: Chu kì T
dinh-nghia: Khoảng thời gian ngắn nhất để con lắc thực hiện một dao động toàn phần, đo bằng giây.
loi: Chu kì là khoảng thời gian ngắn nhất để con lắc thực hiện một dao động toàn phần.
```

## Công thức

Mã loại: `cong-thuc`.

| Trường | Bắt buộc | Giới hạn |
|---|---|---|
| `bieu-thuc` | có | 90 ký tự |
| `giai-thich` | lặp 0–4 dòng | mỗi dòng 60 ký tự |

Cách hiện: biểu thức được viết dần, các dòng giải thích hiện lần lượt theo từng câu của lời.

```
## Cảnh 3
loai: cong-thuc
bieu-thuc: T = 2π√(l/g)
giai-thich: l là chiều dài dây, đơn vị mét
giai-thich: g là gia tốc trọng trường, khoảng 9,8 m/s^2^
loi: Chu kì bằng hai pi nhân căn của l chia g. Chữ l là chiều dài dây. Chữ g là gia tốc trọng trường.
```

## Ý từng ý

Mã loại: `y-tung-y`.

| Trường | Bắt buộc | Giới hạn |
|---|---|---|
| `tieu-de` | có | 90 ký tự |
| `y` | lặp 1–6 dòng | mỗi ý 60 ký tự |

Cách hiện: tiêu đề viết trước, mỗi ý được viết ra khi lời nói tới.

```
## Cảnh 4
loai: y-tung-y
tieu-de: Phản ứng trao đổi ion xảy ra khi
y: Tạo thành chất kết tủa
y: Tạo thành chất khí
y: Tạo thành chất điện li yếu
loi: Thứ nhất, sản phẩm có chất kết tủa. Thứ hai, sản phẩm có chất khí. Thứ ba, sản phẩm có chất điện li yếu.
```

## Quy trình

Mã loại: `quy-trinh`.

| Trường | Bắt buộc | Giới hạn |
|---|---|---|
| `tieu-de` | có | 90 ký tự |
| `buoc` | lặp 2–5 dòng | mỗi bước 50 ký tự |

Cách hiện: các bước và mũi tên nối được vẽ nối tiếp nhau.

```
## Cảnh 5
loai: quy-trinh
tieu-de: Cách đo chu kì
buoc: Treo quả nặng vào dây
buoc: Kéo lệch một góc nhỏ rồi thả
buoc: Đo thời gian 10 dao động
loi: Bước một, treo quả nặng vào dây. Bước hai, kéo lệch một góc nhỏ rồi thả. Bước ba, đo thời gian mười dao động rồi chia cho mười.
```

## So sánh

Mã loại: `so-sanh`.

| Trường | Bắt buộc | Giới hạn |
|---|---|---|
| `tieu-de` | có | 90 ký tự |
| `trai`, `phai` | có | tên cột, mỗi tên 24 ký tự |
| `y-trai` | lặp 1–4 dòng | mỗi ý 60 ký tự |
| `y-phai` | lặp 1–4 dòng | mỗi ý 60 ký tự |

Cách hiện: hai cột được viết lần lượt, cột trái trước rồi cột phải; các ý hiện theo thứ tự câu của lời (hết ý cột trái mới tới ý cột phải).

```
## Cảnh 6
loai: so-sanh
tieu-de: Yếu tố nào ảnh hưởng đến chu kì
trai: Có ảnh hưởng
phai: Không ảnh hưởng
y-trai: Chiều dài dây
y-trai: Gia tốc trọng trường
y-phai: Khối lượng quả nặng
loi: Chiều dài dây ảnh hưởng đến chu kì. Gia tốc trọng trường cũng vậy. Còn khối lượng quả nặng thì không.
```

## Đồ thị

Mã loại: `do-thi`.

| Trường | Bắt buộc | Giới hạn |
|---|---|---|
| `tieu-de` | có | 90 ký tự |
| `truc-ngang`, `truc-doc` | có | tên trục, mỗi tên 40 ký tự |
| `diem` | lặp 2–12 dòng | dạng `x, y`, hai số, dấu thập phân là dấu chấm |

Cách hiện: hai trục được vẽ trước, rồi từng điểm hiện dần và được nối thành đường.

```
## Cảnh 7
loai: do-thi
tieu-de: Chu kì theo chiều dài dây
truc-ngang: Chiều dài l (m)
truc-doc: Chu kì T (s)
diem: 0.25, 1.0
diem: 0.5, 1.42
diem: 1.0, 2.01
diem: 2.0, 2.84
loi: Khi chiều dài tăng từ một phần tư mét đến hai mét, chu kì tăng từ một giây đến gần ba giây.
```

## Thí nghiệm ảo

Mã loại: `thi-nghiem`.

| Trường | Bắt buộc | Giới hạn |
|---|---|---|
| `mau` | có | mã một trong tám mẫu dưới đây |
| `do` | không | tối đa 3 mã đại lượng đo, cách nhau bằng dấu phẩy; bỏ trống thì hiện hai đại lượng đầu của mẫu |
| `tham-so` | lặp | dạng `<giây> <mã> <giá trị>`, tối đa 3 mã tham số khác nhau |

Tám mã mẫu:

| Môn | Mã | Thí nghiệm |
|---|---|---|
| Vật lí | `li-nem-xien` | Chuyển động ném xiên |
| Vật lí | `li-con-lac-don` | Con lắc đơn |
| Vật lí | `li-mach-ohm` | Đoạn mạch nối tiếp và song song |
| Hoá học | `hoa-chuan-do` | Chuẩn độ acid – base |
| Hoá học | `hoa-can-bang-no2` | Cân bằng N~2~O~4~ ⇌ 2NO~2~ |
| Hoá học | `hoa-toc-do` | Tốc độ phản ứng |
| Toán | `toan-ham-so` | Khảo sát hàm số |
| Toán | `toan-xac-suat` | Xác suất thực nghiệm |

Mã tham số, khoảng cho phép và mã đại lượng đo của từng mẫu nằm ở `docs/vi/tro-ly/mo-hinh-thi-nghiem.md` (mục "Danh mục mẫu"). Video chỉ dùng mẫu có sẵn, không dùng `mau: moi`.

Cách dùng `tham-so: <giây> <mã> <giá trị>`:

- `<giây>` tính từ đầu cảnh. Giá trị của một tham số tại mỗi thời điểm là nội suy tuyến tính giữa hai mốc liền kề của chính tham số đó.
- Trước mốc đầu, tham số giữ giá trị mặc định của mẫu; sau mốc cuối, giữ giá trị của mốc cuối. Tham số không có dòng nào thì giữ mặc định suốt cảnh.
- Giá trị phải nằm trong khoảng cho phép của mẫu (sai là lỗi `parse` nêu đúng dòng). Chỉ dùng được tham số dạng số; tham số lựa chọn (ví dụ `loai-acid`, `kieu-mac`) giữ mặc định.
- Mốc giây vượt thời lượng cảnh là lỗi `canh`. Thời lượng cảnh xấp xỉ thời gian đọc lời cộng khoảng 1,3 giây, nên đặt mốc cuối trước khi lời kết thúc.
- Hai mẫu chạy một lần là `li-nem-xien` và `hoa-toc-do`: mỗi mốc `tham-so` (của bất kỳ tham số nào) làm chuyển động chạy lại từ đầu. Nên đặt mỗi mốc sau khi chuyển động trước đã dừng, để học sinh xem trọn một lần rồi mới thấy lần mới.

Cách hiện: hình vẽ của mẫu chạy trong khung, tham số đổi dần theo các mốc, các đại lượng trong `do` hiện thành số bên cạnh và luôn khớp công thức của mẫu.

```
## Cảnh 8
loai: thi-nghiem
mau: li-con-lac-don
tham-so: 0 chieu-dai 0.4
tham-so: 6 chieu-dai 1.6
do: chu-ki
loi: Hãy quan sát. Khi ta tăng chiều dài dây, chu kì dao động tăng theo.
```
