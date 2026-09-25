# Cảnh video giải thích: danh mục mười loại cảnh

File dành cho AI. Đọc cùng `docs/vi/tro-ly/video-giai-thich.md`. Mỗi cảnh trong `video.md` có `loai:`, `loi:` và các trường của loại cảnh dưới đây; trường không có trong danh sách của loại cảnh là lỗi `parse`.

Quy ước chung:

- Giới hạn tính theo số ký tự hiện ra, không tính dấu quy ước `~`, `^`, `**`. Vượt giới hạn là lỗi `canh`.
- Trường lặp viết nhiều dòng cùng khoá, theo đúng thứ tự muốn hiện.
- Cảnh có danh sách: ý thứ k hiện khi câu thứ k của `loi` bắt đầu; lời ít câu hơn số ý thì các ý hiện đều nhau theo thời lượng giọng.
- Không chèn địa chỉ web vào bất kỳ trường nào.
- Bốn loại `tieu-de`, `khai-niem`, `cong-thuc`, `y-tung-y` có thêm trường tuỳ chọn `hinh` (tên một biểu tượng `tabler-outline`, ví dụ `hinh: flask`) hoặc `anh` (tên file trong `anh/`, ví dụ `anh: con-lac.jpg`); một cảnh chỉ được có một trong hai, không cả hai (có cả hai là lỗi `parse`).
- Cảnh `tieu-de` có hình: hình 180×180 được vẽ dần ở giữa phía trên, rồi tiêu đề viết bên dưới. Ba loại còn lại: hình hoặc ảnh nằm ở cột phải, chữ thu hẹp về bên trái; giới hạn ký tự giữ nguyên, chữ dài thì cỡ chữ nhỏ lại một chút và công cụ vẫn bắt lỗi tràn khung.
- Tên biểu tượng là tiếng Anh, không có tiền tố thư viện: tra ở mục "Bảng tra biểu tượng" cuối file. Viết `tabler-outline/flask`, `Flask` hay `flask.svg` vẫn được nhận; tên sai hay tên tiếng Việt là lỗi `canh` kèm tối đa 5 tên gần đúng.
- Ảnh thật do AI tải về `anh/` của thư mục video bằng `image_search.py` trước khi dựng: xem mục "Ảnh thật".
- Mỗi hình được vẽ dần từng nét như bút vẽ trên bảng; bàn tay cầm bút đi theo nét và chữ đang viết, máy quay phóng vào phần đang nói rồi thu về toàn cảnh trước khi hết cảnh. Tắt được bằng khoá đầu `ban-tay`, `may-quay`, `chuyen-canh` (xem `docs/vi/tro-ly/video-giai-thich.md`).

## Tiêu đề

Mã loại: `tieu-de`.

| Trường | Bắt buộc | Giới hạn |
|---|---|---|
| `chu` | có | 90 ký tự |
| `phu` | không | 90 ký tự |
| `hinh` | không | tên biểu tượng `tabler-outline` |
| `anh` | không | tên file trong `anh/` |

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
| `hinh` | không | tên biểu tượng `tabler-outline` |
| `anh` | không | tên file trong `anh/` |

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
| `hinh` | không | tên biểu tượng `tabler-outline` |
| `anh` | không | tên file trong `anh/` |

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
| `hinh` | không | tên biểu tượng `tabler-outline` |
| `anh` | không | tên file trong `anh/` |

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
| `do` | không | tối đa 3 mã đại lượng đo, cách nhau bằng dấu phẩy; không ghi dòng `do` thì hiện hai đại lượng đầu của mẫu |
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

## Minh hoạ

Mã loại: `minh-hoa`.

| Trường | Bắt buộc | Giới hạn |
|---|---|---|
| `tieu-de` | có | 90 ký tự |
| `hinh` | lặp 1–3 dòng | dạng `tên \| nhãn`, tên là biểu tượng `tabler-outline`, nhãn tối đa 30 ký tự |

Cách hiện: tiêu đề viết trước; hình thứ k được vẽ dần từng nét tại mốc câu thứ k của `loi`, nhãn hiện ngay sau khi hình vẽ xong. Các hình dàn đều theo chiều ngang. Viết lời có đúng một câu cho mỗi hình, theo đúng thứ tự. Thiếu dấu `|` hoặc thiếu nhãn là lỗi `canh`.

```
## Cảnh 9
loai: minh-hoa
tieu-de: Dụng cụ đo chu kì
hinh: clock | Đồng hồ bấm giây
hinh: ruler-measure | Thước đo chiều dài
hinh: weight | Quả nặng
loi: Ta cần đồng hồ bấm giây. Thước đo chiều dài dây. Và một quả nặng.
```

## Ảnh thật

Mã loại: `anh`.

| Trường | Bắt buộc | Giới hạn |
|---|---|---|
| `anh` | có | tên file `.jpg`/`.jpeg`/`.png`/`.webp` trong `anh/`, tối đa 8 MB |
| `chu-thich` | có | 90 ký tự |
| `nguon` | không | dòng ghi nguồn; không ghi thì lấy từ bản ghi cùng tên trong `anh/image_sources.json` (tải bằng `image_search.py`); không có nguồn nào là lỗi `canh` |

Cách hiện: ảnh hiện trong khung vẽ tay, vừa khung và không méo (ảnh dọc hay ngang đều được), phóng hoặc lướt chậm suốt cảnh; dòng nguồn nhỏ ở góc dưới; chú thích được viết dưới ảnh, chữ dài thì cỡ chữ nhỏ lại. Ảnh ở cột phải của bốn loại cảnh trên thì dòng nguồn nằm ngay dưới khung ảnh.

```
## Cảnh 10
loai: anh
anh: con-lac-foucault.jpg
chu-thich: Con lắc Foucault ở Paris
loi: Đây là con lắc Foucault, dài 67 mét, dao động rất chậm.
```

Trường `nguon` chỉ có ở loại cảnh `anh`. Ảnh ở trường `anh:` của bốn loại cảnh kia luôn lấy nguồn từ `anh/image_sources.json`.

Ảnh thầy cô tự chụp: chép file vào `anh/`. Ở cảnh `anh`, ghi `nguon:` trong cảnh, ví dụ `nguon: Ảnh: cô Lan chụp`. Ở cột phải của bốn loại cảnh kia, thêm vào danh sách `items` của `anh/image_sources.json` (chưa có file thì tạo) một bản ghi dạng `{"filename": "<tên file>", "author": "cô Lan", "license_name": "Ảnh tự chụp", "provider": "Trường THPT ..."}`.

Tải ảnh từ kho ảnh mở (Openverse, Wikimedia, không cần khoá), mỗi ảnh một lệnh, sau khi thầy cô đã duyệt danh sách cảnh:

```
python skills\ppt-master\scripts\image_search.py "Foucault pendulum" --filename con-lac-foucault.jpg --orientation landscape -o projects\_video\<tên_video>\anh
```

- Từ khoá viết bằng tiếng Anh, cụ thể (tên sự vật, địa danh); `--orientation` là `landscape` cho cảnh `anh`, `portrait` cho ảnh ở cột phải.
- Lệnh ghi ảnh vào `anh/` và thêm bản ghi tác giả, giấy phép vào `anh/image_sources.json`; công cụ dựng lấy dòng nguồn từ đó nên không cần ghi `nguon:`.
- Lệnh còn ghi một bản thu nhỏ để xem vào `anh/.review/<tên>.jpg` (cạnh dài 1024 px). Mở bản này xem ảnh có đúng nội dung không; sai thì tải lại với từ khoá khác.
- Ảnh gốc nặng quá 8 MB là lỗi `canh`. Khi đó chép `anh/.review/<tên>.jpg` đè lên file gốc cùng tên (tên gốc có đuôi `.jpg`); tên gốc có đuôi khác thì chép bản thu nhỏ thành `anh/<tên>.jpg`, sửa dòng `anh:` theo tên mới, và thêm vào `items` của `anh/image_sources.json` một bản ghi chép nguyên bản ghi của file gốc, chỉ đổi `filename` thành tên mới. Không có bản `.review` thì chọn ảnh khác.
- Công cụ dựng không bao giờ tự lên mạng tìm ảnh; không chèn địa chỉ web vào `anh:` hay `nguon:`.

## Bảng tra biểu tượng

Tên ở cột cuối là tên file (bỏ đuôi `.svg`) trong `skills/ppt-master/templates/icons/tabler-outline/`; viết đúng tên đó sau `hinh:`. Chọn hình nói đúng sự vật trong lời, không chọn hình chỉ để trang trí. Khái niệm không có trong bảng thì tìm thêm theo từ khoá tiếng Anh, ví dụ:

```
rg --files skills/ppt-master/templates/icons/tabler-outline -g "*flask*"
```

| Môn | Khái niệm | Tên |
|---|---|---|
| Toán | hàm số | `math-function` |
| Toán | đồ thị đường | `chart-line` |
| Toán | biểu đồ cột | `chart-bar` |
| Toán | biểu đồ tròn, tỉ lệ | `chart-pie` |
| Toán | tập điểm, số liệu | `chart-dots` |
| Toán | phần trăm | `percentage` |
| Toán | máy tính cầm tay | `calculator` |
| Toán | tích phân | `math-integral` |
| Toán | số pi | `math-pi` |
| Toán | căn bậc hai | `square-root` |
| Toán | vô cực | `infinity` |
| Toán | tổng, dấu sigma | `sum` |
| Toán | phép chia | `divide` |
| Toán | phân số | `math-x-divide-y` |
| Toán | lũy thừa, số mũ | `superscript` |
| Toán | tam giác | `triangle` |
| Toán | góc | `angle` |
| Toán | hình vuông | `square` |
| Toán | đường tròn | `circle` |
| Toán | khối lập phương | `cube` |
| Toán | hình cầu | `sphere` |
| Toán | hình trụ | `cylinder` |
| Toán | hình nón | `cone` |
| Toán | vectơ | `vector` |
| Toán | trục hoành | `axis-x` |
| Toán | trục tung | `axis-y` |
| Toán | xác suất, xúc xắc | `dice-5` |
| Toán | thống kê, kiểm đếm | `tallymarks` |
| Toán | thước kẻ | `ruler` |
| Vật lí | nguyên tử | `atom` |
| Vật lí | nam châm | `magnet` |
| Vật lí | điện, tia sét | `bolt` |
| Vật lí | pin, acquy | `battery` |
| Vật lí | bóng đèn, ý tưởng | `bulb` |
| Vật lí | phích cắm, nguồn điện | `plug` |
| Vật lí | sóng hình sin, dao động | `wave-sine` |
| Vật lí | sóng trên mặt nước | `ripple` |
| Vật lí | âm thanh, loa | `volume` |
| Vật lí | tai, nghe | `ear` |
| Vật lí | mắt, nhìn | `eye` |
| Vật lí | lăng kính | `prism` |
| Vật lí | tán sắc, cầu vồng | `rainbow` |
| Vật lí | nhiệt kế | `thermometer` |
| Vật lí | nhiệt độ | `temperature` |
| Vật lí | đồng hồ | `clock` |
| Vật lí | đồng hồ bấm giây | `stopwatch` |
| Vật lí | đồng hồ cát | `hourglass` |
| Vật lí | quả nặng, khối lượng | `weight` |
| Vật lí | cân, cân bằng | `scale` |
| Vật lí | thước đo chiều dài | `ruler-measure` |
| Vật lí | trọng lực, rơi xuống | `arrow-big-down-lines` |
| Vật lí | chuyển động nảy | `bounce-right` |
| Vật lí | áp kế, đồng hồ đo | `gauge` |
| Vật lí | tên lửa | `rocket` |
| Vật lí | hành tinh | `planet` |
| Vật lí | Mặt Trời | `sun` |
| Vật lí | Mặt Trăng | `moon` |
| Vật lí | vệ tinh | `satellite` |
| Vật lí | kính thiên văn | `telescope` |
| Vật lí | phóng xạ | `radioactive` |
| Vật lí | sóng vô tuyến | `wifi` |
| Vật lí | ô tô | `car` |
| Vật lí | xe đạp | `bike` |
| Vật lí | thuyền buồm | `sailboat` |
| Hoá học | bình tam giác | `flask` |
| Hoá học | bình cầu | `flask-2` |
| Hoá học | ống nghiệm | `test-pipe` |
| Hoá học | ống nghiệm (kiểu 2) | `test-pipe-2` |
| Hoá học | ngọn lửa, đốt cháy | `flame` |
| Hoá học | giọt chất lỏng | `droplet` |
| Hoá học | nhiều giọt, dung dịch | `droplets` |
| Hoá học | lọc, phễu lọc | `filter` |
| Hoá học | chai hoá chất | `bottle` |
| Hoá học | muối | `salt` |
| Hoá học | mô hình nguyên tử | `atom-2` |
| Hoá học | bình xịt | `spray` |
| Hoá học | bình chữa cháy | `fire-extinguisher` |
| Hoá học | tái chế | `recycle` |
| Hoá học | nhà máy hoá chất | `building-factory` |
| Sinh học | tế bào | `cell` |
| Sinh học | ADN | `dna` |
| Sinh học | ADN (kiểu 2) | `dna-2` |
| Sinh học | kính hiển vi | `microscope` |
| Sinh học | lá cây | `leaf` |
| Sinh học | cây non, nảy mầm | `seedling` |
| Sinh học | cây | `plant` |
| Sinh học | cây gỗ | `tree` |
| Sinh học | rừng | `trees` |
| Sinh học | hoa | `flower` |
| Sinh học | xương rồng | `cactus` |
| Sinh học | tim | `heart` |
| Sinh học | nhịp tim | `heartbeat` |
| Sinh học | phổi | `lungs` |
| Sinh học | não | `brain` |
| Sinh học | xương | `bone` |
| Sinh học | răng | `dental` |
| Sinh học | virus | `virus` |
| Sinh học | côn trùng | `bug` |
| Sinh học | bướm | `butterfly` |
| Sinh học | cá | `fish` |
| Sinh học | chim, lông vũ | `feather` |
| Sinh học | động vật, dấu chân | `paw` |
| Sinh học | chó | `dog` |
| Sinh học | mèo | `cat` |
| Sinh học | nấm | `mushroom` |
| Sinh học | quả táo, trái cây | `apple` |
| Sinh học | trứng | `egg` |
| Sinh học | thuốc | `pill` |
| Sinh học | vắc xin | `vaccine` |
| Sinh học | sơ cứu | `first-aid-kit` |
| Địa lí | quả địa cầu | `globe` |
| Địa lí | thế giới | `world` |
| Địa lí | bản đồ | `map` |
| Địa lí | vị trí trên bản đồ | `map-pin` |
| Địa lí | toạ độ | `current-location` |
| Địa lí | la bàn, phương hướng | `compass` |
| Địa lí | núi | `mountain` |
| Địa lí | núi lửa | `volcano` |
| Địa lí | hải đăng, bờ biển | `building-lighthouse` |
| Địa lí | mây | `cloud` |
| Địa lí | mưa | `cloud-rain` |
| Địa lí | giông bão | `cloud-storm` |
| Địa lí | sương mù | `cloud-fog` |
| Địa lí | tuyết, băng giá | `snowflake` |
| Địa lí | gió | `wind` |
| Địa lí | lốc xoáy | `tornado` |
| Địa lí | bình minh | `sunrise` |
| Địa lí | hoàng hôn | `sunset` |
| Địa lí | mùa mưa | `umbrella` |
| Địa lí | nông nghiệp, máy kéo | `tractor` |
| Địa lí | lương thực, lúa | `wheat` |
| Địa lí | khu dân cư, đô thị | `building-community` |
| Địa lí | công nghiệp | `building-factory-2` |
| Chung | sách | `book` |
| Chung | nhiều sách, thư viện | `books` |
| Chung | vở ghi | `notebook` |
| Chung | bút chì | `pencil` |
| Chung | viết bài | `writing` |
| Chung | trường học | `school` |
| Chung | nhóm học sinh | `users` |
| Chung | một người | `user` |
| Chung | câu hỏi | `question-mark` |
| Chung | mục tiêu | `target` |
| Chung | danh sách việc | `clipboard-list` |
| Chung | kiểm tra từng mục | `list-check` |
| Chung | lịch | `calendar` |
| Chung | báo thức | `alarm` |
| Chung | cờ, đích đến | `flag` |
| Chung | cúp, thành tích | `trophy` |
| Chung | ghép hình, giải đố | `puzzle` |
| Chung | bài trình chiếu | `presentation` |
| Chung | trò chuyện, thảo luận | `message-circle` |
| Chung | tìm kiếm | `search` |
| Chung | thông tin | `info-circle` |
| Chung | cảnh báo, nguy hiểm | `alert-triangle` |
| Chung | đúng | `check` |
| Chung | sai | `x` |
| Chung | mũi tên, dẫn tới | `arrow-right` |
| Chung | chu trình, lặp lại | `refresh` |
| Chung | tăng trưởng | `trending-up` |
| Chung | nhà | `home` |
