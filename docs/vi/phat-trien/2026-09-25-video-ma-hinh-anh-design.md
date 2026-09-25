# Thiết kế: Hình ảnh, bàn tay cầm bút và máy quay động cho video giải thích (v6.3.2-vi.10)

Ngày 2026-09-25. Gói này sửa lỗi font tiếng Việt và làm video giải thích (bản v6.3.2-vi.9) trực quan hơn. Chủ repo nhận xét video vi.9 "bị lỗi font tiếng Việt, chỉ có chữ, không có hình ảnh hay chuyển động gì trực quan để học sinh xem". Chủ repo đã chọn font Itim và cả bốn hướng: hình vẽ nét, ảnh thật, bàn tay cầm bút, máy quay động.

Kiểu Vox và khổ dọc 9:16 (trước đây dự kiến là vi.10) lùi sang vi.11.

## 1. Tiêu chí thành công

1. Mọi chữ tiếng Việt trong khung hình và phụ đề được vẽ bằng **một** font duy nhất, không chữ nào lấy từ font khác. Điều này phải được kiểm bằng test cho cả 134 chữ có dấu.
2. Mỗi video có hình minh hoạ được vẽ dần từng nét như bút vẽ trên bảng. Hình lấy từ 5.138 biểu tượng nét `tabler-outline` có sẵn trong repo và không cần mạng.
3. Có loại cảnh ảnh thật: ảnh chụp chuyển động phóng to hoặc lướt chậm, luôn kèm dòng ghi nguồn khi giấy phép yêu cầu.
4. Một bàn tay cầm bút chạy đúng theo nét đang vẽ và chữ đang viết. Khi chuyển sang cảnh mới, bàn tay lau bảng cũ.
5. Máy quay phóng vào phần đang được nói, lia sang phần kế tiếp, rồi thu về toàn cảnh trước khi hết cảnh. Chuyển động mượt vì chụp ở 30 khung/giây.
6. Thời gian dựng video 5 phút không quá 8 phút trên máy chủ repo.
7. Kịch bản vi.9 cũ vẫn dựng được, không phải sửa.
8. Mọi thứ vẫn xác định theo thời điểm: gọi lại `datThoiDiem(t)` cho cùng khung hình.
9. Toàn bộ test lớp Việt xanh; không sửa `skills/` (chỉ đọc biểu tượng), `video.py`, `video_parts/`, `thi_nghiem_parts/`.

## 2. Hiện trạng đã kiểm

- **Lỗi font:** Segoe Print thiếu 92 chữ tiếng Việt dựng sẵn (ơ ư ạ ả ấ ầ ẩ ẫ ậ ắ ằ ẳ ẵ ặ ẹ ẻ ẽ ế ề ể ễ ệ ỉ ị ọ ỏ ố ồ ổ ỗ ộ ớ ờ ở ỡ ợ ụ ủ ứ ừ … và chữ hoa), đo bằng bảng `cmap` của `segoepr.ttf`. Ink Free thiếu 80 chữ, Comic Sans thiếu 92. Chromium lấy từng chữ thiếu từ font khác, nên một từ bị trộn hai kiểu nét ("Chiều" có "ều" khác nét). Phép thử font của vi.9 chỉ dùng `document.fonts.check`, cách này không kiểm độ phủ chữ nên đã nhận sai.
- **Font thay thế:** Itim, Patrick Hand, Pangolin và Mali (Google Fonts, giấy phép SIL OFL 1.1) đều đủ 134 chữ. Chủ repo chọn **Itim**. Giấy phép OFL cho phép đóng gói kèm phần mềm nếu có file giấy phép đi kèm.
- **Phụ đề in:** đang dùng `FontName=Segoe UI` ở cỡ 12, chữ trông sít và đậm.
- **Biểu tượng:** `skills/ppt-master/templates/icons/tabler-outline/*.svg` là nét mảnh `stroke`, `fill="none"`, viewBox 24×24, gồm các phần tử `path`, `circle`, `rect`, `line`, `polyline`, `ellipse`. Mọi phần tử này vẽ dần được bằng `pathLength` và `stroke-dashoffset`. Tên biểu tượng là tiếng Anh.
- **Ảnh:** `skills/ppt-master/scripts/image_search.py "<từ khoá>" --filename x.jpg -o <thư_mục>` tải ảnh có giấy phép mở (Openverse, Wikimedia, không cần khoá) và ghi `image_sources.json`. Mỗi bản ghi có `filename`, `author`, `license`, `provider`, `attribution_required`, `attribution_text`.
- **Tốc độ:** chụp PNG 45 ms/khung trên máy 6 lõi. Chụp 30 khung/giây một luồng mất khoảng 6,8 phút cho video 5 phút; chụp song song là cách giữ thời gian dưới 8 phút.
- **Lịch vi.9:** mỗi cảnh có 0,7 giây dẫn đầu trước khi tiếng bắt đầu; các mục (chữ, nét) có `batDau`, `thoiLuong` cố định; `kiemTran` đo ở trạng thái cuối.

## 3. Quyết định

| # | Quyết định | Lý do |
|---|---|---|
| Q1 | Đóng gói `Itim-Regular.ttf` và `OFL.txt` trong `tools/vi/video_ma_parts/runtime/fonts/`, nhúng vào trang bằng `@font-face` dạng `data:`. Bỏ Segoe Print. | Máy nào cũng hiện giống nhau; trang vẫn tự chứa, không cần mạng |
| Q2 | Phụ đề in dùng chính Itim qua `fontsdir` của bộ lọc `subtitles` (đường dẫn tương đối `.khung/fonts`). Cỡ chữ và khoảng cách chỉnh lại cho dễ đọc. | Cùng font với khung hình; không phụ thuộc font máy |
| Q3 | Hình vẽ nét chỉ lấy từ `tabler-outline`. Tên sai bị bắt ở bước kiểm, kèm tối đa 5 tên gần đúng. | Chỉ bộ này là nét mảnh vẽ dần được; bộ tô đặc không có nét để vẽ |
| Q4 | Thêm trường `hinh:` (một biểu tượng) cho cảnh `tieu-de`, `khai-niem`, `cong-thuc`, `y-tung-y`, và loại cảnh mới `minh-hoa` (1–3 hình, mỗi hình có nhãn). | Các cảnh này còn chỗ cho một cột hình; `minh-hoa` là cảnh mà hình là nội dung chính |
| Q5 | Ảnh thật là loại cảnh mới `anh`, và cũng dùng được ở trường `anh:` thay cho `hinh:`. File ảnh do AI tải về `anh/` bằng `image_search.py`; công cụ dựng không tự lên mạng. | Giữ bộ dựng chạy không cần mạng; tải ảnh là bước riêng, thầy cô xem được ảnh trước |
| Q6 | Ảnh phải có nguồn: lấy từ `anh/image_sources.json`, hoặc trường `nguon:` do thầy cô ghi (ảnh tự chụp). Không có nguồn thì báo lỗi `canh`. | Giấy phép CC BY bắt buộc ghi tác giả; không ghi là vi phạm |
| Q7 | Bàn tay là hình SVG do repo tự vẽ, không lấy ảnh ngoài. Bàn tay chỉ hiện khi đang vẽ hoặc viết; nghỉ thì lướt ra góc phải dưới. | Không vướng giấy phép; xác định theo thời điểm |
| Q8 | Chuyển cảnh "lau bảng" trong 0,5 giây đầu mỗi cảnh (trừ cảnh 1). Nền là khung cuối của cảnh trước; bàn tay cầm giẻ quét từ trái sang phải. Dẫn đầu của mọi cảnh tăng từ 0,7 lên 1,0 giây. | Không cần FFmpeg ghép chồng; thời lượng cảnh vẫn khớp tiếng |
| Q9 | Máy quay là phép `scale` và `translate` trên lớp bảng, tính thuần theo `t` từ danh sách mục. Phóng vào mục đang vẽ tối đa 1,35 lần; thu về 1,0 trong 1,2 giây cuối cảnh; phần đang vẽ luôn nằm trên vạch phụ đề (y ≤ 620 sau biến đổi). | Xác định, test được bằng Node; phụ đề in bằng FFmpeg nên không bị phóng theo |
| Q10 | Chụp 30 khung/giây, chia các cảnh cho tối đa `min(4, số lõi // 2)` tiến trình Chromium chạy song song. Mỗi tiến trình ghi khung vào đúng số thứ tự đã tính từ lịch. | Chuyển động bàn tay và máy quay ở 15 khung/giây bị giật; song song giữ thời gian dựng dưới 8 phút |
| Q11 | Khoá đầu video mới, đều không bắt buộc: `ban-tay: co\|khong` (mặc định `co`), `may-quay: co\|khong` (`co`), `chuyen-canh: lau-bang\|khong` (`lau-bang`). | Thầy cô tắt được khi thấy rối; kịch bản cũ không cần sửa |
| Q12 | Cảnh `thi-nghiem` không có bàn tay; máy quay chỉ đẩy chậm từ 1,0 lên 1,06. | Canvas mô hình đã chuyển động; thêm tay và lia máy sẽ che số đo |

## 4. Kiến trúc và file

```
tools/vi/video_ma_parts/runtime/fonts/Itim-Regular.ttf, OFL.txt     font đóng gói (Q1)
tools/vi/video_ma_parts/runtime/ban-tay.js                          bàn tay: vị trí theo t, lau bảng (Q7, Q8)
tools/vi/video_ma_parts/runtime/may-quay.js                         máy quay theo t (Q9)
tools/vi/video_ma_parts/runtime/hinh.js                             vẽ dần một biểu tượng; ảnh Ken Burns
tools/vi/video_ma_parts/runtime/canh/minh-hoa.js, anh.js            hai loại cảnh mới
tools/vi/video_ma_parts/hinh.py                                     đọc biểu tượng tabler-outline, gợi ý tên gần đúng
tools/vi/video_ma_parts/anh.py                                      đọc ảnh trong anh/, nguồn ảnh, nhúng data:
```

Sửa: `parse.py` (loại cảnh và khoá mới), `kiem.py` (kiểm hình, ảnh, nguồn), `lich.py` (dẫn đầu 1,0 s; FPS 30; dữ liệu cảnh thêm hình, ảnh, nền cảnh trước, cờ), `trang.py` (font, các file runtime mới), `chup.py` (chụp song song theo dải cảnh; khung cuối cảnh trước), `ghep.py` (`fontsdir`, 30 khung/giây vào, 30 ra), `khung-video.js` (lớp bảng cho máy quay; điểm ngòi bút cho bàn tay; font), `viet-tay.css`, sáu file cảnh có cột hình. Tài liệu `docs/vi/tro-ly/video-giai-thich.md`, `canh-video.md` (bảng tra biểu tượng theo khái niệm, loại cảnh mới, cách tải ảnh), `docs/vi/video-giai-thich.md`, `AGENTS.vi.md` §15, luật Antigravity (chỉ khi cần thêm dòng), `CHANGELOG-VI.md`.

## 5. Ngữ pháp mới trong `video.md`

Khối đầu thêm ba khoá tuỳ chọn ở Q11.

Trường `hinh:` là tên một biểu tượng `tabler-outline`, không có tiền tố thư viện, ví dụ `hinh: flask`. Trường `anh:` là tên file trong thư mục `anh/`, ví dụ `anh: con-lac.jpg`. Một cảnh có tối đa một trong hai trường. Cảnh có cột hình thì chữ thu hẹp lại, giới hạn ký tự giữ nguyên, và `kiemTran` vẫn bắt chữ tràn.

```
## Cảnh 3
loai: minh-hoa
tieu-de: Dụng cụ đo chu kì
hinh: clock | Đồng hồ bấm giây
hinh: ruler-measure | Thước đo chiều dài
hinh: weight | Quả nặng
loi: Ta cần đồng hồ bấm giây. Thước đo chiều dài dây. Và một quả nặng.

## Cảnh 4
loai: anh
anh: con-lac-foucault.jpg
chu-thich: Con lắc Foucault ở Paris
loi: Đây là con lắc Foucault, dài 67 mét, dao động rất chậm.
```

| Loại | Trường | Giới hạn | Cách hiện |
|---|---|---|---|
| `minh-hoa` | `tieu-de`, `hinh` (lặp, dạng `tên \| nhãn`) | 1–3 hình; nhãn ≤ 30 | hình thứ k vẽ dần từng nét tại mốc câu k, nhãn viết sau |
| `anh` | `anh`, `chu-thich`, `nguon` (tuỳ chọn) | chú thích ≤ 90; ảnh jpg/png/webp ≤ 8 MB | ảnh hiện trong khung vẽ tay, phóng hoặc lướt chậm suốt cảnh; dòng nguồn nhỏ ở góc |

## 6. Bàn tay, lau bảng, máy quay

- **Ngòi bút:** với nét, điểm ngòi là `getPointAtLength(p × L)` của phần tử đang vẽ. Với chữ, ngòi là góc phải dưới của ký tự vừa hiện: `catDanhDau` chèn một thẻ đánh dấu rỗng tại điểm cắt. Toạ độ đổi sang hệ toạ độ của lớp bảng.
- **Bàn tay theo t:** tại `t`, nếu có mục đang vẽ (0 < p < 1) thì tay ở ngòi của mục đó. Giữa hai mục cách nhau không quá 0,6 giây, tay lướt tuyến tính từ điểm cuối mục trước tới điểm đầu mục sau. Còn lại, tay lướt về góc nghỉ trong 0,4 giây.
- **Lau bảng:** khung PNG cuối của cảnh trước được nhúng làm lớp nền. Trong 0–0,5 s, một mặt nạ quét trái sang phải xoá lớp nền, bàn tay cầm giẻ đi theo mép quét. Nội dung cảnh mới bắt đầu từ 0,5 s.
- **Máy quay theo t:** mỗi mục có hộp bao ở trạng thái cuối, đo một lần sau khi trang nạp. Mục tiêu tại `t` là hộp của mục đang vẽ, hoặc mục vừa vẽ xong gần nhất. Độ phóng Z = min(1,35; 0,6 × 1280 / rộng hộp, 0,6 × 560 / cao hộp), không nhỏ hơn 1. Máy chuyển giữa hai mục tiêu trong 0,6 giây, làm mượt bằng hàm easeInOut. 1,2 giây cuối cảnh thu về Z = 1. Phép dịch được kẹp để khung không lộ ra ngoài bảng và hộp mục tiêu nằm trong y ≤ 620.
- `kiemTran` và hộp bao luôn đo khi máy quay ở Z = 1 và bàn tay ẩn.

## 7. Chụp và ghép

- Lịch cho biết số khung của từng cảnh và số thứ tự khung đầu tiên. Các cảnh được chia thành `N` dải liên tiếp có tổng số khung gần bằng nhau. Mỗi tiến trình con (Python `multiprocessing`, một Chromium riêng) chụp một dải.
- Khung cuối của cảnh k−1 cần cho cảnh k. Mỗi tiến trình tự dựng khung đó bằng cách mở trang cảnh k−1 và chụp tại thời điểm cuối. Như vậy không tiến trình nào phải chờ tiến trình khác.
- Ghép: đầu vào khung 30 khung/giây, đầu ra 30 khung/giây. Tiếng đệm im lặng 1,0 giây đầu cảnh. Phụ đề in bằng Itim qua `fontsdir=.khung/fonts`.
- Lỗi ở bất kỳ tiến trình con nào đều báo `error.step` là `dung`, kèm số cảnh.

## 8. Kiểm thử

- **Font:** một test thuần Python đọc bảng `cmap` của `Itim-Regular.ttf` (bộ đọc nhỏ chỉ dùng thư viện chuẩn) và khẳng định đủ 134 chữ. Một test Chromium đo độ rộng từng chữ có dấu với `font-family: Itim, monospace` và `Itim, serif`; hai độ rộng phải bằng nhau, tức là không chữ nào rơi sang font dự phòng.
- **Node:** vị trí bàn tay và máy quay tại mọi `t` là hàm thuần. Test gồm: xác định; tay ở ngòi khi đang vẽ; tay nghỉ khi không vẽ; Z ∈ [1; 1,35]; Z = 1 ở cuối cảnh; hộp mục tiêu sau biến đổi nằm trong khung và trên y = 620; lau bảng che toàn bộ ở 0 s và hết ở 0,5 s; hình thứ k của `minh-hoa` bắt đầu tại mốc câu k.
- **Python:** đọc và kiểm `hinh` (tên sai kèm gợi ý), `anh` (thiếu file, sai định dạng, quá 8 MB, thiếu nguồn), khoá mới, kịch bản vi.9 cũ vẫn qua; chia dải chụp có tổng khung đúng và không chồng lấn.
- **Chromium:** mọi loại cảnh có hình hoặc ảnh không tràn chữ ở giới hạn tối đa; biểu tượng vẽ dần (khung giữa khác khung cuối); ảnh Ken Burns khác nhau giữa đầu và cuối cảnh; dòng nguồn có mặt.
- **Tích hợp:** dựng video 3 cảnh (một cảnh có hình, một cảnh ảnh, một cảnh thí nghiệm) với tiếng giả, song song 2 tiến trình. `ffprobe` phải cho 1280×720, 30/1, có tiếng, đúng thời lượng.
- **Bằng mắt:** chủ repo xem video demo con lắc đơn dựng lại với giọng thật.

## 9. Ngoài phạm vi

Kiểu Vox, khổ dọc 9:16 (vi.11); nhạc nền; biểu tượng tô màu hay nhiều màu; tự tìm ảnh bên trong công cụ dựng; video AI sinh; nhận diện chữ viết tay thật.

## 10. Rủi ro

- **Tên biểu tượng là tiếng Anh:** AI có thể chọn sai hình. Có bảng tra theo khái niệm trong `canh-video.md` và gợi ý tên gần đúng khi sai.
- **Ảnh tải về có thể sai nội dung.** Hướng dẫn yêu cầu AI chạy `--xem-truoc` và thầy cô xem ảnh trước khi dựng thật.
- **Máy yếu (2 lõi):** chạy một tiến trình, video 5 phút mất khoảng 7 phút.
- **Bàn tay che chữ đang viết:** ngòi bút đặt ở góc dưới phải của ký tự và tay vẽ lệch xuống dưới phải. Chủ repo nghiệm thu bằng mắt.
- **Kích thước trang tăng:** font khoảng 370 KB và ảnh nhúng dạng `data:`. Thời gian nạp trang tăng một lần mỗi cảnh, không phải mỗi khung.
