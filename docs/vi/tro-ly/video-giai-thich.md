# Loại việc: Video giải thích dựng bằng mã

File dành cho AI. Luôn đọc `docs/vi/tro-ly/quy-trinh-hoi.md` trước file này, và đọc `docs/vi/tro-ly/canh-video.md` trước khi viết cảnh. Đầu ra của loại việc này là một video MP4 kiểu viết tay (chữ và hình được viết dần ra theo lời đọc), không phải PPTX.

## Khi nào dùng

Thầy cô muốn một video giải thích một bài học từ nội dung chữ: video viết tay, video whiteboard, video hoạt hình chữ. Công cụ vẽ từng cảnh bằng mã (tiêu đề, khái niệm, công thức, ý từng ý, quy trình, so sánh, đồ thị, thí nghiệm ảo), đọc lời bằng giọng máy tiếng Việt hoặc giọng thầy cô thu sẵn, rồi ghép thành video có phụ đề.

Ví dụ câu lệnh:
- "Làm video giải thích bài Con lắc đơn bằng kiểu viết tay"
- "Làm video viết tay giải thích phản ứng trao đổi ion"
- "Làm video whiteboard về hàm số bậc ba cho Toán 12"

Khác loại việc "Video bài giảng": loại đó dựng video từ một bài giảng slide đã có. Câu lệnh chỉ có "làm video" hoặc "xuất video" (không nói rõ từ slide hay video giải thích) thì hỏi đúng một câu trước khi làm gì khác: "Thầy cô muốn làm video từ bài giảng slide đã có, hay dựng video giải thích mới từ nội dung chữ?". Trả lời slide thì dùng video-bai-giang.md; trả lời video mới thì dùng file này.

Không dùng loại việc này cho video do AI sinh hình (Veo, Flow) hay gói câu lệnh cho các công cụ đó. Bản này chỉ có một phong cách `viet-tay`, khổ ngang 16:9; chưa có kiểu Vox, chưa có khổ dọc, không chèn ảnh hay video tư liệu, không có nhạc nền.

## Câu hỏi bắt buộc

1. Môn, lớp, tên bài và nội dung cần giải thích?
   Gợi ý: thầy cô dán nội dung hoặc dàn ý (vài gạch đầu dòng là đủ); em chia thành cảnh để thầy cô duyệt.
2. Sau video, học sinh cần hiểu hoặc làm được gì?
   Gợi ý: em đề xuất một câu theo yêu cầu cần đạt của bài để thầy cô sửa.
3. Video dài khoảng bao lâu?
   Gợi ý: 2–4 phút, tức 6–10 cảnh.
4. Giọng đọc nam hay nữ, tốc độ chậm, vừa hay nhanh?
   Gợi ý: giọng nữ, tốc độ vừa.
5. Có muốn chèn một cảnh thí nghiệm ảo không? Có sẵn 8 mẫu: con lắc đơn, ném xiên, mạch điện nối tiếp và song song (Vật lí); chuẩn độ acid – base, cân bằng N~2~O~4~ ⇌ 2NO~2~, tốc độ phản ứng (Hoá học); khảo sát hàm số, xác suất thực nghiệm (Toán).
   Gợi ý: có, nếu bài trùng một mẫu; không thì bỏ qua.
6. Phụ đề in lên hình hay để thành file riêng?
   Gợi ý: in lên hình.
7. Thầy cô có file giọng thu sẵn cho từng cảnh không?
   Gợi ý: không, dùng giọng máy (cần mạng khi dựng).

## Câu hỏi tuỳ chọn

- Có ý nào cần nhấn mạnh, nhắc lại ở cảnh cuối không? Chỉ hỏi khi video dài trên 4 phút.
- Tên video hiện ở cảnh mở đầu? Chỉ hỏi khi thầy cô muốn khác tên bài.
- Học sinh xem là lớp đại trà, lớp chuyên hay học sinh cần hỗ trợ? Chỉ hỏi khi nội dung có nhiều mức khó.

## Tạo nhanh

1. Môn, lớp, tên bài và nội dung cần giải thích (câu hỏi bắt buộc 1).
2. Video dài khoảng bao lâu (câu hỏi bắt buộc 3).

Tạo nhanh vẫn ghi brief; các câu không hỏi thì dùng gợi ý và ghi kèm "(AI đề xuất, chưa duyệt)".

## Cấu trúc video.md

Viết một file `video.md` trong `projects/_video/<tên_video>/`. File mở đầu bằng khối thông tin giữa hai dòng `---`, rồi mỗi cảnh một mục `## Cảnh N`, đánh số liên tiếp từ 1.

Khối thông tin:

| Khoá | Bắt buộc | Giá trị |
|---|---|---|
| `tieu-de` | có | tên video |
| `mon`, `lop` | có | môn và lớp |
| `phong-cach` | không | `viet-tay` (mặc định, duy nhất ở bản này) |
| `giong` | không | `nu` (mặc định) hoặc `nam` |
| `toc-do` | không | `cham`, `vua` (mặc định) hoặc `nhanh` |
| `phu-de` | không | `hinh` (mặc định, in lên hình), `file` (file riêng) hoặc `khong` |

Mỗi cảnh:

- Dòng `loai:` là một trong tám loại cảnh ở `docs/vi/tro-ly/canh-video.md`.
- Dòng `loi:` là lời đọc của cảnh, viết trên một dòng, ít nhất một câu.
- Các trường của loại cảnh đó, mỗi dòng dạng `khoá: giá trị`. Trường lặp (ví dụ `y:`, `buoc:`) thì viết nhiều dòng cùng khoá.

Giới hạn chữ (đếm ký tự hiện ra, không tính dấu quy ước): tiêu đề cảnh tối đa 90; `khai-niem` thuật ngữ 60, định nghĩa 220; `cong-thuc` biểu thức 90, mỗi giải thích 60; mỗi ý của `y-tung-y` và `so-sanh` 60; mỗi bước của `quy-trinh` 50; tên cột `trai`, `phai` 24; tên trục 40. Vượt giới hạn là lỗi `canh`. Lời dài quá 700 ký tự chỉ bị cảnh báo, nhưng nên tách cảnh.

Quy ước chữ: `H~2~SO~4~` cho chỉ số dưới, `m/s^2^` cho chỉ số trên, `**in đậm**`. Không chèn địa chỉ web vào file: công cụ báo lỗi `parse`.

Cách chia bài thành cảnh:

- Mỗi cảnh một ý. Lời 1–4 câu; câu đầu nêu ý của cảnh.
- Cảnh có danh sách (`y-tung-y`, `quy-trinh`, `so-sanh`, giải thích của `cong-thuc`): ý thứ k hiện ra khi câu thứ k của lời bắt đầu, nên viết mỗi ý ứng với đúng một câu, theo đúng thứ tự.
- Viết lời để đọc thành tiếng: số và kí hiệu viết thành chữ khi cần ("hai pi", "mười dao động"), không viết kí hiệu mà giọng máy đọc sai.
- Mở đầu bằng một cảnh `tieu-de`; cảnh cuối nên tóm tắt điều học sinh cần nhớ.

Giọng thu sẵn: thầy cô đưa file MP3 thì đặt vào `giong/canh-1.mp3`, `giong/canh-2.mp3`… trong thư mục video, đúng số cảnh. Cảnh có file thì dùng file đó và không cần mạng; công cụ không bao giờ ghi đè file thầy cô đặt. Cảnh không có file thì tạo bằng giọng máy (cần mạng). File giọng thu sẵn không có mốc câu nên các ý hiện theo ước lượng và có cảnh báo "mốc câu ước lượng".

Ví dụ (Hoá học, bốn cảnh, có một cảnh thí nghiệm ảo):

```
---
tieu-de: Chuẩn độ acid – base
mon: Hoá học
lop: 11
giong: nu
toc-do: vua
phu-de: hinh
---

## Cảnh 1
loai: tieu-de
chu: Chuẩn độ acid – base
phu: Hoá học 11
loi: Hôm nay chúng ta tìm hiểu cách xác định nồng độ một dung dịch acid bằng phép chuẩn độ.

## Cảnh 2
loai: khai-niem
thuat-ngu: Điểm tương đương
dinh-nghia: Thời điểm lượng NaOH thêm vào vừa đủ phản ứng hết với lượng acid có trong bình.
loi: Điểm tương đương là lúc lượng NaOH thêm vào vừa đủ phản ứng hết với acid trong bình.

## Cảnh 3
loai: y-tung-y
tieu-de: Ba bước chuẩn độ
y: Lấy chính xác thể tích acid vào bình tam giác
y: Nhỏ từ từ dung dịch NaOH từ burette
y: Dừng khi chất chỉ thị vừa đổi màu
loi: Bước một, lấy chính xác thể tích acid vào bình tam giác. Bước hai, nhỏ từ từ dung dịch NaOH từ burette. Bước ba, dừng lại khi chất chỉ thị vừa đổi màu.

## Cảnh 4
loai: thi-nghiem
mau: hoa-chuan-do
tham-so: 1 the-tich-base 0
tham-so: 6 the-tich-base 30
do: ph
loi: Hãy quan sát pH khi nhỏ thêm NaOH. Gần điểm tương đương, pH tăng vọt rất nhanh.
```

## Đầu ra

Thư mục của mỗi video là `projects\_video\<tên_video>\`, chứa `video.md`. Có `venv\Scripts\python.exe` ở thư mục gốc repo thì dùng nó thay cho `python` trong các lệnh dưới.

1. Kiểm cú pháp và giới hạn: `python tools\vi\video_ma.py projects\_video\<tên_video> --plan-only`.
2. Dựng thử: `python tools\vi\video_ma.py projects\_video\<tên_video> --xem-truoc`. Công cụ chụp mỗi cảnh một ảnh ở trạng thái cuối cảnh vào thư mục `xem-truoc`, không tạo giọng, không ghép. Nên chạy trước khi dựng thật và mở xem từng ảnh: chữ chồng lên nhau hay tràn khung thì rút gọn nội dung hoặc chia cảnh rồi chạy lại.
3. Dựng thật: `python tools\vi\video_ma.py projects\_video\<tên_video>`. Dựng thật mất vài phút (video 5 phút cần khoảng 4 phút chụp khung, chưa kể tạo giọng); báo trước thầy cô một dòng.

Công cụ ghi vào chính thư mục đó:

- `video.mp4`: video 1280×720.
- `phu-de.srt`: chỉ có khi `phu-de: file`.
- `giong\canh-N.mp3`: giọng từng cảnh. Lần dựng sau, cảnh có lời không đổi dùng lại giọng cũ; chỉ cảnh bị sửa lời mới tạo giọng lại.

Đọc dòng JSON ở stdout. `ready` là `true` thì báo thầy cô đường dẫn `video.mp4`, số cảnh, thời lượng (`thoi_luong_giay`), nguồn giọng (`giong`: `may` là giọng máy, `co-san` là file thầy cô đưa, `hon-hop` là cả hai), nơi để phụ đề, và đọc nguyên văn các dòng `warnings`.

| `error.step` | Xử lý |
|---|---|
| `input` | Chưa có thư mục hoặc `video.md`, hoặc sai tham số lệnh: viết file rồi chạy lại. |
| `parse` | Sửa đúng dòng `error.message` nêu rồi chạy lại. |
| `canh` | Nội dung cảnh sai (chữ quá dài, tràn khung, mã mẫu hay mã tham số lạ, mốc `tham-so` vượt thời lượng cảnh): rút gọn hoặc sửa cảnh đó theo `error.message`. |
| `giong` | Không tạo được giọng máy, thường do mất mạng: kiểm mạng rồi chạy lại, tối đa một lần; hoặc đặt file giọng thầy cô đưa vào `giong/canh-N.mp3`. |
| `chromium` | Chưa cài Chromium: hỏi thầy cô trước (tải 150–300 MB), rồi chạy `powershell -NoProfile -ExecutionPolicy Bypass -File tools\vi\pptmaster.ps1 -Action tool -Name chromium`. |
| `ffmpeg` | Chưa có FFmpeg: cài theo mục "Công cụ tuỳ chọn" của `docs/vi/cai-dat-bang-ai.md`. |
| `dung` | Chụp khung hoặc ghép hỏng: báo nguyên `error.message` cho thầy cô, không tự sửa. |
| `write` | Không ghi được file: xin thầy cô đóng `video.mp4` nếu đang mở, kiểm ổ đĩa còn chỗ, rồi chạy lại. |
| `internal` | Lỗi ngoài dự kiến: dán nguyên `error.message` để báo cho người bảo trì, không tự đoán cách sửa. |

Không tự chạy FFmpeg, không tự ghép hay cắt video theo cách riêng. Không viết HTML hay ảnh cảnh bằng tay, không sửa khung hình trong thư mục tạm: mọi thay đổi đi qua `video.md` rồi chạy lại công cụ.

## Ghi vào brief

- Loại việc: Video giải thích dựng bằng mã.
- Thầy cô yêu cầu: môn, lớp, bài, nội dung cần giải thích, điều học sinh cần đạt, độ dài, giọng, tốc độ, cảnh thí nghiệm, phụ đề, giọng thu sẵn — ghi đúng lời thầy cô.
- AI đề xuất (thầy cô đã đồng ý): danh sách cảnh (loại cảnh và ý chính của từng cảnh), mẫu thí nghiệm nếu có, và các gợi ý thầy cô chấp nhận.
- Loại việc này không tạo PPTX: ghi brief vào `projects/_video/<tên_video>/brief.md`; không chạy `import-sources`, không có bước xác nhận của upstream.
- Tin nhắn hỏi vẫn kết thúc bằng dòng chốt cách xác nhận như các loại việc khác, nhưng luôn ở dạng khung chat và không mở trang web xác nhận: "Trước khi dựng, em sẽ gửi danh sách cảnh trong khung chat để thầy cô duyệt. Thầy cô đồng ý nhé?". Tạo nhanh thì bỏ bước duyệt này.
- Viết theo mẫu `docs/vi/tro-ly/mau-brief.md`.
