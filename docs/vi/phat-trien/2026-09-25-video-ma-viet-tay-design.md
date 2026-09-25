# Thiết kế: Video giải thích dựng bằng mã, kiểu viết tay (v6.3.2-vi.9)

Ngày 2026-09-25. Gói này thêm loại việc thứ 10 cho lớp Việt: từ một kịch bản chữ, dựng ra video giải thích kiểu viết tay (whiteboard) có giọng đọc tiếng Việt và phụ đề, gồm cả cảnh quay thí nghiệm ảo. Video dựng bằng mã (HTML/SVG vẽ theo thời gian, Chromium chụp khung, FFmpeg ghép), không dùng AI sinh video.

Gói này là bản đầu của nhóm "video dựng bằng mã". Kiểu Vox và khổ dọc 9:16 cho truyền thông tách sang gói v6.3.2-vi.10 (spec riêng), dùng lại bộ dựng của gói này. Video AI sinh (Veo 3.1, Gemini Omni Flash, Flow) là nhóm khác, chưa làm.

## 1. Tiêu chí thành công

1. Thầy cô đưa nội dung một bài (đoạn văn hoặc dàn ý); sau một lượt hỏi, nhận được `video.mp4` 1280×720 có giọng đọc, cảnh viết tay khớp lời, phụ đề in lên hình.
2. Tám loại cảnh dựng đúng: tiêu đề, khái niệm, công thức, ý từng ý, quy trình, so sánh, đồ thị, thí nghiệm.
3. Cảnh thí nghiệm chạy **đúng 8 mô hình đã kiểm** của thí nghiệm ảo: số đo hiện trong video là số do mô hình tính, không do AI vẽ.
4. Chữ tiếng Việt hiển thị đủ dấu, kể cả dấu chồng, bằng font viết tay có sẵn trên Windows (Segoe Print hoặc Ink Free).
5. Ý thứ k hiện khi câu thứ k của lời bắt đầu; hình và tiếng khớp nhau, phụ đề khớp tiếng.
6. Có sẵn `giong/canh-N.mp3` thì công cụ dùng file đó thay cho giọng máy, để dùng giọng thu sẵn (kể cả giọng clone) và dựng không cần mạng.
7. Kịch bản sai bị chặn **trước khi** tốn thời gian dựng, nêu đúng dòng; chữ tràn khung bị bắt ở bước dựng thử, nêu đúng cảnh.
8. AI trong Antigravity đi đúng luồng: cổng hỏi và dòng loại việc thứ 10 nằm thẳng trong file luật Antigravity.
9. Kịch bản, giọng và video của thầy cô không bao giờ lên GitHub.
10. Toàn bộ test của lớp Việt vẫn xanh; không sửa `skills/`, `requirements.txt` gốc hay `attribution_guard.py`; không thêm thư viện Python mới ngoài `playwright` đã có đường cài theo yêu cầu.

## 2. Hiện trạng đã kiểm

- `tools/vi/video.py` (v6.3.2-vi.4) chỉ dựng video **từ slide** của một dự án PPTX: chụp ảnh slide rồi ghép bằng FFmpeg, hoặc xuất bằng PowerPoint. Có sẵn `video_parts/srt.py` (`Cue`, `parse_srt`, `render_srt`, `cumulative_offsets`), `video_parts/media.py` (`probe_duration`, `build_concat_text`, `escape_subtitles_filter`, `scale_filter`, `MediaError`) và `burn_subtitles` in phụ đề lên hình.
- `video.py` có `has_chromium()` (cần thư mục `ms-playwright` và `import playwright` chạy được bằng Python của repo) và cách cài theo yêu cầu: `tools\vi\pptmaster.ps1 -Action tool -Name chromium` (khoảng 150–300 MB, phải hỏi thầy cô trước). `playwright` không nằm trong `requirements.txt` và **chưa có** trong `venv` của máy chủ repo.
- `edge-tts` 7.2.8 đã có trong `venv` (dòng `edge-tts>=7.2.8` trong requirements của upstream); giọng `vi-VN-HoaiMyNeural` và `vi-VN-NamMinhNeural`. **Cần mạng.**
- FFmpeg cài sẵn qua winget trên máy chủ repo; đã có lệnh cài trong `pptmaster.ps1`.
- Máy chủ repo có font `segoepr.ttf` (Segoe Print), `Inkfree.ttf` (Ink Free), `comic.ttf`. Đây là font chuẩn của Windows nên không cần nhúng hay xin giấy phép; chưa kiểm chúng có hiển thị đủ dấu chồng tiếng Việt.
- Tám mô hình thí nghiệm ảo nằm ở `tools/vi/thi_nghiem_parts/mo_hinh/` (mỗi mô hình một `.json` khai báo và một `.js` có `tinh`, `ve`); khung `tools/vi/thi_nghiem_parts/runtime/khung.js` có logic thuần (`taoNgauNhien`, `dinhDang`, `danhDau`…) dùng chung được.
- Bài học vi.7: Antigravity không chép nội dung file nhắc bằng `@`; quy tắc cần tới Antigravity phải nằm thẳng trong `.agents/rules/ppt-master-vi.md` (tối đa 12.000 ký tự, hiện khoảng 6.600). Test khoá bảng loại việc trong file luật khớp `quy-trinh-hoi.md` và khoá chữ "9 loại".
- Loại việc "Video bài giảng" (thứ 6) đã dùng các từ "làm video", "xuất video", "lồng tiếng". Từ "làm video" một mình sẽ trùng giữa hai loại việc.
- Chưa đo tốc độ chụp khung của Chromium: đây là điều chưa biết, là việc đầu tiên của kế hoạch (mục 12 của spec này).

## 3. Quyết định

| # | Quyết định | Lý do |
|---|---|---|
| Q1 | Nhóm "dựng bằng mã" làm trước; AI sinh video làm sau | Chính xác khoa học, miễn phí, sửa từng chữ, không cần khoá API |
| Q2 | Mục đích trước mắt: giảng bài cho học sinh tự học; truyền thông (Vox, dọc) là gói sau | Chủ repo chọn "cả hai, giảng bài trước" |
| Q3 | Cảnh mẫu đã kiểm; AI chỉ chia cảnh và điền nội dung | Chất lượng ổn định kể cả khi AI yếu; kiểm được và test được; cùng cách với thí nghiệm ảo |
| Q4 | Bản đầu chỉ phong cách viết tay, khổ ngang 1280×720 | Làm cho chắc; Vox và dọc cần 30 khung/giây và bố cục khác |
| Q5 | Mỗi cảnh là hàm `datThoiDiem(t)` xác định theo thời điểm | Khung hình lặp lại được, test được, chụp ở fps tuỳ ý |
| Q6 | Cảnh thí nghiệm chỉ dùng 8 mô hình trong thư viện | Giữ bảo đảm khoa học đã kiểm; mô hình do AI viết cần thầy cô soát nên không vào video bản đầu |
| Q7 | Giọng: dùng `giong/canh-N.mp3` nếu có, không thì edge-tts | Cho phép giọng thu sẵn và giọng clone; edge-tts cần mạng nên phải có đường không cần mạng |
| Q8 | Không có mốc câu từ file giọng có sẵn thì chia câu theo số ký tự và cảnh báo | Trung thực về độ chính xác của mốc |
| Q9 | Chữ tràn khung là lỗi, không tự cắt chữ | Thầy cô sửa nội dung; cắt chữ âm thầm làm sai bài |
| Q10 | Chromium cài theo yêu cầu, hỏi trước | Nặng 150–300 MB; cùng cách với chức năng làm video từ slide |
| Q11 | Ghép bằng `video_parts` hiện có, `video.py` không sửa | Không phá chức năng đã phát hành |
| Q12 | Font Segoe Print hoặc Ink Free; chỉ chạy trên Windows | Đúng phạm vi bản Việt; không nhúng font |

## 4. Kiến trúc và file

```
tools/vi/video_ma.py                    CLI: đọc → kiểm → giọng → dựng thử → chụp khung → ghép → JSON
tools/vi/video_ma_parts/
    parse.py                            đọc video.md, lỗi nêu đúng dòng
    kiem.py                             kiểm giới hạn chữ, cảnh báo thời lượng
    giong.py                            edge-tts hoặc giong/canh-N.mp3, mốc câu, thời lượng
    lich.py                             chia mốc hiện ý theo câu; thời lượng cảnh; bảng thời gian cả video
    trang.py                            ghép trang HTML của một cảnh (khung + phong cách + cảnh + dữ liệu)
    chup.py                             Chromium chụp khung theo thời điểm
    ghep.py                             FFmpeg ghép khung + tiếng + phụ đề
    runtime/khung-video.js              vòng thời gian: datThoiDiem(t), tiện ích vẽ chữ dần, nét dần
    runtime/viet-tay.css, viet-tay.js   phong cách viết tay: nền, font, bút
    runtime/canh/<loại>.js              tám loại cảnh
tools/vi/tests/test_video_ma_*.py       và tools/vi/tests/js/test_canh.js
docs/vi/tro-ly/video-giai-thich.md      loại việc thứ 10: câu hỏi, ngữ pháp video.md, thứ tự làm
docs/vi/tro-ly/canh-video.md            danh mục 8 loại cảnh và trường của từng loại
docs/vi/video-giai-thich.md             tài liệu cho thầy cô
```

Sửa: `AGENTS.vi.md` (mục 3 câu lệnh kích hoạt, mục 10 bảng 10 loại, mục 15 mới), `.agents/rules/ppt-master-vi.md`, `docs/vi/tro-ly/quy-trinh-hoi.md`, `docs/vi/tro-ly/mau-brief.md`, `docs/vi/xu-ly-loi.md`, `docs/vi/cau-lenh-mau.md`, `docs/vi/bat-dau-nhanh.md`, `README.md`, `CHANGELOG-VI.md`, `tools/vi/tests/test_vi_layer.py`. Không sửa `tools/vi/video.py`, `video_parts/`, `thi_nghiem_parts/` (chỉ đọc file mô hình và `khung.js` từ đó).

Đầu ra ở `projects/_video/<tên>/`: `video.md`, `giong/` (mp3 từng cảnh, ghi đè được), `video.mp4`, `phu-de.srt` khi `phu-de: file`, `xem-truoc/canh-N.png` khi dùng `--xem-truoc`. Không chạy `project_manager.py init`, không tạo SVG của upstream.

Ràng buộc: Python chỉ dùng thư viện chuẩn, `edge-tts` (đã có) và `playwright` (cài theo yêu cầu); phần vẽ là JavaScript thuần, không thư viện ngoài, không địa chỉ web.

## 5. Ngữ pháp `video.md`

Khối thông tin giữa hai dòng `---`, rồi mỗi cảnh một mục `## Cảnh <số>` theo thứ tự.

| Khoá | Bắt buộc | Giá trị |
|---|---|---|
| `tieu-de` | có | tên video |
| `mon`, `lop` | có | môn và lớp |
| `phong-cach` | không | `viet-tay` (mặc định và duy nhất ở bản đầu) |
| `giong` | không | `nu` (mặc định, `vi-VN-HoaiMyNeural`) hoặc `nam` (`vi-VN-NamMinhNeural`) |
| `toc-do` | không | `cham` (−10%), `vua` (mặc định), `nhanh` (+15%) |
| `phu-de` | không | `hinh` (mặc định, in lên hình), `file` (rời), `khong` |

Mỗi cảnh có `loai:` và `loi:` (lời đọc, ít nhất một câu) cùng các trường của loại cảnh. Quy ước chữ như các công cụ trước: `H~2~SO~4~`, `m/s^2^`, `**in đậm**`. Không chứa địa chỉ web.

Ví dụ:

```
---
tieu-de: Chuẩn độ acid – base
mon: Hoá học
lop: 11
---

## Cảnh 1
loai: tieu-de
chu: Chuẩn độ acid – base
phu: Hoá học 11
loi: Hôm nay chúng ta tìm hiểu cách xác định nồng độ một dung dịch acid.

## Cảnh 2
loai: y-tung-y
tieu-de: Ba bước chuẩn độ
y: Đo chính xác thể tích acid vào bình tam giác
y: Nhỏ từ từ dung dịch NaOH từ buret
y: Dừng khi chất chỉ thị đổi màu
loi: Bước một, đo chính xác thể tích acid. Bước hai, nhỏ từ từ dung dịch NaOH. Bước ba, dừng khi chất chỉ thị đổi màu.
```

## 6. Tám loại cảnh

| Loại | Trường (ngoài `loi`) | Giới hạn | Cách hiện |
|---|---|---|---|
| `tieu-de` | `chu`, `phu` | `chu` ≤ 90 ký tự | chữ lớn được viết ra |
| `khai-niem` | `thuat-ngu`, `dinh-nghia` | thuật ngữ ≤ 60, định nghĩa ≤ 220 | khung vẽ nét, chữ viết vào |
| `cong-thuc` | `bieu-thuc`, `giai-thich` (lặp) | biểu thức ≤ 90; ≤ 4 giải thích, mỗi giải thích ≤ 80 | biểu thức viết dần, giải thích hiện sau |
| `y-tung-y` | `tieu-de`, `y` (lặp) | ≤ 6 ý, mỗi ý ≤ 80 | mỗi ý viết ra khi lời nói tới |
| `quy-trinh` | `tieu-de`, `buoc` (lặp) | 2–5 bước, mỗi bước ≤ 50 | bước và mũi tên vẽ nối tiếp |
| `so-sanh` | `tieu-de`, `trai`, `phai`, `y-trai`, `y-phai` (lặp) | mỗi cột ≤ 4 ý, mỗi ý ≤ 60 | hai cột viết lần lượt |
| `do-thi` | `tieu-de`, `truc-ngang`, `truc-doc`, `diem` (lặp, dạng `x, y`) | 2–12 điểm | trục rồi từng điểm vẽ dần, nối đường |
| `thi-nghiem` | `mau`, `tham-so` (lặp), `do` | mẫu thuộc 8 mô hình | xem mục 7 |

Giới hạn bổ sung: tiêu đề cảnh ≤ 90 ký tự, `trai` và `phai` ≤ 24, `truc-ngang` và `truc-doc` ≤ 40.

Ý thứ k của các loại có danh sách hiện khi câu thứ k của lời bắt đầu; ít câu hơn số ý thì chia đều theo thời lượng cảnh (mục 8).

## 7. Cảnh thí nghiệm

```
## Cảnh 5
loai: thi-nghiem
mau: li-con-lac-don
tham-so: 0 chieu-dai 0.4
tham-so: 6 chieu-dai 1.6
tham-so: 12 goc-lech 12
do: chu-ki
loi: Khi tăng chiều dài dây, chu kì dao động tăng theo.
```

- `mau` là mã một trong 8 mô hình; mã khác thì lỗi `canh`.
- Mỗi dòng `tham-so:` là `<giây> <mã tham số> <giá trị>`. Giá trị của một tham số tại thời điểm `t` là nội suy tuyến tính giữa hai mốc liền kề của tham số đó; trước mốc đầu giữ mặc định của mô hình; sau mốc cuối giữ giá trị của mốc cuối. Giá trị phải nằm trong khoảng cho phép của tham số (nếu không là lỗi `parse` nêu đúng dòng). Mốc vượt thời lượng cảnh chỉ biết được sau khi có giọng, nên bị bắt ở bước lịch với lỗi `canh`.
- `do:` liệt kê các đại lượng đo hiện thành số bên cạnh hình vẽ; mỗi giá trị lấy từ `tinh()` của mô hình tại thời điểm đó.
- Cảnh chạy `ve(ctx, p, t, kt, d)` của mô hình trong canvas với `t` là thời gian trong cảnh; mô hình `hoatHinh: mot-lan` chạy một lần trong khoảng `thoiLuong`.
- Dùng lại `khung.js` của thí nghiệm ảo cho `MAU`, `PHONG`, `dinhDang`, `danhDau`.
- Tối đa 3 mã tham số khác nhau trong `tham-so:` và tối đa 3 đại lượng trong `do:` (mặc định hai đại lượng đầu của mô hình); số giây của `tham-so:` tính từ đầu cảnh.

## 8. Giọng, mốc thời gian và phụ đề

1. Với mỗi cảnh: có `giong/canh-N.mp3` thì dùng; không thì gọi edge-tts với `lời` của cảnh, lưu `giong/canh-N.mp3` (lần dựng sau dùng lại nếu `lời` không đổi; công cụ ghi sổ `giong/canh-N.json` gồm mã băm của lời, giọng, tốc độ và mốc câu; có `canh-N.mp3` mà không có `canh-N.json` là file thầy cô đặt sẵn, không bao giờ bị ghi đè).
2. Thời lượng cảnh = 0,7 giây dẫn đầu (tiêu đề viết trước khi tiếng bắt đầu) + thời lượng file giọng (`probe_duration`) + 0,6 giây, tối thiểu 2,5 giây, làm tròn lên bội của 1/15 giây. Tiếng được đệm 0,7 giây im lặng ở đầu; mốc câu và mốc hiện ý tính theo thời gian cảnh.
3. Mốc câu: edge-tts trả mốc câu; file có sẵn không có mốc thì chia các câu của `lời` theo tỉ lệ số ký tự trên thời lượng giọng, và ghi cảnh báo "mốc câu ước lượng".
4. Ý thứ k hiện tại mốc câu thứ k; ít câu hơn ý thì hiện đều nhau trên thời lượng giọng.
5. Phụ đề: mỗi câu là một `Cue` tại mốc câu cộng độ lệch của cảnh (`cumulative_offsets`); `phu-de: hinh` in bằng phần in phụ đề dùng chung; `file` ghi `phu-de.srt`.

## 9. Dựng khung và ghép

- `trang.py` ghép một trang HTML cho mỗi cảnh: khung + phong cách + file cảnh + dữ liệu cảnh (JSON, thoát `<` như `build_html.py`); trang chỉ đọc thời điểm qua `window.datThoiDiem(t)`.
- `chup.py` mở trang trong Chromium 1280×720, gọi `datThoiDiem(số_khung / fps_chụp)`, chụp. Trước khi chụp, đo phần tử có tràn khung không (dùng cho `--xem-truoc` và cho lỗi `canh`).
- `ghep.py` ghép chuỗi khung của các cảnh thành hình, ghép tiếng các cảnh (đệm im lặng cuối cảnh), xuất MP4 30 khung/giây (khung chụp thấp hơn được nhân lên), rồi in phụ đề.
- `fps_chụp = 15`, ảnh PNG (đã đo ở biên bản kiểm thử: 45,3 ms/khung).

## 10. Công cụ

```
python tools/vi/video_ma.py <thư_mục> [--plan-only] [--xem-truoc]
```

stdout đúng một dòng JSON: `ready`, `files`, `so_canh`, `thoi_luong_giay`, `phong_cach`, `giong` (`may` hoặc `co-san`), `warnings`, `error`. Mã thoát 0 hoặc 1. `--plan-only` chỉ đọc và kiểm; `--xem-truoc` dựng thử mỗi cảnh một ảnh ở `xem-truoc/`, không tạo giọng, không ghép.

| `error.step` | Nghĩa | Xử lý |
|---|---|---|
| `input` | thiếu thư mục hoặc `video.md`, sai tham số | viết file rồi chạy lại |
| `parse` | sai ngữ pháp, nêu đúng dòng | sửa dòng đó |
| `canh` | nội dung cảnh sai (quá dài, mã mẫu lạ, chữ tràn khi dựng thử) | rút gọn nội dung theo `error.message` |
| `giong` | edge-tts lỗi, thường do mất mạng | có mạng rồi chạy lại, hoặc đặt sẵn `giong/canh-N.mp3` |
| `chromium` | chưa cài Chromium hoặc `playwright` | hỏi thầy cô rồi chạy lệnh cài |
| `ffmpeg` | chưa có FFmpeg | cài theo mục "Công cụ tuỳ chọn" |
| `dung` | chụp khung hoặc ghép hỏng | báo `error.message` |
| `write` | không ghi được file | đóng file đang mở, kiểm ổ đĩa |
| `internal` | lỗi ngoài dự kiến | dán nguyên `error.message` cho người bảo trì |

Cảnh báo (không chặn): cảnh dài quá 40 giây, video dài quá 8 phút, mốc câu ước lượng, `lời` dài quá 700 ký tự.

## 11. Lượt hỏi và nối vào AI

Câu hỏi bắt buộc (tối đa 7, mỗi câu có gợi ý): (1) môn, lớp, bài và nội dung cần giải thích; (2) học sinh cần hiểu hoặc làm được gì sau video; (3) độ dài mong muốn; (4) giọng nam hay nữ, tốc độ; (5) có thí nghiệm ảo trong danh mục nào để đưa vào không; (6) phụ đề in lên hình hay file rời; (7) có file giọng thu sẵn không. Tạo nhanh hỏi câu 1 và 3.

**Trùng từ "làm video":** câu lệnh có "video giải thích", "video viết tay", "video whiteboard", "video hoạt hình chữ" thì vào loại việc này. Câu chỉ có "làm video", "xuất video" thì hỏi đúng một câu: "Thầy cô muốn làm video từ bài giảng slide đã có, hay dựng video giải thích mới từ nội dung chữ?". Trả lời slide thì theo mục 11 của `AGENTS.vi.md` (đã có); trả lời video mới thì theo mục 15 của `AGENTS.vi.md` (mới).

Cổng hỏi, bảng 10 loại và các bước nằm thẳng trong file luật Antigravity; file phải giữ dưới 12.000 ký tự.

## 12. Việc đầu tiên của kế hoạch: đo và kiểm

Trước khi viết bộ dựng, làm hai phép thử bỏ đi và ghi kết quả vào biên bản kiểm thử:

1. **Font:** hiển thị các câu có dấu chồng ("Nhờ ướt nhẫm quyết định", "Đường thẳng khớp") bằng Segoe Print và Ink Free trong Chromium; chọn font nào đọc rõ hơn; nếu cả hai lỗi dấu thì dừng và báo chủ repo chọn font khác.
2. **Tốc độ chụp:** chụp 100 khung liên tiếp của một cảnh mẫu ở 1280×720 bằng PNG và JPEG; ghi giây mỗi khung, từ đó tính thời gian dựng video 5 phút ở 12, 15 và 30 khung/giây; chọn `fps_chụp` và định dạng ảnh.

Nếu thời gian dựng 5 phút quá dài (quá 20 phút) thì đề xuất giảm khung/giây hoặc chụp song song nhiều tab trước khi làm tiếp.

## 13. Kiểm thử

- **Bộ đọc:** mọi lỗi ở mục 5, 6, 7 nêu đúng dòng; ví dụ trong file hướng dẫn chạy qua bộ đọc thật.
- **Cảnh (Node):** mỗi loại có `datThoiDiem(t)` cho cùng kết quả khi gọi lại; `t = 0` chưa hiện ý nào; cuối cảnh hiện đủ; cảnh thí nghiệm ở mọi `t` cho số đo bằng `tinh()` của mô hình.
- **Lịch:** ý k hiện tại mốc câu k; ít câu hơn ý thì chia đều; thời lượng cảnh tối thiểu; mốc câu ước lượng có cảnh báo.
- **Công cụ:** mọi nhánh ra đúng một dòng JSON; giọng thật **không bao giờ** gọi trong test (dùng tiếng giả tạo bằng FFmpeg); lỗi `giong`, `chromium`, `ffmpeg` mô phỏng bằng giả lập.
- **Tích hợp** (bỏ qua có ghi lý do nếu máy thiếu Chromium hoặc FFmpeg): dựng video 2 cảnh với tiếng giả, kiểm bằng `ffprobe` khớp thời lượng, kích thước 1280×720, có tiếng.
- **Lớp hướng dẫn:** `test_vi_layer.py` cho loại việc thứ 10, mục 15, file luật, tài liệu thầy cô, mục xử lý lỗi; chữ "9 loại" thành "10 loại"; bảng file luật và `quy-trinh-hoi.md` vẫn phải khớp (10 dòng); mục cuối của `AGENTS.vi.md` mở rộng thêm mục 15.
- **Bằng mắt:** sau khi dựng thật, chủ repo xem ảnh chụp từng loại cảnh và một video mẫu.

## 14. Ngoài phạm vi

Kiểu Vox và khổ dọc 9:16 (gói vi.10); video AI sinh (Veo, Omni Flash, Flow) và gói prompt cho chúng; mô hình thí nghiệm do AI viết trong video; ảnh và video tư liệu chèn vào; nhạc nền; giọng clone tích hợp sẵn trong bộ công cụ (chỉ hỗ trợ file có sẵn); dựng trên macOS hoặc Linux; sửa `video.py`.

## 15. Rủi ro

- **Tốc độ dựng:** chưa đo; video dài có thể mất nhiều phút; mục 12 là cổng quyết định.
- **Font:** chưa kiểm dấu chồng tiếng Việt; nếu cả hai font lỗi phải chọn font khác, ảnh hưởng phong cách.
- **Cần mạng cho giọng máy:** đường `giong/canh-N.mp3` là lối thoát duy nhất khi mất mạng.
- **Chromium 150–300 MB** phải tải lần đầu; máy trường chặn tải sẽ không dựng được.
- **Nội dung do thầy cô chịu trách nhiệm:** công cụ chỉ đảm bảo số đo của cảnh thí nghiệm; không kiểm đúng sai của lời giảng và công thức thầy cô viết.
- **Mốc câu ước lượng** khi dùng giọng có sẵn: hình có thể lệch tiếng vài trăm mili giây; ghi cảnh báo.
- **Bản phát hành lớn** vì gồm cả bộ dựng, tám loại cảnh, cảnh thí nghiệm và tài liệu.
