# Thiết kế: Hiệu ứng giáo dục cho video giải thích (v6.3.2-vi.11)

Ngày 2026-09-26. Chủ repo nhận xét video vi.10 "hơi đơn giản" và muốn học những gì tốt cho giáo dục từ HyperFrames (HeyGen, Apache 2.0). Quyết định: không thay bộ dựng mà học ý tưởng. Bộ dựng của repo đã cùng kiến trúc với HyperFrames: trang HTML chạy theo thời gian, Chromium chụp khung, FFmpeg ghép. Nó đã được kiểm với tám mô hình thí nghiệm, font Itim và phụ đề, và chỉ cần Python, không cần Node hay mạng. Chủ repo chọn đủ tám nhóm: nhấn ý chính, biểu đồ và số động, chuyển cảnh đa dạng, chữ động và phụ đề tô từng từ, cảnh câu hỏi nhanh, hiệu ứng âm thanh, nhạc nền, sơ đồ tư duy và dòng thời gian.

Kiểu Vox và khổ dọc 9:16 lùi sang vi.12.

## 1. Tiêu chí thành công

1. **Nhấn ý chính:** ý chính được tô, khoanh hoặc gạch đúng lúc giọng đọc nói tới từ đó, lệch không quá 0,15 giây với giọng máy.
2. **Biểu đồ:** biểu đồ cột, đường, tròn mọc dần theo lời, cùng số liệu chạy từ 0 lên giá trị thật. Công thức hiện từng phần.
3. **Chuyển cảnh:** có năm kiểu (lau bảng, lật trang, trượt, phóng xuyên, mở màn), chọn được cho cả video hoặc từng cảnh.
4. **Chữ động:**
   - Tiêu đề cảnh mở đầu nảy vào từng chữ; từ khoá nảy nhẹ khi được nhấn.
   - Phụ đề tô màu từng từ theo giọng đọc (kiểu karaoke).
   - Mọi chữ vẫn là Itim, đủ dấu.
5. **Câu hỏi nhanh:** cảnh `cau-hoi` hiện câu hỏi 2–4 lựa chọn, đếm ngược cho học sinh nghĩ, rồi hiện đáp án đúng kèm lời giải thích có giọng đọc.
6. **Hiệu ứng âm thanh:**
   - Có tiếng bút viết, tiếng "ting" khi ý hiện, tiếng chuyển cảnh, tiếng tích tắc khi đếm ngược và tiếng chuông khi hiện đáp án.
   - Tất cả được tạo bằng FFmpeg, không dùng file âm thanh ngoài.
   - Luôn nhỏ hơn giọng đọc.
7. **Nhạc nền:** nhạc nhẹ tự hạ nhỏ khi có giọng. File do thầy cô đưa, hoặc do AI tìm bằng công cụ mới (Openverse, giấy phép CC0/CC BY, không cần khoá). Nguồn nhạc luôn được ghi.
8. **Sơ đồ tư duy và dòng thời gian:** nút và nhánh vẽ dần theo lời.
9. **Tương thích và chất lượng:**
   - Kịch bản vi.9 và vi.10 dựng được mà không phải sửa.
   - Mỗi hiệu ứng mới tắt được.
   - Mọi khung vẫn là hàm xác định của `t`.
   - Thời gian dựng video 5 phút vẫn không quá 8 phút trên máy chủ repo.
10. **Ranh giới:** không sửa `skills/`, `video.py`, `video_parts/`, `thi_nghiem_parts/`; không thêm thư viện Python hay JavaScript bên ngoài; trang dựng không dùng địa chỉ web.

## 2. Hiện trạng đã kiểm

- **Mốc từng từ:** edge-tts 7.2.8 trả `WordBoundary` cho giọng tiếng Việt với mốc từng từ (đã chạy thử: "Chu" 0,125 s, "kì" 0,275 s…). Vi.10 chỉ dùng `SentenceBoundary`.
- **FFmpeg bản máy chủ repo** có sẵn:
  - `sidechaincompress`, để hạ nhạc khi có giọng;
  - `anoisesrc`, `aevalsrc`, `afade`, `aloop`, `amix`, `adelay`, để tạo và trộn hiệu ứng âm thanh;
  - bộ lọc `ass`/`subtitles` (libass), hỗ trợ thẻ karaoke `\kf`.
- **Openverse:** `https://api.openverse.org/v1/audio/?q=...&license=cc0,by` trả kết quả không cần khoá, gồm tiêu đề, giấy phép, tác giả, thời lượng và đường tải.
- **Runtime vi.10:** mỗi cảnh là danh sách mục có `batDau`/`thoiLuong`. Hai kế hoạch thuần `ban-tay.js` và `may-quay.js` đã có. Chuyển cảnh lau bảng dùng `nenTruoc`, là khung cuối của cảnh trước. Tiếng mỗi cảnh = 1,0 s dẫn đầu + giọng + đuôi, trộn nhiễu nền −57 dBFS.
- Phụ đề hiện là `.srt` tạo từ mốc câu, in bằng libass với Itim.

## 3. Quyết định

| # | Quyết định | Lý do |
|---|---|---|
| Q1 | Giọng máy chuyển sang `WordBoundary`. Mốc câu tính từ từ đầu của mỗi câu. Sổ giọng lưu cả `tu` (mốc từng từ). File giọng có sẵn thì ước lượng mốc từ theo tỉ lệ ký tự, kèm cảnh báo như mốc câu. | Nhấn ý và phụ đề tô từng từ đều cần mốc từ; mốc câu vẫn giữ nguyên nghĩa |
| Q2 | Cú pháp nhấn trong mọi trường chữ: `==chữ==` tô bút dạ vàng, `((chữ))` khoanh tròn, `__chữ__` gạch chân. Hiệu ứng nổ ra khi giọng đọc tới từ đầu tiên của cụm (so khớp bỏ dấu, không phân biệt hoa thường, trong lời của cảnh). Không tìm thấy thì nổ 0,3 s sau khi cụm được viết xong. | Người soạn đánh dấu ngay trong chữ; không thêm trường mới |
| Q3 | Loại cảnh mới `bieu-do` gồm `kieu: cot\|duong\|tron`, `du-lieu: nhãn \| số` (2–8 dòng), `don-vi`, `truc-ngang`, `truc-doc`. Phần tử k mọc dần tại mốc câu k (ít câu hơn thì chia đều); số chạy từ 0 lên giá trị trong 0,8 s. | Biểu đồ số liệu xuất hiện trong mọi môn |
| Q4 | Trong mọi trường chữ, `{{số}}` là số chạy từ 0 tới số đó khi chữ quanh nó được viết. Định dạng số kiểu Việt (dấu phẩy thập phân). `bieu-thuc` của `cong-thuc` tách phần bằng ` \| `; phần thứ k hiện tại mốc câu k. | Số và công thức động, dùng lại cú pháp đã quen |
| Q5 | Khoá đầu `chuyen-canh` nhận thêm `lat-trang`, `truot`, `phong`, `mo-man`, `luan-phien` (xoay vòng năm kiểu). Trường `chuyen:` trong từng cảnh ghi đè. Mọi kiểu dùng `nenTruoc` trong 0,5 s đầu cảnh. | Cùng cơ chế với lau bảng, không cần ghép chồng bằng FFmpeg |
| Q6 | Khoá đầu `chu-dong: co\|khong` (mặc định `co`): tiêu đề cảnh `tieu-de` nảy từng chữ có đàn hồi thay cho bút viết; từ được nhấn nảy nhẹ 1 → 1,12 → 1. Các chữ khác vẫn viết tay. | Nảy từng chữ hợp tiêu đề lớn; viết tay giữ chất bảng trắng cho nội dung |
| Q7 | `phu-de` thêm giá trị `karaoke` (mặc định mới khi in lên hình): tạo file `.ass` với thẻ `\kf` theo mốc từ, tô vàng từng từ. `hinh` giữ kiểu cũ; `file` vẫn ra `.srt`. | Học sinh theo được từng từ, có lợi khi học từ vựng khoa học |
| Q8 | Loại cảnh mới `cau-hoi` gồm `cau-hoi`, `lua-chon` (2–4, tự đánh A–D), `dap-an` (chữ cái), `giai-thich`, `cho` (giây đếm ngược, 3–10, mặc định 5), `loi` (đọc câu hỏi), `loi-giai` (đọc sau khi hiện đáp án). Tiếng cảnh = dẫn đầu + giọng câu hỏi + `cho` giây + 0,4 s + giọng lời giải + đuôi. Giọng lời giải là file riêng `giong/canh-N-giai.mp3`. | Dừng để học sinh tự nghĩ là giá trị lớn nhất của video tự học |
| Q9 | Hiệu ứng âm thanh tạo bằng FFmpeg từ công thức cố định, nên không có bản quyền và xác định được. Gồm: bút (nhiễu lọc thông dải, điều biên), ting (sine 1318 Hz tắt dần), chuyển cảnh (nhiễu quét tần), tích tắc, chuông đúng. Mức −20 dB so với giọng, tổng tiếng bút trong một cảnh không vượt 40% thời gian cảnh. Khoá đầu `am-thanh: co\|khong` (mặc định `co`). | Không vướng giấy phép; không lấn giọng |
| Q10 | Thời điểm âm thanh lấy từ runtime: sau khi nạp trang, `THI_VIDEO.suKien()` trả `[{t, loai, dai}]` theo mục. Tiến trình chụp đọc danh sách này; `ghep.py` trộn vào tiếng cảnh. | Một nguồn thời gian duy nhất: đúng khung hình nào thì đúng âm thanh đó |
| Q11 | Nhạc nền: khoá đầu `nhac-nen: <file trong nhac/>`. Nguồn lấy từ `nhac/nguon.json` hoặc khoá `nguon-nhac:`; thiếu nguồn thì lỗi `canh`. Nhạc được lặp cho đủ dài, vào và ra dần 1,5 s, nền −24 dB, hạ thêm bằng `sidechaincompress` khi có giọng. Nguồn nhạc hiện ở góc dưới trong 4 s cuối video. Công cụ mới `tools/vi/tim_nhac.py "<từ khoá>" -o <thư_mục>/nhac` tìm trên Openverse (CC0, CC BY), tải bản xem trước mp3, ghi `nguon.json`. | Có nhạc mà vẫn đúng giấy phép; bộ dựng không tự lên mạng |
| Q12 | Loại cảnh mới `so-do` gồm `trung-tam` và `nhanh` (2–6; nhánh k vẽ dần tại mốc câu k, đường cong kiểu vẽ tay tới ô nhãn), `hinh` tuỳ chọn ở nút trung tâm. Loại cảnh mới `dong-thoi-gian` gồm `moc: <nhãn> \| <mô tả>` (2–6, trục vẽ trước, mốc k tại câu k). | Sơ đồ tư duy và tiến trình lịch sử là dạng tóm tắt bài phổ biến nhất |
| Q13 | Không nhúng HyperFrames, GSAP hay thư viện nào. Hàm làm mượt (easeOutBack, spring) viết thuần trong `runtime/dong.js`. | Giữ ràng buộc "JS thuần, không thư viện" của vi.9 |

## 4. Kiến trúc và file

```
tools/vi/video_ma_parts/runtime/dong.js              easing, spring, nảy chữ, số chạy (thuần, chạy được trong Node)
tools/vi/video_ma_parts/runtime/nhan.js              tô / khoanh / gạch cho cụm được đánh dấu
tools/vi/video_ma_parts/runtime/chuyen-canh.js       5 kiểu chuyển cảnh trên nenTruoc
tools/vi/video_ma_parts/runtime/canh/bieu-do.js, cau-hoi.js, so-do.js, dong-thoi-gian.js
tools/vi/video_ma_parts/am_thanh.py                  tạo và trộn hiệu ứng âm thanh theo sự kiện; nhạc nền và hạ nhạc
tools/vi/video_ma_parts/karaoke.py                   file .ass với \kf theo mốc từ
tools/vi/tim_nhac.py                                 tìm và tải nhạc Openverse, ghi nhac/nguon.json
```

Sửa: `parse.py`, `kiem.py` (cú pháp và giới hạn mới), `giong.py` (WordBoundary, giọng lời giải), `lich.py` (mốc từ, thời lượng cảnh câu hỏi, dữ liệu cảnh), `trang.py`, `chup.py` (đọc `suKien`), `ghep.py` (trộn hiệu ứng, nhạc, karaoke), `video_ma.py`, `khung-video.js` (`catDanhDau` hiểu `==`, `((`, `__`, `{{}}`), tài liệu và test.

## 5. Giới hạn chữ mới

| Trường | Giới hạn |
|---|---|
| `bieu-do` nhãn | ≤ 16 ký tự; giá trị là số (dấu chấm thập phân), 2–8 dòng; `tron` chỉ nhận số dương |
| `cau-hoi` câu hỏi | ≤ 160 |
| `cau-hoi` mỗi lựa chọn | ≤ 60 |
| `cau-hoi` giải thích | ≤ 180 |
| `so-do` trung tâm | ≤ 30 |
| `so-do` mỗi nhánh | ≤ 40 |
| `dong-thoi-gian` nhãn mốc | ≤ 12 |
| `dong-thoi-gian` mô tả | ≤ 60 |
| Cụm nhấn | không lồng nhau; tối đa 3 cụm mỗi trường |

## 6. Kiểm thử

- **Node** (hàm thuần):
  - easing và spring xác định, bắt đầu tại 0 và kết thúc tại 1;
  - số chạy đi qua giá trị đầu và giá trị cuối;
  - `catDanhDau` với `==`/`((`/`__`/`{{}}` cho đúng chữ hiển thị và thoát ký tự đặc biệt;
  - từng kiểu chuyển cảnh: tại 0 s che toàn bộ bằng nền cũ, tại 0,5 s không còn;
  - các loại cảnh mới: mục thứ k bắt đầu tại mốc câu k, không mục nào vượt cuối cảnh, cảnh 2,5 s vẫn đúng;
  - `suKien()` xác định và nằm trong cảnh.
- **Python:**
  - parse và kiem cho mọi cú pháp mới (lỗi nêu đúng dòng);
  - mốc từ ước lượng khi dùng file giọng có sẵn;
  - thời lượng cảnh câu hỏi;
  - file `.ass` hợp lệ với `\kf` cộng lại bằng thời lượng câu;
  - lệnh trộn âm thanh (lệnh giả), tổng thời lượng không đổi;
  - `tim_nhac.py` với HTTP giả (không lên mạng trong test);
  - nhạc thiếu nguồn thì lỗi `canh`.
- **Chromium:**
  - mọi loại cảnh mới ở giới hạn tối đa không tràn chữ và ở trên y = 620;
  - cụm nhấn chưa hiện trước mốc và đã hiện sau mốc;
  - khung đầu của mỗi kiểu chuyển cảnh khớp nền cũ ở vùng chưa lộ.
- **FFmpeg thật:**
  - tiếng cảnh có hiệu ứng vẫn không có đoạn im lặng tuyệt đối;
  - đỉnh hiệu ứng thấp hơn đỉnh giọng ít nhất 12 dB;
  - nhạc nền bị hạ khi có giọng: mức nhạc trong đoạn có giọng thấp hơn đoạn không giọng ít nhất 6 dB;
  - phụ đề karaoke in được.
- **Tích hợp:** dựng một video có đủ loại cảnh mới, hiệu ứng âm thanh và nhạc nền giả (sine), rồi kiểm bằng `ffprobe`.
- **Bằng mắt và tai:** dựng lại demo con lắc đơn có câu hỏi, sơ đồ, biểu đồ, nhạc thật và giọng thật; chủ repo xem và nghe.

## 7. Ngoài phạm vi

Kiểu Vox và khổ dọc 9:16 (vi.12); bản đồ; 3D; hạt và pháo giấy; shader WebGL; tự động chọn nhạc theo cảm xúc; ghép video quay thật.

## 8. Rủi ro

- **Mốc từ ước lượng:** với file giọng thầy cô tự thu, nhấn ý và karaoke có thể lệch vài trăm mili giây. Công cụ sẽ ghi cảnh báo.
- **Hiệu ứng âm thanh tự tạo** có thể nghe "máy". Chủ repo cần nghe thử; có khoá `am-thanh: khong` để tắt.
- **Openverse:** nhiều bản ghi là bản xem trước chất lượng vừa phải. Nhạc CC BY bắt buộc ghi tác giả. Công cụ chỉ nhận CC0 và CC BY.
- **Sai lệch nội dung câu hỏi:** công cụ chỉ đảm bảo đáp án là một trong các lựa chọn; tính đúng của câu hỏi do thầy cô chịu trách nhiệm.
- **Gói lớn:** chia hai đợt trong kế hoạch. Đợt A là hình (Q2–Q7, Q12); đợt B là câu hỏi, âm thanh và nhạc (Q1 dùng chung, Q8–Q11). Cả hai phát hành cùng lúc.
