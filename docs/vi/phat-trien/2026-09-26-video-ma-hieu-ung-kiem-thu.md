# Biên bản kiểm thử: hiệu ứng giáo dục cho video giải thích (v6.3.2-vi.11)

Ngày 2026-09-26. Máy chủ repo: Windows 11, 6 lõi, Chromium headless shell của Playwright, FFmpeg bản Gyan 8.1.2. Nhánh `feat/vi-video-hieu-ung`. Spec: `docs/vi/phat-trien/2026-09-26-video-ma-hieu-ung-design.md`. Kế hoạch: `docs/vi/phat-trien/2026-09-26-video-ma-hieu-ung-plan.md`.

## 1. Ba lần chạy test cuối

Chạy sau commit tài liệu, trên máy chủ repo:

- `venv\Scripts\python.exe -m unittest discover -s tools/vi/tests`: **1078 test, OK, bỏ qua 62, 79,6 giây** (các test cần Chromium tự bỏ qua vì venv chính không có playwright).
- `C:/Users/ADMIN/vmt/v/Scripts/python.exe -m unittest discover -s tools/vi/tests -p "test_video_ma_*.py"`: **402 test, OK, 369,6 giây** (Chromium và FFmpeg chạy thật, không mở cửa sổ; chạy cùng lúc với bộ venv).
- `node --test` từng file trong `tools/vi/tests/js/`: **116 test, pass 116, fail 0**.

| File | Kết quả |
|---|---|
| `test_am_thanh.js` | 8/8 |
| `test_bieu_do.js` | 10/10 |
| `test_canh.js` | 24/24 |
| `test_cau_hoi.js` | 8/8 |
| `test_chuyen_dong.js` | 34/34 |
| `test_dong.js` | 17/17 |
| `test_khung.js` | 15/15 |

Lúc chạy ba lần này, cây làm việc còn thay đổi chưa commit của một việc song song (sửa dấu câu của phụ đề karaoke trong `karaoke.py`, `giong.py` và `test_video_ma_karaoke.py`); kết quả trên đã gồm các thay đổi đó.

Test lớp Việt mới (`ExplainerEffectsDocsTest` trong `test_vi_layer.py`):

- hai file hướng dẫn cho AI nêu đủ `==`, `((`, `__`, `{{`, `cau-hoi`, `bieu-do`, `so-do`, `dong-thoi-gian`, `tim_nhac.py`, `nhac-nen`, `am-thanh`, `karaoke`, `luan-phien`;
- bảng khoá đầu có mọi khoá và mọi giá trị trong `parse.META_CHOICES`, `META_FREE`, `META_REQUIRED`; `karaoke` là mặc định;
- bảng chuyển cảnh có mọi giá trị của `chuyen-canh`, và `SCENE_KIEU_CHUYEN` đúng bằng các giá trị đó trừ `luan-phien`;
- mọi kịch bản mẫu bắt đầu bằng `---` trong hai file hướng dẫn và tài liệu cho thầy cô chạy qua `parse` và `kiem` thật (ảnh PNG/JPEG và nhạc WAV giả tạo theo kịch bản, kèm `nhac/nguon.json` khi không có `nguon-nhac:`); mọi cảnh mẫu trong `canh-video.md` (đủ 14 loại) cũng vậy, sau khi gắn một khối thông tin và một cảnh mở đầu;
- ví dụ trong hướng dẫn có `cau-hoi`, `bieu-do`, `so-do`, `dong-thoi-gian`, bốn loại dấu nhấn, `chuyen:`, `nhac-nen` và `luan-phien`;
- bước `error.step` của `tim_nhac.py` đọc từ mã nguồn (`mang`, `input`, `write`) khớp mục "Tìm nhạc nền thất bại" của Xử lý lỗi;
- `AGENTS.vi.md` mục 15, luật Antigravity và tài liệu cho thầy cô có bước tìm nhạc và lời nhắc nghe thử.

Đã thử phá một kịch bản mẫu (bỏ dấu đóng `==`): test báo đúng `ParseError` ở dòng đó.

Luật `.agents/rules/ppt-master-vi.md`: **7.208 ký tự** (`len` của Python trên nội dung UTF-8), dưới ngưỡng 12.000; bảng loại việc không đổi, vẫn khớp `quy-trinh-hoi.md`.

## 2. Demo giọng thật

Thư mục `projects/_video/con-lac-don-demo4/` (không commit), dựa trên demo3:

- 11 cảnh: tiêu đề có `((chu kì))`, ảnh thật, khái niệm có `==ngắn nhất==` và `__một dao động toàn phần__`, minh hoạ, quy trình có `{{10}}`, công thức từng phần có `{{2.01}}`, ý từng ý có ba cụm nhấn, biểu đồ đường 5 điểm (0,25 → 1,0 … 2,0 → 2,84), thí nghiệm con lắc, sơ đồ tư duy 4 nhánh, câu hỏi 4 lựa chọn (đáp án B, có `loi-giai`).
- `chuyen-canh: luan-phien`, phụ đề karaoke (mặc định), tiếng hiệu ứng (mặc định).
- Nhạc: `tim_nhac.py "calm piano"` tải `calm-piano.mp3`, "Calm Piano Melody – Last hope" · Robbnix · CC BY 4.0 (84 giây, 1,9 MB); dòng nguồn hiện 4 giây cuối.
- Giọng edge-tts thật (`giong: may`), 0 cảnh báo, tức mốc từng từ thật ở mọi cảnh.
- Kết quả: `ready: true`, **155,93 giây**, **12.124.435 byte (12,1 MB)**, 1280×720, 30 khung/giây, 4.678 khung, 3 tiến trình Chromium. Dựng lần đầu **335 giây** kể cả tạo giọng; dựng lại sau một lần sửa **172 giây** (giọng đã có).

Mức tiếng đo trên bản dựng lại có giữ WAV trung gian (tách lớp: hiệu ứng = tiếng cảnh − giọng; nhạc = tiếng cuối − tiếng cảnh):

| Tín hiệu | Mức |
|---|---|
| Giọng (kèm nhiễu nền) | đỉnh −2,5 dBFS |
| Hiệu ứng | đỉnh −22,5 dBFS, thấp hơn đỉnh giọng đúng 20,0 dB ở cả 11 cảnh |
| Nhạc sau khi hạ | đỉnh −22,4 dBFS; RMS −54,1 dBFS lúc có giọng, −39,7 dBFS lúc không giọng, tức hạ khoảng 14 dB |
| Tiếng cuối | đỉnh −2,5 dBFS, RMS −19,4 dBFS |

Không có cửa sổ 50 ms nào im lặng tuyệt đối trong tiếng MP4; lớp nhiễu nền của vi.10 vẫn giữ.

Đã xem `--xem-truoc` cả 11 cảnh và 16 khung trích từ video. Khung tiêu biểu trong `projects/_video/con-lac-don-demo4/khung/`: tô vàng "dài hơn" nổ đúng lúc karaoke tới "dài hơn" (giây 67,2); biểu đồ đường đang mọc, nhãn số đang chạy (giây 89,0); sơ đồ 4 nhánh trọn khung, phụ đề không đè (giây 122,0); đồng hồ đếm ngược ở số 3 (giây 143,5); đáp án B viền xanh có ✓, các lựa chọn khác mờ, giải thích đang viết (giây 148,5); lật trang giữa chừng (giây 19,75).

Tiêu chí tốc độ (spec mục 1.9, video 5 phút dựng không quá 8 phút): bản dựng lại 172 giây cho 155,93 giây video, khoảng 1,1 lần thời lượng; suy ra video 5 phút khoảng 5,5 phút chưa kể tạo giọng. Đây là số suy ra, chưa dựng video 5 phút thật.

## 3. Fixture và đo bằng FFmpeg thật

- Fixture `tools/vi/fixtures/video-hieu-ung/` (7 cảnh, 20 KB, dựng không cần mạng): mọi loại cảnh mới, ba dấu nhấn, `{{2.01}}`, `{{4}}`, `{{2}}`, `luan-phien` đi qua đủ năm kiểu ở cảnh 2–6, cảnh 7 ghi `chuyen: lat-trang`, `am-thanh: co`, nhạc là WAV sine 150 Hz tự tạo kèm `nguon-nhac:`. Test `test_video_ma_hieu_ung.py` dựng thật hai lần (`am-thanh: co` và `khong`): 1280×720, 30/1, có tiếng; thời lượng 34,6 giây đúng công thức; tiếng hiệu ứng nghe được (sau lọc thông cao 1,2 kHz: −57,3 so với −71,1 dBFS); nhạc bị hạ khi có giọng (−46,6 dBFS lúc có giọng so với −33,1 dBFS lúc đếm ngược); không có đoạn 50 ms im tuyệt đối.
- Hiệu ứng trên fixture `video-hinh` (giọng sine): đỉnh hiệu ứng −25,88 dBFS so với giọng −5,57 dBFS, thấp hơn 20,3 dB (yêu cầu ít nhất 12 dB); thời lượng không đổi (27,6 giây).
- Hạ nhạc với giọng sine 440 Hz và nhạc sine 2000 Hz: nhạc lúc có giọng −57,5 dBFS, lúc không giọng −33,5 dBFS, hạ khoảng 24 dB (yêu cầu ít nhất 6 dB); nhạc 5 giây lặp đủ video 12,000 giây.
- Cảnh câu hỏi dựng thật với mp3 sine: giọng lời giải bắt đầu ở 10,500 giây, đúng đầu cảnh cộng `bat_dau_giai`; lúc đếm ngược không có phụ đề.
- Phụ đề karaoke in bằng libass: chiều cao chữ đo được từ 13 px (lỗi, xem mục 4) lên đạt ngưỡng ít nhất 28 px, lệch không quá 25% so với phụ đề `hinh`.

## 4. Lỗi tìm ra trong lúc làm và đã sửa

Mỗi mục có test trượt trước khi sửa và đạt sau khi sửa.

- **Mốc câu trôi dần** (`a628dde2`): bản đầu chia từ của edge-tts theo số từ của câu, nên số "1500" đọc thành bốn từ làm lệch mọi câu sau mà không báo gì. Nay so chữ của lời với chữ edge-tts trả về; không chắc thì ước lượng kèm một cảnh báo gộp cho cảnh đó.
- **Vòng khoanh đè chữ bên cạnh** (`a5d6ea77`): elip và cú nảy 1,12 lần lấn sang từ bên cạnh. Nay elip ôm sát hộp dòng, mỗi dòng một elip, biên độ nảy tính theo khoảng trống thật; cụm có ngoặc lệch và dấu `**`/`~`/`^` cắt ngang ranh giới cụm là lỗi `parse`.
- **Test lỗi song song chập chờn** (`7e1c4b18`): báo dải chụp hỏng sớm nhất, không phải dải hỏng nhanh nhất.
- **Chữ karaoke quá nhỏ** (`47d2cc9e`): file `.ass` khai khổ 1280×720 nhưng cỡ chữ vẫn theo khổ 384×288, chữ chỉ cao 13 px. Nhân các số theo tỉ lệ 2,5 (cỡ 40, viền 3, lề dưới 55).
- **Thiếu file sự kiện âm thanh** (`708e71b6`): cảnh thiếu hoặc hỏng file sự kiện nay dựng không có hiệu ứng thay vì hỏng cả video.
- **Tải nhạc qua giao thức không an toàn** (`b462bdfc`): `tim_nhac.py` chỉ nhận địa chỉ https lấy từ Openverse.
- **Sơ đồ tư duy trượt xuống vùng phụ đề** (`de19c71e`): máy quay phóng vào nhánh trên bên trái đẩy các ô dưới xuống y 590, phụ đề đè lên. Nay mọi mục của `so-do` giữ máy quay đứng yên như biểu đồ; test Node kiểm 2–6 nhánh ở 321 thời điểm, không mục nào xuống dưới y 620.

## 4b. Sửa sau review toàn nhánh

Mỗi mục có test trượt trước khi sửa và đạt sau khi sửa.

- **Kịch bản NFD** (`0b2a4b88`): Unikey "Unicode tổ hợp", Mac hay PDF cho chữ NFD; `\w` của Python bỏ dấu tổ hợp nên khoá của "kì" thành "ki", còn `nhan.js` dùng "kì", cụm nhấn rơi về "không tìm thấy". Nay `parse` chuẩn hoá cả `video.md` về NFC một lần khi đọc, và khoá so khớp là một hàm dùng chung (`lich.khoa_so_khop`, cũng là `giong._chuan_hoa`) tự chuẩn hoá NFC. Test: khoá của từ NFD bằng khoá NFC; `nhan.js` chạy bằng Node trên mốc từ do Python tạo tìm cụm NFD đúng giờ của từ; mốc câu khớp khi kịch bản NFD mà chữ edge-tts NFC; tiền tố "Nhạc:" gõ NFD không bị thêm lần nữa.
- **Gạch ngược trong karaoke** (`ec3f10c4`): libass không có thoát cho `\`, nên `a\Nb`, `\h` trong chữ thầy cô thành mã điều khiển. Nay đổi thành ⧵ (U+29F5).
- **Dấu thô trong phụ đề `hinh`/`file`** (`ec3f10c4`): chỉ bỏ `**`, `~`, `^`; nay bỏ đủ bộ dấu như giọng đọc.
- **Hiệu ứng lấn giọng nhỏ** (`8fc60573`): sàn −50 dBFS làm hiệu ứng chỉ thấp hơn giọng −45 dBFS 5 dB; nay sàn −80 dBFS, chỉ chạm tới ở cảnh gần như im lặng.
- **Chuyển hướng sang http** (`33f93155`): `tim_nhac.py` kiểm địa chỉ cuối trước khi đọc nội dung, không phải https thì dừng với `error.step: "mang"`.
- **Cách sửa lỗi nhạc** (`c7d7d7fe`): lỗi nhạc nền (Cảnh 0) và dòng nguồn nhạc tràn khung có `fix` riêng; ffprobe chỉ đo nhạc một lần mỗi lượt dựng.
- Tài liệu: bỏ mục `magnt`→`magnet` khỏi vi.11 (đã phát hành ở vi.10, `2374b95c`); thêm rủi ro dự án vi.10 dùng mốc ước lượng tới khi xoá `giong/*.json`; ghi rõ dòng nguồn nhạc hiện suốt cảnh cuối khi cảnh đó ngắn hơn 4 giây.

Ba lần chạy test sau các sửa này:

- `venv\Scripts\python.exe -m unittest discover -s tools/vi/tests`: **1092 test, OK, bỏ qua 62, 55,5 giây**.
- `C:/Users/ADMIN/vmt/v/Scripts/python.exe -m unittest discover -s tools/vi/tests -p "test_video_ma_*.py"`: **415 test, OK, 264,7 giây**.
- `node --test` từng file: **117 test, pass 117, fail 0** (`test_am_thanh.js` 8/8, `test_bieu_do.js` 10/10, `test_canh.js` 24/24, `test_cau_hoi.js` 8/8, `test_chuyen_dong.js` 34/34, `test_dong.js` 18/18, `test_khung.js` 15/15).

## 5. Chưa kiểm

Các mục sau **chưa** được kiểm, không suy diễn là đạt:

- **Phiên Antigravity thật với luật mới** (có thêm bước tìm nhạc và nhắc nghe thử).
- **Cài từ đầu trên máy sạch**: máy kiểm đã có sẵn Playwright, Chromium và FFmpeg.
- **Thời gian dựng trên máy 2 lõi**.
- **Chủ repo nghe thử** tiếng hiệu ứng và nhạc nền của demo: các con số ở mục 2 là đo máy, chưa phải đánh giá bằng tai. Tiếng hiệu ứng tự tạo có thể nghe "máy".
- Tiêu chí lệch nhấn ý không quá 0,15 giây với giọng máy (spec mục 1.1): cụm nổ tại mốc từ của edge-tts và đã xem đúng ở demo, nhưng chưa đo độ lệch so với tiếng bằng số.

Các điểm còn để lại, không chặn phát hành:

- Phụ đề karaoke bỏ dấu câu và có thể để một từ lẻ ở dòng dưới; đang sửa trong một việc riêng, chưa thuộc biên bản này.
- Bàn tay có thể lấn vào vùng phụ đề khi viết giải thích ở cảnh câu hỏi.
- Hộp công thức giữ chữ ở phần trên, công thức một dòng nằm cao; gạch chân tiêu đề có bề rộng cố định, ngắn hơn tiêu đề dài (có từ vi.9 và vi.10).
- Máy quay chỉ giữ mục đang nhắm ở trên y 620; loại cảnh khác có nội dung thấp về lý thuyết vẫn có thể bị đẩy xuống dưới phụ đề khi phóng; mới thấy ở `so-do` và đã sửa.
- Nhạc được hạ theo cả tiếng chính (giọng, nhiễu nền, hiệu ứng), nên một tiếng "ting" cũng làm nhạc nhỏ đi một chút.
- Dòng nguồn nhạc chỉ nằm ở cảnh cuối; cảnh cuối ngắn hơn 4 giây thì dòng nguồn hiện suốt cảnh đó.
- Elip khoanh chữ chỉ có khoảng 2,6 px khoảng hở trên dưới nên chạm nhẹ dấu thanh cao ở hai đầu cụm.
- Thời gian chạy bộ test Chromium dao động theo tải máy (một lần 858 giây so với khoảng 250 giây thường lệ, khi máy đang chạy việc khác).
