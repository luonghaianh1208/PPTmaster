# Hiệu ứng giáo dục cho video giải thích — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development. Steps use checkbox (`- [ ]`) syntax.

**Goal:** Thêm nhấn ý chính, số động, biểu đồ, sơ đồ tư duy, dòng thời gian, năm kiểu chuyển cảnh, chữ động, phụ đề karaoke, cảnh câu hỏi nhanh, hiệu ứng âm thanh tự tạo và nhạc nền cho `tools/vi/video_ma.py` (bản v6.3.2-vi.11).

**Architecture:** Giữ nguyên đường ống vi.10 (parse → kiem → giong → lich → trang → chup → ghep). Có năm thay đổi chính:
- Giọng máy trả mốc từng từ.
- Runtime có thêm các module thuần `dong.js`, `nhan.js`, `chuyen-canh.js` và bốn loại cảnh mới.
- Runtime xuất danh sách sự kiện âm thanh `suKien()`.
- `ghep.py` trộn hiệu ứng âm thanh, nhạc nền và phụ đề karaoke.
- Công cụ mới `tim_nhac.py` tải nhạc giấy phép mở.

**Tech Stack:** Python thư viện chuẩn + edge-tts + playwright (lười); JS thuần; FFmpeg/libass.

**Spec:** `docs/vi/phat-trien/2026-09-26-video-ma-hieu-ung-design.md` (có quyền quyết định cuối; Q1–Q13 là các quyết định bắt buộc).

**Cách viết:** mỗi task nêu giao diện và khẳng định test cụ thể. Người thực hiện tự viết mã theo TDD: test trước, thấy trượt, rồi viết mã. Tên và số trong kế hoạch là bắt buộc.

## Global Constraints

- Không sửa `skills/` (chỉ đọc), `tools/vi/video.py`, `tools/vi/video_parts/`, `tools/vi/thi_nghiem_parts/`, `LICENSE`, `SPONSORS*`, `requirements.txt`, `attribution_guard.py`. Không commit gì trong `projects/`.
- Python chỉ thư viện chuẩn + `edge-tts` + `playwright` (import lười). JS thuần, không thư viện, không địa chỉ web trong trang. Mọi hiệu ứng là hàm xác định của `t`.
- Kịch bản vi.9 và vi.10 dựng được mà không phải sửa. Fixture `tools/vi/fixtures/video-mau/` và `video-hinh/` không đổi, vẫn qua test.
- Khoá đầu mới:
  - `chu-dong: co|khong` (mặc định `co`);
  - `am-thanh: co|khong` (`co`);
  - `nhac-nen: <file>` (không có thì không có nhạc);
  - `nguon-nhac: <chữ>`.
- Khoá có thêm giá trị mới:
  - `chuyen-canh` thêm `lat-trang|truot|phong|mo-man|luan-phien`;
  - `phu-de` thêm `karaoke` và `karaoke` trở thành mặc định (kịch bản cũ ghi rõ `hinh` thì giữ kiểu cũ).
- Trường cảnh mới: `chuyen:` (mọi loại cảnh, ghi đè khoá đầu).
- Mọi CLI in đúng một dòng JSON; tập `error.step` giữ nguyên. Test không bao giờ gọi mạng (edge-tts, Openverse) thật.
- Không có đoạn im lặng tuyệt đối trong tiếng. Hiệu ứng âm thanh thấp hơn giọng ít nhất 12 dB ở đỉnh. Nhạc nền bị hạ khi có giọng.
- Test:
  - `venv\Scripts\python.exe -m unittest discover -s tools/vi/tests`;
  - `C:/Users/ADMIN/vmt/v/Scripts/python.exe -m unittest discover -s tools/vi/tests -p "test_video_ma_*.py"` (Chromium + FFmpeg thật, không mở cửa sổ);
  - `node --test tools/vi/tests/js/*.js` (liệt kê từng file trên Windows).
- Trailer commit: `Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>`. Không push.

## Review Focus

1. **Cụm nhấn trùng từ xuất hiện nhiều lần trong lời, hoặc không có trong lời, hoặc có dấu khác lời** ("Chu kì" vs "chu kỳ"). Hiệu ứng nổ ở lần xuất hiện đầu tiên sau khi mục bắt đầu viết; không tìm thấy thì nổ 0,3 s sau khi viết xong. Pin ở Task 2.
2. **Cảnh câu hỏi mà giọng lời giải lỗi mạng hoặc thầy cô chỉ đặt sẵn file câu hỏi:** báo `giong` nêu đúng file. `canh-N-giai.mp3` có sẵn thì dùng, không ghi đè. Pin ở Task 6.
3. **Biểu đồ có giá trị âm, 0 hoặc rất chênh lệch** (1 và 100000). Cột âm vẽ xuống dưới trục; biểu đồ tròn từ chối số không dương; nhãn giá trị không chồng nhau. Pin ở Task 4.
4. **Nhạc ngắn hơn video, file nhạc hỏng, thiếu nguồn, tên có dấu:** nhạc lặp đủ dài; file hỏng hoặc thiếu nguồn báo lỗi `canh` với tên file. Pin ở Task 8.
5. **Video có mọi hiệu ứng bật trên cảnh 2,5 giây:** không mục nào, không sự kiện âm thanh nào vượt cuối cảnh; tổng tiếng bút ≤ 40% thời gian cảnh. Pin ở Task 7.

---

### Task 1: Mốc từng từ

**Files:** sửa `giong.py`, `lich.py`, `video_ma.py` (nếu cần); test `test_video_ma_giong.py`, `test_video_ma_lich.py`.

**Interfaces:**
- `giong.tong_hop_edge(text, voice, rate, out_path) -> dict` trả `{"cau": [giây...], "tu": [{"t": giây, "d": giây, "chu": str}]}` dùng `boundary="WordBoundary"`. Mốc câu = thời điểm của từ đầu mỗi câu (tách câu bằng `lich.tach_cau` trên văn bản đã bỏ đánh dấu, ghép từ theo thứ tự). Giữ tương thích: `lay_giong` nhận hàm tổng hợp trả `list` (kiểu cũ, chỉ mốc câu) hoặc `dict`.
- Sổ giọng `canh-N.json` thêm `"tu"`. Sổ cũ không có `tu` vẫn dùng được, mốc từ khi đó được ước lượng.
- `lich.GiongInfo` thêm trường `moc_tu: list` (mặc định `[]`) và `uoc_luong_tu: bool`.
- `lich.moc_tu_uoc_luong(loi, moc_cau_giong, giay) -> list[dict]`: chia mỗi câu theo tỉ lệ ký tự của từng từ.
- `lich.CanhLich` thêm `moc_tu` (thời gian cảnh, đã cộng `DAN_DAU`).
- `du["tu"] = [{"t", "d", "chu", "khoa"}]`, trong đó `khoa` là chữ thường, bỏ dấu câu, giữ dấu thanh.
- Văn bản gửi tới giọng máy bỏ mọi đánh dấu mới `==`, `((`, `))`, `__`, `{{`, `}}` (giữ nội dung bên trong).

- [ ] **Tests:**
  - FakeTts trả dict: mốc câu khớp từ đầu mỗi câu.
  - Sổ ghi `tu`; lần dựng sau dùng lại.
  - Sổ cũ không có `tu` → `uoc_luong_tu=True` kèm cảnh báo.
  - File giọng có sẵn → mốc từ ước lượng, tăng dần, nằm trong thời lượng giọng.
  - Bỏ đánh dấu trước khi gửi giọng (FakeTts nhận chữ sạch); mã băm tính trên chữ sạch.
  - `du["tu"]` đã cộng 1,0 s.
- [ ] Commit `feat(vi): word timings for explainer narration`.

### Task 2: Nhấn ý chính, số chạy, chữ động

**Files:** tạo `runtime/dong.js`, `runtime/nhan.js`; sửa `khung-video.js` (`catDanhDau`, đo hộp cụm), `canh/tieu-de.js`, `trang.py` (thứ tự nạp: `dong.js`, `khung-video.js`, `nhan.js`, …), `kiem.py` (luật cụm), `viet-tay.css`; test `tools/vi/tests/js/test_dong.js` (mới), bổ sung `test_canh.js`, `test_video_ma_parse.py`, Chromium.

**Interfaces:**
- `THI_DONG` (thuần):
  - `easeOutBack(x)`, `easeInOut(x)`, `lo_xo(x, cung=12, tat=0.35)`: nhận 0..1, trả 0 tại 0 và 1 tại 1.
  - `soChay(t, batDau, dai, gtri, chuSo)`: trả chuỗi số kiểu Việt, dấu phẩy thập phân, giữ số chữ số thập phân của giá trị gốc.
  - `nayChu(i, n, t, batDau)` → `{s, y, a}`.
- `THI_NHAN` (thuần):
  - `tachCum(chu)` → danh sách `{kieu: 'to'|'khoanh'|'gach', noiDung, viTri}`.
  - `thoiDiemNhan(cum, tu, batDauMuc, xongMuc)`: so khớp dãy khoá của cụm với `du.tu` từ `batDauMuc` trở đi; trả `t` của từ đầu, hoặc `xongMuc + 0.3` nếu không có.
- `catDanhDau(chu, n)`:
  - hiểu `==x==`, `((x))`, `__x__`, `{{số}}`;
  - đánh dấu và ký tự đặc biệt không đếm vào số ký tự hiện;
  - cụm bọc `<span class="cum to|khoanh|gach" data-cum="k">`;
  - số bọc `<span class="so" data-so="k">`;
  - `**`, `~`, `^` và cách thoát ký tự vẫn như cũ.
- Hiệu ứng khi nổ (trong 0,45 s):
  - `to`: nền vàng `#fde047` quét trái→phải phía sau chữ (dùng `background-size`);
  - `khoanh`: một vòng elip nét đỏ vẽ tay quanh hộp cụm (SVG path, `pathLength=1`);
  - `gach`: nét gạch dưới vẽ dần;
  - cụm nảy 1 → 1,12 → 1 nếu `chu-dong`.
- Số `{{…}}` chạy 0 → giá trị trong 0,8 s kể từ khi chữ tới vị trí số. Trong lúc chạy vẫn giữ bề rộng (dùng `min-width` theo chữ số cuối) để bố cục không nhảy.
- `tieu-de` khi `du.co.chuDong`: chữ `chu` nảy từng ký tự (`nayChu`, lệch nhau 0,04 s); không có bàn tay cho mục này (mục có `tay:false`).
- `du["co"]["chuDong"]` lấy từ khoá đầu (Task 1 hoặc Task này thêm vào `lich.du_lieu_canh`).
- `kiem`: cụm lồng nhau, cụm không đóng, quá 3 cụm mỗi trường, `{{x}}` không phải số (dấu chấm thập phân) → `ParseError` hoặc `CanhError` nêu đúng dòng. Giới hạn ký tự tính trên chữ hiển thị.

- [ ] **Tests Node:**
  - hàm làm mượt;
  - `soChay` tại batDau ("0"), giữa, và sau khi xong ("1500,5": không dấu phân cách nghìn, dấu phẩy thập phân, như `dinhDang` của thí nghiệm ảo);
  - `thoiDiemNhan`: khớp lần đầu sau `batDauMuc`; so khớp không phân biệt hoa thường, bỏ dấu câu nhưng **giữ dấu thanh** ("Chu kì" khớp "chu kì", "chu kỳ" không khớp "chu kì"); không có trong lời → `xongMuc + 0.3`;
  - `catDanhDau` hiển thị đúng chữ, thoát `<`, không lồng thẻ sai khi cắt giữa cụm.
- [ ] **Tests Chromium:**
  - khung trước thời điểm nhấn không có nền vàng hay vòng khoanh, khung sau có (đọc style hoặc pixel);
  - cụm dài nhất ở giới hạn ký tự không tràn;
  - tiêu đề nảy: chữ hiển thị đủ ở cuối cảnh;
  - số `{{1500}}` hiện "1500" ở cuối cảnh.
- [ ] Commit `feat(vi): highlight key ideas when the narrator says them`.

### Task 3: Năm kiểu chuyển cảnh

**Files:** tạo `runtime/chuyen-canh.js`; sửa `khung-video.js` (thay đoạn lau bảng cứng bằng `THI_CHUYEN`), `ban-tay.js` (giẻ chỉ khi `lau-bang`), `parse.py` (giá trị mới, trường `chuyen:`), `lich.py` (`du["co"]["chuyen"]` là chuỗi kiểu hoặc `null`; `luan-phien` xoay vòng `lau-bang, lat-trang, truot, phong, mo-man` theo số cảnh; cảnh 1 luôn `null`; giữ `lauBang` tương thích = `chuyen === 'lau-bang'`); test Node `test_chuyen_dong.js`, Chromium, `test_video_ma_lich.py`, `test_video_ma_parse.py`.

**Interfaces:** `THI_CHUYEN.trangThai(kieu, t, dai)` → `{nen: {transform, opacity, clipPath}, moi: {transform, opacity, clipPath}, loe: số 0..1}`:
- `lau-bang`: như vi.10;
- `lat-trang`: nền xoay quanh mép trái `rotateY` 0 → −100° với bóng tối dần;
- `truot`: nền `translateX` 0 → −1280, mới từ +1280 → 0;
- `phong`: nền `scale` 1 → 1,6 với opacity 1 → 0, cộng một lớp loé trắng đỉnh 0,6 tại giữa;
- `mo-man`: hai nửa nền tách sang hai bên.

Mọi kiểu trong `[0, 0,5]` s. Tại `t=0` khung hình bằng nền cũ ở vùng hiển thị; tại `t ≥ 0,5` không còn nền cũ.

- [ ] **Tests:** Node `trangThai` xác định và đúng biên cho cả năm kiểu; `luan-phien` cho đúng dãy; Chromium mỗi kiểu tại `t = 1/30` có ≥ 90% điểm ảnh vùng giữa khớp khung cuối cảnh trước và tại 0,55 s không còn; khoá `chuyen-canh: khong` và cảnh 1 → không nền cũ.
- [ ] Commit `feat(vi): page turn, slide, zoom-through and curtain scene transitions`.

### Task 4: Biểu đồ, sơ đồ tư duy, dòng thời gian, công thức từng phần

**Files:** tạo `runtime/canh/bieu-do.js`, `so-do.js`, `dong-thoi-gian.js`; sửa `canh/cong-thuc.js`, `parse.py`, `kiem.py`, `lich.py` (`so_muc`, dữ liệu `du["duLieu"]` dạng `[[nhãn, số]]`), `trang.py`; test như các task trước.

**Interfaces và giới hạn:** theo spec §5 và Q3, Q4, Q12.
- `bieu-do`:
  - `cot`: cột mọc từ trục; cột âm mọc xuống; nhãn giá trị ở đầu cột, số chạy.
  - `duong`: điểm và đoạn nối vẽ dần.
  - `tron`: các lát mở dần theo góc; nhãn phần trăm ở ngoài lát.
  - Phần tử k bắt đầu tại `du.moc[k]` (đã có cách chia đều khi ít câu).
  - Trục và thang số: tự tính mốc tròn (1, 2, 5 × 10ⁿ), 4–6 vạch.
- `so-do`: ô trung tâm (có thể kèm `hinh`) ở giữa. 2–6 nhánh rải quanh theo góc cố định theo số nhánh. Mỗi nhánh là đường cong `duongQua` tới ô nhãn; nhánh k bắt đầu tại `moc[k]`.
- `dong-thoi-gian`: trục ngang vẽ trước (0,6 s). Mốc k có chấm, nhãn trên và mô tả dưới, xen kẽ cao thấp nếu > 4 mốc.
- `cong-thuc`: `bieu-thuc` có ` | ` → các phần hiện nối tiếp; phần k tại `moc[k]`. Không có ` | ` thì như cũ.

- [ ] **Tests:**
  - parse/kiem cho mọi trường và lỗi (giới hạn, số sai, `tron` có số ≤ 0, số dòng ngoài khoảng);
  - Node: phần tử k tại mốc câu k; cảnh 2,5 s không vượt cuối; thang số tròn cho các bộ dữ liệu (0,3–7,2; −40–25; 1–100000);
  - Chromium: mỗi loại ở giới hạn tối đa (8 cột nhãn 16 ký tự; 6 nhánh 40 ký tự; 6 mốc) không tràn, ở trên y = 620; nhãn giá trị cột không chồng nhau (so hộp).
- [ ] Commit `feat(vi): animated charts, mind maps and timelines`.

### Task 5: Phụ đề karaoke

**Files:** tạo `karaoke.py`; sửa `parse.py` (`phu-de` thêm `karaoke`, mặc định `karaoke`), `ghep.py` (dùng `.ass` với bộ lọc `subtitles=.khung/phu-de.ass:fontsdir=.khung/fonts` khi `karaoke`); test `test_video_ma_karaoke.py`, `test_video_ma_ghep.py`.

**Interfaces:**
- `karaoke.tao_ass(cac_lich, rong=1280, cao=720) -> str`:
  - `[Script Info]` có `PlayResX/Y`, `WrapStyle: 0`.
  - Style `Itim`, `PrimaryColour` vàng `&H0000D7FF` cho phần đã đọc, `SecondaryColour` trắng, viền đen 2, `MarginV` 22.
  - Mỗi câu là một `Dialogue` với `{\kfN}` cho từng từ, N tính bằng centi-giây từ mốc từ; tổng N bằng thời lượng câu (±1 cs).
  - Thoát `{`, `}`, `\` trong chữ; bỏ đánh dấu.
- `ghep`: với `karaoke`, ghi `.khung/phu-de.ass`; với `file` vẫn ghi `phu-de.srt`.

- [ ] **Tests:** cấu trúc `.ass` hợp lệ (đủ ba mục); tổng `\kf` của mỗi dòng khớp thời lượng; chữ có `{}` được thoát; FFmpeg thật (`skipUnless`) in được file `.ass` có tiếng Việt vào một video màu 2 s, lệnh trả mã 0.
- [ ] Commit `feat(vi): karaoke subtitles that follow each spoken word`.

### Task 6: Cảnh câu hỏi nhanh

**Files:** tạo `runtime/canh/cau-hoi.js`; sửa `parse.py`, `kiem.py`, `giong.py` (giọng thứ hai), `lich.py` (thời lượng, mốc), `video_ma.py` (lấy giọng lời giải), `ghep.py` (tiếng cảnh ghép hai giọng); test tương ứng.

**Interfaces:**
- `parse`: loại `cau-hoi` với `cau-hoi`, `lua-chon` (2–4), `dap-an` (A–D, phải trong số lựa chọn), `giai-thich`, `cho` (3–10, mặc định 5), `loi`, `loi-giai` (bắt buộc).
- `giong.lay_giong(..., ten="canh-N")`: tổng quát hoá tên file; lời giải dùng `canh-N-giai`.
- `lich`:
  - `CanhLich` thêm `giay_giai` và `bat_dau_giai`.
  - Thời lượng = dẫn đầu + giọng câu hỏi + `cho` + 0,4 + giọng lời giải + 0,6, làm tròn lên bội 1/30.
  - `du["cauHoi"] = {"batDauDem", "cho", "batDauGiai", "dapAn"}`.
  - Phụ đề và karaoke gồm cả câu của lời giải, đúng mốc.
- Runtime:
  - câu hỏi và các lựa chọn viết theo mốc câu;
  - đồng hồ vòng tròn đếm ngược, cung giảm dần và số giây lớn;
  - tại `batDauGiai`: lựa chọn đúng viền xanh, có dấu ✓ và nảy; các lựa chọn khác mờ còn 0,35;
  - giải thích viết ra.
- `ghep`: tiếng cảnh = adelay giọng câu hỏi + adelay giọng lời giải tại `bat_dau_giai`, vẫn có nhiễu nền.

- [ ] **Tests:**
  - parse lỗi (thiếu `loi-giai`, `dap-an` E, `cho` 12);
  - thời lượng đúng công thức;
  - `canh-N-giai.mp3` có sẵn được dùng, không ghi đè;
  - giọng lời giải lỗi → `giong` nêu `canh-N-giai.mp3`;
  - Node: đáp án hiện tại `batDauGiai`, không sớm hơn;
  - Chromium: ở giới hạn tối đa không tràn; khung trước lúc giải không có dấu ✓, khung sau có;
  - FFmpeg thật: tiếng cảnh có giọng thứ hai đúng vị trí (±50 ms).
- [ ] Commit `feat(vi): quick quiz scenes with a countdown and spoken answer`.

### Task 7: Hiệu ứng âm thanh

**Files:** tạo `am_thanh.py`; sửa `khung-video.js` (`THI_VIDEO.suKien()`), các loại cảnh (gắn loại sự kiện), `chup.py` (đọc `suKien` của mỗi cảnh trong lượt kiểm tràn, trả về cùng kết quả), `video_ma.py`, `ghep.py`, `parse.py` (`am-thanh`); test `test_video_ma_am_thanh.py`, Node, CLI.

**Interfaces:**
- `THI_VIDEO.suKien()` → `[{t, loai, dai}]` với `loai`:
  - `but`: mục chữ hoặc nét đang vẽ, `dai` = thời lượng vẽ;
  - `ting`: mục ý, bước, nhánh, mốc, cột hiện ra;
  - `chuyen`: đầu cảnh có chuyển cảnh;
  - `tictac`: mỗi giây khi đếm ngược;
  - `dung`: lúc hiện đáp án;
  - `nhan`: lúc nổ nhấn.
- Sự kiện được sắp theo `t`, không trùng, và nằm trong `[0, thoiLuong − 0,1]`. Tổng `dai` của `but` ≤ 40% `thoiLuong`: gộp và cắt các đoạn chồng nhau rồi co lại nếu vượt.
- `am_thanh.tao_mau(thu_muc) -> dict[loai, Path]`: tạo WAV mẫu bằng `ffmpeg -f lavfi` từ công thức cố định (spec Q9), lưu `.khung/am/<loai>.wav`.
- `am_thanh.lenh_tron(wav_giong, su_kien, mau, wav_ra, thoi_luong)`: `adelay` từng sự kiện rồi `amix`. Mức hiệu ứng −20 dB so với giọng; `but` lặp hoặc cắt theo `dai`.
- `ghep`: khi `am-thanh: co`, tiếng cảnh = giọng (có nhiễu nền) trộn hiệu ứng.

- [ ] **Tests:**
  - Node: `suKien` xác định, trong cảnh, luật 40%, cảnh 2,5 s với mọi hiệu ứng (Review Focus 5);
  - Python: lệnh trộn (lệnh giả) có đủ đầu vào và độ trễ; `am-thanh: khong` → không trộn;
  - FFmpeg thật: đỉnh hiệu ứng thấp hơn đỉnh giọng ≥ 12 dB (dùng giọng sine −6 dBFS), không có đoạn im lặng tuyệt đối, thời lượng không đổi.
- [ ] Commit `feat(vi): pen, ting, transition and quiz sound effects`.

### Task 8: Nhạc nền và công cụ tìm nhạc

**Files:** tạo `tools/vi/tim_nhac.py`, `video_ma_parts/nhac.py`; sửa `parse.py` (`nhac-nen`, `nguon-nhac`), `kiem.py`, `ghep.py`, `lich.py` hoặc runtime (dòng nguồn nhạc 4 s cuối video ở cảnh cuối); test `test_video_ma_nhac.py`, `test_tim_nhac.py`.

**Interfaces:**
- `tim_nhac.py "<từ khoá>" -o <thư_mục_nhạc> [--so 3]`:
  - gọi `https://api.openverse.org/v1/audio/?q=...&license=cc0,by&page_size=20` bằng `urllib`;
  - lọc bản có thời lượng ≥ 60 s;
  - tải bản đầu (hoặc `--so` bản) vào `<thư_mục>/<slug>.mp3`;
  - thêm bản ghi `{filename, title, creator, license, license_url, source_url}` vào `nguon.json`;
  - in đúng một dòng JSON;
  - lỗi mạng → `error.step: "mang"`.
  - Hàm `tim(tu_khoa, lay=urllib.request.urlopen)` có thể thay bằng hàm giả trong test.
- `nhac.doc(thu_muc_du_an, ten, nguon_tay) -> {"duong_dan", "nguon", "giay"}`:
  - cùng luật chặn đường dẫn như `anh.doc` (chỉ trong `nhac/`, NFC);
  - nguồn từ `nhac/nguon.json` hoặc `nguon-nhac:`;
  - thiếu nguồn, không phải mp3/m4a/wav/ogg, hoặc ffprobe không đọc được → `NhacError` → `CanhError` (số cảnh 0, nêu dòng khoá đầu).
- `ghep`: sau khi nối tiếng các cảnh:
  - nhạc `-stream_loop -1`, cắt đúng tổng thời lượng;
  - `afade` vào/ra 1,5 s; `volume` −24 dB;
  - `sidechaincompress` (threshold 0.02, ratio 8, attack 20, release 400) với giọng làm tín hiệu điều khiển;
  - `amix` với tiếng chính (normalize 0).
- Dòng nguồn nhạc "Nhạc: <title> · <creator> · <license>" hiện 4 s cuối ở góc dưới trái, trên vùng phụ đề, cỡ 14 px.

- [ ] **Tests:**
  - `tim_nhac` với phản hồi giả: lọc thời lượng, ghi `nguon.json`, slug có dấu tiếng Việt thành không dấu, lỗi mạng → một dòng JSON;
  - `nhac.doc`: thoát đường dẫn, thiếu nguồn, tên có dấu;
  - FFmpeg thật: giọng sine ngắt quãng + nhạc sine khác tần số → mức nhạc trong đoạn có giọng thấp hơn đoạn không giọng ≥ 6 dB (đo bằng lọc thông dải quanh tần số nhạc); nhạc 5 s lặp đủ video 12 s; thời lượng đúng.
- [ ] Commit `feat(vi): background music with ducking and an Openverse finder`.

### Task 9: Dựng thật và demo

- [ ] Fixture mới `tools/vi/fixtures/video-hieu-ung/` dựng được không cần mạng:
  - dùng mọi loại cảnh mới, cụm nhấn, `{{số}}`, `luan-phien`, `am-thanh: co`;
  - nhạc là file WAV sine tự tạo nhỏ kèm `nguon-nhac:`.
- [ ] Test tích hợp có ba thứ:
  - dựng `--plan-only`;
  - dựng thật 2 tiến trình, kiểm `ffprobe`;
  - phát hiện được tiếng hiệu ứng (so với bản `am-thanh: khong`, năng lượng khác nhau).
- [ ] `--xem-truoc` và trích khung giữa cảnh; xem **từng** ảnh bằng Read. Sửa lỗi bố cục kèm test.
- [ ] Demo thật `projects/_video/con-lac-don-demo4/` (không commit):
  - dựa trên demo3, thêm cụm nhấn, `{{}}`, một `bieu-do` (chu kì theo chiều dài), một `so-do`, một `cau-hoi`, `luan-phien`;
  - nhạc tải bằng `tim_nhac.py "calm piano"`;
  - giọng thật;
  - báo thời gian dựng, dung lượng, thời lượng; trích 6 khung tiêu biểu.
- [ ] Commit `test(vi): end-to-end explainer video with every educational effect`.

### Task 10: Tài liệu, luật, biên bản, changelog

- [ ] Cập nhật tài liệu:
  - `docs/vi/tro-ly/video-giai-thich.md`: câu hỏi gộp để vẫn ≤ 7 câu; ví dụ có cụm nhấn, `{{}}` và `cau-hoi`, qua `parse` + `kiem`.
  - `docs/vi/tro-ly/canh-video.md`: bốn loại cảnh mới, cú pháp nhấn và số, năm kiểu chuyển cảnh, các khoá mới.
  - `docs/vi/video-giai-thich.md`: thêm phần cho thầy cô.
  - `AGENTS.vi.md` §15: bước tìm nhạc, nhắc nghe thử.
  - `.agents/rules/ppt-master-vi.md`: dưới 12.000 ký tự, bảng không đổi.
  - `docs/vi/xu-ly-loi.md`: thêm `tim_nhac.py` và bước `mang`.
- [ ] Test lớp Việt:
  - các cụm bắt buộc `==`, `((`, `__`, `{{`, `cau-hoi`, `bieu-do`, `so-do`, `dong-thoi-gian`, `tim_nhac.py`, `nhac-nen`, `am-thanh`, `karaoke`, `luan-phien`;
  - ví dụ trong hướng dẫn chạy qua bộ đọc thật.
- [ ] Biên bản `docs/vi/phat-trien/2026-09-26-video-ma-hieu-ung-kiem-thu.md`: số liệu thật, ghi rõ việc chưa kiểm.
- [ ] Mục `6.3.2-vi.11` trong `CHANGELOG-VI.md` và dòng phiên bản `README.md`.
- [ ] Commit `docs(vi): document educational effects in explainer videos`, rồi `docs(vi): record the vi.11 acceptance run and changelog`.
