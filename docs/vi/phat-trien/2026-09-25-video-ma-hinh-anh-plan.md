# Hình ảnh, bàn tay cầm bút và máy quay cho video giải thích — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development. Steps use checkbox (`- [ ]`) syntax.

**Goal:** Sửa lỗi font tiếng Việt (đóng gói Itim) và thêm hình vẽ nét, ảnh thật, bàn tay cầm bút, lau bảng, máy quay động, chụp 30 khung/giây song song cho `tools/vi/video_ma.py` (bản v6.3.2-vi.10).

**Architecture:** Giữ nguyên đường ống vi.9 (parse → kiem → giong → lich → trang → chup → ghep). Dữ liệu cảnh `du` có thêm hình, ảnh, cờ và nền cảnh trước. Runtime JS có thêm ba loại mục `hinh`, `anh` và `dong`, cùng hai module thuần (`ban-tay.js`, `may-quay.js`) tính vị trí tay và máy quay theo `t`. `chup.py` chia cảnh cho nhiều tiến trình Chromium.

**Tech Stack:** Python thư viện chuẩn (+ `multiprocessing`/`concurrent.futures`), edge-tts, playwright (import lười), JavaScript thuần, FFmpeg/libass.

**Spec:** `docs/vi/phat-trien/2026-09-25-video-ma-hinh-anh-design.md` (có quyền quyết định cuối). Spec vi.9: `docs/vi/phat-trien/2026-09-25-video-ma-viet-tay-design.md`.

**Cách viết kế hoạch này:** mỗi task nêu giao diện chính xác và các khẳng định test cụ thể. Người thực hiện tự viết mã theo TDD: viết test trước, thấy test trượt, viết mã, thấy test đạt. Các con số và tên trong kế hoạch là bắt buộc.

## Global Constraints

- Không sửa `skills/` (chỉ đọc `skills/ppt-master/templates/icons/tabler-outline/*.svg`), `tools/vi/video.py`, `tools/vi/video_parts/`, `tools/vi/thi_nghiem_parts/`, `LICENSE`, `SPONSORS*`, `requirements.txt`, `attribution_guard.py`. Không commit gì trong `projects/`.
- Python chỉ thư viện chuẩn, `edge-tts` và `playwright` (import lười). JS thuần, không địa chỉ web trong trang (trừ namespace SVG). Trang dựng tự chứa: font và ảnh nhúng `data:`.
- Font khung hình và phụ đề: **Itim**, đóng gói ở `tools/vi/video_ma_parts/runtime/fonts/Itim-Regular.ttf` kèm `OFL.txt`. Nguồn: `https://github.com/google/fonts/raw/main/ofl/itim/Itim-Regular.ttf` và `https://github.com/google/fonts/raw/main/ofl/itim/OFL.txt`.
- Hằng số mới: `FPS = 30`, `DAN_DAU = 1.0`, `LAU_BANG = 0.5` (giây), phóng máy quay tối đa `1.35`, thu về trong `1.2` s cuối cảnh, vùng an toàn phụ đề `y ≤ 620`, số tiến trình chụp `min(4, os.cpu_count() // 2)` (tối thiểu 1).
- Khoá đầu mới: `ban-tay: co|khong` (mặc định `co`), `may-quay: co|khong` (`co`), `chuyen-canh: lau-bang|khong` (`lau-bang`).
- Kịch bản vi.9 (ví dụ `tools/vi/fixtures/video-mau/video.md`) vẫn phải qua `--plan-only` và dựng được mà không sửa.
- Mọi đường CLI vẫn in đúng một dòng JSON; mọi `error.step` giữ nguyên tập cũ.
- Test: `venv\Scripts\python.exe -m unittest discover -s tools/vi/tests` (repo venv, các test Chromium/FFmpeg tự bỏ qua); `C:/Users/ADMIN/vmt/v/Scripts/python.exe -m unittest discover -s tools/vi/tests -p "test_video_ma_*.py"` (Python có playwright + edge-tts, chạy thật Chromium/FFmpeg, không mở cửa sổ); `node --test tools/vi/tests/js/test_canh.js` và các file Node mới trong `tools/vi/tests/js/`.
- Trailer commit: `Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>`. Không push.

## Review Focus

1. **Chữ có dấu hiếm** ("ỷ", "ỹ", "Ử", "ẵ") trong tiêu đề, ý, nhãn hình, chú thích ảnh và phụ đề: phải cùng font Itim. Pin ở Task 1 bằng test độ rộng Chromium cho đủ 134 chữ và test cmap.
2. **Tên biểu tượng gần đúng hoặc có tiền tố** (`tabler-outline/flask`, `Flask`, `flask.svg`, `bình thí nghiệm`): chấp nhận dạng có tiền tố hoặc đuôi `.svg`; tên tiếng Việt thì báo lỗi `canh` kèm gợi ý. Pin ở Task 2.
3. **Ảnh dọc, ảnh rất lớn, ảnh không có trong `image_sources.json`, tên file có dấu cách hoặc dấu tiếng Việt:** ảnh vừa khung (contain, không méo); quá 8 MB thì lỗi `canh`; thiếu nguồn thì lỗi `canh`; tên có dấu vẫn đọc được. Pin ở Task 2 và Task 6.
4. **Cảnh rất ngắn (2,5 s) mà vẫn có lau bảng, bàn tay và máy quay:** không mục nào vượt cuối cảnh; máy quay vẫn thu về 1,0 ở khung cuối; tay nghỉ ở khung cuối. Pin ở Task 4.
5. **Một tiến trình chụp lỗi giữa chừng, hoặc máy chỉ có 1–2 lõi:** báo `dung` kèm số cảnh, dọn `.khung/`; một tiến trình vẫn chạy đúng. Pin ở Task 5.

---

### Task 1: Font Itim cho khung hình và phụ đề

**Files:** tạo `tools/vi/video_ma_parts/runtime/fonts/Itim-Regular.ttf`, `OFL.txt`; tạo `tools/vi/video_ma_parts/phong.py` (bộ đọc cmap thư viện chuẩn); sửa `trang.py`, `runtime/viet-tay.css`, `ghep.py`; test `tools/vi/tests/test_video_ma_phong.py`, bổ sung `test_video_ma_chup.py`, `test_video_ma_ghep.py`.

**Interfaces:**
- `phong.FONT = <Path tới Itim-Regular.ttf>`, `phong.TEN = "Itim"`, `phong.CHU_VIET: str` (134 chữ có dấu: `ạảãàáâậầấẩẫăặằắẳẵẹẻẽèéêệềếểễịỉĩìíọỏõòóôộồốổỗơợờớởỡụủũùúưựừứửữỵỷỹỳýđ` và bản hoa), `phong.bang_ma(path) -> set[int]` (đọc bảng `cmap` định dạng 4 và 12 bằng `struct`), `phong.font_css() -> str` (khối `@font-face{font-family:'Itim';src:url(data:font/ttf;base64,...)}`).
- `trang.dung_trang` chèn `phong.font_css()` vào `<style>`; CSS `#khung { font-family: 'Itim', sans-serif; }` (bỏ Segoe Print/Ink Free/Comic Sans).
- `ghep.ghep_video`: sao `Itim-Regular.ttf` vào `.khung/fonts/`; bộ lọc `subtitles=.khung/phu-de.srt:fontsdir=.khung/fonts:force_style='FontName=Itim,FontSize=16,Outline=1.5,Shadow=0,Spacing=0.5,MarginV=22'`.

- [ ] Tải hai file font về đúng chỗ (curl), kiểm `OFL.txt` là giấy phép SIL OFL 1.1.
- [ ] Test trượt → viết `phong.py`: `CHU_VIET` có đúng 134 ký tự khác nhau; `bang_ma(FONT)` chứa mọi `ord(c)` của `CHU_VIET` và "ĐƯƠ"; `bang_ma` của `C:/Windows/Fonts/segoepr.ttf` (nếu có, `skipUnless`) thiếu `ord("ề")` (chứng minh bộ đọc thật sự phân biệt).
- [ ] Test Chromium (trong `test_video_ma_chup.py`): với một trang chứa `phong.font_css()`, cho mỗi chữ trong `CHU_VIET`, độ rộng `measureText`/`getBoundingClientRect` với `font-family:'Itim',monospace` bằng với `'Itim',serif` (sai lệch < 0,01 px); trang cảnh dựng bằng `dung_trang` có `getComputedStyle(.chu).fontFamily` bắt đầu bằng `Itim`.
- [ ] Test ghép (lệnh giả): lệnh video chứa `fontsdir=.khung/fonts` và `FontName=Itim`; `.khung/fonts/Itim-Regular.ttf` tồn tại sau `ghep_video` với `phu_de="hinh"`.
- [ ] Sửa các test cũ đang khẳng định "Segoe Print" (tìm bằng Grep) sang Itim. Chạy lại toàn bộ Chromium test hiện có: mọi cảnh ở giới hạn tối đa vẫn không tràn chữ (Itim rộng hơn Segoe Print ở một số chữ; nếu tràn thì chỉnh ô/cỡ chữ trong file cảnh, không nới giới hạn).
- [ ] Thêm dòng ghi công font vào `tools/vi/video_ma_parts/runtime/fonts/README.md` (một dòng: tên, tác giả Cadson Demak, giấy phép OFL 1.1, nguồn).
- [ ] Commit `fix(vi): draw every Vietnamese letter in the bundled Itim font`.

### Task 2: Ngữ pháp và kiểm cho hình, ảnh, khoá mới

**Files:** sửa `parse.py`, `kiem.py`; tạo `hinh.py`, `anh.py`; test `test_video_ma_parse.py` (thêm), `test_video_ma_hinh.py` (mới).

**Interfaces:**
- `parse.META_CHOICES` thêm `"ban-tay": ("co","khong")`, `"may-quay": ("co","khong")`, `"chuyen-canh": ("lau-bang","khong")`; mặc định `co`, `co`, `lau-bang`.
- `parse.SCENE_SPEC`: thêm trường đơn tuỳ chọn `hinh`, `anh` cho `tieu-de`, `khai-niem`, `cong-thuc`, `y-tung-y`; loại mới `"minh-hoa": (("tieu-de",), (), {"hinh": (1, 3)})`; `"anh": (("anh", "chu-thich"), ("nguon",), {})`. Một cảnh có cả `hinh` và `anh` → `ParseError` tại dòng thứ hai.
- `hinh.THU_MUC = <repo>/skills/ppt-master/templates/icons/tabler-outline`; `hinh.chuan_ten(s) -> str` (bỏ tiền tố `tabler-outline/`, đuôi `.svg`, về chữ thường, bỏ khoảng trắng hai đầu); `hinh.doc(ten) -> dict` trả `{"ten", "viewBox": "0 0 24 24", "phanTu": [{"the": "path"|"circle"|"rect"|"line"|"polyline"|"polygon"|"ellipse", "thuocTinh": {...}}]}` (bỏ phần tử có `stroke="none"` và `<path d="M0 0h24v24H0z" .../>`); lỗi `hinh.HinhError(message)` với gợi ý `difflib.get_close_matches(ten, danh_sach, n=5, cutoff=0.5)`; `hinh.tach_minh_hoa(value) -> (ten, nhan)` tách tại `|` (nhãn bắt buộc, ≤ 30 ký tự hiển thị).
- `anh.DINH_DANG = (".jpg", ".jpeg", ".png", ".webp")`, `anh.TOI_DA = 8 * 1024 * 1024`; `anh.doc(thu_muc_du_an, ten_file, nguon_tay) -> dict` trả `{"dataUrl", "nguon", "rong", "cao"}`; kích thước đọc từ phần đầu file bằng thư viện chuẩn (PNG IHDR, JPEG SOF0/SOF2, WEBP VP8/VP8L/VP8X); nguồn = `nguon_tay` nếu có, không thì dựng từ bản ghi có `filename == ten_file` trong `anh/image_sources.json` dạng `"Ảnh: <author> · <license> · <provider>"` (bỏ phần trống); không có nguồn → `anh.AnhError`.
- `kiem.kiem`: gọi `hinh.doc`/`anh.doc` cho mọi cảnh có trường tương ứng, đổi `HinhError`/`AnhError` thành `CanhError(so, ...)` kèm số dòng; giới hạn: `("minh-hoa","tieu-de")`: 90, `("anh","chu-thich")`: 90, nhãn hình 30.

- [ ] Test: kịch bản vi.9 mẫu vẫn parse và `kiem` không cảnh báo; meta mới có mặc định đúng; `hinh: tabler-outline/Flask.svg` chấp nhận thành `flask`; `hinh: binh-thi-nghiem` → `CanhError` có "flask" không bắt buộc nhưng có ít nhất một gợi ý và có số dòng; `doc("flask")` có 3 phần tử `path`; `doc("atom")` không chứa phần tử khung `M0 0h24v24H0z`; cảnh có cả `hinh` và `anh` → `ParseError` đúng dòng; `minh-hoa` 0 hoặc 4 hình → lỗi; nhãn 31 ký tự → `CanhError`; thiếu `|` → lỗi nêu dạng `tên | nhãn`.
- [ ] Test ảnh (tạo ảnh thật nhỏ bằng thư viện chuẩn: PNG 1×1 viết tay bằng `zlib`+`struct`; JPEG lấy từ FFmpeg nếu có, `skipUnless`): đọc đúng kích thước; file không tồn tại, đuôi `.gif`, file 8 MB + 1 byte, không có nguồn → `CanhError`; có bản ghi `image_sources.json` → nguồn chứa tác giả và giấy phép; `nguon:` tay thắng manifest; tên file `ảnh con lắc.png` đọc được.
- [ ] Commit `feat(vi): pictures and photos in explainer video scripts`.

### Task 3: Lịch và dữ liệu cảnh

**Files:** sửa `lich.py`, `video_ma.py` (chỉ chỗ dựng `du`), test `test_video_ma_lich.py`, `test_video_ma_cong_cu.py`.

**Interfaces:**
- `lich.FPS = 30`, `lich.DAN_DAU = 1.0`, `lich.LAU_BANG = 0.5`. `thoi_luong_canh` giữ công thức với hằng mới: `max(2.5, 1.0 + giọng + 0.6)` làm tròn lên bội 1/30.
- `lich.so_muc` thêm `"minh-hoa": len(hinh)`.
- `lich.du_lieu_canh(scene, cl, model=None, tai_nguyen=None)`: `tai_nguyen` là dict do `video_ma` chuẩn bị: `{"hinh": dict|None, "anh": dict|None, "hinhs": [dict]}`; `du` thêm khoá `"hinh"`, `"anh"`, `"hinhs"` (mỗi phần tử `hinh.doc(...)` cộng `"nhan"`), `"co": {"banTay": bool, "mayQuay": bool, "lauBang": bool}` (`lauBang` là `True` chỉ khi `chuyen-canh` là `lau-bang` và `scene.so > 1`), `"nenTruoc": None` (tiến trình chụp điền sau).
- `video_ma._tai_nguyen(scene, thu_muc) -> dict` gọi `hinh.doc` / `anh.doc`; `_trang` truyền `tai_nguyen` và `video.meta`.

- [ ] Test: thời lượng tối thiểu `ceil(2.5*30)/30`, bội 1/30; mốc câu cộng 1,0; `du["co"]` đúng cho cảnh 1 và cảnh 2, và khi `chuyen-canh: khong`; `du["hinhs"]` giữ thứ tự và nhãn; `du["moc"]` của `minh-hoa` có đúng số hình; phụ đề (`ghep.cues_phu_de`) dùng 1,0 dẫn đầu và `lenh_am_canh` dùng `adelay=1000`.
- [ ] Commit `feat(vi): schedule explainer scenes at 30 fps with a one-second lead`.

### Task 4: Runtime vẽ: hình nét, ảnh, cột hình, minh-hoa, anh, bàn tay, máy quay, lau bảng

**Files:** sửa `runtime/khung-video.js`, `runtime/viet-tay.css`, `runtime/canh/{tieu-de,khai-niem,cong-thuc,y-tung-y}.js`; tạo `runtime/hinh.js`, `runtime/ban-tay.js`, `runtime/may-quay.js`, `runtime/canh/minh-hoa.js`, `runtime/canh/anh.js`; sửa `trang.py` (nạp file mới theo thứ tự `khung-video.js`, `hinh.js`, `ban-tay.js`, `may-quay.js`, file cảnh); test `tools/vi/tests/js/test_canh.js` (bổ sung), `tools/vi/tests/js/test_chuyen_dong.js` (mới), `test_video_ma_canh.py` (chạy cả hai file Node), Chromium test trong `test_video_ma_chup.py`.

**Interfaces (JS):**
- Mục mới trong `muc(du)`: `{kieu:'hinh', id, phanTu, viewBox, x, y, kich, batDau, thoiLuong, mau}` (vẽ lần lượt từng phần tử, mỗi phần tử một phần bằng nhau của `thoiLuong`, tối thiểu 1,2 s cho cả hình); `{kieu:'anh', id, dataUrl, nguon, x, y, rong, cao, batDau, thoiLuong}` (ảnh hiện dần 0,4 s rồi Ken Burns: `scale` 1,0 → 1,12 và dịch tối đa 3% theo hướng xác định bằng hạt giống `du.so`, suốt từ `batDau` tới cuối cảnh; ảnh `object-fit: contain`, khung vẽ tay quanh ảnh; dòng nguồn cỡ 14 px ở góc dưới phải trong khung).
- `THI_VIDEO.tao(du)` thêm `B.hinh(id, h, x, y, kich, batDau)`, `B.anh(id, a, x, y, rong, cao, batDau)`; khi `du.co.lauBang`, mọi `batDau` < `LAU_BANG + 0.05` được đẩy lên `LAU_BANG + 0.05`.
- **Cột hình:** khi `du.hinh` hoặc `du.anh` có mặt ở `khai-niem`, `cong-thuc`, `y-tung-y`: mọi ô chữ thu còn chiều rộng tối đa `x ≤ 860`; hình/ảnh ở ô `x=900, y=200, rộng 320, cao 380`, bắt đầu tại `du.moc[0]` hoặc 1,0 nếu không có mốc. `tieu-de`: hình 180×180 ở giữa phía trên tiêu đề (y 40–220) và tiêu đề dời xuống vừa đủ.
- `minh-hoa.js`: tiêu đề như các cảnh khác; n hình (1–3) chia đều bề ngang, mỗi hình 240×240 ở y 230, nhãn cỡ 28 ngay dưới; hình k bắt đầu tại `du.moc[k]`, nhãn k bắt đầu khi hình k xong.
- `anh.js`: ảnh chiếm vùng `x 80–1200, y 70–560`, chú thích cỡ 30 ở y 575–625 bắt đầu tại `du.moc[0]` hoặc 1,0.
- `THI_VIDEO.catDanhDau(chu, n)` chèn `<span class="ngoi"></span>` ngay sau ký tự thứ n khi 0 < n < tổng (đánh dấu ngòi bút); chữ hiển thị không đổi.
- `ban-tay.js` (thuần, chạy được trong Node): `THI_BAN_TAY.viTri(ds, t, ngoi, nghi)` với `ds` là danh sách mục (có `id, kieu, batDau, thoiLuong, dong?`), `ngoi(muc, p) -> {x,y}` do trình duyệt cung cấp, `nghi = {x: 1220, y: 760}`; trả `{x, y, hien: bool, kieu: 'but'|'gie'}`. Quy tắc: mục đang vẽ (0<p<1, bỏ mục `dong` và `anh`) → ngòi của mục đó; khoảng trống ≤ 0,6 s giữa hai mục → nội suy tuyến tính từ `ngoi(trước,1)` tới `ngoi(sau,0)`; ngoài ra → lướt tới `nghi` trong 0,4 s rồi đứng ở `nghi` với `hien=false`. Trong `[0, LAU_BANG]` khi lau bảng: `kieu:'gie'`, `x = -120 + (1400)*t/LAU_BANG`, `y = 380`.
- `may-quay.js` (thuần): `THI_MAY_QUAY.tinh(ds, hop, t, gh, cau_hinh) -> {z, tx, ty}` với `hop[id] = {x,y,w,h}` (đo ở Z=1). Mục tiêu = mục đang vẽ, không có thì mục xong gần nhất; `z = clamp(min(0.6*1280/w, 0.6*560/h), 1, 1.35)`; chuyển giữa mục tiêu trong 0,6 s bằng easeInOut; trong `[gh-1.2, gh]` nội suy về `{1,0,0}`; kẹp `tx, ty` để lớp bảng phủ kín khung 1280×720 và hộp mục tiêu sau biến đổi có đáy ≤ 620 và nằm trong khung. `cau_hinh.day = true` (cảnh `thi-nghiem`) → chỉ đẩy chậm `z = 1 + 0.06*t/gh`, `tx, ty` giữ tâm. `mayQuay=false` → luôn `{1,0,0}`.
- `khoiDong(du)`: toàn bộ nội dung nằm trong `#bang` (con của `#khung`); nền cảnh trước `#nen-truoc` (ảnh `du.nenTruoc`) nằm dưới `#bang` và bị mặt nạ quét trái → phải trong `[0, LAU_BANG]`; bàn tay `#ban-tay` (SVG do repo tự vẽ: bàn tay nét đen tô da nhạt cầm bút dạ xanh; bản `gie` cầm giẻ) nằm trong `#bang` để đi theo máy quay; đo `hop` khi `z=1`, tay ẩn. `datThoiDiem(t)` áp: mục → máy quay (transform `#bang`) → tay. `kiemTran()` đặt `z=1`, ẩn tay, không nền cũ. Cảnh `thi-nghiem`: không tay.

- [ ] Test Node `test_chuyen_dong.js`: `viTri` xác định; tay tại ngòi giả khi đang vẽ; nội suy trong khoảng trống 0,4 s; `hien=false` khi rảnh 2 s; `kieu:'gie'` trong lau bảng và `x` tăng dần; `tinh` xác định; `z ∈ [1, 1.35]` tại 200 điểm `t`; `z=1, tx=0, ty=0` tại `t = gh`; mọi điểm: hộp mục tiêu biến đổi có đáy ≤ 620 + 1e-6 và lớp bảng phủ kín khung; `day=true` cho `z(gh) ≈ 1.06`; `mayQuay` tắt → `{1,0,0}`.
- [ ] Test Node `test_canh.js` (bổ sung): `minh-hoa` hình k bắt đầu tại `moc[k]`; mọi loại cảnh (gồm `minh-hoa`, `anh`, bốn cảnh có cột hình) với `lauBang` và `gh = 2.5667`: không mục nào bắt đầu trước 0,55 s hay kết thúc sau `gh − 0,2`; mục `hinh` có `thoiLuong ≥ 1.2` khi cảnh đủ dài; `catDanhDau` có đúng một `span.ngoi` khi 0<n<tổng và không có khi n=0 hoặc n=tổng; chữ hiển thị không đổi.
- [ ] Test Chromium: `khai-niem`/`y-tung-y`/`cong-thuc`/`tieu-de` có hình ở giới hạn chữ tối đa không tràn; `minh-hoa` 3 hình nhãn 30 ký tự không tràn; `anh` với ảnh dọc 600×1200 và ngang 2000×800 nằm trong khung, tỉ lệ đúng (đo kích thước `img` hiển thị); khung giữa hình (t = batDau + 0,5·thoiLuong) khác khung cuối; tại `t` giữa một mục chữ, `#ban-tay` hiện và tâm ngòi cách `span.ngoi` < 40 px (sau transform); lau bảng: tại 0,05 s `#nen-truoc` còn thấy, tại 0,6 s không còn; transform của `#bang` tại `t = gh − 1/30` là ma trận đơn vị.
- [ ] Commit (có thể 2–3 commit theo phần): `feat(vi): draw icons and photos stroke by stroke in explainer scenes`, `feat(vi): pen hand, board wipe and camera motion for explainer scenes`.

### Task 5: Chụp song song 30 khung/giây và nền cảnh trước

**Files:** sửa `chup.py`, `video_ma.py`, `ghep.py`; test `test_video_ma_chup.py`, `test_video_ma_cong_cu.py`, `test_video_ma_ghep.py`.

**Interfaces:**
- `chup.chia_dai(so_khung_moi_canh: list[int], so_tien_trinh: int) -> list[tuple[int, int]]` — các dải cảnh liên tiếp `[dau, cuoi)` phủ đủ, không chồng, tổng khung mỗi dải gần đều (thuật toán tham lam theo tổng tích luỹ), số dải ≤ số tiến trình và ≤ số cảnh.
- `chup.so_tien_trinh() -> int` = `max(1, min(4, (os.cpu_count() or 2) // 2))`.
- `chup.chup_dai(cong_viec: dict) -> int` — hàm cấp module (chạy được trong tiến trình con Windows `spawn`): nhận `{"cac_du": [...], "models_js": {so: js}, "dau": i, "cuoi": j, "khung_dau": [...], "fps": 30, "thu_muc_anh": str}`; tự thêm `tools/vi` vào `sys.path`; mở Chromium; nếu `dau > 0` và `cac_du[dau]["co"]["lauBang"]` thì dựng trang cảnh `dau-1`, chụp tại `thoiLuong − 1/fps` thành PNG trong bộ nhớ, đặt `nenTruoc` (data URL) cho cảnh `dau`; với mỗi cảnh k trong dải: chụp khung, rồi (nếu cảnh k+1 cần) giữ PNG khung cuối làm `nenTruoc` cho k+1. Trả số khung đã ghi.
- `chup.chup_song_song(cac_du, models_js, so_khung, fps, thu_muc_anh, so_tt) -> None` — `concurrent.futures.ProcessPoolExecutor(max_workers=so_tt)`; `so_tt == 1` chạy thẳng `chup_dai` trong tiến trình hiện tại; lỗi ở dải nào → `MediaError("dung", "Chụp khung lỗi ở cảnh X–Y: ...")`.
- `trang.dung_trang(du, model=None)` giữ chữ ký; `nenTruoc` đã nằm trong `du`.
- `video_ma._dung`: kiểm tràn (một Chromium) → giọng → lịch → `du` → `chup_song_song` → ghép. Nhật ký stderr báo số tiến trình.
- `ghep.lenh_video`: `-framerate 30`, `-r 30`.

- [ ] Test thuần: `chia_dai([30]*8, 4)` → 4 dải liền nhau; `chia_dai([300, 30, 30, 30], 4)` → dải đầu chỉ cảnh 0; một cảnh hay `so_tt=1` → một dải; tổng khung khớp.
- [ ] Test Chromium thật (vmt Python): 3 cảnh ngắn (2,5–3 s, có lau bảng), `so_tt=2`: đúng `sum(so_khung)` file `f%06d.png` liên tục không thiếu; khung đầu cảnh 2 (t=0,05 s) giống khung cuối cảnh 1 ở nửa phải (vùng chưa lau) — so sánh một vùng pixel bằng cách đọc lại PNG qua Chromium hoặc so byte của ảnh cắt bằng FFmpeg; `so_tt=1` cho cùng số khung.
- [ ] Test lỗi: `chup_dai` giả ném lỗi → `MediaError` step `dung`, thông báo có "cảnh"; `.khung/` bị dọn (CLI test).
- [ ] Đo: dựng video 60 giây (kịch bản mẫu rút gọn, tiếng giả) bằng vmt Python, ghi thời gian với `so_tt` mặc định và `so_tt=1`; ngoại suy cho 5 phút; ghi số vào báo cáo. Nếu vượt 8 phút cho 5 phút thì báo lại, không tự đổi hằng số.
- [ ] Commit `feat(vi): capture explainer frames at 30 fps across parallel browsers`.

### Task 6: Dựng thật, xem bằng mắt, demo có hình

**Files:** tạo fixture mới `tools/vi/fixtures/video-hinh/video.md` (giữ nguyên fixture vi.9) gồm: `tieu-de` có `hinh`, `khai-niem` có `hinh`, `y-tung-y` có `anh`, `minh-hoa` 3 hình, `anh`, `thi-nghiem`; ảnh là PNG nhỏ sinh sẵn bằng thư viện chuẩn, commit ở `tools/vi/fixtures/video-hinh/anh/`, cảnh ảnh ghi `nguon:` để fixture dựng được không cần mạng. Sửa `test_video_ma_tich_hop.py` (thêm cảnh hình, ảnh; `so_tt=2`); sửa file cảnh nếu thấy lỗi bố cục.

- [ ] Chạy `--xem-truoc` trên `tools/vi/fixtures/video-hinh`, xem **từng** ảnh bằng Read: chữ Itim đủ dấu, hình nằm đúng cột, không chồng chữ, ảnh không méo, dòng nguồn đọc được.
- [ ] Dựng thật `video-hinh` (tiếng giả `sine`), trích khung giữa 5 cảnh bằng FFmpeg, xem: bàn tay đúng ngòi, máy quay đang phóng, lau bảng đang quét, phụ đề Itim.
- [ ] Tải 1 ảnh thật bằng `skills/ppt-master/scripts/image_search.py "Foucault pendulum" --filename con-lac-foucault.jpg --orientation landscape -o projects/_video/con-lac-don-demo3/anh` (cần mạng), viết kịch bản demo có đủ loại cảnh mới vào `projects/_video/con-lac-don-demo3/video.md`, dựng với giọng thật (vmt Python có edge-tts), xem khung, báo số liệu (thời gian dựng, dung lượng, thời lượng). Không commit gì trong `projects/`.
- [ ] Mọi lỗi bố cục thấy được: sửa ở file cảnh kèm test, commit riêng `fix(vi): ...`.
- [ ] Commit `test(vi): end-to-end explainer video with pictures, photo and motion`.

### Task 7: Tài liệu, luật và biên bản

**Files:** sửa `docs/vi/tro-ly/video-giai-thich.md`, `docs/vi/tro-ly/canh-video.md`, `docs/vi/video-giai-thich.md`, `AGENTS.vi.md` §15, `.agents/rules/ppt-master-vi.md` (chỉ nếu cần một dòng về tải ảnh), `docs/vi/xu-ly-loi.md`, `CHANGELOG-VI.md` (mục `6.3.2-vi.10`), `README.md` (dòng phiên bản như commit phát hành trước), `tools/vi/tests/test_vi_layer.py`; tạo `docs/vi/phat-trien/2026-09-25-video-ma-hinh-anh-kiem-thu.md`.

- [ ] `canh-video.md`: hai loại cảnh mới (`minh-hoa`, `anh`) với trường, giới hạn, ví dụ; trường `hinh`/`anh` ở bốn loại cảnh; **bảng tra biểu tượng** tối thiểu 80 dòng "khái niệm tiếng Việt → tên `tabler-outline`" cho Toán, Vật lí, Hoá, Sinh, Địa, chung (mọi tên phải tồn tại — test khoá); cách tìm thêm: `rg --files skills/ppt-master/templates/icons/tabler-outline -g "*từ-khoá*"`.
- [ ] `video-giai-thich.md` (cho AI): câu hỏi thêm "có muốn ảnh thật không" (gộp vào câu hỏi sẵn có để không vượt 7 câu); bước tải ảnh bằng `image_search.py ... -o projects\_video\<tên>\anh` trước khi dựng; bắt buộc chạy `--xem-truoc` và cho thầy cô xem ảnh khi có ảnh thật; ba khoá mới; ví dụ kịch bản (khối bắt đầu bằng `---`) có `hinh`, `minh-hoa` và `anh` với `nguon:` (không cần file thật để qua `parse`, nhưng test ví dụ gọi `kiem` → ví dụ dùng `anh:` phải đặt trong khối riêng không được test `kiem`, hoặc test tạo file ảnh giả; chọn cách thứ hai).
- [ ] `AGENTS.vi.md` §15: thêm bước tải ảnh và `--xem-truoc` khi có ảnh; nhắc thời gian dựng ~1,5× thời lượng video.
- [ ] `test_vi_layer.py`: mọi tên trong bảng tra tồn tại trong `tabler-outline`; ví dụ trong hướng dẫn qua `parse` + `kiem` (tạo ảnh giả trong thư mục tạm); các cụm bắt buộc mới (`minh-hoa`, `image_search.py`, `ban-tay`, `may-quay`, `chuyen-canh`, `Itim`).
- [ ] Biên bản kiểm thử: font (bằng chứng cmap + đo độ rộng), số test ba lần chạy, số đo tốc độ Task 5, kết quả xem bằng mắt Task 6, demo giọng thật, việc chưa kiểm.
- [ ] CHANGELOG `6.3.2-vi.10`: sửa lỗi font (nêu rõ nguyên nhân), hình vẽ nét, ảnh thật, bàn tay, lau bảng, máy quay, 30 khung/giây song song; Vox/9:16 sang vi.11.
- [ ] Commit `docs(vi): document pictures and motion in explainer videos`, rồi `docs(vi): record the vi.10 acceptance run and changelog`.
