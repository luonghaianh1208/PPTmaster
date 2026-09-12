# Thiết kế: Video bài giảng có lời giảng (`v6.3.2-vi.4`)

- **Ngày:** 2026-09-12
- **Người phụ trách:** Lương Hải Anh — 2Anh AI Education
- **Trạng thái:** Đã duyệt thiết kế; chờ chủ repo đọc spec
- **Nền:** `main` @ `615c87b2` (bản Việt `v6.3.2-vi.3`); upstream vẫn là `v6.3.2`
- **Nhánh:** `feat/vi-video`
- **Gói tiếp theo (ngoài phạm vi file này):** Remotion và video ngắn dọc cho Zalo/Facebook/TikTok

---

## 1. Mục tiêu

Thầy cô đã có bài giảng trong PPT Master thì nhắn một câu là ra **video có lời giảng tiếng Việt và phụ đề**, đăng được lên LMS hoặc YouTube, kể cả khi máy không có PowerPoint.

### Tiêu chí thành công

1. Dự án đã có slide và ghi chú thuyết minh: thầy cô nhắn "làm video bài giảng này" → nhận file MP4 có tiếng khớp từng slide và một file phụ đề `.srt`. Kiểm bằng chạy thật (§8.2) và nghiệm thu (§8.3).
2. Máy **không có PowerPoint** vẫn dựng được video (đường FFmpeg). Kiểm bằng chạy thật.
3. Máy **có PowerPoint** thì mặc định dùng đường PowerPoint để giữ hiệu ứng chuyển cảnh, trừ khi thầy cô yêu cầu khác. Kiểm bằng test chọn cách dựng và chạy thật.
4. Dự án chưa có tiếng thuyết minh: AI chạy bước tạo thuyết minh sẵn có của upstream (giọng Việt `edge-tts`) rồi mới dựng video; lớp Việt không viết lại phần TTS.
5. Phụ đề cuối khớp mốc thời gian của video, dựng từ các file SRT theo slide, không cần thư viện nhận dạng giọng nói.
6. Không file nào của upstream bị sửa: `git diff v6.3.2 -- skills/` rỗng; guard bản quyền exit 0; toàn bộ test `tools/vi/tests` đạt.
7. Thứ nặng (Chromium) chỉ cài khi thật sự cần, và hỏi thầy cô trước khi tải.

---

## 2. Bối cảnh & phát hiện

| # | Phát hiện | Nguồn | Hệ quả thiết kế |
|---|---|---|---|
| 1 | Upstream đã có TTS đầy đủ: `edge-tts` (miễn phí, có giọng `vi-VN-HoaiMyNeural` và `vi-VN-NamMinhNeural`), ElevenLabs, MiniMax, Qwen, CosyVoice. Mỗi slide một MP3 + một SRT, kèm `audio/manifest.json`. | `workflows/stages/generate-audio.md`; `scripts/notes_to_audio.py:417-433`; `AGENTS.vi.md` mục 7 | Không viết lại TTS. Gói này chỉ gọi lại và ghép. |
| 2 | Bước thuyết minh cần `notes/*.md` (tách từ `notes/total.md` ở Step 7.1 của Generate). Một ghi chú → một file tiếng. | `workflows/stages/generate-audio.md:15-16` | Thiếu ghi chú thì AI chạy bước ghi chú của upstream trước, không tự chia nhỏ. |
| 3 | Xuất video sẵn có gọi PowerPoint bản cài trên Windows qua API `CreateVideo`; tham số `-o`, `--resolution` (mặc định 1080), `--fps` (30), `--quality` (85); `--check` để dò PowerPoint. | `scripts/powerpoint_video.py:10-20,214-247` | Đường PowerPoint dùng nguyên lệnh này, không sửa. |
| 4 | Upstream nói rõ: phụ đề là file rời, không nhúng vào PPTX và không in lên MP4. | `docs/audio-narration.md:28` | Việc in phụ đề lên hình do lớp Việt làm bằng FFmpeg. |
| 5 | `video_subtitles.py` canh phụ đề theo tiếng nói nhưng cần `stable-ts` (nặng, kéo theo mô hình nhận dạng). | `scripts/video_subtitles.py:16-18` | Không dùng. Phụ đề cuối dựng bằng cách cộng dồn thời lượng từng slide. |
| 6 | Có sẵn bộ chụp ảnh slide bằng Playwright/Chromium, ghi ra `<dự án>/.preview/<tên>.png`, tự dò máy chủ xem trước, mã thoát 0/2/3/4. | `scripts/visual_review.py:19-25,256-301` | Đường FFmpeg dùng lại bộ này thay vì tự viết bộ chuyển SVG sang ảnh. |
| 7 | Máy chủ xem trước bật bằng `server.py <dự án> --daemon --live --no-browser`, tắt bằng `server.py <dự án> --shutdown`; file trạng thái ở `<dự án>/live_preview/`. | `scripts/svg_editor/server.py:1197-1222,1247-1286` | Đường FFmpeg tự bật rồi tự tắt máy chủ, không để lại tiến trình. |
| 8 | `playwright` **không** nằm trong `requirements.txt`; `ffmpeg`/`ffprobe` là công cụ tuỳ chọn đã có lệnh cài trong lớp Việt. | `skills/ppt-master/requirements.txt`; `tools/vi/pptmaster.ps1` (`-Action tool`) | Thêm `-Name chromium` cho lệnh cài tuỳ chọn; hỏi thầy cô trước vì tải nặng. |
| 9 | Upstream ghi rõ không làm nhạc nền và không chèn video rời. | `docs/roadmap.md` (bảng Layer 1, dòng "Arbitrary video & background music") | Gói này giữ đúng ranh giới đó. |
| 10 | Remotion không phải mã nguồn mở tự do: miễn phí cho cá nhân, tổ chức phi lợi nhuận và công ty từ 3 nhân sự trở xuống; cấm bán sản phẩm phái sinh từ mã Remotion. Từ bản 4.0 tự kèm FFmpeg và tự tải Chrome Headless Shell. | `https://github.com/remotion-dev/remotion/blob/main/LICENSE.md`; `remotion.dev/docs/ffmpeg`; `remotion.dev/docs/miscellaneous/chrome-headless-shell` | Remotion để ở gói sau, không đưa vào bộ cài mặc định; khi làm sẽ kèm dòng nêu điều kiện giấy phép. |

---

## 3. Quyết định đã chốt

| Chủ đề | Quyết định |
|---|---|
| Loại video làm trước | Bài giảng ngang có lời giảng; video ngắn dọc để gói sau |
| Cách dựng | Hai đường trong gói này: PowerPoint và FFmpeg + ảnh slide. Kiến trúc chừa sẵn đường thứ ba (Remotion) |
| Chọn cách dựng | Theo yêu cầu của thầy cô; không nêu thì: có PowerPoint → PowerPoint, không có → FFmpeg |
| TTS | Dùng nguyên bước thuyết minh của upstream, giọng Việt mặc định `edge-tts` |
| Phụ đề | Dựng bằng cộng dồn thời lượng; thầy cô chọn file `.srt` rời (mặc định) hoặc in lên hình |
| Cài thêm | `chromium` thêm vào `-Action tool`; hỏi thầy cô trước khi tải |
| Ranh giới | Không sửa `skills/`; không nhạc nền, không chèn video rời, không người dẫn ảo |
| Phiên bản | `v6.3.2-vi.4`, nhánh `feat/vi-video` |

---

## 4. Kiến trúc

**Nguyên tắc:** lớp Việt chỉ **điều phối và ghép**. Mọi việc tạo nội dung (slide, ghi chú, tiếng nói, xuất PPTX) vẫn do upstream làm.

### 4.1 File thêm/sửa

```text
tools/vi/video.py                 # (mới) cổng vào duy nhất: chọn cách dựng, ghép video, in JSON
tools/vi/video_backends/          # (mới) mỗi cách dựng một file
├── __init__.py
├── powerpoint.py                 # gọi scripts/powerpoint_video.py
└── ffmpeg_slides.py              # ảnh slide + tiếng + FFmpeg
tools/vi/pptmaster.ps1            # (sửa) -Action tool -Name chromium
docs/vi/lam-video.md              # (mới) hướng dẫn cho thầy cô
docs/vi/tro-ly/video-bai-giang.md # (mới) bộ câu hỏi trước khi dựng, theo khuôn 5 loại việc
docs/vi/tro-ly/quy-trinh-hoi.md   # (sửa) thêm loại việc thứ 6 vào bảng nhận diện
AGENTS.vi.md                      # (sửa) mục 3 (câu lệnh video), mục 11 mới cho video
docs/vi/bat-dau-nhanh.md          # (sửa) một mục ngắn "Làm video bài giảng"
docs/vi/xu-ly-loi.md              # (sửa) mục "Dựng video thất bại"
tools/vi/tests/test_video.py      # (mới) test chọn cách dựng, phụ đề, kế hoạch lệnh
tools/vi/tests/test_vi_layer.py   # (sửa) test nội dung tài liệu và quy tắc
CHANGELOG-VI.md                   # (sửa khi phát hành)
```

### 4.2 Ảnh hưởng tới test đang khoá bộ 5 loại việc

"Video bài giảng" là **loại việc thứ 6** của quy trình hỏi đã có, nên phải cập nhật đúng các chỗ sau trong `tools/vi/tests/test_vi_layer.py`:

| Chỗ cần sửa | Việc phải làm |
|---|---|
| `GUIDE_FILES` (dòng 176) | Thêm `video-bai-giang.md` |
| `test_guides_have_required_sections_in_order` (dòng 266) | File mới phải có đúng 8 mục cấp 2 theo thứ tự như 5 file kia |
| `test_guides_have_numbered_required_questions` (dòng 272) | 1–7 câu đánh số, mỗi câu có "Gợi ý:" |
| `test_guides_have_quick_items` (dòng 280) | Mục "Tạo nhanh" có 2–3 mục đánh số |
| `test_guides_use_registered_canvas_keys` (dòng 287) | Trong mục `## Khổ slide` chỉ đặt key canvas trong dấu backtick (`ppt169`). Độ phân giải và số khung hình viết chữ thường, **không** đặt trong backtick, nếu không test sẽ coi `1080` là key canvas lạ |
| `test_guides_and_templates_have_no_markdown_links` (dòng 293) | File mới không có link Markdown |
| `test_common_rules_link_all_guides` (dòng 388) | `quy-trinh-hoi.md` phải trỏ tới file mới |
| `test_assistant_section_links_common_rules_and_all_guides` (dòng 403) | Bảng ở mục 10 của `AGENTS.vi.md` phải có dòng cho loại việc mới |
| `test_agents_vi_has_assistant_section_last` (dòng 399) | Đổi thành: mục 10 đứng ngay trước mục 11, và mục 11 là mục cuối |

Trong `quy-trinh-hoi.md`, thêm một dòng vào bảng nhận diện với dấu hiệu: "làm video", "xuất video", "lồng tiếng", "video bài giảng". Quy tắc "chỉ xếp vào một loại khi câu lệnh có dấu hiệu trong bảng" giữ nguyên.

---

## 5. Luồng

### 5.1 Cổng vào

Thầy cô nhắn "làm video bài giảng này", "lồng tiếng rồi xuất video", "xuất video bài giảng". AI đọc `docs/vi/tro-ly/video-bai-giang.md`, hỏi một lượt ngắn (tối đa 4 câu: giọng đọc, tốc độ, phụ đề, cách dựng nếu thầy cô có ý riêng), rồi chạy:

```
venv\Scripts\python.exe tools\vi\video.py <đường_dẫn_dự_án> [--cach auto|powerpoint|ffmpeg] [--phu-de file|hinh|khong] [--do-phan-giai 1080|720]
```

Tiến trình ra stderr; stdout đúng một dòng JSON, như bộ cài.

### 5.2 Chọn cách dựng

1. Thầy cô nêu rõ → theo đúng ý.
2. `--cach auto`: có `exports/*_narrated.pptx` (hoặc dựng được nó) **và** `powerpoint_video.py --check` báo có PowerPoint → `powerpoint`; ngược lại → `ffmpeg`.
3. Không đủ điều kiện cho cách được chọn → dừng với `error`, nêu cách sửa; không tự đổi sang cách khác khi thầy cô đã chỉ định.

### 5.3 Đường FFmpeg

1. Kiểm `audio/*.mp3`; thiếu → `error.step = "audio"`, cách sửa là chạy bước thuyết minh của upstream.
2. Bật máy chủ xem trước: `server.py <dự án> --daemon --live --no-browser`.
3. Chụp ảnh: `visual_review.py <dự án>` → `.preview/<tên>.png`.
4. Tắt máy chủ: `server.py <dự án> --shutdown`, kể cả khi bước 3 lỗi.
5. Đọc thời lượng từng file tiếng bằng `ffprobe`.
6. Ghép bằng một lệnh FFmpeg: mỗi ảnh giữ đúng thời lượng tiếng của slide đó, chuyển cảnh mờ dần 0,4 giây, nối tiếng theo thứ tự, xuất H.264 + AAC.
7. Phụ đề theo §5.5.

### 5.4 Đường PowerPoint

1. Cần `exports/<tên>_narrated.pptx` (bản đã gắn tiếng do upstream tạo). Chưa có → `error.step = "narrated_pptx"`, nêu bước upstream cần chạy.
2. Gọi `powerpoint_video.py <pptx> -o exports\<tên>_video.mp4 --resolution <1080|720>`.
3. Phụ đề theo §5.5.

### 5.5 Phụ đề

- Mốc bắt đầu của slide thứ n bằng tổng thời lượng các slide trước; cộng mốc đó vào từng dòng trong `audio/<tên>.srt`, ghép thành `exports/<tên>_video.srt`.
- `--phu-de file` (mặc định): chỉ ghi file `.srt`.
- `--phu-de hinh`: in phụ đề lên video bằng FFmpeg, giữ nguyên file `.srt` bên cạnh.
- `--phu-de khong`: không tạo phụ đề.
- Đường PowerPoint có sai số vì PowerPoint tự thêm thời gian chuyển cảnh. Khi in phụ đề lên hình ở đường này, script so tổng thời lượng video với tổng thời lượng tiếng; lệch quá 2% thì ghi cảnh báo vào `warnings` và vẫn xuất file `.srt` rời.

### 5.6 Báo thầy cô

AI nêu: đường dẫn video, thời lượng, dung lượng, cách dựng đã dùng, phụ đề nằm ở đâu, và các cảnh báo nếu có. Nếu có cài thêm Chromium thì nói rõ đã cài gì.

---

## 6. Chi tiết thành phần

### 6.1 `tools/vi/video.py`

**Tham số:** `<đường_dẫn_dự_án>`, `--cach auto|powerpoint|ffmpeg` (mặc định `auto`), `--phu-de file|hinh|khong` (mặc định `file`), `--do-phan-giai 1080|720` (mặc định `1080`), `--plan-only`.

**JSON ở stdout:**

```json
{
  "ready": true,
  "backend": "ffmpeg",
  "video": "projects/<dự án>/exports/<tên>_video.mp4",
  "subtitle": "projects/<dự án>/exports/<tên>_video.srt",
  "duration_seconds": 512.4,
  "size_mb": 148.2,
  "slides": 12,
  "installed": [],
  "warnings": [],
  "error": null
}
```

- `error` là `null` hoặc `{ "step", "message", "fix" }`; `step` thuộc `audio`, `narrated_pptx`, `chromium`, `ffmpeg`, `render`, `powerpoint`.
- `--plan-only` in `{ "backend", "steps": [...], "warnings": [...] }` mà không dựng gì, để test.
- Mã thoát 0 khi `ready`, ngược lại 1.
- Chỉ dùng thư viện chuẩn Python; gọi `ffmpeg`, `ffprobe` và các script upstream bằng tiến trình con.

### 6.2 `-Action tool -Name chromium`

- Đã có `ffmpeg` và `pandoc`. Thêm `chromium`: cài `playwright` vào `venv` rồi chạy `playwright install chromium`.
- Kiểm đã có chưa bằng `venv\Scripts\python.exe -c "import playwright"` cộng với thư mục `%LOCALAPPDATA%\ms-playwright`.
- JSON giữ nguyên khuôn hiện tại: `tool`, `found`, `installed`, `dir`, `error`.

### 6.3 `docs/vi/lam-video.md`

Cho thầy cô: cần gì trước (bài giảng và ghi chú), thời gian dựng, dung lượng video, chọn phụ đề rời hay in lên hình, cách đăng YouTube kèm file `.srt`, và lời nhắc nghe lại tiếng máy đọc trước khi giao cho học sinh.

### 6.4 `docs/vi/tro-ly/video-bai-giang.md`

Theo đúng khuôn 8 mục của 5 loại việc hiện có (xem §4.2). Câu hỏi bắt buộc 4 câu, mỗi câu kèm "Gợi ý:": giọng đọc (nữ `vi-VN-HoaiMyNeural` hay nam `vi-VN-NamMinhNeural`), tốc độ đọc, phụ đề, độ phân giải. Mục "Tạo nhanh" 2 câu: giọng đọc và phụ đề.

Mục `## Khổ slide` của file này chỉ nêu `ppt169` trong backtick; độ phân giải (1080 hoặc 720) và số khung hình viết chữ thường để không vỡ test key canvas. File không chứa link Markdown; mọi đường dẫn đặt trong backtick.

### 6.5 `AGENTS.vi.md`

- Mục 3: thêm "làm video bài giảng", "lồng tiếng", "xuất video" vào danh sách câu lệnh tiếng Việt.
- Mục 11 mới: khi nào đọc `docs/vi/tro-ly/video-bai-giang.md`, thứ tự chạy (ghi chú → thuyết minh → video), và nhắc rằng mọi lệnh Python dùng Python của `venv`.

---

## 7. Trường hợp biên

| Tình huống | Cách xử lý |
|---|---|
| Dự án chưa có `notes/*.md` | `error.step = "audio"`, cách sửa: chạy bước ghi chú rồi bước thuyết minh của upstream |
| Có ghi chú nhưng chưa có tiếng | AI chạy `notes_to_audio.py` với giọng Việt rồi dựng tiếp |
| Số file tiếng khác số slide | Dừng, nêu tên slide thiếu tiếng |
| Máy không có PowerPoint mà thầy cô chọn `--cach powerpoint` | Dừng, gợi ý dùng `--cach ffmpeg` |
| Chưa có Chromium | `error.step = "chromium"`, cách sửa là chạy `-Action tool -Name chromium`; AI hỏi thầy cô trước khi tải |
| Máy chủ xem trước đang chạy sẵn | Dùng lại, không tắt của người khác; chỉ tắt cái do mình bật |
| Chụp ảnh lỗi một trang | Dừng, nêu tên trang; không dựng video thiếu trang |
| Ổ đĩa còn dưới 2 GB | Cảnh báo trước khi dựng |
| Đường dẫn dự án quá dài | Cảnh báo như bộ cài, vì FFmpeg cũng ghi file tạm |
| Video đã tồn tại | Ghi thêm hậu tố thời gian, không ghi đè |

---

## 8. Kiểm thử

### 8.1 Test tự động (`tools/vi/tests/test_video.py`)

Chạy nền, không mở cửa sổ, không cần mạng, không dựng video thật:

- Chọn cách dựng: có PowerPoint, không có PowerPoint, thầy cô chỉ định, thiếu `_narrated.pptx`.
- Phụ đề: cộng dồn mốc thời gian từ nhiều file SRT giả, kiểm từng mốc và định dạng `HH:MM:SS,mmm`.
- `--plan-only`: các bước và cảnh báo đúng cho từng tình huống (thiếu tiếng, thiếu Chromium, ổ đĩa đầy giả lập).
- Lệnh FFmpeg sinh ra: đúng số ảnh, đúng thời lượng từng ảnh, có `xfade`, có `-c:v libx264`.
- Nội dung tài liệu và quy tắc: `test_vi_layer.py` như các gói trước.

### 8.2 Chạy thật trên máy này (chủ repo đã đồng ý)

Máy này có PowerPoint, FFmpeg, `ffprobe`; chưa có Playwright/Chromium; ổ C còn khoảng 20 GB.

1. Lấy dự án `projects/thpt_gioi_thieu_ppt169_20260911` (3 slide, đã có PPTX), tạo ghi chú và tiếng thuyết minh bằng `edge-tts` (cần mạng).
2. Cài Chromium bằng `-Action tool -Name chromium` (tải khoảng 150–300 MB, chạy nền).
3. Dựng đường `ffmpeg`, kiểm: video mở được, tiếng khớp slide, phụ đề đúng mốc, chữ tiếng Việt đúng dấu.
4. Dựng đường `powerpoint`. **Bước này gọi PowerPoint thật nên sẽ có cửa sổ hiện lên**; chỉ chạy khi không làm phiền chủ repo.
5. Chạy lại lần hai: không cài lại gì, không để lại tiến trình máy chủ xem trước.

### 8.3 Nghiệm thu của chủ repo

Một bài giảng thật, có thuyết minh, xuất video cả hai đường, mở bằng trình phát, kiểm tiếng khớp slide và phụ đề đúng chỗ; thử đăng thử lên YouTube với file `.srt` rời.

---

## 9. Phát hành

Sau nghiệm thu: cập nhật `CHANGELOG-VI.md` và dòng phiên bản README thành `6.3.2-vi.4`; gộp vào `main`; tag `v6.3.2-vi.4`; `gh release create … --repo luonghaianh1208/PPTmaster`.

## 10. Ngoài phạm vi

- Remotion và video ngắn dọc (gói sau).
- Nhạc nền, chèn video rời, người dẫn ảo.
- Dựng video trên macOS/Linux.
- Cắt ghép lại video đã có, hoặc chỉnh sửa video ngoài luồng bài giảng.

## 11. Rủi ro

| Rủi ro | Giảm thiểu |
|---|---|
| Xuất video qua PowerPoint chiếm máy vài phút và hiện cửa sổ | Báo trước cho thầy cô trong tin nhắn; đường FFmpeg không có cửa sổ |
| Chromium nặng, mạng trường tải lâu | Hỏi trước khi tải; máy có PowerPoint thì không cần |
| Video 10 phút nặng 100–300 MB | Cảnh báo dung lượng trống trước khi dựng |
| Tiếng máy đọc sai tên riêng, thuật ngữ tiếng Anh | Tài liệu nhắc thầy cô nghe lại; có thể sửa ghi chú rồi dựng lại |
| Phụ đề lệch ở đường PowerPoint do thời gian chuyển cảnh | So tổng thời lượng, lệch quá 2% thì cảnh báo và giữ phụ đề rời |
| Máy chủ xem trước còn chạy sau khi lỗi | Luôn tắt trong khối `finally`; test kiểm không còn file khoá |
