# AGENTS.vi.md — Quy tắc bổ sung cho bản Việt

File này bổ sung ngữ cảnh Việt Nam cho [AGENTS.md](AGENTS.md). Nó không thay thế quy tắc nào của dự án gốc.

## 1. Thứ tự ưu tiên

- `skills/ppt-master/SKILL.md` và `AGENTS.md` luôn được ưu tiên khi có mâu thuẫn với file này.
- Không bao giờ sửa, bỏ qua hay tìm cách "sửa chữa" `skills/ppt-master/scripts/attribution_guard.py`. Nếu kiểm tra toàn vẹn thất bại, dừng lại và hướng dẫn người dùng tải lại bản đầy đủ.
- `tools/vi/` và `docs/vi/` là lớp Việt hoá của bản phân phối này; `tools/vi/tests/` là test riêng của lớp đó.

## 2. Ngôn ngữ

- Giữ quy tắc ngôn ngữ trong `SKILL.md`: trả lời theo ngôn ngữ người dùng đang dùng.
- Khi người dùng viết tiếng Việt và không yêu cầu ngôn ngữ khác, đặt ngôn ngữ nội dung của dự án (`primary_language`) là `vi`.

## 3. Câu lệnh tiếng Việt kích hoạt skill `ppt-master`

"tạo PPT", "làm slide", "làm bài giảng", "tạo bài thuyết trình", "làm poster", "làm báo cáo", "thêm thuyết minh", "làm đẹp slide", "làm video bài giảng", "lồng tiếng", "xuất video", "soạn đề", "làm đề kiểm tra", "đề tiếng Anh".

Các cụm "tạo nhanh", "làm nhanh", "không cần hỏi lại" là yêu cầu Quick tường minh (với bài tạo mới thông thường là hồ sơ `workflows/profiles/quick-generate.md` của upstream). Việc chọn hồ sơ và các bước thực hiện vẫn theo đúng `SKILL.md`.

## 4. Chạy lệnh trên Windows

- Có `venv\Scripts\python.exe` ở thư mục gốc repo thì dùng nó cho mọi lệnh `python3 …` hoặc `python …` của repo, ví dụ `venv\Scripts\python.exe skills/ppt-master/scripts/project_manager.py init <tên_dự_án>`.
- Nếu lệnh `python3 ...` báo không tìm thấy, trả mã 49 hoặc mở Microsoft Store, chạy lại đúng lệnh đó với `python`.
- Không dùng `cp`, `mkdir -p`, `/tmp` hay heredoc của bash trong PowerShell. Dùng lệnh tương đương (`Copy-Item`, `New-Item -ItemType Directory -Force`, `$env:TEMP`) hoặc Python.
- Khi ghi file có chữ tiếng Việt từ PowerShell 5.1, không dùng `Set-Content` hay `Out-File` mặc định. Dùng Python hoặc `[IO.File]::WriteAllText(path, text, (New-Object Text.UTF8Encoding $false))`.

## 5. Khổ canvas theo cách gọi của người Việt

<!-- format-map:start -->
| Người dùng nói | Key canvas |
|---|---|
| Slide 16:9, trình chiếu | `ppt169` |
| Slide 4:3, máy chiếu cũ | `ppt43` |
| Bài đăng Facebook/TikTok dọc 3:4 | `xiaohongshu` |
| Ảnh vuông Zalo/Facebook/Instagram | `moments` |
| Story, Reels, TikTok 9:16 | `story` |
| Ảnh bìa bài viết Zalo OA (2.35:1) | `wechat` |
<!-- format-map:end -->

## 6. Font

Với chữ tiếng Việt, chỉ dùng font có sẵn trên Windows hiển thị đủ dấu: Segoe UI, Arial, Calibri, Times New Roman (văn bản hành chính). Không dùng PingFang SC hay Microsoft YaHei cho chữ tiếng Việt.

## 7. Giọng thuyết minh (edge-tts)

- Giọng nữ: `vi-VN-HoaiMyNeural`
- Giọng nam: `vi-VN-NamMinhNeural`

## 8. Kết quả

Sau khi xuất, cho người dùng biết đường dẫn file PPTX trong thư mục dự án dưới `projects/`.

## 9. Môi trường: tự kiểm tra, tự cài và xử lý lỗi

- Mỗi cuộc trò chuyện, trước lệnh Python đầu tiên của repo (ví dụ `project_manager.py init`, `source_to_md.py`): chạy `tools/vi/doctor.py --no-smoke --json` (bằng Python của `venv` nếu có). Kết quả có `"ready": true` thì làm tiếp; chưa sẵn sàng, hoặc không chạy được Python, thì làm theo [docs/vi/cai-dat-bang-ai.md](docs/vi/cai-dat-bang-ai.md) rồi mới chạy lệnh đó.
- Yêu cầu thuộc mục 10: chạy kiểm tra trên (chỉ đọc, vài giây) trước khi gửi tin nhắn hỏi thầy cô. Chưa sẵn sàng thì thêm đúng một dòng ở cuối tin nhắn hỏi (ngay trước dòng chốt cách xác nhận, nếu có): "Máy chưa cài xong bộ công cụ; sau khi thầy cô trả lời, em sẽ cài trước (khoảng 5–10 phút) rồi làm bài." Không chạy được lệnh kiểm tra (công cụ chạy lệnh bị tắt hoặc bị từ chối) thì không kết luận máy chưa cài, mà dùng dòng này thay thế: "Em chưa kiểm tra được máy; sau khi thầy cô trả lời, em sẽ chạy kiểm tra `tools/vi/doctor.py` và cài nếu cần (khoảng 5–10 phút) rồi làm bài." Thầy cô trả lời xong thì cài ngay theo [docs/vi/cai-dat-bang-ai.md](docs/vi/cai-dat-bang-ai.md), rồi mới làm tiếp.
- Lệnh Python nào của repo báo không tìm thấy Python (kể cả sau khi chạy lại với `python` theo mục 4) hoặc báo `ModuleNotFoundError`: dừng, làm theo [docs/vi/cai-dat-bang-ai.md](docs/vi/cai-dat-bang-ai.md); không tự cài Python hay thư viện theo cách khác.
- Người dùng nhờ cài đặt, kiểm tra máy, hoặc dán link repo để cài: làm theo [docs/vi/cai-dat-bang-ai.md](docs/vi/cai-dat-bang-ai.md).
- Cần FFmpeg (thuyết minh, video) hoặc Pandoc (tài liệu định dạng cũ như `.doc`, `.odt`, `.rtf`) mà máy chưa có: làm theo mục "Công cụ tuỳ chọn" của file đó.
- Người dùng gặp lỗi môi trường: đề nghị chạy `KIEM-TRA.bat` (Windows) hoặc `python tools/vi/doctor.py`, rồi đối chiếu với [docs/vi/xu-ly-loi.md](docs/vi/xu-ly-loi.md).

## 10. Hỗ trợ thầy cô trước khi làm bài

Khi người dùng viết tiếng Việt và yêu cầu thuộc một trong 7 loại việc dưới đây, đọc [docs/vi/tro-ly/quy-trinh-hoi.md](docs/vi/tro-ly/quy-trinh-hoi.md) trước, rồi đọc file của loại việc đó. Hỏi thầy cô một lượt và chờ trả lời trước khi khởi tạo dự án.

| Loại việc | File hướng dẫn |
|---|---|
| Bài giảng | [docs/vi/tro-ly/bai-giang.md](docs/vi/tro-ly/bai-giang.md) |
| Báo cáo – tổng kết | [docs/vi/tro-ly/bao-cao-tong-ket.md](docs/vi/tro-ly/bao-cao-tong-ket.md) |
| Hoạt động Đoàn – sự kiện | [docs/vi/tro-ly/hoat-dong-doan.md](docs/vi/tro-ly/hoat-dong-doan.md) |
| Poster/ấn phẩm Zalo – Facebook | [docs/vi/tro-ly/poster-mang-xa-hoi.md](docs/vi/tro-ly/poster-mang-xa-hoi.md) |
| Tập huấn/workshop | [docs/vi/tro-ly/tap-huan-workshop.md](docs/vi/tro-ly/tap-huan-workshop.md) |
| Video bài giảng | [docs/vi/tro-ly/video-bai-giang.md](docs/vi/tro-ly/video-bai-giang.md) |
| Soạn đề KHTN tiếng Anh | [docs/vi/tro-ly/de-khtn-tieng-anh.md](docs/vi/tro-ly/de-khtn-tieng-anh.md) |

- `SKILL.md` vẫn được ưu tiên. Lượt hỏi này chỉ tạo thêm tài liệu nguồn; bước xác nhận của upstream vẫn bắt buộc, trừ khi người dùng yêu cầu tạo nhanh (xem mục 3) hoặc thuộc loại việc "Soạn đề KHTN tiếng Anh" (mục 12) — loại việc đó không có bước xác nhận của upstream, xem mục 12. Khi tạo nhanh, kể cả với "không cần hỏi lại", vẫn đọc `docs/vi/tro-ly/quy-trinh-hoi.md` và làm theo mục "Tạo nhanh" của file đó: có thể không hỏi câu nào, nhưng vẫn ghi brief.
- Yêu cầu không thuộc 7 loại (bối cảnh trường học hay Đoàn một mình không đủ để xếp loại), hoặc người dùng không viết tiếng Việt: làm theo `SKILL.md` như bình thường, không tìm hồ sơ đơn vị và không dùng bộ câu hỏi Việt.

## 11. Làm video bài giảng

Khi người dùng yêu cầu làm video từ một bài giảng đã có, đọc [docs/vi/tro-ly/video-bai-giang.md](docs/vi/tro-ly/video-bai-giang.md), hỏi một lượt theo file đó, rồi làm đúng thứ tự sau. Như mục 4: có `venv\Scripts\python.exe` ở thư mục gốc repo thì dùng nó cho mọi lệnh Python dưới đây, không có thì dùng `python`.

1. Chưa có `notes/*.md`: viết ghi chú lời giảng cho từng slide theo quy trình của upstream.
2. Chưa có `audio/*.mp3`: chạy `skills/ppt-master/scripts/notes_to_audio.py <đường_dẫn_dự_án> --voice vi-VN-HoaiMyNeural` (hoặc `vi-VN-NamMinhNeural`). Tốc độ đọc: chậm → thêm `--rate -10%`; vừa → không thêm cờ nào; nhanh → thêm `--rate +15%`.
3. Chưa có `exports/*_narrated.pptx`: xuất bản PPTX đã gắn tiếng bằng `skills/ppt-master/scripts/svg_to_pptx.py <đường_dẫn_dự_án> --recorded-narration audio`; dự án tạo nhanh (không có `spec_lock.md`) thì thêm `--quick-generate --with-notes`. Bỏ bước này thì đường PowerPoint không chạy được, AI buộc phải ghép bằng FFmpeg và phải tải Chromium. Chi tiết ở [docs/audio-narration.md](docs/audio-narration.md).
4. Dựng video: `venv\Scripts\python.exe tools\vi\video.py <đường_dẫn_dự_án>` kèm các cờ chọn theo đúng câu trả lời của thầy cô.

| Thầy cô trả lời | Cờ thêm vào |
|---|---|
| Phụ đề để thành file riêng | `--phu-de file` |
| Phụ đề in lên hình | `--phu-de hinh` |
| Không cần phụ đề | `--phu-de khong` |
| Độ phân giải 1080 | `--do-phan-giai 1080` |
| Độ phân giải 720 | `--do-phan-giai 720` |
| Muốn giữ hiệu ứng chuyển cảnh | `--cach powerpoint` |
| Không muốn mở PowerPoint | `--cach ffmpeg` |
| Không nêu cách dựng | `--cach auto` |

5. Đọc dòng JSON ở stdout. `ready` là `true` thì báo thầy cô đường dẫn video, thời lượng, dung lượng và nơi để phụ đề; `error` khác `null` thì làm theo `error.fix`, tối đa một lần, rồi báo thầy cô.

- `error.step` là `chromium`: hỏi thầy cô trước rồi chạy `powershell -NoProfile -ExecutionPolicy Bypass -File tools\vi\pptmaster.ps1 -Action tool -Name chromium`, vì bước này tải khoảng 150–300 MB.
- `error.step` là `audio`: quay lại bước 1 nếu dự án chưa có `notes/*.md`, quay lại bước 2 nếu đã có ghi chú.
- Cách dựng `powerpoint` sẽ mở **cửa sổ PowerPoint** và chiếm máy vài phút; báo trước cho thầy cô một dòng.
- Đường FFmpeg ghép từ ảnh chụp slide nên độ phân giải cao nhất bằng khổ slide, tức 1280×720 với `ppt169`; chọn 1080 ở đường này cũng không nét hơn.
- Không tự cài phần mềm nào khác, không tự chạy FFmpeg theo cách riêng.

## 12. Soạn đề KHTN bằng tiếng Anh

Khi người dùng cần đề kiểm tra KHTN bằng tiếng Anh, đọc [docs/vi/tro-ly/de-khtn-tieng-anh.md](docs/vi/tro-ly/de-khtn-tieng-anh.md) và [docs/vi/tro-ly/tieng-anh-khoa-hoc.md](docs/vi/tro-ly/tieng-anh-khoa-hoc.md) rồi làm đúng thứ tự dưới. Như mục 4: có `venv\Scripts\python.exe` ở thư mục gốc repo thì dùng nó cho mọi lệnh Python dưới đây.

1. Luồng A (đã có đề tiếng Việt): đọc đề thầy cô đưa bằng `python skills/ppt-master/scripts/source_to_md.py <file> -o <thư_mục_tạm>` trước khi hỏi gì. Thầy cô đưa **ảnh** thì nói rõ không đọc được ảnh, xin bản PDF hoặc Word. Luồng B (đề mới hoàn toàn) không có gì để đọc, bỏ qua bước này.
2. Sau bước 1 (luồng A) hoặc ngay từ đầu (luồng B), hỏi một lượt theo file hướng dẫn: luồng A lấy câu 3, 4 và 5 (chủ đề, số câu mỗi phần, tỉ lệ mức độ) trực tiếp từ đề vừa đọc, chỉ hỏi phần đề không trả lời được; luồng B hỏi đủ theo file đó.
3. Tạo `projects/_de-thi/<tên_đề>/`. Luồng B soạn ma trận đặc tả trước (chủ đề × mức độ × số câu) theo câu trả lời của thầy cô, viết câu hỏi trực tiếp bằng tiếng Anh, ghi bản tiếng Việt của từng câu vào `vi:` để thầy cô soát. Viết `de.md` theo đúng ngữ pháp trong file hướng dẫn.
4. Chạy `python tools\vi\de_thi.py projects\_de-thi\<tên_đề>`; thêm `--phan de,dap-an` khi thầy cô không cần bản song ngữ; thêm `--plan-only` khi chỉ muốn kiểm cú pháp.
5. Đọc dòng JSON ở stdout. `ready` là `true` thì báo thầy cô đường dẫn các file thực sự sinh ra (hai file khi chạy `--phan de,dap-an`, ba file với các trường hợp còn lại), số câu mỗi phần, tổng điểm, đọc nguyên văn các dòng `warnings`, và báo cho thầy cô các mục trong "Cần thầy cô soát" của file đáp án.

| `error.step` | Xử lý |
|---|---|
| `input` | Chưa có `de.md`, viết file rồi chạy lại. |
| `parse` | Sửa đúng dòng `error.message` nêu rồi chạy lại. |
| `docx` | Có `venv\Scripts\python.exe` ở thư mục gốc repo thì chạy `venv\Scripts\python.exe -m pip install -r tools/vi/requirements-vi.txt`; không thì chạy `python -m pip install -r tools/vi/requirements-vi.txt` (thư viện `python-docx`). Chạy lại lệnh xuất, tối đa một lần. |
| `write` | Xin thầy cô đóng file Word đang mở rồi chạy lại. |
| `internal` | Lỗi ngoài dự kiến; dán nguyên `error.message` để báo cho người bảo trì, không tự đoán cách sửa. |

Điều cấm: không tự sửa số liệu hay đáp án của đề gốc; không chạy `project_manager.py init`; không tạo SVG; không chạm `skills/`; không commit gì trong `projects/`.
