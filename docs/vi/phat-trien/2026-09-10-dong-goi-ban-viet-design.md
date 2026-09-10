# Thiết kế: Đóng gói PPT Master bản Việt ("clone về là dùng")

- **Ngày:** 2026-09-10
- **Người phụ trách:** Lương Hải Anh — 2Anh AI Education
- **Trạng thái:** Đã duyệt; đang triển khai M1 (plan: `2026-09-10-m1-plan.md`)
- **Repo:** `luonghaianh1208/PPTmaster` (công khai, miễn phí)
- **Nền tảng:** upstream `hugohe3/ppt-master` tag `v6.3.2` (MIT, © 2025-2026 Hugo He)

---

## 1. Mục tiêu

Khách hàng (giáo viên, nhân viên, người làm nội dung — phần lớn dùng Windows, một phần đã quen AI coding) tải repo về, bấm cài đặt một lần, mở thư mục trong AI editor (Claude Code / Cursor / Antigravity) và tạo được bài thuyết trình PPTX tiếng Việt — không cần biết lập trình.

### Tiêu chí thành công

1. Mô phỏng trên máy phát triển 3 kịch bản PATH — (a) không có Python, (b) chỉ có lối tắt Microsoft Store, (c) có Python ≥ 3.10 trong môi trường ảo tách biệt: `CAI-DAT` cho đúng thông báo và mã thoát; ở (c) `KIEM-TRA` báo toàn bộ mục **bắt buộc** ✅. (Máy phát triển không chạy được Windows Sandbox vì ảo hoá firmware đang tắt; chủ repo chọn chỉ mô phỏng.)
2. Mở thư mục trong Claude Code, Cursor và Antigravity; nhập "Tạo 3 slide giới thiệu trường THPT" → AI trả lời tiếng Việt, xuất PPTX mở được trong PowerPoint, chữ tiếng Việt hiển thị đúng dấu.
3. `python skills/ppt-master/scripts/attribution_guard.py` trả về exit 0 trên mọi commit của `main`.
4. `git diff v6.3.2 main -- skills/` chỉ gồm: 5 thư mục template tiếng Việt + `decks_index.json` + `brands_index.json`.
5. Người đã clone bản cũ chạy `git pull` → fast-forward, không lỗi. 29 fork hiện có không bị ảnh hưởng (không force push).
6. Diễn tập đồng bộ: áp lớp Việt hoá lên nhánh thử dựng từ `v6.3.1`, chạy `tools/vi/sync_upstream.ps1 v6.3.2` → hoàn tất không cần sửa tay, test pass.

---

## 2. Bối cảnh & phát hiện

| # | Phát hiện | Hệ quả |
|---|---|---|
| 1 | Repo hiện tại là fork của `hugohe3/ppt-master` (khớp gần nhất tag `v2.3.0`), nhưng `LICENSE` đã đổi copyright sang tên người fork; ghi công gốc chỉ còn trong slide ví dụ. | Vi phạm điều khoản MIT khi phân phối lại → phải khôi phục ghi công. |
| 2 | Phần tuỳ biến của fork mỏng: đổi tên định dạng (WeChat→Zalo OA, Xiaohongshu→Facebook/TikTok, Moments→Zalo/Instagram), trigger tiếng Việt trong SKILL.md, sửa UTF-8 ở `config.py`, docs tiếng Việt. `svg_to_shapes.py`/`svg_to_pptx.py` của fork **cũ hơn** v2.3.0 (thoái lui). | Chuyển sang upstream mới rẻ; không mất tính năng riêng nào. |
| 3 | Upstream đang ở `v6.3.2`: xuất PPTX native editable, 203 hiệu ứng, thuyết minh (edge-tts), `update_repo.py`, `.env` loader, `console_encoding.py` (UTF-8 stdout), examples tách ra repo riêng. | Không tự viết lại những thứ upstream đã có. |
| 4 | Upstream có `scripts/attribution_guard.py`: dừng mọi script (exit 78) nếu đổi frontmatter `copyright`/`license`/`official_repository`/`sponsors` trong `SKILL.md`, đổi `LICENSE` (so hash), xoá `SPONSORS.md`/`SPONSORS_CN.md`, hoặc gỡ gate. | Lớp Việt hoá **chỉ được cộng thêm**, không sửa file danh tính của skill. |
| 5 | Discovery template chỉ đọc `templates/<kind>/<kind>s_index.json`, không quét thư mục. `register_template.py --rebuild-all --kind <k>` dựng lại index từ mọi thư mục có `templates/design_spec.md`. | Thư mục template mới không xung đột; xung đột index giải bằng `--rebuild-all`. |
| 6 | Upstream không có installer, doctor, cấu hình Cursor/Antigravity; tài liệu và SKILL.md dùng `python3` (SKILL.md có dòng fallback `python` trên Windows). | Giá trị riêng của bản Việt: bộ cài Windows, kiểm tra môi trường, docs tiếng Việt, template Việt. |

---

## 3. Quyết định đã chốt

| Chủ đề | Quyết định |
|---|---|
| Nền tảng | Dựng lại trên upstream `v6.3.2`, thêm lớp Việt hoá |
| Phân phối | Repo công khai, miễn phí |
| Repo | Giữ `PPTmaster`; merge lịch sử upstream bằng commit mới, **không force push** |
| Đối tượng | Cả người không rành kỹ thuật (Windows) lẫn người quen AI coding |
| Phạm vi giai đoạn 1 | Bộ cài + kiểm tra + cập nhật; kết nối AI editor; tài liệu tiếng Việt; template tiếng Việt |
| Template | 3 deck (`bai_giang_thpt`, `bao_cao_tong_ket`, `hoat_dong_doan`) + 2 brand (`doan_thanh_nien`, `2anh_ai`) |
| Font | Font hệ thống Windows hỗ trợ tiếng Việt: Segoe UI, Arial, Calibri, Times New Roman |
| Phát hành | 2 mốc: M1 `v6.3.2-vi.1` (nền tảng + bộ cài + docs), M2 `v6.3.2-vi.2` (template) |
| Tài nguyên | Chủ repo cung cấp logo 2Anh AI và huy hiệu Đoàn (SVG hoặc PNG nền trong suốt ≥ 1000px) |
| Nơi lưu spec | `docs/vi/phat-trien/` |

---

## 4. Nguyên tắc kiến trúc

**Upstream nguyên vẹn + lớp Việt hoá cộng thêm.**

File upstream được phép thay đổi (danh sách đóng):

| File | Thay đổi | Cách giữ khi đồng bộ upstream |
|---|---|---|
| `README.md` | Thay toàn bộ bằng README tiếng Việt | `.gitattributes`: `README.md merge=ours` |
| `CLAUDE.md` | Thêm 1 dòng `@AGENTS.vi.md` sau `@AGENTS.md` | Merge thường (xung đột hiếm, 1 dòng) |
| `skills/ppt-master/templates/decks/decks_index.json` | Thêm 3 entry qua `register_template.py` | Lấy bản upstream → `--rebuild-all --kind deck` |
| `skills/ppt-master/templates/brands/brands_index.json` | Thêm 2 entry qua `register_template.py` | Lấy bản upstream → `--rebuild-all --kind brand` |

Mọi file khác của upstream — đặc biệt `skills/ppt-master/SKILL.md`, `LICENSE`, `SPONSORS*.md`, `AGENTS.md`, `scripts/` — **không được sửa**.

### Cấu trúc lớp Việt hoá

```text
PPTmaster/
├── README.md                     # (thay) README tiếng Việt
├── CLAUDE.md                     # (sửa 1 dòng) @AGENTS.md + @AGENTS.vi.md
├── AGENTS.vi.md                  # (mới) quy tắc bổ sung cho AI khi dùng bản Việt
├── NOTICE                        # (mới) ghi công bản gốc + bản Việt hoá
├── CHANGELOG-VI.md               # (mới) lịch sử phiên bản bản Việt
├── .gitattributes                # (mới) README.md merge=ours
├── CAI-DAT.bat                   # (mới) bấm đúp để cài
├── KIEM-TRA.bat                  # (mới) bấm đúp để kiểm tra môi trường
├── CAP-NHAT.bat                  # (mới) bấm đúp để cập nhật
├── .cursor/rules/ppt-master-vi.mdc   # (mới) nạp AGENTS.vi.md cho Cursor
├── .agents/rules/ppt-master-vi.md    # (mới) nạp AGENTS.vi.md cho Antigravity (xem §5.4)
├── tools/vi/
│   ├── pptmaster.ps1             # trình khởi chạy Windows: -Action setup | check | update
│   ├── setup.sh                  # bộ cài tối giản macOS/Linux
│   ├── doctor.py                 # kiểm tra môi trường + smoke test
│   ├── sync_upstream.ps1         # (chủ repo) đồng bộ bản upstream mới
│   ├── fixtures/smoke/           # dữ liệu smoke test (1 SVG tiếng Việt)
│   └── tests/                    # unittest cho doctor + tính nhất quán lớp Việt
├── docs/vi/
│   ├── bat-dau-nhanh.md
│   ├── cai-dat-windows.md
│   ├── cau-lenh-mau.md
│   ├── xu-ly-loi.md
│   ├── lay-api-key.md
│   └── phat-trien/
│       ├── bao-tri.md            # quy trình đồng bộ upstream, phát hành
│       └── 2026-09-10-dong-goi-ban-viet-design.md
└── skills/ppt-master/templates/
    ├── decks/{bai_giang_thpt,bao_cao_tong_ket,hoat_dong_doan}/
    └── brands/{doan_thanh_nien,2anh_ai}/
```

`.gitignore` của upstream `v6.3.2` chỉ chặn `.claude/`, không chặn `.cursor/` hay `.agents/` → không cần sửa `.gitignore`. `test_vi_layer.py` kiểm tra 2 file rule vẫn được git theo dõi sau mỗi lần đồng bộ.

---

## 5. Thành phần

### 5.1 Chuyển repo sang nền upstream (M1)

Trên nhánh `feat/vi-packaging`:

1. Gắn tag `v2-vi-legacy` vào commit `87e16ca` (trạng thái fork cũ) để tham chiếu.
2. `git merge -s ours --no-commit --allow-unrelated-histories v6.3.2` — ghi nhận `v6.3.2` là parent.
3. `git read-tree --reset -u v6.3.2` — thay toàn bộ cây file bằng cây upstream (xoá `examples/`, `viewer.html`, `svg_to_shapes.py` cũ…).
4. Khôi phục riêng `docs/vi/phat-trien/` từ `HEAD` trước merge (spec này).
5. Commit merge. Kiểm tra: `git diff v6.3.2 HEAD --stat` chỉ còn `docs/vi/phat-trien/`.
6. Thêm lớp Việt hoá thành các commit riêng theo §5.2–§5.8.
7. Khi M1 đạt DoD: fast-forward `main`, push `main` + tag `v2-vi-legacy` + tag `v6.3.2-vi.1`, tạo GitHub Release.

Các thay đổi riêng của fork cũ **không** mang sang: đổi key định dạng (thay bằng bảng ánh xạ ở `AGENTS.vi.md`), `examples/`, `index.html`/`viewer.html` cũ, bản script thoái lui.

### 5.2 Ghi công — `NOTICE`

Nội dung bắt buộc:
- Tên dự án gốc, tác giả Hugo He, link `https://github.com/hugohe3/ppt-master`, giấy phép MIT, trỏ tới `LICENSE`.
- "Bản Việt hoá và đóng gói: Lương Hải Anh — 2Anh AI Education", link repo `PPTmaster`, mô tả phần đã thêm (bộ cài, tài liệu tiếng Việt, template Việt).
- Ghi chú phụ thuộc: PyMuPDF (AGPL) được cài từ PyPI bởi người dùng, không đóng gói kèm; icon/âm thanh xem các `THIRD_PARTY_NOTICES.md` của upstream.
- Song ngữ Việt/Anh.

README ghi công ở đầu trang (ngay dưới tiêu đề) và mục "Giấy phép & Ghi công".

### 5.3 Bộ cài, kiểm tra, cập nhật

Ba file `.bat` chỉ chứa ký tự ASCII và cùng gọi `powershell -NoProfile -ExecutionPolicy Bypass -File tools\vi\pptmaster.ps1 -Action <setup|check|update>` (tránh lỗi execution policy). Mọi thông báo tiếng Việt nằm trong `pptmaster.ps1` (UTF-8 **có BOM**, đặt `[Console]::OutputEncoding` UTF-8) để Windows PowerShell 5.1 hiển thị đúng. Tham số `-NonInteractive`: không hỏi Y/N, không cài phần mềm (dùng khi kiểm thử).

**`CAI-DAT.bat` → `pptmaster.ps1 -Action setup`**

1. Tìm `python` trong PATH. Nếu không có, hoặc đường dẫn nằm trong `WindowsApps` và `python --version` thất bại (alias Microsoft Store), hoặc phiên bản < 3.10:
   - Nếu có `winget`: hỏi Y/N rồi `winget install -e --id Python.Python.3.12`; sau đó yêu cầu đóng/mở lại cửa sổ và chạy lại `CAI-DAT.bat`.
   - Nếu không có `winget`: in link python.org + nhắc tick "Add python.exe to PATH", mở `docs/vi/cai-dat-windows.md`, thoát mã 1.
   - Nếu phát hiện alias Store: hướng dẫn tắt tại *Settings → Apps → Advanced app settings → App execution aliases*.
2. `python -m pip install --upgrade pip` rồi `python -m pip install -r requirements.txt`. Thất bại → giữ nguyên log pip trên màn hình, in gợi ý trong `docs/vi/xu-ly-loi.md`, thoát mã 1.
3. Nếu chưa có `.env` ở thư mục gốc: sao chép từ `.env.example`.
4. Hỏi Y/N cài công cụ tuỳ chọn qua `winget` (nếu có): Git (`Git.Git`), Pandoc (`JohnMacFarlane.Pandoc`), FFmpeg (`Gyan.FFmpeg`). ID gói được xác minh bằng `winget show` khi triển khai.
5. Chạy `python tools/vi/doctor.py` và thoát theo mã của doctor.

Chạy lại nhiều lần an toàn (idempotent).

**`tools/vi/setup.sh`** (macOS/Linux, cho người rành kỹ thuật): kiểm tra `python3` ≥ 3.10, `python3 -m pip install -r requirements.txt`, sao chép `.env` nếu thiếu, chạy doctor. Không tự cài phần mềm hệ thống.

**`KIEM-TRA.bat` → `pptmaster.ps1 -Action check` → `python tools/vi/doctor.py`**

Chỉ dùng thư viện chuẩn Python; ngay đầu gọi `sys.stdout.reconfigure(encoding="utf-8", errors="replace")` và tương tự cho `sys.stderr`. In bảng ✅ / ⚠️ / ❌ bằng tiếng Việt, mỗi lỗi kèm một dòng cách sửa.

| Kiểm tra | Mức | Cách kiểm | Khi lỗi |
|---|---|---|---|
| Python ≥ 3.10 | Bắt buộc | `sys.version_info` | Chạy `CAI-DAT.bat` |
| Thư viện Python | Bắt buộc | Đọc `skills/ppt-master/requirements.txt`: bỏ dòng trống/comment, bỏ phần phiên bản (`>=`, `==`, `~=`…), marker (`; …`) và extras (`[…]`) → với mỗi tên gói, `importlib.metadata.distribution(<tên>)` phải tồn tại | Liệt kê gói thiếu, chạy `CAI-DAT.bat` |
| Tính toàn vẹn skill | Bắt buộc | `attribution_guard.py` exit 0 | Tải lại bản đầy đủ; không sửa `LICENSE`/`SKILL.md`/`SPONSORS*` |
| Smoke test xuất PPTX | Bắt buộc | Chép `tools/vi/fixtures/smoke/01_smoke.svg` (1280×720, chứa "Kiểm tra tiếng Việt") vào `<tmp>/svg_output/`, chạy `finalize_svg.py <tmp> -q` rồi `svg_to_pptx.py <tmp> -s final -o <tmp>/smoke.pptx --no-notes --no-animations -q` (chế độ chẩn đoán của upstream, không cần `spec_lock.md`); pass khi file là zip hợp lệ, có đúng 1 `ppt/slides/slideN.xml` chứa chuỗi trên. Đã chạy thật trên v6.3.2: ~2,5 giây | In bước lỗi + 3 dòng log cuối, trỏ `docs/vi/xu-ly-loi.md` |
| Git | Khuyến nghị | `shutil.which("git")` | Cần để dùng `CAP-NHAT.bat` |
| Pandoc | Tuỳ chọn | `shutil.which("pandoc")` | Chỉ cần cho định dạng tài liệu cũ |
| FFmpeg / FFprobe | Tuỳ chọn | `shutil.which` | Chỉ cần cho thuyết minh/video |
| API key tạo ảnh | Tuỳ chọn | Đọc `.env`/biến môi trường: có ít nhất một key ảnh (`GEMINI_API_KEY`, `OPENAI_API_KEY`, …) — **chỉ báo có/không, không in giá trị** | Xem `docs/vi/lay-api-key.md` |

- Mã thoát: `0` khi mọi mục bắt buộc pass; `1` nếu có mục bắt buộc lỗi.
- Tuỳ chọn `--no-smoke`: bỏ smoke test (dùng trong `sync_upstream.ps1` khi cần nhanh).
- Smoke test chỉ chạy khi mục Thư viện và Tính toàn vẹn đều đạt.

**`CAP-NHAT.bat` → `pptmaster.ps1 -Action update`**: `python skills\ppt-master\scripts\update_repo.py` (upstream: kiểm tra git + cây sạch, `git pull --ff-only`, cài lại thư viện nếu requirements đổi), sau đó `python tools\vi\doctor.py --no-smoke`. Nếu không có `.git` (tải ZIP): in hướng dẫn tải bản mới, thoát mã 1.

### 5.4 Kết nối AI editor — `AGENTS.vi.md`

Cơ chế nạp (không cần người dùng thao tác):

| Công cụ | Cơ chế |
|---|---|
| Claude Code | `CLAUDE.md` chứa `@AGENTS.md` và `@AGENTS.vi.md` |
| Cursor | Tự đọc `AGENTS.md`; thêm `.cursor/rules/ppt-master-vi.mdc` với `alwaysApply: true` tham chiếu `@AGENTS.vi.md` |
| Antigravity | `.agents/rules/ppt-master-vi.md` (thư mục mặc định; vẫn hỗ trợ `.agent/rules`) tham chiếu `@../../AGENTS.md` và `@../../AGENTS.vi.md` (đường dẫn tính từ file rule); frontmatter `trigger: always_on` — tài liệu chính thức chưa nêu cú pháp frontmatter nên phải kiểm chứng |

Mỗi cơ chế phải được kiểm chứng bằng tay ở M1 (hỏi AI "Bạn đang áp dụng quy tắc nào của bản Việt?"). Nếu cơ chế của một công cụ không nạp được, phương án dự phòng: thêm **một dòng** cuối `AGENTS.md` trỏ tới `AGENTS.vi.md` (bổ sung vào danh sách file upstream được sửa ở §4).

Nội dung bắt buộc của `AGENTS.vi.md` (bổ sung, không mâu thuẫn `SKILL.md`):

1. **Ưu tiên:** quy tắc trong `skills/ppt-master/SKILL.md` luôn thắng; file này chỉ bổ sung ngữ cảnh Việt Nam. Không bao giờ sửa, bỏ qua hay "sửa chữa" `attribution_guard.py`.
2. **Ngôn ngữ:** không ghi cứng ngôn ngữ trả lời (quy ước `docs/rules/language.md` của upstream) — giữ quy tắc của `SKILL.md` là trả lời theo ngôn ngữ người dùng; khi người dùng viết tiếng Việt, đặt `primary_language` của dự án là `vi`.
3. **Từ khoá kích hoạt:** "tạo PPT", "làm slide", "làm bài giảng", "tạo bài thuyết trình", "làm poster", "làm báo cáo", "thêm thuyết minh", "làm đẹp slide" → dùng skill `ppt-master`.
4. **Windows:** nếu `python3` lỗi (không tìm thấy / mã 49 / mở Microsoft Store) → dùng `python`. Không dùng `cp`, `mkdir -p`, `/tmp`, heredoc bash trên Windows — dùng lệnh PowerShell tương đương hoặc Python.
5. **Ánh xạ định dạng** (tên người Việt hay nói → key upstream):

   | Người dùng nói | Key |
   |---|---|
   | Slide 16:9, trình chiếu | `ppt169` |
   | Slide 4:3, máy chiếu cũ | `ppt43` |
   | Bài đăng Facebook/TikTok dọc 3:4 | `xiaohongshu` |
   | Ảnh vuông Zalo/Facebook/Instagram | `moments` |
   | Story, Reels, TikTok 9:16 | `story` |
   | Ảnh bìa bài viết Zalo OA (2.35:1) | `wechat` |

6. **Font:** chỉ dùng font có sẵn trên Windows hỗ trợ tiếng Việt — Segoe UI, Arial, Calibri, Times New Roman (văn bản hành chính). Không dùng PingFang SC / Microsoft YaHei cho chữ tiếng Việt.
7. **Giọng thuyết minh (edge-tts):** nữ `vi-VN-HoaiMyNeural`, nam `vi-VN-NamMinhNeural`.
8. **Template Việt:** khi phù hợp, gợi ý `bai_giang_thpt` (bài giảng), `bao_cao_tong_ket` (báo cáo, sơ kết, tổng kết), `hoat_dong_doan` (sự kiện Đoàn, học sinh), brand `doan_thanh_nien`, brand `2anh_ai` — chỉ khi đã có trong index (từ M2).
9. **Nơi lưu kết quả:** nhắc người dùng file PPTX nằm trong `projects/<tên_dự_án>/`.

### 5.5 Tài liệu tiếng Việt

| File | Nội dung |
|---|---|
| `README.md` | Giới thiệu 1 đoạn; ghi công bản gốc; 3 bước bắt đầu (tải → `CAI-DAT.bat` → mở trong AI editor); bảng tính năng; link docs; cập nhật; giấy phép & ghi công |
| `docs/vi/bat-dau-nhanh.md` | Từ lúc cài xong đến PPTX đầu tiên: mở thư mục trong từng AI editor, câu lệnh đầu tiên, trả lời 8 câu xác nhận, tìm file kết quả |
| `docs/vi/cai-dat-windows.md` | Cài Python đúng cách (Add to PATH, tắt alias Store), tải repo (Git hoặc ZIP), `CAI-DAT.bat`, đọc kết quả `KIEM-TRA.bat` |
| `docs/vi/cau-lenh-mau.md` | Câu lệnh mẫu: bài giảng, báo cáo tổng kết, sự kiện Đoàn, poster Zalo, chuyển Word/PDF thành slide, thêm thuyết minh, làm đẹp PPTX có sẵn |
| `docs/vi/xu-ly-loi.md` | Lỗi thường gặp → cách sửa: `python3` mở Store, pip lỗi mạng, thiếu thư viện, guard báo lỗi toàn vẹn, `CAP-NHAT` báo cây không sạch, font lỗi dấu |
| `docs/vi/lay-api-key.md` | Lấy Gemini API key, điền vào `.env`, kiểm tra bằng `KIEM-TRA.bat`; lưu ý không chia sẻ `.env` |
| `docs/vi/phat-trien/bao-tri.md` | Quy trình `sync_upstream.ps1`, chạy test, cập nhật `CHANGELOG-VI.md`, gắn tag, tạo Release |
| `CHANGELOG-VI.md` | Mục cho `v6.3.2-vi.1`, `v6.3.2-vi.2` |

M1 không kèm ảnh chụp màn hình (không có máy sạch để chụp); bổ sung sau nếu cần.

### 5.6 Template tiếng Việt (M2)

Mỗi gói được tạo theo đúng quy trình upstream `workflows/create-template.md` (create-deck / create-brand): phân tích tham chiếu → brief → **chủ repo duyệt brief** → dựng → `svg_quality_checker.py <workspace>/templates --template-mode` → `template_preview_pptx.py` xuất PPTX xem thử → chủ repo duyệt hình → `register_template.py <id> --kind <kind>`.

| ID | Kind | Canvas | Roster dự kiến (chốt ở brief) | Nhận diện |
|---|---|---|---|---|
| `bai_giang_thpt` | deck | `ppt169` | Bìa (bài, môn, lớp, GV, trường) · Mục tiêu (kiến thức/năng lực/phẩm chất) · Mở đầu/khởi động · Tiêu đề hoạt động · Nội dung kiến thức 2 cột · Câu hỏi/trắc nghiệm · Hoạt động nhóm/phiếu học tập · Luyện tập · Vận dụng/bài về nhà · Tổng kết sơ đồ · Kết thúc | Học đường, sáng rõ; Segoe UI |
| `bao_cao_tong_ket` | deck | `ppt169` | Bìa (đơn vị, tên báo cáo, thời gian) · Mục lục · Kết quả nổi bật (KPI) · Bảng số liệu · Biểu đồ · Hạn chế – nguyên nhân · Phương hướng nhiệm vụ · Kiến nghị · Kết thúc | Hành chính trang trọng (đỏ/xanh); Times New Roman |
| `hoat_dong_doan` | deck | `ppt169` | Bìa sự kiện · Giới thiệu chương trình · Thể lệ/cơ cấu · Lịch trình · Đội thi/gương mặt · Hình ảnh hoạt động · Khen thưởng · Cảm ơn | Tích hợp nhận diện Đoàn; trẻ trung |
| `doan_thanh_nien` | brand | — | — | Màu, typography, huy hiệu Đoàn, giọng văn; dùng kèm layout upstream (vd. `moments_square` cho poster vuông) |
| `2anh_ai` | brand | — | — | Màu từ logo 2Anh AI, typography, logo, giọng văn thân thiện – thực tiễn cho giáo viên |

- ID `2anh_ai` bắt đầu bằng chữ số: xác minh với validator của `register_template.py` trước khi tạo; nếu bị từ chối dùng `hai_anh_ai`.
- Brand `doan_thanh_nien`: `design_spec.md` ghi lưu ý sử dụng huy hiệu đúng quy định nhận diện của Đoàn TNCS Hồ Chí Minh (không biến dạng, đúng màu).
- Logo do chủ repo cung cấp; lưu trong `images/` của workspace.

### 5.7 Đồng bộ upstream — `tools/vi/sync_upstream.ps1 -Tag <tag> [-Python <đường dẫn>] [-SkipSmoke]` (dành cho chủ repo)

1. Dừng nếu cây làm việc không sạch.
2. `git fetch upstream --tags`; dừng nếu `<tag>` không tồn tại.
3. `git config merge.ours.driver true` (kích hoạt `README.md merge=ours`).
4. `git merge --no-ff <tag>`.
5. Nếu xung đột:
   - `decks_index.json` / `brands_index.json` / `layouts_index.json` / `styles_index.json` → `git checkout --theirs`, rồi `python skills/ppt-master/scripts/register_template.py --rebuild-all --kind <kind>`, `git add`.
   - File khác → liệt kê, dừng với mã 1 để xử lý tay (không tự commit).
6. Commit merge (nếu còn mở). Merge luôn được commit trước khi kiểm tra, vì test ranh giới (§5.8) xác định tag upstream bằng `git describe` từ HEAD — trước khi commit, HEAD còn trỏ tag cũ.
7. Kiểm tra: `attribution_guard.py` exit 0 → `python -m unittest discover -s tools/vi/tests` pass → `doctor.py` pass. Nếu một kiểm tra thất bại: dừng, in hướng dẫn "merge đã commit nhưng chưa push — sửa lớp Việt rồi commit tiếp, hoặc hoàn tác bằng `git reset --keep ORIG_HEAD`". Không tự push.

### 5.8 Kiểm tra tính nhất quán lớp Việt — `tools/vi/tests/test_vi_layer.py`

Phát hiện lớp Việt hoá "trôi" khi upstream đổi:
- Mọi key định dạng trong bảng ánh xạ của `AGENTS.vi.md` tồn tại trong `CANVAS_FORMATS` của `skills/ppt-master/scripts/config.py`.
- (Từ M2) Mọi ID template nêu trong `AGENTS.vi.md` tồn tại trong index tương ứng.
- `CLAUDE.md` chứa `@AGENTS.vi.md`; `NOTICE` tồn tại và nhắc "Hugo He".
- `git diff` so với tag upstream gần nhất chỉ chạm các file trong danh sách §4 (bỏ qua nếu không có `.git`).

---

## 6. Luồng sử dụng

**Khách hàng**

```text
Tải repo (git clone hoặc ZIP)
  → CAI-DAT.bat  → (cài Python/thư viện, tạo .env) → doctor
  → Mở thư mục trong Claude Code / Cursor / Antigravity
  → Chat "Tạo bài giảng…"
      AI nạp CLAUDE.md/AGENTS.md + AGENTS.vi.md → SKILL.md → attribution_guard → routing → (template Việt)
  → PPTX trong projects/<tên_dự_án>/
  → Định kỳ: CAP-NHAT.bat
```

**Chủ repo**

```text
Upstream ra tag mới
  → tools/vi/sync_upstream.ps1 <tag>
  → Sửa tay nếu script dừng → test + doctor pass
  → Cập nhật CHANGELOG-VI.md → tag v<upstream>-vi.<n> → push → GitHub Release
```

---

## 7. Xử lý lỗi (nguyên tắc)

- Mọi script của lớp Việt in thông báo tiếng Việt, mỗi lỗi kèm **một** hành động cụ thể và link tài liệu `docs/vi/xu-ly-loi.md#<mục>`.
- Không in giá trị API key, token hay nội dung `.env`.
- Không bao giờ tự sửa/bỏ qua `attribution_guard.py`; khi guard lỗi chỉ hướng dẫn tải lại bản đầy đủ.
- Script chủ repo (`sync_upstream.ps1`) dừng an toàn khi gặp tình huống không tự giải được, không commit dở, không push.

---

## 8. Kiểm thử

| Loại | Nội dung | Khi nào |
|---|---|---|
| Unit (`unittest`, stdlib) | `tools/vi/tests/test_doctor.py`: từng check với mock (`shutil.which`, `importlib.metadata`, `subprocess`), mã thoát, không lộ key | M1 |
| Nhất quán | `tools/vi/tests/test_vi_layer.py` (§5.8) | M1, mở rộng ở M2 |
| Toàn vẹn | `attribution_guard.py` exit 0 | Mọi commit |
| Mô phỏng – máy phát triển | 3 kịch bản PATH (tiêu chí #1) chạy `pptmaster.ps1 -NonInteractive` và các `.bat`; bản sao không có `.git` cho `-Action update`; `setup.sh` qua Git Bash | M1 |
| Kiểm chứng AI editor | Claude Code headless xác nhận đã nạp `AGENTS.vi.md`; chủ repo tự tạo thử 3 slide trong Claude Code/Cursor/Antigravity (tiêu chí #2) | M1 |
| Diễn tập đồng bộ | Nhánh thử từ `v6.3.1` + lớp Việt → `sync_upstream.ps1 v6.3.2` (tiêu chí #6) | M1 |
| Template | `--template-mode` pass; PPTX xem thử được chủ repo duyệt; tạo 1 bài thật với mỗi gói | M2 |
| Review code | `code-reviewer` + `qa` subagent cho `doctor.py`, `setup.ps1`, `sync_upstream.ps1` trước khi phát hành | M1 |

---

## 9. Mốc phát hành

### M1 — `v6.3.2-vi.1`: nền tảng, bộ cài, tài liệu

Gồm §5.1–§5.5, §5.7, §5.8.
**DoD:** tiêu chí thành công #1, #2, #3, #5, #6 đạt; ở M1 `git diff v6.3.2 main -- skills/` phải **rỗng** (chưa có template); unit test pass; review code không còn lỗi mức cao; `CHANGELOG-VI.md` có mục vi.1; Release trên GitHub.

### M2 — `v6.3.2-vi.2`: template tiếng Việt

Gồm §5.6 và mở rộng §5.4 mục 8, §5.8.
**DoD:** 5 gói đăng ký trong index, `--template-mode` pass, chủ repo duyệt brief và PPTX xem thử từng gói; tiêu chí #4 đạt đầy đủ; `docs/vi/cau-lenh-mau.md` có câu lệnh cho từng template; Release vi.2.

Nếu upstream ra tag mới trước khi M2 xong: đồng bộ bằng `sync_upstream.ps1` rồi đổi số phiên bản M2 theo tag mới.

**Lập plan:** mỗi mốc một implementation plan riêng. Plan M1 lập ngay sau khi spec được duyệt; plan M2 lập sau khi M1 phát hành (khi đã có logo và kinh nghiệm chạy thực tế).

---

## 10. Ngoài phạm vi

- Dịch `SKILL.md`, `references/`, `workflows/` của upstream (giữ tiếng Anh — AI đọc tốt, tránh xung đột).
- Plugin marketplace riêng cho bản Việt (dùng cơ chế clone + mở thư mục).
- Bộ cài đồ hoạ (.exe/.msi), video hướng dẫn, trang gallery ví dụ tiếng Việt.
- Thêm layout trung tính mới; template ngoài 5 gói ở §5.6.
- Hỗ trợ macOS/Linux ngoài `setup.sh` tối giản.

---

## 11. Rủi ro & giảm thiểu

| Rủi ro | Giảm thiểu |
|---|---|
| Upstream đổi guard, cấu trúc template hoặc key định dạng | `sync_upstream.ps1` chạy guard + `test_vi_layer.py`; test đỏ → cập nhật lớp Việt trước khi phát hành |
| Alias `python3`/`python` của Microsoft Store | `setup.ps1` phát hiện và hướng dẫn tắt; `AGENTS.vi.md` quy định fallback `python` |
| Máy không có `winget` (Windows cũ/LTSC) | Hướng dẫn cài tay trong `cai-dat-windows.md` |
| Không có máy Windows sạch để thử thật | Mô phỏng PATH; `xu-ly-loi.md` bao phủ lỗi cài đặt; thu thập phản hồi khách sau vi.1 |
| Cơ chế nạp rule của Cursor/Antigravity không hoạt động | Kiểm chứng ở M1; dự phòng thêm 1 dòng cuối `AGENTS.md` |
| Khách sửa file trong repo → `update_repo.py` từ chối pull | `xu-ly-loi.md` hướng dẫn đưa dự án vào `projects/` (đã gitignore) và khôi phục file |
| Người dùng bản cũ mất `examples/` sau khi pull | `CHANGELOG-VI.md` + README nêu rõ, trỏ tới repo ví dụ của upstream và tag `v2-vi-legacy` |
| Dùng huy hiệu Đoàn sai quy định | Ghi chú trong brand spec; logo do chủ repo cung cấp và chịu trách nhiệm |
| PyMuPDF (AGPL) | Không đóng gói kèm; nêu trong `NOTICE` |
