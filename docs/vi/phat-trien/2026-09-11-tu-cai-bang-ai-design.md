# Thiết kế: AI tự cài đặt bộ công cụ (`v6.3.2-vi.3`)

- **Ngày:** 2026-09-11
- **Người phụ trách:** Lương Hải Anh — 2Anh AI Education
- **Trạng thái:** Đã duyệt thiết kế; chờ chủ repo đọc spec
- **Nền:** nhánh `feat/vi-tro-ly` @ `73e9cec7` (M2 — `v6.3.2-vi.2`, đang chờ nghiệm thu). Chỉ triển khai sau khi M2 đã lên `main`; khi đó đặt lại nhánh `feat/vi-tu-cai` lên `main` trước khi code.
- **Liên quan:** bộ cài M1 trong [`2026-09-10-dong-goi-ban-viet-design.md`](2026-09-10-dong-goi-ban-viet-design.md); rủi ro M1 còn treo "nạp lại PATH sau winget chưa chạy thật".

---

## 1. Mục tiêu

Thầy cô không rành kỹ thuật chỉ cần **đưa link repo (hoặc mở thư mục đã tải) cho AI trong Antigravity**. AI tự đọc hướng dẫn, kiểm tra máy, cài những gì còn thiếu rồi báo "sẵn sàng tạo slide" — thầy cô không phải gõ lệnh hay bấm file `.bat`.

### Tiêu chí thành công

1. Trên máy Windows chưa có Python: mở Antigravity với một thư mục trống, dán câu lệnh mẫu có link repo → AI tải repo, cài, rồi báo sẵn sàng với mục "Xuất thử PPTX" đạt. Thầy cô chỉ phải bấm cho phép chạy lệnh khi Antigravity hỏi. Kiểm bằng nghiệm thu máy thật (§8.3).
2. Thư mục đã tải nhưng chưa cài: yêu cầu tạo slide đầu tiên (hoặc "cài đặt giúp em") → AI kiểm tra, cài, báo lại, rồi mới làm tiếp yêu cầu. Kiểm bằng headless (§8.2) và máy thật.
3. Python được cài cho riêng tài khoản đang dùng, không cần quyền quản trị, và việc cài hoàn tất mà không phải mở lại terminal hay Antigravity.
4. Chạy cài lần thứ hai không cài lại gì. Kiểm bằng test `-PlanOnly` (§8.1) và máy thật.
5. Lúc cài không cài Git, Pandoc, FFmpeg; chỉ cài khi thầy cô dùng tới tính năng cần chúng.
6. `CAI-DAT.bat`, `KIEM-TRA.bat`, `CAP-NHAT.bat` vẫn dùng được như cũ. `git diff v6.3.2 -- skills/` rỗng; guard bản quyền exit 0; toàn bộ test `tools/vi/tests` đạt.
7. Khi máy chặn (mạng, chính sách cài đặt, PowerShell), AI dừng và báo cách xử lý kèm một đoạn để thầy cô gửi bộ phận IT; không tìm cách vượt qua.

---

## 2. Bối cảnh & phát hiện

| # | Phát hiện | Nguồn | Hệ quả thiết kế |
|---|---|---|---|
| 1 | Bộ cài hỏi Y/N bằng `Read-Host`; AI chạy lệnh trong terminal không trả lời được. `-NonInteractive` trả "không" cho mọi câu, kể cả cài Python. | `tools/vi/pptmaster.ps1:33-37,93` | Cần chế độ `-Auto`: tự quyết theo quy tắc cố định, không hỏi. |
| 2 | Cài Python bằng winget xong, bộ cài bảo đóng cửa sổ và chạy lại để PATH có hiệu lực. | `tools/vi/pptmaster.ps1:94-96` | `-Auto` tự tìm đường dẫn `python.exe` vừa cài, không dựa vào PATH. |
| 3 | `doctor.py` chỉ in chữ, mã thoát 0/1. | `tools/vi/doctor.py:291-344` | Thêm `--json` để AI đọc chính xác. |
| 4 | Quy tắc Antigravity (`trigger: always_on`) chỉ đọc `AGENTS.md` và `AGENTS.vi.md` khi thư mục repo là thư mục đang mở. | `.agents/rules/ppt-master-vi.md` | Khi tải bằng link, hướng dẫn cài bắt AI đọc hai file này ngay trong cuộc trò chuyện. |
| 5 | Script upstream tìm FFmpeg và Pandoc qua PATH. | `narration_sync.py:1679`, `video_subtitles.py:256`, `video_sound_mix.py:150`, `source_to_md/doc_to_md.py:1320` | Cài khi cần xong phải thêm thư mục chương trình vào PATH trong cùng lệnh. |
| 6 | Không script upstream nào gọi Python theo tên (`"python"`, `"python3"`). | tìm kiếm `skills/ppt-master/scripts/*.py` | Chạy mọi thứ bằng Python của `venv` là an toàn. |
| 7 | `.gitignore` đã có `venv/`. | `.gitignore:129` | Thư mục `venv\` ở gốc repo không lên GitHub, không phải sửa `.gitignore`. |
| 8 | `requirements.txt` ở gốc chỉ trỏ `-r skills/ppt-master/requirements.txt`; `doctor.py` kiểm theo file trong skill. | `requirements.txt`; `tools/vi/doctor.py:310` | Cài bằng file ở gốc, kiểm bằng doctor như hiện nay. |
| 9 | Test M2 khoá mục `## 10. Hỗ trợ thầy cô trước khi tạo PPTX` là mục cuối của `AGENTS.vi.md`. | `tools/vi/tests/test_vi_layer.py:394-400` | Quy tắc mới đặt vào mục 4 và mục 9, không thêm mục mới. |
| 10 | Upstream yêu cầu tách stderr khỏi JSON, không đặt `2>&1` trước bộ đọc JSON. | `AGENTS.md` ("Repository execution anchor") | Kết quả JSON đi ra stdout; mọi tiến trình khác đi ra stderr. |

---

## 3. Quyết định đã chốt

| Chủ đề | Quyết định |
|---|---|
| Cách vào | Cả hai: dán link vào Antigravity đang mở thư mục trống; hoặc thư mục đã tải, nhắn "cài đặt giúp em" hay yêu cầu tạo slide đầu tiên |
| Phạm vi cài | Python, thư viện, `.env`: tự cài, chỉ báo một dòng, không hỏi. Git, Pandoc, FFmpeg: chỉ cài khi dùng tới tính năng cần |
| Hướng làm | A: AI điều phối theo hướng dẫn; việc cài nằm trong `pptmaster.ps1 -Auto` và `doctor.py --json` |
| Môi trường Python | `venv\` riêng ở gốc repo, cài thư viện vào đó |
| Cài Python | Cho riêng tài khoản: `winget --scope user`, không được thì bộ cài python.org chạy ngầm, phiên bản và SHA256 ghi cố định |
| Kiểm chứng | Test tự động trên máy này + chủ repo thử trên một máy Windows chưa có Python |
| Nền tảng | Windows. macOS/Linux giữ `tools/vi/setup.sh` như cũ |
| Phiên bản | `v6.3.2-vi.3`, nhánh `feat/vi-tu-cai` |

---

## 4. Kiến trúc

**Nguyên tắc:** phần cài nằm trong script (chạy giống nhau trên mọi máy, test được); AI chỉ tải repo, gọi đúng một lệnh, đọc JSON và báo thầy cô. Không file nào của upstream bị sửa.

### 4.1 File thêm/sửa

```text
docs/vi/cai-dat-bang-ai.md        # (mới) hướng dẫn cài dành cho AI
tools/vi/pptmaster.ps1            # (sửa) -Auto, -PlanOnly, -Action tool -Name; check/update ưu tiên venv
tools/vi/doctor.py                # (sửa) --json
AGENTS.vi.md                      # (sửa) mục 4 (Python của venv), mục 9 (tự kiểm tra, tự cài)
README.md                         # (sửa) khối "Dành cho AI agent", cách dán link
docs/vi/bat-dau-nhanh.md          # (sửa) câu lệnh mẫu để AI tự cài, lời dặn cho phép chạy lệnh
docs/vi/cai-dat-windows.md        # (sửa) "Cách nhanh: nhờ AI cài"
docs/vi/xu-ly-loi.md              # (sửa) mục "Máy trường chặn cài đặt" kèm đoạn gửi IT
tools/vi/tests/test_doctor.py     # (sửa) test --json
tools/vi/tests/test_installer.py  # (mới) test -PlanOnly và -Action tool bằng PATH giả
tools/vi/tests/test_vi_layer.py   # (sửa) test nội dung hướng dẫn và quy tắc mới
CHANGELOG-VI.md                   # (sửa khi phát hành)
```

---

## 5. Luồng

### 5.1 Dán link

1. Thầy cô mở Antigravity với một thư mục trống, dán: `Cài PPT Master từ https://github.com/luonghaianh1208/PPTmaster vào thư mục này rồi báo khi sẵn sàng tạo slide`.
2. AI đọc README trên GitHub; khối "Dành cho AI agent" trỏ tới `docs/vi/cai-dat-bang-ai.md`.
3. AI tải repo theo §6.3 rồi đọc `AGENTS.md`, `AGENTS.vi.md`.
4. AI báo một dòng trước khi cài, chạy `-Auto`, đọc JSON, báo sẵn sàng hoặc báo lỗi.

### 5.2 Thư mục có sẵn

1. Thầy cô mở thư mục repo và gửi yêu cầu tạo slide (hoặc "cài đặt giúp em").
2. Theo `AGENTS.vi.md` mục 9, trước lệnh Python đầu tiên của repo trong cuộc trò chuyện (ví dụ `project_manager.py init`, `source_to_md.py`), AI chạy `doctor.py --no-smoke --json` bằng Python của `venv` (không có `venv` thì bằng `python`). Với yêu cầu thuộc mục 10 (lượt hỏi thầy cô), AI chạy kiểm tra này trước khi gửi tin nhắn hỏi; chưa sẵn sàng thì thêm một dòng cuối tin nhắn báo sẽ cài (khoảng 5–10 phút) ngay sau khi thầy cô trả lời.
3. `ready` là `true` → làm tiếp yêu cầu. Ngược lại hoặc không chạy được Python → làm theo `docs/vi/cai-dat-bang-ai.md` từ bước cài, rồi làm tiếp yêu cầu ban đầu. Với "cài đặt giúp em" trên máy đã cài sẵn, hướng dẫn cũng chạy `doctor.py --no-smoke --json` trước và không cài lại khi `ready` là `true`.
4. Lưới an toàn: lệnh Python nào của repo báo không tìm thấy Python hoặc `ModuleNotFoundError` → AI dừng, làm theo `docs/vi/cai-dat-bang-ai.md`, không tự cài.

### 5.3 Công cụ tuỳ chọn khi cần

- Thuyết minh, video → cần FFmpeg; tài liệu `.doc`, `.odt`, `.rtf` và các định dạng cũ khác → cần Pandoc.
- Thiếu thì AI chạy `pptmaster.ps1 -Action tool -Name ffmpeg` (hoặc `pandoc`), đọc `dir` trong JSON, rồi chạy lệnh cần công cụ đó trong cùng một lệnh PowerShell: `$env:Path = "<dir>;$env:Path"; <lệnh>`.
- Mở lại Antigravity thì PATH mới có hiệu lực, không cần thêm tiền tố nữa.

### 5.4 Tin nhắn cho thầy cô

- **Trước khi cài:** "Em sẽ cài Python và thư viện cho PPT Master, mất khoảng 5–10 phút. Nếu Antigravity hỏi cho phép chạy lệnh, thầy cô bấm đồng ý giúp em."
- **Sẵn sàng:** đã cài những gì; kết quả xuất thử PPTX; một câu lệnh mẫu (ví dụ `Tạo bài giảng Toán 6 bài Phân số`); với cách dán link thêm "Lần sau thầy cô mở thư mục <đường dẫn> trong Antigravity là dùng được ngay."
- **Bị chặn:** lý do ngắn gọn, cách tự xử lý (nếu có). Có cảnh báo đường dẫn dài hoặc OneDrive thì hướng dẫn chuyển thư mục trước; chỉ kèm đoạn gửi IT theo §6.7 khi lỗi tải bộ công cụ, tải hoặc chạy bộ cài Python, lỗi cài thư viện không kèm cảnh báo thư mục, hoặc PowerShell bị chặn.

---

## 6. Chi tiết thành phần

### 6.1 `tools/vi/pptmaster.ps1`

**Tham số:** `-Action setup|check|update|tool`, `-Auto`, `-PlanOnly`, `-Name ffmpeg|pandoc`, giữ `-NonInteractive`.

**`setup -Auto` — các bước (mỗi bước bỏ qua nếu đã xong):**

1. **Thư mục:** thư mục repo nằm dưới `$env:OneDrive`, `$env:OneDriveCommercial` hoặc `$env:OneDriveConsumer`, hoặc đường dẫn dài hơn 80 ký tự, hoặc có đoạn lồng `PPTmaster-main\PPTmaster-main` → ghi cảnh báo vào `warnings` (không dừng), vì quá trình xuất PPTX tạo thêm nhiều thư mục con (xem mục "Đường dẫn quá dài" trong `docs/vi/xu-ly-loi.md`).
2. **Tìm Python ≥ 3.10**, theo thứ tự:
   1. `venv\Scripts\python.exe`;
   2. `python` trong PATH, bỏ qua lối tắt Microsoft Store;
   3. `py -3`;
   4. `%LOCALAPPDATA%\Programs\Python\Python312\python.exe` (hoặc `Python312-arm64` trên máy ARM64).
3. **Cài Python 3.12** nếu bước 2 không thấy:
   1. có winget → `winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements`, rồi tìm lại theo bước 2.4;
   2. vẫn chưa có → tải bộ cài python.org (phiên bản, đường dẫn và SHA256 là hằng số trong script; chọn bản `amd64` hoặc `arm64` theo `$env:PROCESSOR_ARCHITECTURE`) vào `$env:TEMP`, so `Get-FileHash`; khớp mới chạy `/quiet InstallAllUsers=0 PrependPath=1 Include_launcher=1 InstallLauncherAllUsers=0 Include_test=0`, rồi tìm lại;
   3. vẫn không có → dừng với lỗi bước `python`.
4. **`venv`:** chưa có `venv\Scripts\python.exe` → `<python> -m venv venv`. Sau đó kiểm Python của `venv` chạy được (`-c "import sys"`, bọc try/catch vì file hỏng có thể ném lỗi); không chạy được → lỗi bước `venv` với `fix` "Xoá thư mục venv trong bộ công cụ rồi chạy lại lệnh cài."
5. **Thư viện:** `doctor.py --no-smoke --json` bằng Python của `venv` báo mục "Thư viện Python" đạt → bỏ qua; ngược lại kiểm `-m pip --version` (lỗi → lỗi bước `venv` như trên), rồi `pip install --upgrade pip` và `pip install -r requirements.txt`, lỗi thì thử lại một lần. `fix` của lỗi bước `packages` trỏ mục "Máy trường chặn cài đặt" (thêm mục "Đường dẫn quá dài" khi có cảnh báo thư mục).
6. **`.env`:** chưa có thì chép từ `.env.example`.
7. **Kiểm tra cuối:** `doctor.py --json` (có xuất thử) bằng Python của `venv`.

**Đầu ra của `-Auto`:**

- stdout chỉ có **một đối tượng JSON**; tiến trình (kể cả đầu ra của winget, pip) đi ra stderr.
- Các trường: `ready` (bool), `python` (đường dẫn Python của `venv`), `installed` (danh sách trong `python`, `venv`, `packages`, `env`), `warnings` (danh sách chuỗi), `checks` (lấy từ `doctor.py --json`), `error` (`null` hoặc `{ "step", "message", "fix" }`).
- Mã thoát: 0 khi `ready` là `true`, ngược lại 1.

**`-PlanOnly`** (dùng cùng `setup -Auto` hoặc `tool`): không tải, không cài, không tạo file. In JSON gồm `python_found` (đường dẫn hoặc `null`), `steps` (danh sách `{ "step", "action", "method" }`, ví dụ `{ "step": "python", "action": "install", "method": "winget" }`), `warnings`. Mã thoát 0.

**`-Action tool -Name ffmpeg|pandoc`:**

- Đã có trong PATH hoặc ở thư mục cài mặc định cho tài khoản → không cài.
- Chưa có và có winget → `winget install -e --id Gyan.FFmpeg` (hoặc `JohnMacFarlane.Pandoc`) `--scope user --silent --accept-package-agreements --accept-source-agreements`, rồi tìm thư mục chứa `ffmpeg.exe` / `pandoc.exe`.
- Đầu ra JSON: `tool`, `found` (bool), `installed` (bool), `dir` (thư mục hoặc `null`), `error`. Mã thoát 0 khi `found` là `true`.

**`check` và `update`:** có `venv\Scripts\python.exe` thì dùng nó; không có thì giữ cách tìm Python hiện tại. **`setup` không có `-Auto`** (từ `CAI-DAT.bat`) giữ nguyên hành vi M1, trừ một điểm: có `venv\Scripts\python.exe` thì dùng nó ở bước kiểm tra Python (không đề nghị cài Python), để `CAI-DAT.bat` cài thư viện vào đúng Python mà `KIEM-TRA.bat` kiểm tra.

### 6.2 `tools/vi/doctor.py`

- Thêm `--json`: in một đối tượng JSON `{ "ready": bool, "python": sys.executable, "checks": [ { "name", "level", "ok", "detail", "fix" } ] }`, `ensure_ascii=False`, UTF-8; không in phần chữ.
- Dùng được cùng `--no-smoke`. Mã thoát giữ nguyên (0/1). Không có `--json` thì in chữ như cũ.

### 6.3 `docs/vi/cai-dat-bang-ai.md`

File tiếng Việt dành cho AI, gồm các mục:

1. **Khi nào dùng:** người dùng nhờ cài bộ công cụ; `AGENTS.vi.md` mục 9 báo môi trường chưa sẵn sàng; hoặc một lệnh Python của repo báo không tìm thấy Python hay `ModuleNotFoundError`.
2. **Tải bộ công cụ** (bỏ qua nếu thư mục đang mở đã có `tools/vi/pptmaster.ps1`):
   - thư mục đang mở trống → tải vào đó; có file khác → tải vào thư mục con `PPTmaster`;
   - có Git → `git clone https://github.com/luonghaianh1208/PPTmaster.git <thư_mục>`;
   - không có Git → tải `https://github.com/luonghaianh1208/PPTmaster/archive/refs/heads/main.zip` bằng `Invoke-WebRequest` vào `$env:TEMP`, giải nén bằng `Expand-Archive`, chuyển nội dung của `PPTmaster-main` vào thư mục đích;
   - người dùng nêu rõ một nhánh (ví dụ khi chủ repo nghiệm thu trước khi phát hành) → dùng nhánh đó: `git clone -b <nhánh>`, hoặc ZIP `archive/refs/heads/<nhánh>.zip`;
   - thư mục đang mở có sẵn `PPTmaster\tools\vi\pptmaster.ps1` → dùng thư mục con đó, bỏ bước tải; kiểm thư mục "trống" bằng `Get-ChildItem -Force` (tính cả mục ẩn);
   - đoạn tải ZIP bật TLS 1.2 (`-bor`), đặt `$ProgressPreference = 'SilentlyContinue'` và xoá file tạm sau khi chuyển; lệnh tải hoặc giải nén lỗi → dừng, không tải chồng;
   - tải xong: đọc `AGENTS.md` và `AGENTS.vi.md` của repo vừa tải.
3. **Cài đặt:** thư mục đã có bộ công cụ từ trước → chạy `doctor.py --no-smoke --json` trước (bằng Python của `venv`, không có thì `python`), `ready` là `true` thì báo sẵn sàng, không cài lại. Ngược lại gửi tin nhắn "trước khi cài" (§5.4), rồi chạy từ thư mục repo: `powershell -NoProfile -ExecutionPolicy Bypass -File tools\vi\pptmaster.ps1 -Action setup -Auto`. Lệnh có thể chạy 5–10 phút: công cụ cho đặt thời gian chờ thì đặt ít nhất 15 phút; lệnh chuyển sang chạy nền thì theo dõi tới khi có dòng JSON; không chạy lệnh cài thứ hai khi lệnh trước chưa xong. JSON ở stdout là dòng cuối, bắt đầu bằng `{"ready"`; không gộp stderr vào.
4. **Đọc kết quả:** `ready` → báo sẵn sàng (§5.4); `error` hoặc mục bắt buộc chưa đạt → chỉ làm những việc trong `fix` mà mục 6 không cấm (ví dụ xoá `venv`), chạy lại đúng lệnh cài tối đa một lần; vẫn lỗi → báo bị chặn (§5.4, §6.7). `error.step = packages` → không chạy lệnh `pip` trong `xu-ly-loi.md`, chỉ chạy lại lệnh cài một lần, rồi giải thích nguyên nhân thường gặp (mạng/proxy, diệt virus, đường dẫn dài) bằng lời, không tự tắt gì. Lệnh cài không chạy được vì PowerShell bị chặn chạy script → xử lý theo dòng tương ứng ở §7.
5. **Công cụ tuỳ chọn:** như §5.3.
6. **Không được làm:** chạy script tải thẳng từ mạng (`iex`, `Invoke-Expression`); xin quyền quản trị; tắt phần mềm diệt virus hay tường lửa; đổi ExecutionPolicy vĩnh viễn; cài phần mềm ngoài Python 3.12, thư viện trong `requirements.txt`, Git (chỉ khi người dùng yêu cầu), Pandoc, FFmpeg; tự viết cách cài khác thay cho lệnh ở mục 3.

### 6.4 `AGENTS.vi.md`

- **Mục 4 — thêm:** có `venv\Scripts\python.exe` ở gốc repo thì dùng nó cho mọi lệnh `python3 …` hoặc `python …` của repo.
- **Mục 9 — đổi tiêu đề** thành `## 9. Môi trường: tự kiểm tra, tự cài và xử lý lỗi`, nội dung:
  - mỗi cuộc trò chuyện, trước lệnh Python đầu tiên của repo (ví dụ `project_manager.py init`, `source_to_md.py`): chạy `doctor.py --no-smoke --json`; chưa sẵn sàng thì làm theo `docs/vi/cai-dat-bang-ai.md` rồi mới chạy lệnh đó;
  - yêu cầu thuộc mục 10: chạy kiểm tra trên (chỉ đọc) trước khi gửi tin nhắn hỏi; chưa sẵn sàng thì thêm một dòng cuối tin nhắn (ngay trước dòng chốt cách xác nhận, nếu có) báo sẽ cài (khoảng 5–10 phút) sau khi thầy cô trả lời, rồi cài ngay khi thầy cô trả lời;
  - lưới an toàn: lệnh Python nào của repo báo không tìm thấy Python hoặc `ModuleNotFoundError` → dừng, làm theo `docs/vi/cai-dat-bang-ai.md`, không tự cài;
  - người dùng nhờ "cài đặt", "kiểm tra máy" → làm theo `docs/vi/cai-dat-bang-ai.md`;
  - cần FFmpeg hoặc Pandoc mà thiếu → như §5.3;
  - giữ câu hiện có về `KIEM-TRA.bat` và `docs/vi/xu-ly-loi.md`.
- Mục 10 vẫn là mục cuối.

### 6.5 `README.md`

- Ngay dưới đoạn giới thiệu: khối trích dẫn "**Dành cho AI agent:** người dùng nhờ cài bộ công cụ này thì làm theo `docs/vi/cai-dat-bang-ai.md`, không tự nghĩ cách cài khác."
- Mục "Bắt đầu trong 3 bước": thêm dòng "Dùng Antigravity: mở một thư mục trống, dán câu lệnh mẫu (xem Bắt đầu nhanh) là AI tự tải và cài."

### 6.6 Tài liệu cho thầy cô

- `docs/vi/bat-dau-nhanh.md`: thêm mục `## Để AI tự cài` (câu lệnh mẫu §5.1, thời gian 5–10 phút, bấm cho phép khi Antigravity hỏi chạy lệnh, bấm từ chối nếu AI xin quyền quản trị hoặc tắt diệt virus, lần sau mở thư mục AI báo trong tin nhắn sẵn sàng); ghi rõ cách này không cần `CAI-DAT.bat`.
- `docs/vi/xu-ly-loi.md` mục "KIEM-TRA báo thiếu thư viện": có thư mục `venv` thì nhờ AI chạy lại lệnh cài, hoặc chạy `venv\Scripts\python.exe -m pip install -r requirements.txt`.
- `docs/vi/cai-dat-windows.md`: thêm đầu file "Cách nhanh: nhờ AI cài" trỏ tới mục trên; các bước bấm đúp giữ nguyên.

### 6.7 `docs/vi/xu-ly-loi.md` — mục "Máy trường chặn cài đặt"

- Dấu hiệu: AI báo không tải được, không chạy được bộ cài Python, hoặc PowerShell bị chặn.
- Đoạn gửi IT (AI chép nguyên đoạn này khi báo bị chặn theo §5.4): cần cho tài khoản này truy cập `python.org`, `pypi.org`, `files.pythonhosted.org`, `github.com`, `codeload.github.com`, thêm `cdn.winget.microsoft.com`, `objects.githubusercontent.com` (khi cần FFmpeg/Pandoc); cho phép cài Python 3.12 cho người dùng hiện tại; cho phép chạy PowerShell với `-ExecutionPolicy Bypass` cho một tiến trình.

---

## 7. Trường hợp biên

| Tình huống | Cách xử lý |
|---|---|
| Thư mục đang mở đã là repo này | Bỏ bước tải, cài như §6.1 |
| Thư mục đang mở có file khác | Tải vào thư mục con `PPTmaster`; tin nhắn sẵn sàng nêu thư mục cần mở lần sau |
| Máy có Python 3.8 trong PATH | Bỏ qua (dưới 3.10), cài Python 3.12 cho tài khoản, `venv` dùng bản 3.12 |
| Chỉ có lối tắt Microsoft Store | Coi như chưa có Python |
| Không có winget | Dùng bộ cài python.org; `-Action tool` báo lỗi kèm đường dẫn tải thủ công |
| SHA256 bộ cài không khớp | Không chạy bộ cài; lỗi bước `python`, AI báo bị chặn |
| Không có mạng hoặc proxy chặn | Lỗi ở bước tải hoặc `packages`; AI báo kèm đoạn gửi IT |
| `-ExecutionPolicy Bypass` bị chính sách nhóm chặn | Máy đã có Python ≥ 3.10: AI chạy các lệnh tay trong mục "PowerShell bị chặn" của `xu-ly-loi.md` (cài vào Python chung, không có `venv`), rồi kiểm bằng `doctor.py --json`. Chưa có Python: AI báo bị chặn kèm đoạn gửi IT |
| Thư mục trong OneDrive | Cảnh báo trong tin nhắn sẵn sàng, gợi ý chuyển sang `D:\PPTmaster`; không dừng |
| Người dùng đã cài bằng `CAI-DAT.bat` (không có `venv`) | `doctor --no-smoke --json` bằng `python` đạt (theo mục 9 hoặc bước đầu mục "Cài đặt" của hướng dẫn, kể cả khi nhắn "cài đặt giúp em") → không cài lại |
| `venv` hỏng hoặc tạo dở (lệnh cài bị dừng giữa chừng, Python gốc đã bị gỡ, `pyvenv.cfg` hỏng) | `-Auto` kiểm Python của `venv` chạy được và có pip trước khi cài thư viện; lỗi bước `venv`, `fix` "Xoá thư mục venv trong bộ công cụ rồi chạy lại lệnh cài."; AI xoá `venv` rồi chạy lại một lần, không gửi đoạn cho IT |
| Công cụ chạy lệnh của AI có giới hạn thời gian hoặc tự chuyển lệnh dài sang chạy nền | Hướng dẫn: đặt thời gian chờ ít nhất 15 phút, theo dõi tới dòng JSON, không chạy lệnh cài thứ hai khi lệnh trước chưa xong |
| Yêu cầu thuộc mục 10 trên máy chưa cài | Kiểm tra trước khi gửi tin nhắn hỏi, báo sẽ cài sau khi thầy cô trả lời; cài xong mới khởi tạo dự án |
| Upstream báo "requires ffmpeg on PATH" dù đã cài | AI chạy lại lệnh với tiền tố PATH theo §5.3 |

---

## 8. Kiểm thử

Mọi test chạy nền trong `tools/vi/tests` bằng `unittest`; test gọi `powershell -NoProfile` như tiến trình con, không mở cửa sổ, không cài và không tải gì.

### 8.1 Test tự động

- **`test_doctor.py`:** `--json` in JSON hợp lệ có đủ trường; `ready` khớp mã thoát; chữ tiếng Việt giữ nguyên dấu; `--json --no-smoke` không có mục xuất thử.
- **`test_installer.py`:** chép `pptmaster.ps1` sang thư mục tạm có cấu trúc repo tối thiểu, dựng PATH và `LOCALAPPDATA` giả (một `python.cmd` giả trả phiên bản), chạy `setup -Auto -PlanOnly` và kiểm `steps`:
  - không có Python, có winget → `python: install/winget`;
  - không có Python, không có winget → `python: install/python-org`;
  - chỉ có lối tắt Store → như không có Python;
  - có Python 3.12 → không cài Python, `venv: create`;
  - đã có `venv` → không bước `python`, không `venv: create`;
  - thư mục dưới `$env:OneDrive` giả → có cảnh báo;
  - `tool -Name ffmpeg -PlanOnly` khi thiếu → `install/winget`;
  - không `-PlanOnly`, `venv\Scripts\python.exe` rỗng → `error.step = venv`, mã thoát 1;
  - không `-PlanOnly`, `venv` thật tạo bằng `--without-pip`, `doctor.py` giả báo "Thư viện Python" đạt, có `.env` → `ready` là `true`, `installed` rỗng, không gọi pip.
- **`test_vi_layer.py`:** `docs/vi/cai-dat-bang-ai.md` tồn tại, có lệnh `-Action setup -Auto`, có mục "Không được làm" với `iex` và quyền quản trị; `AGENTS.vi.md` mục 4 có `venv\Scripts\python.exe`, mục 9 có tiêu đề mới và trỏ tới `cai-dat-bang-ai.md`; mục 10 vẫn là mục cuối; README có khối "Dành cho AI agent".

### 8.2 Kiểm thử headless

Chạy trên bản sao repo trong thư mục tạm, không có `venv`: `claude -p "<câu lệnh>" --disallowedTools Bash Write Edit NotebookEdit PowerShell Agent Monitor Skill`. Danh sách cờ chỉ là lớp phòng ngừa, không phải sandbox.

| Ca | Câu lệnh | Đạt khi |
|---|---|---|
| 1 | `cài đặt giúp em` | Đầu ra nêu lệnh `pptmaster.ps1 -Action setup -Auto` hoặc `cai-dat-bang-ai.md`; không có `pip install` tự viết, không có `iex` |
| 2 | `Tạo bài giảng Toán 6 bài Phân số` | AI kiểm tra hoặc cài môi trường trước khi khởi tạo dự án |

### 8.3 Nghiệm thu máy thật (chủ repo)

1. Máy Windows chưa có Python (hoặc tài khoản Windows mới trên máy chưa cài Python cho mọi người dùng), ưu tiên máy **không có Git** để AI đi qua đường tải ZIP; cài Antigravity.
2. Mở thư mục trống, dán câu lệnh mẫu §5.1 nhưng dùng link `https://github.com/luonghaianh1208/PPTmaster/tree/feat/vi-tu-cai` → AI báo sẵn sàng, "Xuất thử PPTX" đạt. Nghiệm thu diễn ra trước khi phát hành, nên nhánh `feat/vi-tu-cai` được push lên `origin` (không force); link trỏ thẳng vào nhánh để AI đọc README và hướng dẫn của nhánh, không phải của `main`. Ghi lại cách Antigravity xử lý lệnh cài dài: chờ tới khi xong, chuyển sang chạy nền, hay chạy trùng lệnh cài.
3. Nhắn tạo một bài giảng → nhận PPTX, mở bằng PowerPoint.
4. Đóng Antigravity, mở lại thư mục, tạo bài thứ hai → AI không cài lại.
5. Tải ZIP của nhánh vào một thư mục khác, không cài; mở thư mục đó trong Antigravity, gửi `Tạo bài giảng Toán 6 bài Phân số` → AI chạy kiểm tra hoặc cài trước `project_manager.py init` (với lượt hỏi thầy cô: kiểm tra trước khi gửi câu hỏi, cài ngay sau khi thầy cô trả lời). Bước này kiểm tiêu chí #2 trên máy thật.
6. Sau khi phát hành lên `main`: thử nhanh câu lệnh §5.1 với link `https://github.com/luonghaianh1208/PPTmaster`, ít nhất tới lúc AI đọc hướng dẫn và chạy `-Auto`.

---

## 9. Phát hành

- Sau khi M2 lên `main`: đặt lại `feat/vi-tu-cai` lên `main`.
- Sau nghiệm thu §8.3: cập nhật `CHANGELOG-VI.md` và dòng phiên bản README thành `6.3.2-vi.3`; gộp vào `main`; tag `v6.3.2-vi.3`; `gh release create … --repo luonghaianh1208/PPTmaster`.

## 10. Ngoài phạm vi

- Tự cài trên macOS/Linux.
- Nghiệm thu riêng cho Claude Code và Cursor (dùng chung quy tắc nhưng không kiểm riêng).
- Cấu hình proxy, vượt AppLocker hay chính sách nhóm.
- Tự cập nhật bộ công cụ bằng AI; đổi câu ví dụ đầu tiên trong `bat-dau-nhanh.md`.
- Thay đổi cài đặt cho phép chạy lệnh của Antigravity.

## 11. Rủi ro

| Rủi ro | Giảm thiểu |
|---|---|
| Antigravity hỏi cho phép chạy lệnh nhiều lần, thầy cô bối rối | Tin nhắn "trước khi cài" dặn trước; tài liệu nhắc |
| Bộ cài python.org 3.12 cho Windows ngừng cập nhật, hằng số phiên bản/SHA256 cũ | Ghi rõ hằng số trong script và `docs/vi/phat-trien/bao-tri.md`; winget là cách chính |
| Gói winget của Pandoc/FFmpeg không hỗ trợ `--scope user` | Đã kiểm khi lên plan (2026-09-11): `winget show --scope user` có bộ cài cho `Python.Python.3.12` (3.12.10), `Gyan.FFmpeg` (portable zip), `JohnMacFarlane.Pandoc` (MSI). Nếu sau này gói đổi, `-Action tool` báo lỗi kèm đường dẫn tải thủ công |
| Quy tắc Antigravity không tự nạp sau khi tải vào thư mục đang mở | Hướng dẫn cài bắt AI đọc `AGENTS.md`, `AGENTS.vi.md` ngay |
| AI vẫn tự nghĩ cách cài khác | Câu "không tự nghĩ cách cài khác" ở README và hướng dẫn; ca headless 1 |
| AI quên tiền tố PATH khi dùng FFmpeg/Pandoc vừa cài | Lỗi của upstream nêu rõ "on PATH"; quy tắc mục 9 chỉ cách chạy lại |
| Thư mục OneDrive gây khoá file khi cài thư viện | Cảnh báo, gợi ý chuyển thư mục; theo dõi ở nghiệm thu |
| Thư mục bộ công cụ ở đường dẫn dài (khoảng 150 ký tự) làm pip lỗi vì giới hạn đường dẫn của Windows (đã gặp khi thử `-Auto` ở thư mục tạm dài 147 ký tự) | Cảnh báo khi dài hơn 80 ký tự, `fix` của lỗi `packages` trỏ thêm mục "Đường dẫn quá dài"; tài liệu gợi ý `D:\PPTmaster`; ghi trong mục "Rủi ro" của `CHANGELOG-VI.md` |
| Chọn Python gốc chưa chặt: nhận mọi Python từ 3.10 trở lên, kể cả bản mới hơn 3.12 chưa có gói build sẵn cho vài thư viện; không dò bản 3.12 cài cho mọi người dùng mà không có trong PATH hay `py` | Hiếm gặp; lỗi hiện ở bước `packages` hoặc bước `python`; ghi trong mục "Rủi ro" của `CHANGELOG-VI.md` |
