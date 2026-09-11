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
2. Theo `AGENTS.vi.md` mục 9, trước lần tạo slide đầu tiên của cuộc trò chuyện, AI chạy `doctor.py --no-smoke --json` bằng Python của `venv` (không có `venv` thì bằng `python`).
3. `ready` là `true` → làm tiếp yêu cầu. Ngược lại hoặc không chạy được Python → làm theo `docs/vi/cai-dat-bang-ai.md` từ bước cài, rồi làm tiếp yêu cầu ban đầu.

### 5.3 Công cụ tuỳ chọn khi cần

- Thuyết minh, video → cần FFmpeg; tài liệu `.doc`, `.odt`, `.rtf` và các định dạng cũ khác → cần Pandoc.
- Thiếu thì AI chạy `pptmaster.ps1 -Action tool -Name ffmpeg` (hoặc `pandoc`), đọc `dir` trong JSON, rồi chạy lệnh cần công cụ đó trong cùng một lệnh PowerShell: `$env:Path = "<dir>;$env:Path"; <lệnh>`.
- Mở lại Antigravity thì PATH mới có hiệu lực, không cần thêm tiền tố nữa.

### 5.4 Tin nhắn cho thầy cô

- **Trước khi cài:** "Em sẽ cài Python và thư viện cho PPT Master, mất khoảng 5–10 phút. Nếu Antigravity hỏi cho phép chạy lệnh, thầy cô bấm đồng ý giúp em."
- **Sẵn sàng:** đã cài những gì; kết quả xuất thử PPTX; một câu lệnh mẫu (ví dụ `Tạo bài giảng Toán 6 bài Phân số`); với cách dán link thêm "Lần sau thầy cô mở thư mục <đường dẫn> trong Antigravity là dùng được ngay."
- **Bị chặn:** lý do ngắn gọn, cách tự xử lý (nếu có), và đoạn gửi IT theo §6.7.

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
4. **`venv`:** chưa có `venv\Scripts\python.exe` → `<python> -m venv venv`.
5. **Thư viện:** `doctor.py --no-smoke --json` bằng Python của `venv` báo mục "Thư viện Python" đạt → bỏ qua; ngược lại `pip install --upgrade pip` rồi `pip install -r requirements.txt`, lỗi thì thử lại một lần.
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

**`check` và `update`:** có `venv\Scripts\python.exe` thì dùng nó; không có thì giữ cách tìm Python hiện tại. **`setup` không có `-Auto`** (từ `CAI-DAT.bat`) giữ nguyên hành vi M1.

### 6.2 `tools/vi/doctor.py`

- Thêm `--json`: in một đối tượng JSON `{ "ready": bool, "python": sys.executable, "checks": [ { "name", "level", "ok", "detail", "fix" } ] }`, `ensure_ascii=False`, UTF-8; không in phần chữ.
- Dùng được cùng `--no-smoke`. Mã thoát giữ nguyên (0/1). Không có `--json` thì in chữ như cũ.

### 6.3 `docs/vi/cai-dat-bang-ai.md`

File tiếng Việt dành cho AI, gồm các mục:

1. **Khi nào dùng:** người dùng nhờ cài bộ công cụ; hoặc `AGENTS.vi.md` mục 9 báo môi trường chưa sẵn sàng.
2. **Tải bộ công cụ** (bỏ qua nếu thư mục đang mở đã có `tools/vi/pptmaster.ps1`):
   - thư mục đang mở trống → tải vào đó; có file khác → tải vào thư mục con `PPTmaster`;
   - có Git → `git clone https://github.com/luonghaianh1208/PPTmaster.git <thư_mục>`;
   - không có Git → tải `https://github.com/luonghaianh1208/PPTmaster/archive/refs/heads/main.zip` bằng `Invoke-WebRequest` vào `$env:TEMP`, giải nén bằng `Expand-Archive`, chuyển nội dung của `PPTmaster-main` vào thư mục đích;
   - người dùng nêu rõ một nhánh (ví dụ khi chủ repo nghiệm thu trước khi phát hành) → dùng nhánh đó: `git clone -b <nhánh>`, hoặc ZIP `archive/refs/heads/<nhánh>.zip`;
   - tải xong: đọc `AGENTS.md` và `AGENTS.vi.md` của repo vừa tải.
3. **Cài đặt:** gửi tin nhắn "trước khi cài" (§5.4), rồi chạy từ thư mục repo: `powershell -NoProfile -ExecutionPolicy Bypass -File tools\vi\pptmaster.ps1 -Action setup -Auto`. Đọc JSON ở stdout; không gộp stderr vào.
4. **Đọc kết quả:** `ready` → báo sẵn sàng (§5.4); `error` hoặc mục bắt buộc chưa đạt → dùng `fix`, chạy lại đúng lệnh cài tối đa một lần; vẫn lỗi → báo bị chặn (§6.7). Lệnh cài không chạy được vì PowerShell bị chặn chạy script → xử lý theo dòng tương ứng ở §7.
5. **Công cụ tuỳ chọn:** như §5.3.
6. **Không được làm:** chạy script tải thẳng từ mạng (`iex`, `Invoke-Expression`); xin quyền quản trị; tắt phần mềm diệt virus hay tường lửa; đổi ExecutionPolicy vĩnh viễn; cài phần mềm ngoài Python 3.12, thư viện trong `requirements.txt`, Git (chỉ khi người dùng yêu cầu), Pandoc, FFmpeg; tự viết cách cài khác thay cho lệnh ở mục 3.

### 6.4 `AGENTS.vi.md`

- **Mục 4 — thêm:** có `venv\Scripts\python.exe` ở gốc repo thì dùng nó cho mọi lệnh `python3 …` hoặc `python …` của repo.
- **Mục 9 — đổi tiêu đề** thành `## 9. Môi trường: tự kiểm tra, tự cài và xử lý lỗi`, nội dung:
  - mỗi cuộc trò chuyện, trước lần tạo slide đầu tiên: chạy `doctor.py --no-smoke --json`; chưa sẵn sàng thì làm theo `docs/vi/cai-dat-bang-ai.md` rồi mới tạo slide;
  - người dùng nhờ "cài đặt", "kiểm tra máy" → làm theo `docs/vi/cai-dat-bang-ai.md`;
  - cần FFmpeg hoặc Pandoc mà thiếu → như §5.3;
  - giữ câu hiện có về `KIEM-TRA.bat` và `docs/vi/xu-ly-loi.md`.
- Mục 10 vẫn là mục cuối.

### 6.5 `README.md`

- Ngay dưới đoạn giới thiệu: khối trích dẫn "**Dành cho AI agent:** người dùng nhờ cài bộ công cụ này thì làm theo `docs/vi/cai-dat-bang-ai.md`, không tự nghĩ cách cài khác."
- Mục "Bắt đầu trong 3 bước": thêm dòng "Dùng Antigravity: mở một thư mục trống, dán câu lệnh mẫu (xem Bắt đầu nhanh) là AI tự tải và cài."

### 6.6 Tài liệu cho thầy cô

- `docs/vi/bat-dau-nhanh.md`: thêm mục `## Để AI tự cài` (câu lệnh mẫu §5.1, thời gian 5–10 phút, bấm cho phép khi Antigravity hỏi chạy lệnh); ghi rõ cách này không cần `CAI-DAT.bat`.
- `docs/vi/cai-dat-windows.md`: thêm đầu file "Cách nhanh: nhờ AI cài" trỏ tới mục trên; các bước bấm đúp giữ nguyên.

### 6.7 `docs/vi/xu-ly-loi.md` — mục "Máy trường chặn cài đặt"

- Dấu hiệu: AI báo không tải được, không chạy được bộ cài Python, hoặc PowerShell bị chặn.
- Đoạn gửi IT (AI chép nguyên đoạn này khi báo bị chặn): cần cho tài khoản này truy cập `python.org`, `pypi.org`, `files.pythonhosted.org`, `github.com`, `codeload.github.com`; cho phép cài Python 3.12 cho người dùng hiện tại; cho phép chạy PowerShell với `-ExecutionPolicy Bypass` cho một tiến trình.

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
| Người dùng đã cài bằng `CAI-DAT.bat` (không có `venv`) | `doctor --no-smoke --json` bằng `python` đạt → không cài lại |
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
  - `tool -Name ffmpeg -PlanOnly` khi thiếu → `install/winget`.
- **`test_vi_layer.py`:** `docs/vi/cai-dat-bang-ai.md` tồn tại, có lệnh `-Action setup -Auto`, có mục "Không được làm" với `iex` và quyền quản trị; `AGENTS.vi.md` mục 4 có `venv\Scripts\python.exe`, mục 9 có tiêu đề mới và trỏ tới `cai-dat-bang-ai.md`; mục 10 vẫn là mục cuối; README có khối "Dành cho AI agent".

### 8.2 Kiểm thử headless

Chạy trên bản sao repo trong thư mục tạm, không có `venv`: `claude -p "<câu lệnh>" --disallowedTools Bash Write Edit NotebookEdit PowerShell Agent Monitor Skill`. Danh sách cờ chỉ là lớp phòng ngừa, không phải sandbox.

| Ca | Câu lệnh | Đạt khi |
|---|---|---|
| 1 | `cài đặt giúp em` | Đầu ra nêu lệnh `pptmaster.ps1 -Action setup -Auto` hoặc `cai-dat-bang-ai.md`; không có `pip install` tự viết, không có `iex` |
| 2 | `Tạo bài giảng Toán 6 bài Phân số` | AI kiểm tra hoặc cài môi trường trước khi khởi tạo dự án |

### 8.3 Nghiệm thu máy thật (chủ repo)

1. Máy Windows chưa có Python (hoặc tài khoản Windows mới trên máy chưa cài Python cho mọi người dùng); cài Antigravity.
2. Mở thư mục trống, dán câu lệnh mẫu §5.1 → AI báo sẵn sàng, "Xuất thử PPTX" đạt. Nghiệm thu diễn ra trước khi phát hành, nên nhánh `feat/vi-tu-cai` được push lên `origin` (không force) và câu lệnh mẫu thêm "nhánh feat/vi-tu-cai".
3. Nhắn tạo một bài giảng → nhận PPTX, mở bằng PowerPoint.
4. Đóng Antigravity, mở lại thư mục, tạo bài thứ hai → AI không cài lại.

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
