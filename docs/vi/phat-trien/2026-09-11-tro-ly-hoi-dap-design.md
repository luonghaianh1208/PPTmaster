# Thiết kế: Trợ lý hỏi đáp cho thầy cô (M2 — `v6.3.2-vi.2`)

- **Ngày:** 2026-09-11
- **Người phụ trách:** Lương Hải Anh — 2Anh AI Education
- **Trạng thái:** Đã duyệt thiết kế; chờ chủ repo đọc spec
- **Nền:** `main` @ `75562b38` (bản Việt `v6.3.2-vi.1`); upstream mới nhất vẫn là `v6.3.2` → không cần đồng bộ upstream
- **Thay thế:** phạm vi M2 cũ trong [`2026-09-10-dong-goi-ban-viet-design.md`](2026-09-10-dong-goi-ban-viet-design.md) §5.6 (5 gói template SVG) — bỏ, thay bằng bộ file hướng dẫn cho AI

---

## 1. Mục tiêu

Khi thầy cô nhờ AI làm slide, AI **hỏi một lượt ngắn** những thông tin chỉ thầy cô biết (môn, lớp, bộ sách, mục tiêu, học sinh, bối cảnh trình chiếu, đơn vị…), ghi lại thành bản tóm tắt yêu cầu, rồi mới chạy quy trình tạo PPTX của upstream — để nội dung sát thực tế mà không cần template SVG.

### Tiêu chí thành công

1. Với mỗi loại việc trong 5 loại (§6), tin nhắn trả lời đầu tiên của AI là bộ câu hỏi đúng loại (≤ 7 câu bắt buộc, mỗi câu có gợi ý), AI chưa khởi tạo dự án — kiểm bằng Claude Code headless (§8.2).
2. Câu lệnh có "tạo nhanh" chỉ nhận 2–3 câu hỏi tối thiểu rồi đi chế độ Quick của upstream; câu lệnh có thêm "không cần hỏi lại" hoặc "không hỏi gì" thì không hỏi câu nào.
3. Yêu cầu không thuộc 5 loại không bị hỏi bộ câu Việt.
4. Hồ sơ đơn vị được lưu ở lần đầu và dùng lại ở lần sau (không hỏi lại các mục đã có).
5. Bước xác nhận của upstream hiển thị trong khung chat, đã điền sẵn từ brief — không mở trang web khi thầy cô đã đồng ý xác nhận trong chat.
6. Không file nào của upstream bị sửa: `git diff v6.3.2 -- skills/` rỗng; guard bản quyền exit 0; toàn bộ test `tools/vi/tests` đạt.
7. Chủ repo nghiệm thu trên Antigravity: một bài giảng thật đi trọn lượt hỏi → xác nhận trong chat → PPTX mở được trong PowerPoint.

---

## 2. Bối cảnh & phát hiện (upstream `v6.3.2`)

| # | Phát hiện | Nguồn | Hệ quả thiết kế |
|---|---|---|---|
| 1 | Quy trình Generate mặc định có bước xác nhận **bắt buộc** 2 giai đoạn. Giai đoạn 1: `primary_language`, `audience`, `communication_intent`, `audience_outcome`, `core_message`, `delivery_context`, `artifact_afterlife`, `content_divergence`, `canvas`, template/thiết kế tự do. Giai đoạn 2: số trang, phong cách, màu, icon, font, ảnh, ghi chú/hiệu ứng/thuyết minh… | `workflows/generate-pptx.md:91-182`; `references/strategist.md:25-26` | Không được bỏ bước này (trừ Quick); bộ câu hỏi Việt phải **điền sẵn** cho nó chứ không thay nó. |
| 2 | Cách xác nhận: mặc định **mở trang web** (`webbrowser.open`); dùng khung chat khi người dùng *trong lần chạy này* yêu cầu/đồng ý xác nhận trong chat hoặc từ chối trang; "AI tự quyết" → AI chốt và đưa một bản tóm tắt cuối. | `references/confirm-surface.md:7-15`; `scripts/confirm_ui/server.py:1136-1137` | Lượt hỏi Việt phải nhắc lại cách xác nhận để thầy cô **đồng ý ngay trong lần chạy**; lựa chọn lưu trong hồ sơ tự nó chưa đủ. |
| 3 | Mọi `.md` trong `projects/<dự án>/sources/` là bằng chứng hợp lệ để Strategist soạn đề xuất giai đoạn 1; câu trả lời trong chat cũng hợp lệ. | `references/artifact-ownership.md:15`; `workflows/generate-pptx.md:55,95` | Ghi câu trả lời thành brief rồi import vào `sources/`. |
| 4 | `import-sources`: file ngoài `projects/` được **chép**; file trong `projects/` bị **chuyển** vào `sources/` trừ khi có `--copy`. | `scripts/docs/project.md:38-56` | Brief đặt trong `projects/` (import sẽ chuyển đi — đúng ý); **không bao giờ** import file hồ sơ; logo import bằng lệnh riêng — `--copy` áp cho cả lệnh, nên không gộp chung với brief. |
| 5 | Không script nào của upstream duyệt thư mục gốc `projects/`; `projects/*` bị `.gitignore`. | tìm kiếm `skills/ppt-master/scripts/*.py`; `.gitignore` | `projects/_ho-so-don-vi.md` an toàn và không lên GitHub. |
| 6 | Upstream tự nghiên cứu nội dung khi thiếu dữ kiện; không hỏi thầy cô về bài dạy. | `workflows/stages/topic-research.md:13-18,36` | Không hỏi kiến thức bài học; chỉ hỏi thông tin sư phạm/bối cảnh. |
| 7 | Quick kích hoạt khi người dùng yêu cầu rõ tạo nhanh/bỏ xác nhận; bỏ Strategist, trang xác nhận, spec. `AGENTS.vi.md:20` đã gắn "tạo nhanh", "làm nhanh", "không cần hỏi lại". | `workflows/profiles/quick-generate.md:9-25` | Tạo nhanh: hỏi 2–3 câu tối thiểu rồi đi Quick. |
| 8 | Dàn ý của người dùng chỉ là điểm khởi đầu, trừ khi ghi rõ là bố cục cuối; điều cấm phải là lời người dùng, gắn `(user)`. | `references/strategist.md:107`; `templates/spec_lock_reference.md:26` | Brief tách "Thầy cô yêu cầu" (lời thầy cô) và "AI đề xuất (thầy cô đã đồng ý)". |
| 9 | Có sẵn mode `instructional`, style `workshop-teaching`, layout 4:3 lớp học `presentation_core_43`; không có hồ sơ riêng cho giáo dục Việt Nam. | `references/modes/_index.md:17`; `templates/styles/styles_index.json:132`; `templates/layouts/layouts_index.json:62` | Hướng dẫn gợi ý các lựa chọn upstream sẵn có thay vì tạo mới. |
| 10 | Quy ước upstream: một ngôn ngữ mỗi file; không ghi cứng ngôn ngữ trả lời; quy tắc có một chủ. | `docs/rules/language.md:7,24,28`; `docs/rules/prompt-layers.md:49` | File hướng dẫn nằm ngoài `skills/`, viết hoàn toàn tiếng Việt, chỉ áp dụng khi người dùng viết tiếng Việt. |

---

## 3. Quyết định đã chốt

| Chủ đề | Quyết định |
|---|---|
| Hướng M2 | Bỏ 5 gói template SVG; làm bộ file hướng dẫn để AI hỏi thầy cô trước khi tạo PPTX |
| Loại việc (5) | Bài giảng · Báo cáo – tổng kết · Hoạt động Đoàn – sự kiện · Poster/ấn phẩm Zalo – Facebook · Tập huấn/workshop |
| Người dùng file tập huấn | Thầy cô tự làm slide tập huấn, bồi dưỡng, sinh hoạt chuyên môn |
| Phối hợp với bước xác nhận upstream | Phương án A: hỏi trước → brief vào `sources/` → bước xác nhận upstream điền sẵn, hiển thị trong chat khi thầy cô đồng ý |
| Tạo nhanh | "tạo nhanh"/"làm nhanh": hỏi 2–3 câu tối thiểu (điều kiện của lớp Việt, hỏi trước khi Quick bắt đầu) rồi đi Quick; "không cần hỏi lại"/"không hỏi gì": không hỏi, mục còn thiếu ghi "(AI đề xuất, chưa duyệt)", rồi đi Quick |
| Hồ sơ đơn vị | Lưu một lần ở `projects/_ho-so-don-vi.md`, dùng lại lần sau |
| Phiên bản | `v6.3.2-vi.2`, nhánh `feat/vi-tro-ly` |

---

## 4. Kiến trúc

**Nguyên tắc:** lớp hướng dẫn chỉ **bổ sung thông tin**, không thay quy trình. `skills/ppt-master/SKILL.md` và `AGENTS.md` luôn thắng. Không file nào của upstream bị sửa.

### 4.1 File thêm/sửa trong repo

```text
AGENTS.vi.md                          # (sửa) thêm mục "## 10. Hỗ trợ thầy cô trước khi tạo PPTX"
docs/vi/tro-ly/
├── quy-trinh-hoi.md                  # (mới) quy tắc chung: nhận diện, hồ sơ, cách hỏi, brief, import, tạo nhanh
├── bai-giang.md                      # (mới) loại 1
├── bao-cao-tong-ket.md               # (mới) loại 2
├── hoat-dong-doan.md                 # (mới) loại 3
├── poster-mang-xa-hoi.md             # (mới) loại 4
├── tap-huan-workshop.md              # (mới) loại 5
├── mau-ho-so-don-vi.md               # (mới) mẫu hồ sơ đơn vị
└── mau-brief.md                      # (mới) mẫu brief
docs/vi/bat-dau-nhanh.md              # (sửa) giải thích lượt hỏi, hồ sơ đơn vị
docs/vi/cau-lenh-mau.md               # (sửa) ví dụ trả lời lượt hỏi, tạo nhanh
CHANGELOG-VI.md                       # (sửa) mục 6.3.2-vi.2
README.md                             # (sửa) dòng phiên bản + 1 dòng giới thiệu trợ lý hỏi đáp
tools/vi/tests/test_vi_layer.py       # (sửa) test cấu trúc bộ hướng dẫn
```

Tất cả nằm trong danh sách đường dẫn đã cho phép của `ALLOWED_CHANGES` (`docs/vi/*`, `AGENTS.vi.md`, `README.md`, `CHANGELOG-VI.md`, `tools/vi/*`) — không cần mở rộng ranh giới.

### 4.2 File sinh ra khi dùng (không commit — `projects/*` bị gitignore)

| File | Nội dung | Vòng đời |
|---|---|---|
| `projects/_ho-so-don-vi.md` | Hồ sơ đơn vị theo mẫu §6.8 | Tạo lần đầu, cập nhật khi thầy cô đổi; không bao giờ import |
| `projects/_brief-<tên_dự_án>.md` | Brief theo mẫu §6.9 | Tạo sau lượt hỏi; import vào `projects/<dự án>/sources/` ngay sau `init` (import chuyển file đi) |

---

## 5. Luồng hoạt động

### 5.1 Quy trình mặc định

1. **Nhận diện loại việc.** AI đọc `AGENTS.vi.md` §10. Nếu người dùng viết tiếng Việt và yêu cầu thuộc 1 trong 5 loại (câu lệnh có dấu hiệu trong bảng từ khoá hoặc nói rõ loại việc; bối cảnh trường học hay Đoàn một mình không đủ để xếp loại) → tiếp bước 2. Không thuộc → quy trình upstream như hiện tại. Không rõ/thuộc hai loại → hỏi **một** câu chọn loại (kèm gợi ý loại gần nhất).
2. **Đọc hướng dẫn.** Đọc `docs/vi/tro-ly/quy-trinh-hoi.md`, file loại việc tương ứng, và `projects/_ho-so-don-vi.md` nếu có.
3. **Hỏi một lượt (bước chặn của lớp Việt).** Một tin nhắn duy nhất gồm:
   - (Lần đầu hoặc hồ sơ thiếu mục) phần hồ sơ đơn vị;
   - (Hồ sơ đã có) một dòng "Hồ sơ hiện tại: <đơn vị> — <người trình bày>. Dùng tiếp?";
   - các câu hỏi bắt buộc của loại việc, đánh số, mỗi câu có gợi ý;
   - dòng chốt cách xác nhận, ví dụ "Ở bước xác nhận, em sẽ tóm tắt trong khung chat để thầy cô duyệt (không mở trang web) — thầy cô đồng ý nhé?" (hoặc theo cách đã lưu trong hồ sơ).
   AI dừng và chờ trả lời; không khởi tạo dự án, không tra cứu, không soạn slide trong lúc chờ.
4. **Ghi lại.** Tạo/cập nhật hồ sơ; viết `projects/_brief-<tên_dự_án>.md`.
5. **Vào quy trình upstream.** Theo `SKILL.md`: xử lý tài liệu nguồn nếu có → `project_manager.py init <tên_dự_án>` → `import-sources <đường_dẫn_dự_án> projects/_brief-<tên_dự_án>.md` (không `--copy`, để brief được chuyển vào `sources/`) → nếu có logo: lệnh `import-sources <đường_dẫn_dự_án> <logo>` riêng, thêm `--copy` chỉ khi logo nằm trong `projects/` → các bước còn lại. Ở bước xác nhận, nếu thầy cô đã đồng ý xác nhận trong chat thì dùng nhánh chat của `confirm-surface.md` §1; đề xuất giai đoạn 1–2 lấy từ brief; thầy cô duyệt/sửa trong chat.
6. **Xuất PPTX** và báo đường dẫn file trong `projects/<dự án>/exports/`.

### 5.2 Tạo nhanh

Áp dụng khi câu lệnh có "tạo nhanh", "làm nhanh", "không cần hỏi lại" hoặc "không hỏi gì". Không hỏi hồ sơ đơn vị; chưa có hồ sơ thì brief ghi "(chưa có hồ sơ đơn vị)" dưới mục Đơn vị.

- **"tạo nhanh", "làm nhanh":** AI hỏi **chỉ** các câu trong mục "Tạo nhanh" của loại việc (2–3 câu, một tin nhắn) — bỏ qua nếu câu lệnh đã có đủ các thông tin đó. Các câu này là điều kiện của lớp Việt, được hỏi **trước khi** hồ sơ Quick của upstream bắt đầu, không phải điểm dừng trong lượt chạy Quick (Quick upstream "decide every unspecified choice directly without asking").
- **"không cần hỏi lại", "không hỏi gì"** (kể cả khi có thêm "tạo nhanh"): AI không hỏi câu nào.

Sau đó: viết brief (mọi mục thầy cô chưa nêu ghi "(AI đề xuất, chưa duyệt)") → chạy chế độ Quick của upstream (không có bước xác nhận).

---

## 6. Nội dung từng file

### 6.1 `quy-trinh-hoi.md` — quy tắc chung (bắt buộc có các mục)

- `## Khi nào áp dụng` — người dùng viết tiếng Việt, yêu cầu thuộc 5 loại và bối cảnh là trường học hoặc Đoàn; bảng từ khoá nhận diện từng loại; câu hỏi chọn loại việc (không phải chọn quy trình của upstream) chỉ khi có dấu hiệu nhưng chưa rõ loại.
- `## Thứ tự ưu tiên` — `SKILL.md`/`AGENTS.md` thắng; không bỏ qua bước xác nhận (trừ Quick); không sửa/bỏ qua `attribution_guard.py`.
- `## Hồ sơ đơn vị` — đường dẫn, các mục, khi nào hỏi/cập nhật, không import hồ sơ, xử lý nhiều người dùng chung máy, logo không tồn tại; không gợi ý sẵn tên đơn vị, họ tên người trình bày hay logo (không lấy từ README, NOTICE, bộ nhớ của AI hay dự án khác), chỉ gợi ý cấp học suy ra từ câu lệnh và màu chủ đạo; không hỏi riêng "Cách xác nhận" mà ghi theo câu trả lời cho dòng chốt.
- `## Cách hỏi` — xưng "em", gọi "thầy cô"; một tin nhắn; tối đa 7 câu bắt buộc (không tính phần hồ sơ lần đầu); đánh số; mỗi câu hỏi đánh số có gợi ý (không áp cho phần hồ sơ); câu hỏi tuỳ chọn có thể thay câu bắt buộc mà câu lệnh đã trả lời đủ; câu cuối chốt cách xác nhận; không hỏi kiến thức bài học; trả lời thiếu → dùng gợi ý, ghi "(AI đề xuất, chưa duyệt)", chỉ hỏi lại câu bắt buộc không thể đoán (tên bài, loại báo cáo…), tối đa một lần; dòng chốt cách xác nhận không áp quy tắc dùng gợi ý — không trả lời thì hỏi lại đúng một câu trong chat trước khi mở trang web.
- `## Ghi brief và đưa vào dự án` — tên file, mẫu, hai lệnh `import-sources` tách riêng (brief: không `--copy`, file bị chuyển vào `sources/`; logo: lệnh riêng, `--copy` chỉ khi logo nằm trong `projects/`), tách lời thầy cô và đề xuất, điều cấm ghi đúng lời thầy cô.
- `## Tạo nhanh` — §5.2.
- `## Đổi ý giữa chừng` — cập nhật brief; nếu đã qua bước xác nhận thì nêu thay đổi để duyệt lại ở bước xác nhận của upstream.

### 6.2 Cấu trúc bắt buộc của mỗi file loại việc

Mỗi file trong 5 file loại việc có đúng các tiêu đề cấp 2 sau, theo thứ tự:

1. `## Khi nào dùng` — mô tả + ví dụ câu lệnh.
2. `## Câu hỏi bắt buộc` — danh sách đánh số (≤ 7), mỗi câu có dòng "Gợi ý: …".
3. `## Câu hỏi tuỳ chọn` — chỉ hỏi khi liên quan (có thể trống kèm "Không có").
4. `## Tạo nhanh` — danh sách đánh số 2–3 câu (tham chiếu tới câu bắt buộc).
5. `## Cấu trúc gợi ý` — mạch trang gợi ý.
6. `## Phong cách gợi ý` — màu/font/cỡ chữ mô tả bằng chữ, và lựa chọn upstream nên đề xuất (mode/style/layout) nếu phù hợp.
7. `## Khổ slide` — key canvas mặc định (phải có trong `CANVAS_FORMATS` của upstream) và khi nào đổi.
8. `## Ghi vào brief` — các mục của brief lấy từ câu trả lời.

### 6.3 `bai-giang.md` — Bài giảng (mọi môn, mọi cấp)

- **Câu hỏi bắt buộc:** (1) môn, lớp, tên bài, chủ đề/chương; (2) bộ sách — Kết nối tri thức với cuộc sống / Chân trời sáng tạo / Cánh diều / khác; (3) số tiết, thời lượng — gợi ý 1 tiết × 45 phút (Tiểu học: 35 phút); (4) mục tiêu kiến thức, năng lực, phẩm chất — gợi ý AI đề xuất theo chương trình GDPT 2018; (5) hoạt động mong muốn — khởi động trò chơi, thảo luận nhóm, phiếu học tập, trắc nghiệm, video; (6) đặc điểm học sinh và thiết bị — trình độ lớp, máy chiếu 16:9/4:3, TV thông minh; (7) tài liệu có sẵn — giáo án Word, ảnh SGK, đề bài.
- **Tạo nhanh:** câu 1 và câu 2.
- **Cấu trúc gợi ý:** THCS/THPT theo 4 hoạt động của Công văn 5512/BGDĐT-GDTrH (Mở đầu → Hình thành kiến thức mới → Luyện tập → Vận dụng), kèm trang mục tiêu và trang tổng kết/dặn dò; Tiểu học theo mạch tiết học thầy cô quen dùng, không áp khung 4 hoạt động nếu thầy cô không yêu cầu.
- **Phong cách gợi ý:** học đường sáng rõ, chữ nội dung ≥ 28pt, Segoe UI hoặc Arial; gợi ý mode `instructional`.
- **Khổ slide:** `ppt169`; `ppt43` khi thầy cô dùng máy chiếu 4:3 (có thể gợi ý layout `presentation_core_43`).

### 6.4 `bao-cao-tong-ket.md` — Báo cáo – tổng kết

- **Câu hỏi bắt buộc:** (1) loại báo cáo — sơ kết học kỳ / tổng kết năm học / chuyên đề / thi đua; (2) kỳ báo cáo và đơn vị/bộ phận; (3) người nghe — hội nghị viên chức, cha mẹ học sinh, đoàn kiểm tra; (4) thời lượng — gợi ý 15 phút, 12–15 slide; (5) nguồn số liệu (Excel/Word đính kèm) và số liệu cần biểu đồ; (6) nội dung cần nhấn mạnh — thành tích, hạn chế.
- **Tạo nhanh:** câu 1 (kèm kỳ báo cáo) và câu 5.
- **Cấu trúc gợi ý:** Kết quả đạt được → Hạn chế, nguyên nhân → Bài học kinh nghiệm → Phương hướng, nhiệm vụ → Kiến nghị, đề xuất.
- **Phong cách gợi ý:** hành chính trang trọng; Times New Roman để giữ phong cách văn bản hành chính (Nghị định 30/2020/NĐ-CP quy định font này cho văn bản; slide không bắt buộc); màu theo hồ sơ đơn vị; biểu đồ rõ số liệu.
- **Khổ slide:** `ppt169`.

### 6.5 `hoat-dong-doan.md` — Hoạt động Đoàn – sự kiện

- **Câu hỏi bắt buộc:** (1) tên chương trình, dịp tổ chức, cấp tổ chức; (2) mục đích slide — giới thiệu / thể lệ / chiếu sân khấu / tổng kết, trao giải; (3) thời gian, địa điểm, đối tượng tham gia; (4) nội dung chính — thể lệ, cơ cấu giải, lịch trình, đơn vị phối hợp, nhà tài trợ; (5) nhận diện — huy hiệu Đoàn hay logo chương trình, màu riêng; (6) ảnh có sẵn hay AI tạo.
- **Tạo nhanh:** câu 1 và câu 2.
- **Cấu trúc gợi ý:** theo mục đích — giới thiệu (bìa → mục đích → đối tượng → nội dung → lịch trình → liên hệ); thể lệ (bìa → điều kiện → vòng thi → tiêu chí → giải thưởng → mốc thời gian); tổng kết (bìa → kết quả → hình ảnh → vinh danh → cảm ơn).
- **Phong cách gợi ý:** trẻ trung, năng động, chữ đậm đọc được từ xa; dùng huy hiệu Đoàn đúng tỉ lệ và màu gốc (không kéo giãn, không đổi màu).
- **Khổ slide:** `ppt169`.

### 6.6 `poster-mang-xa-hoi.md` — Poster/ấn phẩm Zalo – Facebook

- **Câu hỏi bắt buộc:** (1) kênh đăng — nhóm Zalo lớp/phụ huynh, Fanpage Facebook, TikTok; (2) khổ — gợi ý theo kênh; (3) một ảnh hay chuỗi nhiều ảnh (bao nhiêu); (4) thông điệp chính và thông tin bắt buộc — thời gian, địa điểm, link/mã QR, hạn đăng ký; (5) người đọc và hành động mong muốn — đăng ký, tham dự, chia sẻ; (6) ảnh có sẵn hay AI tạo, có dùng logo không.
- **Tạo nhanh:** câu 1 và câu 4.
- **Cấu trúc gợi ý:** một ảnh — tiêu đề lớn → thông tin chính → lời kêu gọi → logo/QR; chuỗi — ảnh 1 thu hút, ảnh giữa mỗi ảnh một ý, ảnh cuối kêu gọi hành động.
- **Phong cách gợi ý:** chữ lớn đọc được trên điện thoại, ≈ 30 chữ mỗi ảnh, tương phản cao; giữ chữ quan trọng cách xa mép trên và mép dưới (bảng tin Facebook cắt ảnh dọc).
- **Khổ slide:** nhóm Zalo/Facebook bài vuông → `moments`; bài dọc Facebook/TikTok → `xiaohongshu`; story/Reels → `story`; ảnh bìa bài viết Zalo OA → `wechat`.

### 6.7 `tap-huan-workshop.md` — Tập huấn/workshop (thầy cô tự làm)

- **Câu hỏi bắt buộc:** (1) chủ đề và hình thức — sinh hoạt tổ chuyên môn / bồi dưỡng thường xuyên / chia sẻ chuyên đề / tập huấn toàn trường; (2) đối tượng và mức độ đã biết; (3) thời lượng, tỉ lệ lý thuyết – thực hành; (4) sau buổi người học làm được gì (2–3 kết quả); (5) có hướng dẫn thao tác từng bước không, có ảnh chụp màn hình không; (6) hoạt động tương tác — thảo luận, bài tập, khảo sát.
- **Tạo nhanh:** câu 1 (chủ đề) kèm câu 2 (đối tượng) và câu 3 (thời lượng).
- **Cấu trúc gợi ý:** Mục tiêu → Vì sao cần → Nội dung/thao tác (bài tập ngay sau mỗi phần) → Thực hành → Hỏi đáp → Tổng kết, tài liệu.
- **Phong cách gợi ý:** gần gũi, rõ ràng; gợi ý mode `instructional` và style `workshop-teaching` của upstream.
- **Khổ slide:** `ppt169`.

> Ghi chú: mục "Tạo nhanh" của loại 5 gộp 3 thông tin vào tối đa 3 câu.

### 6.8 `mau-ho-so-don-vi.md` — mẫu hồ sơ

```markdown
# Hồ sơ đơn vị
- Cập nhật: <YYYY-MM-DD>
- Đơn vị: <tên trường/đơn vị>
- Cấp học: <Tiểu học | THCS | THPT | khác>
- Người trình bày: <họ tên> — <chức danh>
- Logo: <đường dẫn file hoặc "không dùng">
- Màu chủ đạo: <mã màu hoặc mô tả>
- Cách xác nhận: <khung chat | trang web>
```

### 6.9 `mau-brief.md` — mẫu brief

```markdown
# Tóm tắt yêu cầu từ thầy cô
- Loại việc: <1 trong 5 loại>
- Ngày: <YYYY-MM-DD>

## Đơn vị
<chép các mục cần dùng từ hồ sơ đơn vị; tạo nhanh khi chưa có hồ sơ thì ghi "(chưa có hồ sơ đơn vị)">

## Thầy cô yêu cầu
<ghi đúng lời thầy cô, mỗi ý một dòng; điều thầy cô không muốn có thì ghi nguyên văn lời thầy cô>

## AI đề xuất (thầy cô đã đồng ý)
<mỗi dòng một gợi ý, ghi kèm "(thầy cô đồng ý)" hoặc "(AI đề xuất, chưa duyệt)">

## Cấu trúc gợi ý
<mạch trang gợi ý — là điểm khởi đầu, không phải bố cục cuối trừ khi thầy cô yêu cầu giữ nguyên>

## Phong cách gợi ý
<màu, font, cỡ chữ, mode/style upstream gợi ý>

## Khổ slide
<key canvas>
```

> Ghi chú: giữ nguyên tiêu đề `## AI đề xuất (thầy cô đã đồng ý)`, nhưng tiêu đề không tự chứng minh thầy cô đã duyệt. Mỗi dòng dưới tiêu đề phải ghi kèm "(thầy cô đồng ý)" khi thầy cô đã chấp nhận gợi ý, hoặc "(AI đề xuất, chưa duyệt)" khi thầy cô không trả lời câu đó hoặc khi tạo nhanh.

### 6.10 `AGENTS.vi.md` — mục mới

`## 10. Hỗ trợ thầy cô trước khi tạo PPTX` gồm: điều kiện áp dụng (người dùng viết tiếng Việt, yêu cầu thuộc 5 loại); yêu cầu **đọc `docs/vi/tro-ly/quy-trinh-hoi.md` trước**; bảng 5 loại → link file tương ứng; nhắc SKILL.md vẫn thắng và bước xác nhận upstream vẫn bắt buộc (trừ Quick). Mục 3 hiện có ("tạo nhanh"…) giữ nguyên.

---

## 7. Tình huống biên

| Tình huống | Xử lý |
|---|---|
| Trả lời thiếu câu | Dùng gợi ý, ghi "(AI đề xuất, chưa duyệt)"; chỉ hỏi lại câu bắt buộc không thể đoán, tối đa một lần |
| Không trả lời dòng chốt cách xác nhận | Không coi là đồng ý (kể cả khi hồ sơ đã lưu cách xác nhận); trước khi mở trang web, hỏi lại đúng một câu trong chat về cách xác nhận |
| Chưa có hồ sơ đơn vị | Hỏi hồ sơ nhưng không gợi ý sẵn tên đơn vị, họ tên người trình bày hay logo; không lấy tên từ README, NOTICE, bộ nhớ của AI hay dự án khác |
| Không thuộc 5 loại, hoặc có từ khoá nhưng không phải bối cảnh trường học/Đoàn | Không hỏi bộ câu Việt; quy trình upstream như hiện tại |
| Chỉ có bối cảnh trường học hay Đoàn, câu lệnh không có dấu hiệu trong bảng và không nói rõ loại việc (ví dụ giới thiệu trường) | Coi là không thuộc 5 loại: không hỏi bộ câu Việt, không tìm hồ sơ đơn vị; quy trình upstream như hiện tại |
| Không rõ loại / thuộc hai loại | Hỏi một câu chọn loại việc (không phải chọn quy trình của upstream), kèm gợi ý loại gần nhất |
| Đổi ý giữa chừng | Cập nhật brief; nếu đã qua bước xác nhận, nêu thay đổi để duyệt lại ở bước xác nhận upstream |
| Muốn đổi hồ sơ | Sửa hồ sơ, báo lại một dòng tóm tắt |
| Logo không tồn tại | Báo, xin đường dẫn khác hoặc làm tiếp không logo |
| Nhiều người dùng chung máy | Lượt hỏi nêu "Hồ sơ hiện tại là của …, dùng tiếp?" |
| Người dùng viết tiếng Anh | Không áp bộ câu Việt; quy tắc ngôn ngữ upstream |
| Tạo nhanh nhưng câu lệnh đã đủ thông tin tối thiểu | Không hỏi; viết brief rồi chạy Quick |
| Câu lệnh có "không cần hỏi lại" hoặc "không hỏi gì" | Không hỏi câu nào (kể cả hồ sơ, kể cả khi thiếu thông tin); brief ghi mọi mục còn thiếu "(AI đề xuất, chưa duyệt)"; chạy Quick |
| Thầy cô nói "AI tự quyết" ở bước xác nhận | Nhánh "delegated" của `confirm-surface.md` §1 |
| Thầy cô muốn dùng trang web | Nhánh trang web mặc định của upstream; brief vẫn điền sẵn đề xuất |

---

## 8. Kiểm thử

### 8.1 Unit test (`tools/vi/tests/test_vi_layer.py`)

- Đủ 8 file trong `docs/vi/tro-ly/`.
- Mỗi file loại việc có đủ 8 tiêu đề cấp 2 của §6.2 theo đúng thứ tự.
- `## Câu hỏi bắt buộc`: 1–7 mục đánh số, mỗi mục có "Gợi ý:"; `## Tạo nhanh`: 2–3 mục đánh số.
- Mọi key canvas trong `## Khổ slide` (dạng `` `key` ``) có trong `CANVAS_FORMATS` của upstream.
- `quy-trinh-hoi.md` có đủ các mục cấp 2 của §6.1.
- `AGENTS.vi.md` có tiêu đề `## 10. Hỗ trợ thầy cô trước khi tạo PPTX`, link tới `quy-trinh-hoi.md` và đủ 5 file loại việc (link test hiện có kiểm link hợp lệ).
- Toàn bộ test hiện có vẫn đạt.

### 8.2 Kiểm tra hành vi — Claude Code headless

Chạy trong thư mục repo, chặn công cụ ghi và lệnh: `claude -p "<câu lệnh>" --disallowedTools Bash Write Edit NotebookEdit PowerShell Agent Monitor Skill` (câu lệnh đứng trước, cờ đứng sau). Trên Windows, danh sách cờ cũ không chặn được PowerShell và một lượt chạy đã ghi file vào `projects/`, nên danh sách cờ này chỉ là một lớp phòng ngừa, không phải sandbox. 9 ca:

| Ca | Câu lệnh | Đạt khi |
|---|---|---|
| 1 | "Tạo bài giảng Vật lí 10 bài Chuyển động thẳng đều" | Trả lời là danh sách câu hỏi có bộ sách, số tiết, mục tiêu; không tuyên bố đã tạo dự án |
| 2 | "Làm slide báo cáo tổng kết năm học của tổ Toán" | Có câu hỏi kỳ báo cáo/người nghe/nguồn số liệu |
| 3 | "Làm slide giới thiệu cuộc thi Tuổi trẻ sáng tạo của Đoàn trường" | Có câu hỏi mục đích slide, thời gian – địa điểm, thể lệ/giải |
| 4 | "Làm poster Zalo thông báo họp phụ huynh" | Có câu hỏi kênh/khổ, thông tin bắt buộc |
| 5 | "Làm slide tập huấn dùng AI soạn giáo án cho giáo viên" | Có câu hỏi đối tượng, thời lượng, kết quả sau buổi |
| 6 | "Tạo nhanh bài giảng Hoá 11" | Chỉ 2–3 câu hỏi (thiếu tên bài, bộ sách) |
| 7 | "Make a 5-slide deck about our software product" | Không dùng bộ câu hỏi Việt |
| 8 | "Tạo bài thuyết trình 8 slide giới thiệu câu lạc bộ tin học của trường, phong cách trẻ trung" | Việc ngoài 5 loại viết bằng tiếng Việt: không có "Gợi ý:", không hỏi hồ sơ đơn vị |
| 9 | "Tạo bài giảng Hoá học 11 bài Nitrogen, không cần hỏi lại" | Không hỏi câu nào: không có "Gợi ý:", không có dòng chốt cách xác nhận |

Ca 1–5 còn kiểm thêm: không có tên người hay tên trường tự bịa (ví dụ tên lấy từ README, NOTICE hay bộ nhớ của AI). Kết quả từng ca ghi vào báo cáo kiểm thử của lần triển khai.

### 8.3 Nghiệm thu của chủ repo (Antigravity)

Một bài giảng thật: lượt hỏi → trả lời → bước xác nhận hiển thị trong chat, đã điền sẵn → PPTX mở trong PowerPoint, chữ tiếng Việt đúng dấu; chạy lần hai để thấy hồ sơ đơn vị được dùng lại.

---

## 9. Phát hành

- Nhánh `feat/vi-tro-ly` từ `main` @ `75562b38`.
- `CHANGELOG-VI.md` thêm `## 6.3.2-vi.2 — <ngày phát hành>`; `README.md` cập nhật dòng phiên bản và một dòng giới thiệu trợ lý hỏi đáp.
- Sau khi test, review và nghiệm thu đạt: fast-forward `main`, tag `v6.3.2-vi.2`, push `main` + tag (không force), `gh release create v6.3.2-vi.2 --repo luonghaianh1208/PPTmaster` (luôn có `--repo` vì repo có remote `upstream`).

---

## 10. Ngoài phạm vi

- Template SVG (deck/brand) tiếng Việt — bỏ khỏi M2; có thể xét lại sau nếu cần giao diện đồng đều.
- Sửa trang xác nhận, Strategist hay bất kỳ file nào dưới `skills/`.
- Bộ câu hỏi tiếng Anh hoặc cho loại việc ngoài 5 loại.
- Các rủi ro còn treo từ M1 (nạp lại PATH sau winget, gợi ý alias Store khi có `py`, nhãn `[ĐẠT]` trong spec M1, biểu tượng trên console Windows 10, kiểm thử Cursor) — theo dõi riêng.

---

## 11. Rủi ro & giảm thiểu

| Rủi ro | Giảm thiểu |
|---|---|
| AI bỏ qua lượt hỏi, tạo luôn | `AGENTS.vi.md` §10 ghi rõ bước chặn; kiểm headless 5 ca; nghiệm thu Antigravity |
| Thầy cô bị hỏi trùng ở bước xác nhận | Brief điền sẵn đề xuất; xác nhận trong chat chỉ còn duyệt/sửa |
| Trang web vẫn tự mở | Lượt hỏi luôn có dòng chốt xác nhận trong chat → nhánh chat của `confirm-surface.md` |
| Hỏi quá dài, thầy cô ngại | ≤ 7 câu bắt buộc, mỗi câu có gợi ý; test giới hạn số câu |
| Brief bị coi là bố cục cứng hoặc đề xuất AI bị coi là điều cấm | Tách "Thầy cô yêu cầu" / "AI đề xuất"; cấu trúc gợi ý ghi rõ là điểm khởi đầu |
| Hồ sơ của người khác trên máy dùng chung | Lượt hỏi luôn nêu tên hồ sơ và hỏi "dùng tiếp?" |
| Upstream đổi trường xác nhận hoặc cơ chế `sources/` | `sync_upstream.ps1` chạy test; test canvas key; đọc lại §2 khi đồng bộ bản mới |
| File hướng dẫn phình to, tốn ngữ cảnh | Mỗi file loại việc chỉ gồm 8 mục §6.2, viết gọn |
