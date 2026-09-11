# Kiểm tra hành vi trợ lý hỏi đáp (M2) — Claude Code headless

- Ngày chạy: 2026-09-11 (vòng 1 từ 14:33, vòng 2 từ 14:54, vòng 3 từ 15:14, vòng P14 từ 15:44, giờ máy)
- Commit được kiểm: vòng 1 `d126822d`, vòng 2 `c8f7afa9`, vòng 3 `c15e700d` (9 ca); vòng P14 `ca6f2ebb` (ca 8 ba lần, ca 1 và ca 3 một lần). Cây làm việc sạch ở mọi lượt chạy.
- Lệnh vòng 1–3: `claude -p "câu lệnh của ca" --disallowedTools Bash Write Edit NotebookEdit` (Claude Code 2.1.212), chạy từ Git Bash trong thư mục gốc repo; mỗi ca bọc trong `timeout 300`.
- Lệnh vòng P14 (quyết định P16): `claude -p "câu lệnh của ca" --disallowedTools Bash Write Edit NotebookEdit PowerShell Agent`, cũng bọc trong `timeout 300`.
- Chấm: `grep -qi` theo tiêu chí Task 5 Bước 2 của kế hoạch, cộng thêm: ca 1–5 không có "Lương Hải Anh" và "Nguyễn Trãi"; ca 8 và ca 9 mới. Vì `grep -i` ở đây không gấp "Đ"/"đ", từ khoá nào bị báo thiếu đều được grep lại với chữ cái đầu viết hoa trước khi kết luận.
- Số vòng sửa hướng dẫn trong kiểm thử: 2; sau đó thêm một lượt sửa theo quyết định P14 của controller (commit `ca6f2ebb`), không có vòng sửa nào sau vòng P14.
- Thư mục đầu ra (dưới `C:\Users\ADMIN\AppData\Local\Temp\claude\c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster\c49c4a0d-7ba3-41e7-84e8-5f1e4b4ad88f\scratchpad\`): `m2-headless-final` (vòng 1), `m2-headless-final-r2` (vòng 2), `m2-headless-final-r3` (vòng 3), `m2-headless-final-p14` (vòng P14).

## Môi trường chạy

- Các lệnh headless chạy trên máy của chủ repo, bằng tài khoản của chủ repo. Chúng kế thừa file quy tắc cấp người dùng `~/.claude/CLAUDE.md` (bắt trả lời tiếng Việt, thêm quy trình làm việc và nguyên tắc "đừng đoán" riêng của chủ repo) và plugin bộ nhớ OpenViking có tự nhắc lại (auto-recall).
- Bằng chứng: transcript của cả 27 phiên đều có khối `openviking-context` của lượt tự nhắc lại (7 lần mỗi phiên). Ngoài ra, hook `SessionStart` của plugin bộ nhớ (`session-start.mjs`) nạp thêm ngữ cảnh lúc mở phiên, trong đó có tên chủ repo, một tên trường và tên các dự án riêng của chủ repo. Các tên này không đến từ file nào AI đọc trong repo. Đầu ra ca 7 ở cả 3 vòng liệt kê các dự án riêng đó; tên các dự án này đã được lược bỏ khỏi trích đoạn trong báo cáo (thay bằng "‹danh sách dự án riêng của chủ máy — đã lược bỏ›"); ca 8 vòng 1 viện dẫn quy tắc "đừng đoán" của file cấp người dùng.
- Hệ quả: kết quả PASS không thể quy hoàn toàn cho hướng dẫn trong repo. Máy của thầy cô (Antigravity, Cursor, Claude Code mới cài) không có file quy tắc cấp người dùng và bộ nhớ này, nên hành vi có thể khác. Theo quyết định của controller (P12), không chạy lại trong hồ sơ người dùng sạch; nghiệm thu Task 6 nên bắt đầu khi chưa có `projects/_ho-so-don-vi.md` và kiểm tra AI không tự bịa tên trường hay tên người.
- Kiểm tra "không bịa tên": ở cả 3 vòng, đầu ra của ca 1–6, 8 và 9 không có "Lương Hải Anh", "Hải Anh", "Nguyễn Trãi" hay "Trãi", dù các tên này có trong ngữ cảnh được nạp của từng phiên. Phần hồ sơ đơn vị ở ca 1–5 để trống tên đơn vị, người trình bày và logo; chỉ gợi ý cấp học (khi suy ra được) và màu chủ đạo.
- `--disallowedTools Bash Write Edit NotebookEdit` không chặn công cụ PowerShell và Agent trên Windows. Ca 6–9 chạy `attribution_guard.py` qua PowerShell; ca 9 gọi agent con nghiên cứu web. Ở vòng 2, ca 9 ghi `projects/_brief-bai-giang-hoa11-nitrogen.md` bằng PowerShell; bản sao được lưu vào `m2-headless-final-r2` làm bằng chứng và file bị xoá khỏi `projects/` trước vòng 3. Không ca nào tạo `projects/_ho-so-don-vi.md` hay thư mục dự án; sau vòng 3 không có file mới nào trong repo.
- Vòng P14 dùng cờ của quyết định P16 và chặn được PowerShell, Agent. AI vẫn tìm công cụ khác để chạy lệnh: ca 8a và 8b gọi `Monitor` để chạy `attribution_guard.py` (bị từ chối vì cần phê duyệt), ca 8a và 8c gọi công cụ `Skill` (các skill cấp người dùng). Không file nào được ghi; sau vòng P14 `git status` sạch và `projects/` không có file mới.
- `claude -p` chỉ in câu trả lời cuối. Ca 9 không dừng để hỏi mà chạy tiếp chế độ tạo nhanh, nên bị `timeout 300` dừng (exit=124) với file đầu ra rỗng ở cả 3 vòng. Hành vi của ca 9 được chấm thêm bằng transcript phiên trong `~/.claude/projects/c--Users-ADMIN-Downloads-VIBE-CODING-PPTmaster/`.

## Kết quả (vòng 3, commit c15e700d)

| Ca | Câu lệnh | Tiêu chí | Kết quả |
|---|---|---|---|
| 1 | Tạo bài giảng Vật lí 10 bài Chuyển động thẳng đều | Hỏi bộ sách, số tiết, mục tiêu; chưa khởi tạo dự án; không có tên tự bịa | PASS |
| 2 | Làm slide báo cáo tổng kết năm học của tổ Toán | Hỏi kỳ báo cáo, người nghe, số liệu; không có tên tự bịa | PASS |
| 3 | Làm slide giới thiệu cuộc thi Tuổi trẻ sáng tạo của Đoàn trường | Hỏi mục đích, thời gian, thể lệ hoặc giải; không có tên tự bịa | PASS |
| 4 | Làm poster Zalo thông báo họp phụ huynh | Hỏi kênh, khổ ảnh, thông điệp; không có tên tự bịa | PASS |
| 5 | Làm slide tập huấn dùng AI soạn giáo án cho giáo viên | Hỏi đối tượng, thời lượng, kết quả sau buổi; không có tên tự bịa | PASS (chấm thủ công — grep -i dưới C.UTF-8 không gấp Đ/đ; từ khoá duy nhất bị báo thiếu là "đối tượng", có ở dòng 15 dạng "Đối tượng") |
| 6 | Tạo nhanh bài giảng Hoá học 11 | 2–3 câu hỏi, có hỏi bộ sách | PASS (2 câu) |
| 7 | Make a 5-slide deck about our software product | Không dùng bộ câu hỏi Việt | PASS |
| 8 | Tạo bài thuyết trình 8 slide giới thiệu câu lạc bộ tin học của trường, phong cách trẻ trung | Việc ngoài 5 loại: không có "Gợi ý:", không hỏi hồ sơ đơn vị | Vòng 3: FAIL (AI xếp yêu cầu vào "Hoạt động Đoàn – sự kiện" và hỏi bộ câu Việt). Vòng P14 (`ca6f2ebb`): 2/3 theo grep — 8a FAIL vì có "gợi ý:" viết thường trong câu hỏi của luồng upstream, 8b và 8c PASS; cả 3 lần không chạy lượt hỏi Việt và không hỏi câu chọn loại việc (xem "Vòng P14") |
| 9 | Tạo bài giảng Hoá học 11 bài Nitrogen, không cần hỏi lại | Không hỏi câu nào: không có "Gợi ý:", không có dòng chốt cách xác nhận | PASS theo grep trên đầu ra rỗng (hết giờ, exit=124); transcript xác nhận không hỏi câu nào (xem "Mối lo ngại") |

### Kết quả qua các vòng

| Ca | Vòng 1 (d126822d) | Vòng 2 (c8f7afa9) | Vòng 3 (c15e700d) |
|---|---|---|---|
| 1 | PASS | PASS | PASS |
| 2 | PASS | PASS | PASS |
| 3 | PASS | PASS | PASS |
| 4 | PASS | PASS | PASS |
| 5 | PASS (chấm thủ công, Đ/đ) | PASS (chấm thủ công, Đ/đ) | PASS (chấm thủ công, Đ/đ) |
| 6 | PASS | PASS | PASS |
| 7 | PASS | PASS | PASS |
| 8 | PASS | FAIL | FAIL |
| 9 | grep đạt trên đầu ra rỗng; transcript: không hỏi nhưng không đọc `quy-trinh-hoi.md` | grep đạt trên đầu ra rỗng; transcript: không hỏi, đã ghi brief đúng mẫu | grep đạt trên đầu ra rỗng; transcript: không hỏi, đọc đủ quy trình, hết giờ trước khi ghi brief |

Ở cả 3 vòng: ca 1–8 `exit=0`, ca 9 `exit=124`.

Đầu ra chấm của vòng 3:

```
case1 PASS
case2 PASS
case3 PASS
case4 PASS
case5 FAIL
case6 PASS
case7 PASS
case8 FAIL
case9 PASS
qcount6=2
case5 'Đối tượng' lines=1
15:2. Đối tượng tham dự là ai, và họ đã biết về chủ đề đến mức nào?
case8 -i 'Gợi ý:' lines=9
case8 -i 'hồ sơ đơn vị' lines=2
case9.txt: 0 byte
```

### Vòng P14 (commit ca6f2ebb)

Chạy sau quyết định P14: chỉ xếp vào một loại khi câu lệnh có dấu hiệu trong bảng hoặc nói rõ loại việc; bối cảnh trường học hay Đoàn một mình không đủ. Lượt chạy dùng cờ P16, từ 15:44 đến 15:55; ca 8 chạy ba lần (8a, 8b, 8c), ca 1 và ca 3 chạy lại để kiểm hồi quy. Theo chỉ đạo, không có vòng sửa câu chữ nào sau lượt chạy này.

| Ca | exit | Chấm grep | Đọc transcript và đầu ra |
|---|---|---|---|
| 8a | 0 | FAIL | Đọc `quy-trinh-hoi.md`, ghi rằng yêu cầu tương tự ví dụ "giới thiệu trường" nên không thuộc 5 loại, rồi đi luồng Default của `SKILL.md`. Không hỏi hồ sơ, không có brief, không có dòng chốt cách xác nhận, không hỏi câu chọn loại việc. Không đạt tiêu chí vì câu hỏi của luồng upstream có 3 chỗ "gợi ý:" viết thường (dòng 25, 27, 29) và `grep -qi` khớp cả dạng thường. |
| 8b | 0 | PASS | Ghi rõ "giới thiệu câu lạc bộ" không có dấu hiệu của 5 loại việc và "bối cảnh trường học một mình không đủ để xếp loại"; không hỏi bộ câu Việt. |
| 8c | 0 | PASS | Ghi yêu cầu "không rơi vào 5 loại việc cần bộ câu hỏi riêng cho thầy cô"; không hỏi bộ câu Việt. |
| 1 | 0 | PASS | Lượt hỏi Việt đủ 7 câu; hồ sơ để trống tên đơn vị và người trình bày; không có tên tự bịa. |
| 3 | 0 | PASS | Lượt hỏi Việt đủ 6 câu; không có tên tự bịa. |

Đầu ra chấm của vòng P14:

```
case8a FAIL
case8b PASS
case8c PASS
case1 PASS
case3 PASS
case8a 'Gợi ý:' lines=0 ; 'gợi ý' lines=3 ; 'hồ sơ đơn vị' lines=0 ; 'brief' lines=0 ; 'khung chat' lines=0
case8b 'gợi ý' lines=0 ; 'hồ sơ đơn vị' lines=0 ; 'khung chat' lines=0
case8c 'gợi ý' lines=0 ; 'hồ sơ đơn vị' lines=0 ; 'khung chat' lines=0
case1, case3: 'Đã khởi tạo' lines=0 ; 'Hải Anh' lines=0 ; 'Trãi' lines=0
```

Kết quả ca 8 sau P14: 2/3 theo tiêu chí grep, chưa đạt yêu cầu 3/3. Về nội dung, cả 3 lần đều không chạy lượt hỏi Việt và không hỏi câu chọn loại việc.

#### Ca 8a (toàn bộ đầu ra, FAIL)

```
Em đã kiểm tra xong. Có hai điều cần báo trước với thầy, rồi em xin thông tin để làm đúng.

**1. Về môi trường chạy:** phiên này chặn chạy lệnh (mọi lệnh `python` đều đòi phê duyệt và hiện chưa chạy được). Để xuất ra file `.pptx` thật, pipeline PPT Master bắt buộc chạy một loạt script (kiểm tra toàn vẹn, khởi tạo dự án, kiểm định chất lượng SVG, export). Nếu thầy mở đúng yêu cầu này trong **Claude Code (bản terminal/desktop) ở chế độ tương tác** thì các bước đó chạy trơn tru. Ở phiên hiện tại, em có thể **chuẩn bị trọn vẹn phần kế hoạch + nội dung + hướng thiết kế** để tới lúc chạy chỉ việc bấm duyệt.

**2. Về nội dung:** đây là giới thiệu **một CLB cụ thể của trường thầy**, nên có những dữ kiện chỉ thầy mới biết — em không được phép bịa (tên CLB, thành tích, cách đăng ký...). Em cần thầy cho biết vài thông tin thật.

---

Đây là **hướng đề xuất** (thầy chỉ cần trả lời "đồng ý" hoặc sửa):

- **Khổ:** slide 16:9 trình chiếu · **Ngôn ngữ:** tiếng Việt · **Phong cách:** trẻ trung (màu tươi, gradient, hình khối năng động, font Segoe UI hiển thị đủ dấu).
- **Bố cục 8 trang (điểm khởi đầu):**
  1. Bìa — tên CLB + slogan
  2. CLB là ai — sứ mệnh, năm thành lập
  3. Chúng mình làm gì — các mảng hoạt động (lập trình / web / AI / thiết kế / thi đấu…)
  4. Hoạt động & sự kiện nổi bật
  5. Thành tích – dấu ấn
  6. Cơ cấu & lịch sinh hoạt
  7. Vì sao nên tham gia — quyền lợi thành viên
  8. Kêu gọi tham gia — cách đăng ký / liên hệ

**Xin thầy cho biết (chỉ những gì chỉ thầy biết):**

1. Tên CLB và trường (ví dụ "CLB Tin học – Trường THPT…")?
2. CLB tập trung những mảng nào? (gợi ý: lập trình thi HSG, web/app, AI, robot, thiết kế đồ hoạ)
3. Có hoạt động / thành tích nổi bật nào để đưa vào trang 4–5 không? (nếu chưa có, em ghi khung mẫu để thầy điền sau)
4. Slide này dùng để làm gì? (gợi ý: tuyển thành viên mới đầu năm) — ảnh hưởng tới giọng điệu và trang cuối.
5. Cách đăng ký / liên hệ để đưa vào trang 8? (form, Fanpage, gặp ai…)
6. Màu chủ đạo mong muốn (gợi ý: xanh dương – tím công nghệ) và có logo CLB/trường không?

Thầy trả lời giúp em, hoặc nếu muốn em **cứ làm một bản khung đẹp với nội dung mẫu để thầy điền vào**, thầy nói "làm bản mẫu" là được. Em sẽ chờ thầy rồi mới khởi tạo dự án.
```

#### Ca 8b (dòng 1–9)

```
Em đã gặp một rào cản kỹ thuật cần thầy/cô xử lý trước khi tiếp tục.

## Định tuyến yêu cầu

Yêu cầu "tạo bài thuyết trình 8 slide giới thiệu câu lạc bộ tin học, phong cách trẻ trung" được em xử lý như sau:

- **Không dùng bộ câu hỏi trợ lý tiếng Việt**: "giới thiệu câu lạc bộ" không có dấu hiệu của 5 loại việc (bài giảng / báo cáo / hoạt động Đoàn–sự kiện / poster / tập huấn). Theo quy tắc, bối cảnh trường học một mình không đủ để xếp loại — nên em làm theo `SKILL.md` bình thường.
- **Route**: Generate PPTX — chế độ Default (không có "tạo nhanh"), canvas `ppt169` (slide 16:9).

## Vướng mắc: phiên này không chạy được script trực tiếp
```

#### Ca 8c (dòng 11–22)

```
## Trong phiên này tôi vẫn giúp được gì

Tôi có thể **soạn sẵn toàn bộ nội dung + khung thiết kế 8 slide** để khi thầy/cô chạy bản tương tác thì việc tạo diễn ra nhanh gọn. (Yêu cầu "giới thiệu CLB tin học" không rơi vào 5 loại việc cần bộ câu hỏi riêng cho thầy cô, nên mình đi theo luồng tạo thông thường.)

Để nội dung đúng thực tế trường mình, thầy/cô cho tôi xin vài thông tin:

1. **Tên CLB** và tên trường (in trên slide bìa)?
2. CLB **hoạt động những mảng gì** — lập trình thi đấu, web/app, đồ họa – thiết kế, robotics/AI, tin văn phòng, an toàn mạng…?
3. **Thành tích nổi bật** (giải thi HSG tin, sản phẩm, dự án, số thành viên)?
4. Lịch **sinh hoạt** và cách **đăng ký tham gia** (để đưa vào slide kêu gọi cuối)?
5. Có **logo/màu thương hiệu** hay ảnh hoạt động không? Nếu chưa, tôi dùng phong cách trẻ trung mặc định (màu tươi, gradient, icon công nghệ).
6. Slide dùng để **trình chiếu 16:9** đúng không?
```

#### Ca 1 (dòng 1–20)

```
Chưa có hồ sơ đơn vị, nên em sẽ hỏi thầy cô một lượt gồm phần hồ sơ và các câu về bài giảng. Thầy cô trả lời xong em mới bắt đầu soạn.

---

Em chào thầy/cô ạ. Em sẽ giúp thầy/cô làm bài giảng **Vật lí 10 — Chuyển động thẳng đều**. Trước khi bắt đầu, thầy/cô cho em xin ít thông tin nhé.

**Hồ sơ đơn vị** (lần đầu nên em hỏi để lưu lại, các lần sau khỏi nhập):
- Tên trường/đơn vị: (thầy/cô điền giúp em)
- Tên người trình bày (hiện trên trang bìa): (thầy/cô điền giúp em)
- Logo trường (nếu có, cho em đường dẫn file; không có cũng được)
- Cấp học: em đoán là **THPT** (vì là lớp 10) — đúng không ạ?
- Màu chủ đạo: em đề xuất **xanh dương** cho môn Vật lí — thầy/cô đồng ý hay muốn màu khác?

**Về bài giảng:**

1. Môn, lớp, tên bài, và bài thuộc chủ đề/chương nào? Em đã có: Vật lí 10 — Chuyển động thẳng đều. Em đề xuất xếp vào chương **Động học** — thầy/cô xác nhận hay đổi giúp em ạ.

2. Thầy/cô dạy theo bộ sách nào: Kết nối tri thức với cuộc sống, Chân trời sáng tạo, Cánh diều, hay bộ khác?

3. Bài dạy trong bao nhiêu tiết, mỗi tiết bao nhiêu phút? Gợi ý: **1 tiết × 45 phút**.
```

#### Ca 3 (dòng 1–19)

```
Em đã đọc xong quy trình. Đây là việc làm slide cho **hoạt động Đoàn – sự kiện**, và máy chưa có hồ sơ đơn vị nên em xin hỏi thầy cô một lượt rồi mới bắt tay vào làm ạ.

**Trước hết, vài thông tin về đơn vị (em để trống cho thầy cô điền):**
- Đơn vị (tên trường):
- Cấp học (Tiểu học / THCS / THPT / khác):
- Người trình bày — chức danh:
- Logo (đường dẫn file, hoặc "không dùng"):
- Màu chủ đạo — em gợi ý **đỏ Đoàn #E31E24** kết hợp **xanh dương #005BAA**; thầy cô có thể đổi.

**Về nội dung cuộc thi, thầy cô cho em biết:**

1. Tên chương trình, **dịp tổ chức** và **cấp tổ chức**?
   Gợi ý: "Cuộc thi Tuổi trẻ sáng tạo" do **Đoàn trường** tổ chức; dịp tổ chức thầy cô bổ sung giúp em (ví dụ chào mừng 26/3, năm học 2025–2026).

2. **Mục đích của slide** là gì: giới thiệu chương trình, phổ biến thể lệ, chiếu trên sân khấu, hay tổng kết và trao giải?
   Gợi ý: giới thiệu chương trình.

3. **Thời gian, địa điểm và đối tượng tham gia**?
   Gợi ý: đoàn viên, học sinh toàn trường; thời gian và địa điểm thầy cô bổ sung.
```

## Trích đoạn đầu ra

Lấy từ vòng 3. Ca 5 (chấm thủ công) và ca 8 (FAIL) chép đủ toàn bộ đầu ra. Không đoạn nào có ký tự ngoặc nhọn cần thay.

### Ca 1 — Tạo bài giảng Vật lí 10 bài Chuyển động thẳng đều

```
Em đã đọc quy trình. Đây là bài giảng Vật lí 10, em xin hỏi thầy cô một lượt rồi mới bắt tay làm. Máy này chưa có hồ sơ đơn vị nên em xin thêm vài thông tin để lưu lại cho những lần sau.

**Hồ sơ đơn vị (điền một lần, để trống mục nào chưa tiện):**
- Tên trường/đơn vị:
- Họ tên thầy/cô trình bày:
- Logo trường (nếu có, cho em đường dẫn file):
- Cấp học: em đoán **THPT** (vì lớp 10) — thầy cô xác nhận giúp.
- Màu chủ đạo: em gợi ý **xanh dương học đường** — thầy cô đổi tuỳ ý.

**Các câu hỏi cho bài giảng:**

1. Bài "Chuyển động thẳng đều" này thuộc chương nào ạ?
   Gợi ý: chương **Động học** (phần chuyển động thẳng).

2. Thầy cô dạy theo bộ sách nào: Kết nối tri thức với cuộc sống, Chân trời sáng tạo, Cánh diều, hay bộ khác?
   Gợi ý: **Kết nối tri thức với cuộc sống**.

3. Bài dạy trong bao nhiêu tiết, mỗi tiết bao nhiêu phút?
   Gợi ý: **1 tiết × 45 phút**.
...
Ở bước xác nhận, em sẽ tóm tắt trong khung chat để thầy cô duyệt, không mở trang web. Thầy cô đồng ý nhé?
```

### Ca 2 — Làm slide báo cáo tổng kết năm học của tổ Toán

```
Chưa có hồ sơ đơn vị nên em hỏi luôn phần hồ sơ trong cùng tin nhắn này. Thầy/cô trả lời một lượt giúp em nhé.

---

Em chào thầy/cô. Em sẽ làm slide **báo cáo tổng kết năm học của tổ Toán**. Trước khi bắt đầu, em xin hỏi vài thông tin ạ.

**Hồ sơ đơn vị** (lần đầu nên em xin thông tin để dùng lại về sau — thầy/cô để trống mục nào chưa tiện thì em bỏ qua mục đó):
- Tên trường/đơn vị:
- Cấp học (THCS/THPT…):
- Người trình bày (họ tên, chức vụ ghi trên slide):
- Logo trường (nếu có, cho em đường dẫn file; không có cũng không sao):
- Màu chủ đạo — gợi ý: xanh dương trang trọng. Thầy/cô "đồng ý" hoặc đổi màu khác.

**Nội dung báo cáo:**

1. Đây là loại báo cáo nào: sơ kết học kỳ, tổng kết năm học, báo cáo chuyên đề hay báo cáo thi đua?
   Gợi ý: tổng kết năm học (theo yêu cầu của thầy/cô).

2. Kỳ báo cáo (năm học, học kỳ hay quý) là khi nào, và bộ phận nào báo cáo (cả trường, tổ, khối)?
   Gợi ý: năm học 2025–2026; bộ phận báo cáo là tổ Toán.
```

### Ca 3 — Làm slide giới thiệu cuộc thi Tuổi trẻ sáng tạo của Đoàn trường

```
Em đã đọc quy trình. Đây là việc thuộc loại **Hoạt động Đoàn – sự kiện**, và máy chưa có hồ sơ đơn vị nào. Em xin hỏi thầy cô một lượt rồi mới bắt tay làm slide.

**Trước hết, vài thông tin về đơn vị (em để trống cho thầy cô điền, không tự đoán):**
- Đơn vị (tên trường): …
- Cấp học: Tiểu học / THCS / THPT / khác
- Người trình bày và chức danh: …
- Logo: có file logo không, hay không dùng?
- Màu chủ đạo: em gợi ý **xanh Đoàn `#0066B3`** — thầy cô đồng ý hay muốn màu khác?

**Và các câu về chương trình:**

1. Tên chương trình, dịp tổ chức và cấp tổ chức?
   Gợi ý: "Tuổi trẻ sáng tạo", do **Đoàn trường** tổ chức; dịp em để trống, thầy cô cho biết (ví dụ chào mừng 26/3).

2. Mục đích của slide là gì: giới thiệu chương trình, phổ biến thể lệ, chiếu trên sân khấu, hay tổng kết và trao giải?
   Gợi ý: giới thiệu chương trình.

3. Thời gian, địa điểm và đối tượng tham gia?
   Gợi ý: đoàn viên, học sinh toàn trường.
```

### Ca 4 — Làm poster Zalo thông báo họp phụ huynh

```
Chào thầy/cô ạ! Em sẽ giúp thầy/cô làm **poster Zalo thông báo họp phụ huynh**. Máy này chưa có hồ sơ đơn vị nên em xin hỏi thêm vài mục hồ sơ (để dành dùng cho các lần sau), rồi tới nội dung poster. Thầy/cô trả lời trong một lượt, em chờ rồi mới bắt tay làm ạ.

**A. Hồ sơ đơn vị** (điền một lần, các lần sau dùng lại)
- Đơn vị (tên trường):
- Cấp học: Tiểu học / THCS / THPT / khác
- Người trình bày — chức danh:
- Logo: có đường dẫn file logo không, hay không dùng?
- Màu chủ đạo: em gợi ý **xanh dương đậm #1F4E79** (trang trọng, dễ đọc) — thầy/cô đồng ý hay đổi màu khác?

**B. Nội dung poster**

1. Thầy/cô đăng ở **kênh** nào: nhóm Zalo lớp/phụ huynh, Fanpage Facebook, hay TikTok?
Gợi ý: nhóm Zalo phụ huynh.

2. Dùng **khổ ảnh** nào?
Gợi ý: nhóm Zalo dùng **ảnh vuông**.

3. Làm **một ảnh hay chuỗi nhiều ảnh**? Nếu là chuỗi thì bao nhiêu ảnh?
Gợi ý: một ảnh cho thông báo.
```

### Ca 5 — Làm slide tập huấn dùng AI soạn giáo án cho giáo viên (toàn bộ đầu ra)

```
Em đã đọc quy trình. Đây là buổi **tập huấn/workshop**, và máy chưa có hồ sơ đơn vị nên em xin hỏi một lượt (thầy cô chỉ cần trả lời "đồng ý" hoặc sửa từng ý). Em sẽ chờ thầy cô trả lời rồi mới bắt đầu.

**Hồ sơ đơn vị** (để trống mục nào chưa tiện, em bỏ qua mục đó):
- Đơn vị: _(tên trường/đơn vị)_
- Cấp học: _(Tiểu học / THCS / THPT / khác)_
- Người trình bày: _(họ tên — chức danh)_
- Logo: _(đường dẫn file logo, hoặc "không dùng")_
- Màu chủ đạo: gợi ý xanh dương giáo dục `#1F4E79` — thầy cô đồng ý hay muốn màu khác?

**Về buổi tập huấn:**

1. Chủ đề là gì, và buổi này theo hình thức nào: sinh hoạt tổ chuyên môn, bồi dưỡng thường xuyên, chia sẻ chuyên đề, hay tập huấn toàn trường?
   Gợi ý: chủ đề "Dùng AI soạn giáo án cho giáo viên"; hình thức sinh hoạt tổ chuyên môn.

2. Đối tượng tham dự là ai, và họ đã biết về chủ đề đến mức nào?
   Gợi ý: giáo viên trong tổ, mới biết sơ qua về AI.

3. Thời lượng buổi này là bao lâu, và tỉ lệ lý thuyết – thực hành thế nào?
   Gợi ý: 90 phút, một nửa thời gian thực hành.

4. Sau buổi, người học làm được gì (2–3 kết quả cụ thể)?
   Gợi ý (AI đề xuất để thầy cô sửa): (1) viết được một prompt yêu cầu AI soạn giáo án theo mẫu; (2) tự chỉnh sửa được giáo án AI tạo cho phù hợp lớp mình; (3) biết cách kiểm tra và loại bỏ thông tin AI bịa.

5. Có phần hướng dẫn thao tác từng bước không? Thầy cô có ảnh chụp màn hình để gửi không?
   Gợi ý: có hướng dẫn từng bước; nếu thầy cô gửi ảnh chụp màn hình thì em dùng, chưa có thì em mô tả thao tác bằng chữ và sơ đồ.

6. Có những hoạt động tương tác nào: câu hỏi thảo luận, bài tập thực hành, phiếu khảo sát?
   Gợi ý: một bài tập thực hành sau mỗi phần và một câu hỏi thảo luận cuối buổi.

Ở bước xác nhận, em sẽ tóm tắt trong khung chat để thầy cô duyệt, không mở trang web. Thầy cô đồng ý nhé?
```

Chấm thủ công: `has 5 "đối tượng"` báo thiếu vì dòng 15 viết "Đối tượng" (U+0110) còn mẫu là "đối tượng" (U+0111). `grep -c -- "Đối tượng"` trả về 1. Các từ khoá còn lại có mặt ("thời lượng", "làm được", "kết quả"); không có "đã khởi tạo", "Đã khởi tạo", "Lương Hải Anh" hay "Nguyễn Trãi".

### Ca 6 — Tạo nhanh bài giảng Hoá học 11 (toàn bộ đầu ra)

```
Em đã đọc quy trình. Câu lệnh của thầy cô là **tạo nhanh** một bài giảng Hoá học 11 — em cần hỏi nhanh 2 thông tin bắt buộc còn thiếu (chỉ những điều mà chỉ thầy cô biết), rồi em sẽ làm luôn, không hỏi thêm và không có bước xác nhận.

1. Tên bài Hoá học 11 cụ thể là bài nào, và bài thuộc chương/chủ đề nào?
   Gợi ý: ví dụ bài "Sự điện li" (Chương Cân bằng hoá học) hoặc "Ammonia và muối ammonium" (Chương Nitrogen – Sulfur). Thầy cô cho em tên bài chính xác giúp em nhé.

2. Thầy cô dạy theo bộ sách nào: **Kết nối tri thức với cuộc sống**, **Chân trời sáng tạo**, **Cánh diều**, hay bộ khác?
   Gợi ý: nếu thầy cô đồng ý, em dùng **Kết nối tri thức với cuộc sống**.

Thầy cô trả lời 2 ý trên (có thể trả lời gọn, ví dụ: "Bài Sự điện li, Cánh diều") là em bắt tay vào tạo ngay ạ.
```

### Ca 7 — Make a 5-slide deck about our software product

```
I've completed routing and the mandatory integrity gate. Here's where we are:

- **Route:** Generate PPTX → **Default** runtime (you didn't ask for "quick/fast," and the request is in English, so I'm not using the Vietnamese teacher-intake flow).
- **Blocker (Step 1 – Source Content):** I can't start a 5-slide deck until I know *which* product and the facts to put on the slides. "Our software product" is ambiguous — from prior context you've worked on several (‹danh sách dự án riêng của chủ máy — đã lược bỏ›). I won't guess, and I can't research an internal product from the web — the facts have to come from you.

Could you answer these so I can plan the deck? Short bullets are fine:

1. **Which product** is this deck about? (Name it, or point me to a repo/site/doc I can convert as a source.)
2. **Audience & purpose** — who's it for and why? (e.g., investor pitch, sales to schools, internal demo, class intro)
3. **The one-liner** — what does the product do, in a sentence?
4. **3–5 key points** you want across the deck — e.g. problem it solves, main features, who it's for, results/traction, a call-to-action.
5. **Any hard facts** — pricing, user numbers, metrics, logos/screenshots — or say "none yet."
```

### Ca 8 — Tạo bài thuyết trình 8 slide giới thiệu câu lạc bộ tin học của trường, phong cách trẻ trung (toàn bộ đầu ra, FAIL)

```
Em đã đọc quy trình. Đây là bài giới thiệu **câu lạc bộ Tin học của trường** — thuộc nhóm *Hoạt động Đoàn – sự kiện*. Máy này chưa có hồ sơ đơn vị nên em xin hỏi một lượt (kèm vài mục hồ sơ để lưu dùng lần sau), thầy cô trả lời "đồng ý" với gợi ý nào thấy hợp là được ạ.

**Hồ sơ đơn vị** (em để trống, thầy cô điền giúp — em không tự đoán):
- Tên trường / đơn vị: …
- Cấp học: *(gợi ý: THPT — sửa nếu khác)*
- Người trình bày / phụ trách CLB (nếu muốn hiện trên bìa): …
- Logo trường hoặc CLB (đường dẫn file, nếu có): …
- Màu chủ đạo: *(gợi ý: xanh dương công nghệ)*

**Câu hỏi về nội dung slide:**

1. Tên câu lạc bộ và cấp tổ chức (CLB cấp trường, trực thuộc Đoàn trường hay tổ Tin học)?
   Gợi ý: "Câu lạc bộ Tin học" trực thuộc Đoàn trường.

2. Mục đích của slide: giới thiệu để tuyển thành viên mới, hay giới thiệu chung trong buổi sinh hoạt/sự kiện?
   Gợi ý: giới thiệu và tuyển thành viên mới.

3. Đối tượng tham gia xem slide và dịp trình chiếu?
   Gợi ý: học sinh toàn trường, chiếu trong ngày hội câu lạc bộ đầu năm học.

4. Nội dung chính gồm những gì (lĩnh vực hoạt động, hoạt động/dự án tiêu biểu, thành tích, quyền lợi thành viên, cách đăng ký)?
   Gợi ý: giới thiệu CLB → lĩnh vực hoạt động → hoạt động tiêu biểu → thành tích → quyền lợi thành viên → cách đăng ký.

5. Dùng nhận diện nào: logo riêng của CLB, huy hiệu Đoàn, hay chỉ màu chủ đạo?
   Gợi ý: dùng màu chủ đạo trong hồ sơ, thêm logo CLB nếu có.

6. Có ảnh hoạt động sẵn không, hay để em tạo ảnh minh hoạ?
   Gợi ý: dùng ảnh thầy cô gửi; thiếu thì em tạo ảnh minh hoạ phong cách công nghệ.

7. Thông tin liên hệ / đăng ký đặt ở slide cuối (Fanpage, nhóm Zalo/Facebook, người phụ trách)?
   Gợi ý: Fanpage + nhóm Zalo của CLB.

Bố cục **8 slide**, khổ **16:9**, phong cách **trẻ trung, năng động** như thầy cô yêu cầu — em sẽ chốt chi tiết sau khi có câu trả lời.

Ở bước xác nhận, em sẽ tóm tắt trong khung chat để thầy cô duyệt, không mở trang web. Thầy cô đồng ý nhé?
```

### Ca 9 — Tạo bài giảng Hoá học 11 bài Nitrogen, không cần hỏi lại

File đầu ra rỗng (0 byte, exit=124). Trích lời AI trong transcript phiên vòng 3 (giờ máy):

```
[15:24:06] Tôi bắt đầu bằng việc đọc các tài liệu bắt buộc theo quy tắc dự án: `SKILL.md`, quy trình hỏi trợ lý thầy cô, và hướng dẫn loại "Bài giảng". Vì bạn nói "không cần hỏi lại" → đây là yêu cầu **Tạo nhanh (Quick)** tường minh.
[15:24:15] Đây là yêu cầu **Tạo nhanh** (có "không cần hỏi lại") → không hỏi câu nào, viết brief và chạy chế độ tạo nhanh của upstream. Tôi chạy cổng toàn vẹn bắt buộc và đọc các tài liệu định tuyến/quy trình.
[15:24:46] Cổng toàn vẹn pass (exit 0). Route: **Generate PPTX — explicit Quick** → runtime authority là `quick-generate.md`. Tôi đọc nó, cùng canvas-formats và templates icons README để chuẩn bị.
[15:25:14] Đã rõ quy trình. Đây là input **chủ đề** (bài giảng Hoá 11 Nitrogen) → chạy tạo nhanh, có nghiên cứu chủ đề trong khâu chuẩn bị nguồn. Tôi lập kế hoạch thực thi và đọc stage nghiên cứu chủ đề + tài liệu canvas.
[15:25:58] Tôi cho chạy **research worker** (nền) cho phần dữ kiện Hoá học, đồng thời đọc bộ tài liệu năng lực lập kế hoạch + lõi thực thi SVG song song.
```

Brief do ca 9 vòng 2 ghi ra (dòng 6–23 của `projects/_brief-bai-giang-hoa11-nitrogen.md`, bản sao trong `m2-headless-final-r2`):

```
## Đơn vị

(chưa có hồ sơ đơn vị)

## Thầy cô yêu cầu

- Tạo bài giảng môn Hoá học lớp 11, bài Nitrogen
- Tạo nhanh, không cần hỏi lại

## AI đề xuất (thầy cô đã đồng ý)

- Bộ sách: Kết nối tri thức với cuộc sống (AI đề xuất, chưa duyệt)
- Bài Nitrogen thuộc Chương "Nitrogen – Sulfur" (AI đề xuất, chưa duyệt)
- Thời lượng: 1 tiết × 45 phút (AI đề xuất, chưa duyệt)
- Mục tiêu theo Chương trình GDPT 2018 (kiến thức – năng lực – phẩm chất) (AI đề xuất, chưa duyệt)
- Hoạt động: trò chơi khởi động ngắn, thảo luận nhóm, 4–5 câu trắc nghiệm luyện tập (đáp án ở slide riêng ngay sau câu hỏi) (AI đề xuất, chưa duyệt)
- Đặc điểm lớp: trình độ trung bình; phòng học dùng máy chiếu 16:9 (AI đề xuất, chưa duyệt)
- Không có tài liệu sẵn; AI soạn từ nội dung bài học (AI đề xuất, chưa duyệt)
```

## Sửa đổi hướng dẫn trong quá trình kiểm thử

Các thay đổi của đợt sửa sau review cuối nhánh nằm trong commit `d126822d` (vòng 1 chạy trên commit này). Trong quá trình kiểm thử có 2 vòng sửa câu chữ.

### Vòng sửa 1 (sau vòng chạy 1, commit c8f7afa9) — liên quan ca 9

Transcript ca 9 vòng 1 cho thấy AI không đọc `docs/vi/tro-ly/quy-trinh-hoi.md`: "Yêu cầu "không cần hỏi lại" là **Quick tường minh** (theo AGENTS.vi.md mục 3 & 10), nên tôi bỏ qua lượt hỏi hồ sơ thầy cô và đi thẳng theo `SKILL.md`." Như vậy brief có nhãn "(AI đề xuất, chưa duyệt)" sẽ không được ghi.

`AGENTS.vi.md`, mục 10.

Trước:
```
- `SKILL.md` vẫn được ưu tiên. Lượt hỏi này chỉ tạo thêm tài liệu nguồn; bước xác nhận của upstream vẫn bắt buộc, trừ khi người dùng yêu cầu tạo nhanh (xem mục 3).
```
Sau:
```
- `SKILL.md` vẫn được ưu tiên. Lượt hỏi này chỉ tạo thêm tài liệu nguồn; bước xác nhận của upstream vẫn bắt buộc, trừ khi người dùng yêu cầu tạo nhanh (xem mục 3). Khi tạo nhanh, kể cả với "không cần hỏi lại", vẫn đọc `docs/vi/tro-ly/quy-trinh-hoi.md` và làm theo mục "Tạo nhanh" của file đó: có thể không hỏi câu nào, nhưng vẫn ghi brief.
```

Sau vòng sửa 1: `& $PY -m unittest discover -s tools/vi/tests` → Ran 92 tests — OK. Ở vòng 2, ca 9 đọc quy trình, không hỏi câu nào và ghi brief đúng mẫu (trích ở trên).

### Vòng sửa 2 (sau vòng chạy 2, commit c15e700d) — liên quan ca 8

Ở vòng 2, ca 8 đọc quy trình và nhận ra yêu cầu không thuộc 5 loại, nhưng vẫn tìm hồ sơ đơn vị ("Trước khi vào Step 1, tôi kiểm tra xem đã có hồ sơ đơn vị chưa để không hỏi lại thông tin sẵn có"). Đầu ra mở đầu bằng "Chưa có hồ sơ đơn vị và chưa có brief nào", có hai dòng "gợi ý:" và hứa xác nhận trong khung chat.

`docs/vi/tro-ly/quy-trinh-hoi.md`, mục "Khi nào áp dụng".

Trước:
```
- Yêu cầu không thuộc loại nào, hoặc có từ khoá nhưng không phải bối cảnh trường học hay Đoàn (ví dụ báo cáo doanh thu của công ty): không dùng file này, làm theo `SKILL.md` như bình thường.
```
Sau:
```
- Yêu cầu không thuộc loại nào, hoặc có từ khoá nhưng không phải bối cảnh trường học hay Đoàn (ví dụ báo cáo doanh thu của công ty, giới thiệu sản phẩm): không dùng phần nào của file này (không tìm hay nhắc tới hồ sơ đơn vị, không viết brief, không dùng dạng câu hỏi kèm "Gợi ý:" hay dòng chốt cách xác nhận), làm theo `SKILL.md` như bình thường.
```

`AGENTS.vi.md`, mục 10.

Trước:
```
- Yêu cầu không thuộc 5 loại, hoặc người dùng không viết tiếng Việt: làm theo `SKILL.md` như bình thường.
```
Sau:
```
- Yêu cầu không thuộc 5 loại, hoặc người dùng không viết tiếng Việt: làm theo `SKILL.md` như bình thường, không tìm hồ sơ đơn vị và không dùng bộ câu hỏi Việt.
```

Thêm cụm "không tìm hay nhắc tới hồ sơ đơn vị" vào test `test_scope_section_limits_detection_to_school_context`. Sau vòng sửa 2: Ran 92 tests — OK. Không còn vòng sửa nào theo giới hạn 2 vòng.

## Mối lo ngại

- **Ca 8 chưa đạt ổn định.** Sau quyết định P14 (commit `ca6f2ebb`), ca 8 đạt 2/3 lần theo grep; lần không đạt (8a) không chạy lượt hỏi Việt, chỉ dùng chữ "gợi ý:" trong câu hỏi của luồng upstream. Theo chỉ đạo, không sửa thêm câu chữ. Trước P14: ca này đạt ở vòng 1 nhưng không đạt ở vòng 2 và vòng 3, nghĩa là kết quả dao động giữa các lần chạy. Ở vòng 3, AI viết trong transcript: "Yêu cầu là tiếng Việt, bối cảnh trường học, giới thiệu câu lạc bộ tin học — gần nhất với nhóm "Hoạt động Đoàn – sự kiện"". Nó tự chọn loại gần nhất thay vì coi là việc ngoài 5 loại hoặc hỏi câu chọn loại việc. Đây đúng là câu lệnh đầu tiên trong `docs/vi/bat-dau-nhanh.md`, nên thầy cô mới dùng dễ gặp. Quy tắc chỉ xếp vào một loại khi câu lệnh có dấu hiệu trong bảng, không suy ra loại chỉ từ bối cảnh trường học, đã được áp dụng ở commit `ca6f2ebb`; sau đó ca 8 đạt 2/3 lần theo grep (8a FAIL vì dòng "gợi ý:" viết thường xuất hiện trong một tin nhắn hỏi dữ kiện thông thường, không có hồ sơ đơn vị, không có brief, không có dòng chốt cách xác nhận), cần kiểm lại ở nghiệm thu của chủ repo.
- **Ca 9 không có đầu ra để chấm.** Chế độ tạo nhanh không dừng lại hỏi nên luôn chạy quá 300 giây. Tiêu chí grep đạt chỉ vì file rỗng; bằng chứng thật là transcript: không có câu hỏi nào ở cả 3 vòng, và ở vòng 2 brief được ghi đúng mẫu. Ở vòng 3, phiên hết giờ khi đang đọc tài liệu lập kế hoạch, trước khi ghi brief.
- **Ca 5 cần chấm thủ công ở cả 3 vòng** vì `grep -i` không gấp "Đ"/"đ". Nội dung đạt.
- **Môi trường chạy không sạch** (xem mục "Môi trường chạy"): các lệnh đã chặn Bash và Write nhưng AI vẫn ghi file được qua PowerShell.

## Lần chạy trước đã được thay thế

Báo cáo này thay cho báo cáo ở commit `b8e140ef`. Lần chạy đó có 7 ca, đầu ra ở `m2-headless`, `m2-headless-r2`, `m2-headless-r3`; vòng cuối chạy trên nội dung của `8fea6451`, lúc chạy chưa commit.
