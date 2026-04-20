# PPT Master — Hướng dẫn sử dụng

[![Version](https://img.shields.io/badge/version-v2.2.0-blue.svg)]()
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Đưa vào **PDF, DOCX, URL hoặc Markdown** — AI tự động tạo ra **bài thuyết trình đẹp mắt, chỉnh sửa được trong PowerPoint**.

---

## Mục lục

- [Cài đặt](#-cài-đặt)
- [Cách sử dụng](#-cách-sử-dụng)
- [Mẫu có sẵn](#-mẫu-có-sẵn)
- [Lệnh thường dùng](#-lệnh-thường-dùng)
- [Cấu hình nâng cao](#-cấu-hình-nâng-cao)
- [Chỉnh sửa sau xuất](#-chỉnh-sửa-ppt-sau-khi-xuất)
- [Câu hỏi thường gặp](#-câu-hỏi-thường-gặp)

---

## 📦 Cài đặt

### 1. Yêu cầu hệ thống

| Thành phần | Yêu cầu | Ghi chú |
|------------|----------|---------|
| **Python** | 3.8 trở lên | Bắt buộc |
| **Pandoc** | Bất kỳ | Chỉ cần khi chuyển đổi DOCX/EPUB/LaTeX |
| **Node.js** | 18 trở lên | Chỉ cần khi dùng `web_to_md.cjs` cho trang bảo mật cao |

### 2. Cài đặt thư viện Python

```bash
pip install -r requirements.txt
```

> Nếu gặp lỗi quyền, dùng `pip install --user -r requirements.txt` hoặc tạo môi trường ảo.

### 3. Mở trình soạn thảo AI

PPT Master hoạt động thông qua AI agent trong trình soạn thảo. Các công cụ được hỗ trợ:

| Công cụ | Mô tả |
|---------|-------|
| [Claude Code](https://claude.ai/) | CLI chính thức của Anthropic — **khuyến nghị** |
| [Cursor](https://cursor.sh/) | Trình soạn thảo AI phổ biến |
| [VS Code + Copilot](https://code.visualstudio.com/) | Giải pháp Microsoft |
| [Antigravity](https://antigravity.dev/) | Miễn phí, hạn ngạch giới hạn |

Mở thư mục dự án PPT Master trong trình soạn thảo, sau đó sử dụng chat AI để bắt đầu.

---

## 🚀 Cách sử dụng

### Bước 1 — Chuẩn bị tài liệu nguồn

Đưa tài liệu cho AI qua cửa sổ chat. Các định dạng được hỗ trợ:

| Đầu vào | Cách dùng |
|----------|-----------|
| **PDF** | Kéo thả file vào chat hoặc chỉ đường dẫn |
| **DOCX / EPUB / LaTeX** | Kéo thả file vào chat (cần cài Pandoc) |
| **URL trang web** | Dán link vào chat |
| **Markdown** | Dán nội dung hoặc chỉ đường dẫn file |
| **Văn bản tự do** | Gõ trực tiếp nội dung bạn muốn trình bày |

### Bước 2 — Yêu cầu AI tạo PPT

Mô tả yêu cầu bằng ngôn ngữ tự nhiên:

```
Tôi có báo cáo quý 1 năm 2026 (đính kèm file PDF). Hãy tạo PPT trình bày cho ban giám đốc.
```

```
Tạo bài thuyết trình 10 slide về chiến lược marketing năm 2026, phong cách consulting cao cấp.
```

```
Chuyển file Word này thành slide thuyết trình cho hội nghị khoa học.
```

### Bước 3 — Xác nhận với AI

AI sẽ hỏi bạn **8 câu hỏi** để xác nhận thiết kế:

1. **Định dạng canvas** — PPT 16:9, 4:3, hay định dạng mạng xã hội?
2. **Số trang** — Bao nhiêu slide?
3. **Đối tượng** — Ai sẽ xem bài thuyết trình?
4. **Phong cách** — Doanh nghiệp, học thuật, sáng tạo...?
5. **Bảng màu** — Tông màu chủ đạo?
6. **Icon** — Có dùng icon không?
7. **Typography** — Kiểu chữ?
8. **Hình ảnh** — Dùng ảnh có sẵn hay AI tạo ảnh?

> 💡 AI sẽ **đề xuất sẵn** cho từng mục dựa trên nội dung của bạn. Bạn chỉ cần xác nhận hoặc điều chỉnh.

### Bước 4 — Nhận kết quả

Sau khi xác nhận, AI sẽ tự động:
- Tạo tất cả slide (SVG) → Tạo ghi chú thuyết trình → Hậu xử lý → Xuất file PPTX

File PPTX xuất ra nằm trong thư mục `projects/<tên_dự_án>/`.

---

## 🎨 Mẫu có sẵn

PPT Master có **20+ mẫu thiết kế** sẵn sàng sử dụng. Khi khởi tạo, AI sẽ hỏi bạn muốn dùng mẫu hay thiết kế tự do.

| Mẫu | Phong cách |
|------|-----------|
| `mckinsey` | McKinsey — Tư vấn chiến lược |
| `google_style` | Google — Hiện đại, tối giản |
| `anthropic` | Anthropic — Công nghệ AI |
| `academic_defense` | Bảo vệ luận văn — Học thuật |
| `medical_university` | Y khoa — Nghiên cứu y học |
| `government_blue` | Chính phủ — Tông xanh |
| `government_red` | Chính phủ — Tông đỏ |
| `smart_red` | Doanh nghiệp — Đỏ thông minh |
| `pixel_retro` | Retro — Pixel art |
| `exhibit` | Triển lãm — Showroom |
| `psychology_attachment` | Tâm lý học — Gắn bó |
| `ai_ops` | AI Operations — Công nghệ |
| *...và nhiều mẫu khác* | |

> 💡 Bạn cũng có thể chọn **thiết kế tự do** — AI sẽ tạo phong cách riêng phù hợp với nội dung.

---

## 🛠️ Lệnh thường dùng

Thông thường AI sẽ **tự động chạy** các lệnh này. Tuy nhiên, bạn có thể chạy thủ công nếu cần.

### Chuyển đổi tài liệu

```bash
# PDF sang Markdown
python3 skills/ppt-master/scripts/pdf_to_md.py <file_PDF>

# DOCX/EPUB/LaTeX sang Markdown (cần Pandoc)
python3 skills/ppt-master/scripts/doc_to_md.py <file>

# Trang web sang Markdown
python3 skills/ppt-master/scripts/web_to_md.py <URL>
```

### Quản lý dự án

```bash
# Khởi tạo dự án mới
python3 skills/ppt-master/scripts/project_manager.py init <tên_dự_án> --format ppt169

# Nhập file nguồn vào dự án
python3 skills/ppt-master/scripts/project_manager.py import-sources <đường_dẫn> <file_nguồn...> --move

# Kiểm tra cấu trúc dự án
python3 skills/ppt-master/scripts/project_manager.py validate <đường_dẫn>
```

### Hậu xử lý & Xuất PPTX

> ⚠️ Ba lệnh dưới đây **phải chạy theo thứ tự, từng lệnh một**. Xác nhận lệnh trước thành công rồi mới chạy lệnh tiếp.

```bash
# Bước 1: Tách ghi chú thuyết trình
python3 skills/ppt-master/scripts/total_md_split.py <đường_dẫn_dự_án>

# Bước 2: Xử lý SVG (nhúng icon, ảnh, font...)
python3 skills/ppt-master/scripts/finalize_svg.py <đường_dẫn_dự_án>

# Bước 3: Xuất file PPTX
python3 skills/ppt-master/scripts/svg_to_pptx.py <đường_dẫn_dự_án> -s final
```

### Công cụ hình ảnh

```bash
# Phân tích ảnh có sẵn trong dự án
python3 skills/ppt-master/scripts/analyze_images.py <đường_dẫn>/images

# Tạo ảnh bằng AI (cần cấu hình Gemini API)
python3 skills/ppt-master/scripts/nano_banana_gen.py "mô tả ảnh" --aspect_ratio 16:9 --image_size 1K -o <đường_dẫn>/images

# Kiểm tra chất lượng SVG
python3 skills/ppt-master/scripts/svg_quality_checker.py <đường_dẫn>
```

---

## ⚙️ Cấu hình nâng cao

### Định dạng canvas

| Định dạng | viewBox | Dùng cho |
|-----------|---------|----------|
| `ppt169` | `0 0 1280 720` | Slide thuyết trình 16:9 (**mặc định**) |
| `ppt43` | `0 0 1024 768` | Slide thuyết trình 4:3 |
| `xhs` | `0 0 1242 1660` | Facebook / TikTok (3:4) |
| `square` | `0 0 1080 1080` | Zalo / Instagram (1:1) |
| `story` | `0 0 1080 1920` | Story (9:16) |

### Tạo ảnh bằng AI (Gemini API)

Nếu muốn AI tự động tạo ảnh minh họa, cấu hình API key trước:

```bash
# Lấy API key tại https://aistudio.google.com/apikey
export GEMINI_API_KEY="your-api-key"

# Tùy chọn: Endpoint proxy
export GEMINI_BASE_URL="https://your-proxy-url.com/v1beta"
```

---

## ✏️ Chỉnh sửa PPT sau khi xuất

File PPTX xuất ra chứa các trang ở định dạng SVG. Để chỉnh sửa nội dung trong PowerPoint:

1. Mở file PPTX trong PowerPoint (yêu cầu **Office 2016** trở lên)
2. Chọn nội dung trên slide
3. Nhấp chuột phải → **"Chuyển đổi thành Hình dạng"** (Convert to Shape)
4. Giờ bạn có thể tự do chỉnh sửa văn bản, hình dạng, màu sắc

---

## ❓ Câu hỏi thường gặp

<details>
<summary><b>AI mất ngữ cảnh giữa chừng, phải làm sao?</b></summary>

Yêu cầu AI đọc lại file hướng dẫn:
```
Hãy đọc skills/ppt-master/SKILL.md rồi tiếp tục tạo PPT
```
Hoặc dùng `AGENTS.md` làm điểm khởi đầu tổng quan.

</details>

<details>
<summary><b>Ba loại Executor khác nhau thế nào?</b></summary>

- **Executor_General** — Phong cách tổng quát, bố cục linh hoạt
- **Executor_Consultant** — Phong cách tư vấn, trực quan hóa dữ liệu
- **Executor_Consultant_Top** — Phong cách tư vấn cấp cao (MBB), kỹ thuật chuyên sâu

AI sẽ tự động chọn Executor phù hợp dựa trên nội dung và phong cách bạn đã xác nhận.

</details>

<details>
<summary><b>Có bắt buộc dùng Optimizer không?</b></summary>

Không. Optimizer là bước tùy chọn, chỉ dùng khi bạn muốn tối ưu thêm hiệu ứng hình ảnh cho các slide quan trọng.

</details>

<details>
<summary><b>Mô hình AI nào cho kết quả tốt nhất?</b></summary>

Claude Opus cho kết quả tốt nhất, nhưng hầu hết các mô hình AI phổ biến hiện nay đều có thể tạo nội dung chất lượng nhờ kiến trúc Skill-based.

</details>

<details>
<summary><b>Có thể quay lại phiên bản cũ không?</b></summary>

Có. Phiên bản ổn định cũ (kiến trúc legacy) tại [v1.3.0](https://github.com/luonghaianh1208/PPTmaster/tree/v1.3.0).

</details>

---

## 📄 Giấy phép

Dự án này theo [Giấy phép MIT](LICENSE).

---

Tác giả: **LƯƠNG HẢI ANH** — GV Trường THPT Chuyên Nguyễn Trãi, TP. Hải Phòng

[⬆ Về đầu trang](#ppt-master--hướng-dẫn-sử-dụng)
