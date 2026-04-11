# PPT Master — AI tự động tạo bài thuyết trình chỉnh sửa được từ mọi tài liệu

[![Version](https://img.shields.io/badge/version-v2.2.0-blue.svg)]()
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Chỉ cần đưa vào một file PDF, DOCX, URL hoặc Markdown — AI sẽ tự động tạo ra **bài thuyết trình đẹp mắt, có thể chỉnh sửa trực tiếp trong PowerPoint**. Hỗ trợ PPT 16:9, Facebook/TikTok, Zalo/Instagram và hơn 10 định dạng khác.

> 💡 **Cập nhật lớn**: Kiến trúc dự án đã được nâng cấp toàn diện (Skill-based architecture):
> 1. **Giảm tiêu hao Token & phụ thuộc mô hình**: Giờ đây ngay cả các mô hình không phải Opus cũng có thể tạo kết quả chất lượng tốt.
> 2. **Khả năng mở rộng cao**: Thư mục `skills` được tổ chức theo chuẩn Agent Skills, mỗi thư mục con là một Skill hoàn toàn độc lập. Có thể được tích hợp vào các AI client hỗ trợ (Claude Code, Antigravity, GitHub Copilot).
> 3. **Phiên bản ổn định dự phòng**: Nếu gặp sự cố với phiên bản mới, có thể quay lại kiến trúc cũ tại [v1.3.0](https://github.com/luonghaianh1208/PPTmaster/tree/v1.3.0).

---

## 🏗️ Kiến trúc hệ thống

```
Đầu vào (PDF/DOCX/URL/Markdown)
    ↓
[Chuyển đổi nguồn] → pdf_to_md.py / doc_to_md.py / web_to_md.py
    ↓
[Tạo dự án] → project_manager.py init <tên_dự_án> --format <định_dạng>
    ↓
[Lựa chọn mẫu] A) Dùng mẫu có sẵn  B) Không dùng mẫu
    ↓
[Strategist] Chiến lược gia - Tám xác nhận & Thiết kế đặc tả
    ↓
[Image_Generator] Tạo hình ảnh (khi chọn AI tạo ảnh)
    ↓
[Executor] Thực thi - Hai giai đoạn
    ├── Giai đoạn xây dựng hình ảnh: Tạo tất cả trang SVG → svg_output/
    └── Giai đoạn xây dựng logic: Tạo bài thuyết trình hoàn chỉnh → notes/total.md
    ↓
[Hậu xử lý] → total_md_split.py → finalize_svg.py → svg_to_pptx.py
    ↓
Đầu ra: PPTX có thể chỉnh sửa (tự động nhúng ghi chú)
    ↓
[Optimizer_CRAP] Tối ưu hóa (tùy chọn, chỉ dùng khi chưa hài lòng)
```

### 📚 Danh mục tài liệu

| Tài liệu | Mô tả |
|-----------|--------|
| 🧭 [AGENTS.md](./AGENTS.md) | Tổng quan cấp kho lưu trữ cho AI agent |
| 📖 [SKILL.md](./skills/ppt-master/SKILL.md) | Quy trình và quy tắc `ppt-master` |
| 🎨 [Hướng dẫn thiết kế](./skills/ppt-master/references/design-guidelines.md) | Màu sắc, typography và bố cục |
| 📐 [Định dạng canvas](./skills/ppt-master/references/canvas-formats.md) | PPT, Facebook/TikTok, Zalo/Instagram và hơn 10 định dạng |
| 🖼️ [Hướng dẫn nhúng ảnh](./skills/ppt-master/references/svg-image-embedding.md) | Best practices nhúng ảnh SVG |
| 📊 [Thư viện biểu đồ](./skills/ppt-master/templates/charts/) | Các mẫu biểu đồ chuẩn hóa |
| 🔧 [Định nghĩa vai trò](./skills/ppt-master/references/) | Định nghĩa vai trò và tham chiếu kỹ thuật |
| 🛠️ [Bộ công cụ](./skills/ppt-master/scripts/README.md) | Hướng dẫn sử dụng tất cả công cụ |
| 💼 [Chỉ mục ví dụ](./examples/README.md) | Các dự án ví dụ |

---

## 🚀 Bắt đầu nhanh

### 1. Cấu hình môi trường

#### Môi trường Python (Bắt buộc)

Dự án yêu cầu **Python 3.8+** để chạy các công cụ chuyển đổi PDF, hậu xử lý SVG và xuất PPTX.

| Nền tảng | Cài đặt khuyến nghị |
|----------|---------------------|
| **Windows** | Tải từ [Python Official](https://www.python.org/downloads/) |
| **macOS** | Dùng [Homebrew](https://brew.sh/): `brew install python` |
| **Linux** | `sudo apt install python3 python3-pip` (Ubuntu/Debian) |

> 💡 **Xác nhận cài đặt**: Chạy `python3 --version` để xác nhận phiên bản ≥ 3.8

#### Môi trường Node.js (Tùy chọn)

Nếu cần sử dụng công cụ `web_to_md.cjs` (chuyển đổi trang web có bảo mật cao), cài đặt Node.js.

> 💡 **Xác nhận**: Chạy `node --version` để xác nhận phiên bản ≥ 18

#### Pandoc (Tùy chọn)

Nếu cần sử dụng `doc_to_md.py` (chuyển đổi DOCX, EPUB, LaTeX sang Markdown), cài đặt [Pandoc](https://pandoc.org/).

### 2. Cài đặt phụ thuộc

```bash
pip install -r requirements.txt
```

> Nếu gặp vấn đề quyền truy cập, dùng `pip install --user -r requirements.txt` hoặc cài trong môi trường ảo.

### 3. Mở trình soạn thảo AI

Các trình soạn thảo AI được khuyến nghị:

| Công cụ | Đánh giá | Mô tả |
|---------|:--------:|-------|
| **[Claude Code](https://claude.ai/)** | ⭐⭐⭐ | **Khuyến nghị mạnh**! CLI chính thức của Anthropic |
| [Cursor](https://cursor.sh/) | ⭐⭐ | Trình soạn thảo AI phổ biến, trải nghiệm tốt |
| [VS Code + Copilot](https://code.visualstudio.com/) | ⭐⭐ | Giải pháp của Microsoft, hiệu quả chi phí |
| [Antigravity](https://antigravity.dev/) | ⭐ | Miễn phí nhưng hạn ngạch ít và không ổn định |

### 4. Bắt đầu sáng tạo

Mở bảng chat AI trong trình soạn thảo và mô tả nội dung bạn muốn tạo:

```
Người dùng: Tôi có một báo cáo quý 3, cần làm thành PPT

AI: Được. Trước tiên xác nhận có sử dụng mẫu không; sau đó Strategist
   sẽ tiếp tục tám xác nhận và tạo đặc tả thiết kế.
   [Lựa chọn mẫu] [Khuyến nghị] B) Không dùng mẫu
   [Strategist] 1. Định dạng canvas: [Khuyến nghị] PPT 16:9
   [Strategist] 2. Số trang: [Khuyến nghị] 8-10 trang
   ...
```

> 💡 **Khuyến nghị mô hình**: Claude Opus cho kết quả tốt nhất, nhưng hầu hết các mô hình phổ biến hiện nay đều có thể tạo nội dung chất lượng.

> 📝 **Chỉnh sửa sau xuất**: Mỗi trang trong PPTX xuất ra đều ở định dạng SVG. Trong PowerPoint, chọn nội dung trang, nhấp chuột phải chọn **"Chuyển đổi thành Hình dạng"** (Convert to Shape) để tự do chỉnh sửa. Yêu cầu **Office 2016** trở lên.

> 💡 **AI mất ngữ cảnh?** Yêu cầu AI đọc `skills/ppt-master/SKILL.md` trước; dùng `AGENTS.md` làm tổng quan.

### 5. Cấu hình Gemini tạo ảnh (Tùy chọn)

Công cụ `nano_banana_gen.py` có thể tạo ảnh chất lượng cao qua Gemini API. Cấu hình biến môi trường trước khi dùng:

```bash
# Bắt buộc: Gemini API Key (lấy từ https://aistudio.google.com/apikey)
export GEMINI_API_KEY="your-api-key"

# Tùy chọn: Endpoint tùy chỉnh (cho dịch vụ proxy)
export GEMINI_BASE_URL="https://your-proxy-url.com/v1beta"
```

---

## 📁 Cấu trúc dự án

```text
ppt-master/
├── skills/
│   └── ppt-master/                 # Nguồn kỹ năng chính
│       ├── SKILL.md                #   Quy trình làm việc
│       ├── workflows/              #   Quy trình công việc
│       ├── references/             #   Định nghĩa vai trò & tham chiếu
│       ├── scripts/                #   Bộ công cụ
│       └── templates/              #   Mẫu bố cục, biểu đồ, icon
├── examples/                       # Dự án ví dụ
├── projects/                       # Không gian làm việc
├── AGENTS.md                       # Tổng quan cho AI Agent
└── CLAUDE.md                       # Tổng quan cho Claude Code CLI
```

---

## 🛠️ Lệnh thường dùng

```bash
# Khởi tạo dự án
python3 skills/ppt-master/scripts/project_manager.py init <tên_dự_án> --format ppt169

# Nhập nguồn tài liệu vào thư mục dự án
python3 skills/ppt-master/scripts/project_manager.py import-sources <đường_dẫn> <file_nguồn>

# PDF sang Markdown
python3 skills/ppt-master/scripts/pdf_to_md.py <file_PDF>

# DOCX sang Markdown (cần pandoc)
python3 skills/ppt-master/scripts/doc_to_md.py <file_DOCX>

# Hậu xử lý (chạy theo thứ tự)
python3 skills/ppt-master/scripts/total_md_split.py <đường_dẫn_dự_án>
python3 skills/ppt-master/scripts/finalize_svg.py <đường_dẫn_dự_án>
python3 skills/ppt-master/scripts/svg_to_pptx.py <đường_dẫn_dự_án> -s final
# Dự phòng: thêm --native nếu không có PowerPoint
```

> 📖 Hướng dẫn đầy đủ xem [Hướng dẫn sử dụng công cụ](./skills/ppt-master/scripts/README.md)

---

## ❓ Câu hỏi thường gặp

<details>
<summary><b>H: PPT tạo ra có chỉnh sửa được không?</b></summary>

Có! Mỗi trang trong PPTX xuất ra ở định dạng SVG. Trong PowerPoint, chọn nội dung, nhấp chuột phải chọn **"Chuyển đổi thành Hình dạng"** (Convert to Shape) — sau đó mọi văn bản, hình dạng và màu sắc đều có thể chỉnh sửa tự do. Yêu cầu **Office 2016** trở lên.

</details>

<details>
<summary><b>H: Ba loại Executor khác nhau thế nào?</b></summary>

- **Executor_General**: Tình huống tổng quát, bố cục linh hoạt
- **Executor_Consultant**: Tư vấn chung, trực quan hóa dữ liệu
- **Executor_Consultant_Top**: Tư vấn cấp cao (MBB), 5 kỹ thuật cốt lõi

</details>

<details>
<summary><b>H: Có bắt buộc dùng Optimizer_CRAP không?</b></summary>

Không. Chỉ dùng khi cần tối ưu hiệu ứng hình ảnh của các trang quan trọng.

</details>

---

## 📄 Giấy phép

Dự án này theo [Giấy phép MIT](LICENSE).

---

Tác giả: **LƯƠNG HẢI ANH** — GV Trường THPT Chuyên Nguyễn Trãi, TP. Hải Phòng

[⬆ Về đầu trang](#ppt-master--ai-tự-động-tạo-bài-thuyết-trình-chỉnh-sửa-được-từ-mọi-tài-liệu)
