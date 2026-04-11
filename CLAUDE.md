# CLAUDE.md

File này cung cấp tổng quan dự án cho Claude Code. Trước khi thực hiện tạo PPT, **bạn PHẢI đọc `skills/ppt-master/SKILL.md`** để nắm đầy đủ quy trình và quy tắc.

## Tổng quan dự án

PPT Master là hệ thống tạo nội dung SVG đa định dạng điều khiển bởi AI. Thông qua phối hợp đa vai trò (Strategist → Image_Generator → Executor → Optimizer), hệ thống chuyển đổi tài liệu nguồn (PDF/DOCX/URL/Markdown) thành các trang SVG chất lượng cao và xuất ra PPTX.

**Quy trình cốt lõi**: `Tài liệu nguồn → Tạo dự án → Lựa chọn mẫu → Strategist Tám xác nhận → [Image_Generator] → Executor → Hậu xử lý → Xuất PPTX`

## Lệnh thường dùng

```bash
# Chuyển đổi nội dung nguồn
python3 skills/ppt-master/scripts/pdf_to_md.py <file_PDF>
python3 skills/ppt-master/scripts/doc_to_md.py <file_DOCX_hoặc_khác>   # Yêu cầu: pandoc (DOCX/EPUB/HTML/LaTeX/RST/v.v.)
python3 skills/ppt-master/scripts/web_to_md.py <URL>
node skills/ppt-master/scripts/web_to_md.cjs <URL>

# Quản lý dự án
python3 skills/ppt-master/scripts/project_manager.py init <tên_dự_án> --format ppt169
python3 skills/ppt-master/scripts/project_manager.py import-sources <đường_dẫn> <file_nguồn...> --move
python3 skills/ppt-master/scripts/project_manager.py validate <đường_dẫn>

# Công cụ hình ảnh
python3 skills/ppt-master/scripts/analyze_images.py <đường_dẫn>/images
python3 skills/ppt-master/scripts/nano_banana_gen.py "prompt" --aspect_ratio 16:9 --image_size 1K -o <đường_dẫn>/images

# Kiểm tra chất lượng SVG
python3 skills/ppt-master/scripts/svg_quality_checker.py <đường_dẫn>

# Quy trình hậu xử lý (PHẢI chạy tuần tự, từng lệnh một — KHÔNG BAO GIỜ gộp)
python3 skills/ppt-master/scripts/total_md_split.py <đường_dẫn>
# ✅ Xác nhận không lỗi trước khi chạy lệnh tiếp theo
python3 skills/ppt-master/scripts/finalize_svg.py <đường_dẫn>
# ✅ Xác nhận không lỗi trước khi chạy lệnh tiếp theo
python3 skills/ppt-master/scripts/svg_to_pptx.py <đường_dẫn> -s final
# Lưu ý: thêm --native nếu không có PowerPoint
```

## Kiến trúc

- `skills/ppt-master/references/` — Định nghĩa vai trò AI và thông số kỹ thuật
- `skills/ppt-master/scripts/` — 28 công cụ script
- `skills/ppt-master/templates/` — Mẫu bố cục, biểu đồ, 640+ icon vector
- `examples/` — Dự án ví dụ
- `projects/` — Không gian làm việc

## Ràng buộc kỹ thuật SVG (Không thương lượng)

**Tính năng bị cấm**: `clipPath` | `mask` | `<style>` | `class` | CSS ngoài | `<foreignObject>` | `textPath` | `@font-face` | `<animate*>` | `<script>` | `marker-end` | `<iframe>` | `<symbol>+<use>` (`id` trong `<defs>` là tham chiếu hợp lệ và KHÔNG bị cấm)

**Thay thế tương thích PPT**:

| Bị cấm | Thay thế |
|--------|----------|
| `rgba()` | `fill-opacity` / `stroke-opacity` |
| `<g opacity>` | Đặt opacity trên từng phần tử con |
| `<image opacity>` | Phủ lớp mask |
| `marker-end` arrows | `<polygon>` tam giác |

## Tham chiếu nhanh định dạng Canvas

| Định dạng | viewBox |
|-----------|---------|
| PPT 16:9 | `0 0 1280 720` |
| PPT 4:3 | `0 0 1024 768` |
| Facebook/TikTok (3:4) | `0 0 1242 1660` |
| Zalo/Instagram (1:1) | `0 0 1080 1080` |
| Story | `0 0 1080 1920` |

## Lưu ý hậu xử lý

- **KHÔNG BAO GIỜ** dùng `cp` thay thế cho `finalize_svg.py`
- **KHÔNG BAO GIỜ** xuất trực tiếp từ `svg_output/` — PHẢI xuất từ `svg_final/` (dùng `-s final`)
- KHÔNG thêm flag thừa như `--only` vào lệnh hậu xử lý
- **KHÔNG BAO GIỜ** chạy ba bước hậu xử lý trong cùng một code block hoặc một lần gọi shell
