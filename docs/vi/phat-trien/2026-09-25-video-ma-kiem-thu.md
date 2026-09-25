# Biên bản kiểm thử: video giải thích dựng bằng mã (v6.3.2-vi.9)

Ngày 2026-09-25. Máy chủ repo: Windows 11, Chromium headless shell của Playwright, FFmpeg bản Gyan.

## 1. Phép đo trước khi làm (spec mục 12)

### 1.1 Font
Segoe Print, Ink Free và Comic Sans MS đều hiện đủ dấu chồng tiếng Việt. Chọn Segoe Print làm font chính, Ink Free và Comic Sans MS dự phòng.

### 1.2 Tốc độ chụp khung
| Định dạng | ms/khung | 5 phút ở 12 khung/giây | ở 15 | ở 30 |
|---|---|---|---|---|
| PNG | 45,3 | 2,7 phút | 3,4 phút | 6,8 phút |
| JPEG chất lượng 92 | 50,9 | 3,1 phút | 3,8 phút | 7,6 phút |

Quyết định: chụp PNG, 15 khung/giây, ghép ra 30 khung/giây (nhân khung). Ngưỡng 20 phút của spec không bị chạm, không cần chụp song song.

## 2. Kết quả kiểm thử

(Điền ở Task 11.)
