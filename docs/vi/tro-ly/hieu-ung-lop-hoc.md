# Hiệu ứng cho slide lớp học

File dành cho AI. Nó quy định ba mức hiệu ứng thầy cô chọn trong lượt hỏi, cách làm từng kiểu hiệu ứng bằng cơ chế sẵn có của upstream, cách giữ hiệu ứng khi làm video, và cách kiểm file PPTX sau khi xuất. Cú pháp `animations.json` và các bước của upstream vẫn theo `skills/ppt-master/references/animations.md` và `skills/ppt-master/workflows/stages/customize-animations.md`.

## Khi nào áp dụng

- Bài mới thuộc Bài giảng, Báo cáo – tổng kết, Hoạt động Đoàn – sự kiện hoặc Tập huấn/workshop, khi brief có dòng "Mức hiệu ứng".
- Video bài giảng làm từ một bài giảng đã có: chỉ áp mục "Video"; không vẽ thêm trang hay thêm hiệu ứng mới cho bài đó.
- Không áp dụng cho poster, làm đẹp slide có sẵn, sửa PPTX có sẵn, tạo template, đề thi và giáo án Word. Không kiểm hay sửa lại bài cũ thầy cô làm từ trước, trừ khi thầy cô yêu cầu.

## Ba mức

"Trang nội dung" là mọi trang đang hiện trừ trang đầu và trang cuối. "Trang có hiệu ứng" là trang nội dung có ít nhất một hiệu ứng đối tượng (xuất hiện, nhấn mạnh, di chuyển, biến mất), một ô bấm hiện đáp án, hoặc chuyển trang Morph.

| Mức | AI làm gì | Công cụ kiểm chặn khi |
|---|---|---|
| không | Chỉ chuyển trang mờ dần mặc định. Không chạy bước tuỳ chỉnh hiệu ứng. | Có hiệu ứng đối tượng, hoặc chuyển trang khác mờ dần |
| vừa | Hiện từng ý ở trang liệt kê, các bước, công thức. Ô bấm hiện đáp án ở câu trắc nghiệm. Chuyển trang nổi bật ở ranh giới giữa các hoạt động hoặc các phần. | Dưới 30% trang nội dung có hiệu ứng |
| nhiều | Như mức vừa, thêm Morph cho diễn biến (thí nghiệm, quá trình, phóng to chi tiết) và nhấn mạnh từ khoá quan trọng. | Dưới 50% trang nội dung có hiệu ứng |

- Thầy cô chọn "vừa" hoặc "nhiều", hoặc brief ghi "vừa (AI đề xuất, chưa duyệt)", là yêu cầu hiệu ứng của lớp Việt: chạy bước `customize-animations` của upstream, kể cả khi tạo nhanh. Ở luồng có bước xác nhận, ghi mức hiệu ứng vào đề xuất và bật Custom Animations trong `design_spec.md` mục I.
- Tỉ lệ ở bảng là mức sàn để phát hiện việc quên làm hiệu ứng, không phải chỉ tiêu. Mỗi hiệu ứng phải phục vụ một việc có thật trên trang (ý cần nói lần lượt, đáp án cần giấu, diễn biến cần thấy); không thêm chuyển động chỉ để đủ tỉ lệ.
- Morph và mức nhiều phải được tính từ lúc lên danh sách trang, vì cần vẽ sẵn hai trang cho mỗi diễn biến.
- Máy chiếu 4:3 hoặc máy tính cũ: khuyên thầy cô chọn "vừa" thay cho "nhiều".

## Bốn kiểu hiệu ứng

### Hiện từng ý khi bấm

- Mỗi ý, mỗi bước, mỗi dòng công thức là một nhóm `<g id>` riêng ở gốc SVG, có `data-pptx-bounds` không chồng nhau. Trang đang để các hình rời không có nhóm thì gom lại theo mục 2 của `customize-animations.md` trước khi viết `animations.json`.
- Trong `animations.json`: `slides.<trang>.animation.trigger` là `on-click`, mỗi nhóm một dòng `entrance_fade` hoặc `entrance_wipe`. Tiêu đề trang không cần hiệu ứng.
- Mỗi trang không quá 8 bước bấm ở mức vừa, 12 bước ở mức nhiều.

### Ô bấm hiện đáp án

- Đặt hai nhóm cạnh nhau, không chồng lên nhau: một nhóm nút (ví dụ "Xem đáp án") và một nhóm ô đáp án. Bộ kiểm SVG của upstream chặn hai nhóm gốc chồng nhau, nên không đặt ô đáp án đè lên nút.
- Trong `animations.json`: nhóm ô đáp án có `effect` là `entrance_fade` và `trigger_shape` là id của nhóm nút.
- Câu trắc nghiệm để đáp án ở trang riêng ngay sau (theo `bai-giang.md`) thì không cần ô bấm.

### Morph cho diễn biến

- Vẽ hai trang liền nhau là hai trạng thái của cùng một vật (bình trước và sau phản ứng, sơ đồ trước và sau khi phóng to). Việc này làm lúc viết SVG; bước hiệu ứng không tự tạo được trang mới.
- Trên trang sau: `slides.<trang_sau>.transition.effect` là `morph`, và `slides.<trang_sau>.morph` có `from` là trang trước cùng `pairs` nối id nhóm hai trang.

### Chuyển trang nổi bật

- Chỉ đặt ở trang mở đầu một hoạt động hoặc một phần mới, không đặt cho mọi trang. Chọn một kiểu cho cả bài, ví dụ `push`, `wipe` hoặc `cube`, trong `slides.<trang>.transition`.

## Video

Video không có ai bấm chuột, và bản xuất có thuyết minh của upstream từ chối mọi hiệu ứng `on-click`. Làm theo thứ tự:

1. Thầy cô chọn giữ hiệu ứng và dự án có `animations.json`: chép thành `animations_video.json` trong thư mục dự án, rồi chỉ sửa bản chép. Không sửa `animations.json`, để bản trình chiếu trên lớp vẫn bấm được.
   - Đổi mọi `on-click` sang `after-previous` (nối tiếp) hoặc `with-previous` (cùng lúc); thêm `delay` vài giây khi ý sau cần chờ lời giảng.
   - Bỏ `trigger_shape`: ô đáp án hiện bằng `after-previous` với `delay` sau khi đọc xong câu hỏi.
   - Giữ nguyên Morph và chuyển trang.
2. Xuất bản thuyết minh với `--animation-config animations_video.json`, kể cả khi đã có một file `_narrated.pptx` cũ.
3. Thầy cô chọn bỏ hiệu ứng: xuất bản thuyết minh với `--no-animations` (cờ này bỏ cả chuyển trang). Dự án không có `animations.json` mà thầy cô giữ hiệu ứng: không thêm cờ nào, bản xuất giữ chuyển trang mặc định.

## Kiểm sau khi xuất

Có `venv\Scripts\python.exe` ở thư mục gốc repo thì dùng nó thay cho `python`.

- Bài trình chiếu: `python tools\vi\kiem_hieu_ung.py <đường_dẫn_dự_án>\exports\<file>.pptx --muc khong|vua|nhieu`, dùng đúng mức trong brief, kể cả mức "(AI đề xuất, chưa duyệt)".
- Bài làm video (giữ hay bỏ hiệu ứng): `python tools\vi\kiem_hieu_ung.py <file>_narrated.pptx --video`.

Đọc dòng JSON ở stdout.

| Kết quả | Xử lý |
|---|---|
| `ready` là `true` | Báo thầy cô số trang có hiệu ứng trên số trang nội dung, số trang Morph, số ô bấm hiện đáp án, và đọc nguyên văn `warnings`. |
| `error.step` là `muc` | Xem lại trang nội dung nào có ý cần nói lần lượt, đáp án cần giấu hay diễn biến cần thấy mà chưa có hiệu ứng; thêm hiệu ứng cho đúng những trang đó, xuất lại, chạy lại đúng một lần. Thiếu trang Morph thì quay lại bước viết SVG và chạy lại bộ kiểm SVG trước khi xuất. Vẫn không đạt thì báo thầy cô nguyên văn `error.message`, nêu lý do nội dung không cần thêm hiệu ứng và đề xuất hạ mức. |
| `error.step` là `video` | Sửa `animations_video.json` theo mục "Video", xuất lại bản thuyết minh, chạy lại đúng một lần. Vẫn không đạt thì dừng, không dựng video, báo thầy cô. |
| `error.step` là `input` | Sửa lệnh theo `error.fix` rồi chạy lại. |
| `error.step` là `parse` | Xuất lại PPTX bằng `svg_to_pptx.py`, không sửa tay file PPTX. |
| `error.step` là `internal` | Báo thầy cô nguyên văn `error.message` để gửi người bảo trì. |
