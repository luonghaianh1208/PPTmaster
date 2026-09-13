# Loại việc: Soạn giáo án tích hợp năng lực số và năng lực AI

File dành cho AI. Luôn đọc docs/vi/tro-ly/quy-trinh-hoi.md trước file này, và đọc docs/vi/tro-ly/nang-luc-so-va-ai.md trước khi viết mục tiêu năng lực.

## Khi nào dùng

Thầy cô cần file Word kế hoạch bài dạy theo Công văn 5512 có tích hợp năng lực số và năng lực AI, môn nào cũng được.

Có hai luồng:
- Luồng A — thầy cô đã có giáo án của bài này, cần nâng cấp thêm phần năng lực số, năng lực AI và rubric.
- Luồng B — thầy cô chưa có gì, cần soạn giáo án mới hoàn toàn.

Luồng A: đọc giáo án cũ trước rồi mới hỏi thầy cô; giáo án không đọc được (ví dụ ảnh chụp) thì báo ngay, không hỏi hết bảy câu bắt buộc rồi mới phát hiện.

Ví dụ câu lệnh:
- "Nâng cấp giáo án Ammonia lớp 11 này thành giáo án tích hợp năng lực số và AI"
- "Soạn giáo án mới bài Ancol lớp 11 có tích hợp AI"

Câu lệnh có chữ "giáo án" là ca mơ hồ đã biết: hỏi đúng một câu định tuyến trước khi làm gì khác (xem docs/vi/tro-ly/quy-trinh-hoi.md). Câu lệnh có "kế hoạch bài dạy", "KHBD", "giáo án Word" hoặc "giáo án 5512" là rõ ràng, đi thẳng vào loại việc này, không hỏi Word hay slide.

**Thầy cô đưa ảnh chụp giáo án thì không đọc được.** Xin bản PDF hoặc Word, không tự đoán nội dung.

## Câu hỏi bắt buộc

1. Môn, lớp, tên bài, và bài dạy trong bao nhiêu tiết?
   Gợi ý: lấy từ câu lệnh nếu đã có; chỉ hỏi phần còn thiếu.
2. Thầy cô có giáo án cũ của bài này không, hay em soạn mới?
   Gợi ý: có thì gửi file Word hoặc PDF; ảnh chụp thì em không đọc được.
3. Có file kế hoạch dạy học hoặc phân phối chương trình để em lấy tuần, tiết thứ và mã năng lực số đã khai không?
   Gợi ý: không có thì em tự chọn mã theo khung của Bộ.
4. Lớp có thiết bị gì để làm hoạt động số (máy tính, điện thoại, mạng, phòng máy)?
   Gợi ý: học sinh dùng điện thoại theo nhóm, có wifi.
5. Thầy cô muốn dùng công cụ AI nào với học sinh?
   Gợi ý: một trợ lý chat để dự đoán, kèm PhET hoặc thí nghiệm thật để kiểm chứng.
6. Bài này có thí nghiệm thật không, và có phiếu học tập không?
   Gợi ý: có một thí nghiệm đối chứng và một phiếu học tập.
7. Trường mình yêu cầu thể thức riêng nào không (ký duyệt, quốc hiệu, năm học)?
   Gợi ý: giữ thể thức A4 dọc, Times New Roman 14pt, giãn dòng 1,3.

## Câu hỏi tuỳ chọn

- Trường có khung ký duyệt riêng (Hiệu trưởng, tổ trưởng) không? Chỉ hỏi khi thầy cô nhắc tới.
- Lớp có học sinh cần hỗ trợ riêng (khuyết tật, hoà nhập) không? Chỉ hỏi khi thầy cô nhắc tới.

## Tạo nhanh

1. Môn, lớp, tên bài, số tiết (câu hỏi bắt buộc 1).
2. Thầy cô có giáo án cũ hay soạn mới (câu hỏi bắt buộc 2).

Hỏi xong hai câu trên thì viết giao-an.md ngay và chạy lệnh; loại việc này không theo quick-generate.md của upstream (hồ sơ đó dành cho PPTX).

## Cấu trúc giáo án

Ngữ pháp giao-an.md dưới đây khớp đúng với tools/vi/giao_an_parts/parse.py và frameworks.py; hướng dẫn này tự đủ, không cần tra thêm tài liệu thiết kế. Sai một chỗ, công cụ báo lỗi kèm số dòng; viết đúng ngay từ đầu thay vì đoán.

Cần trích nội dung SGK để tham khảo: chuyển SGK sang Markdown một lần, đặt file đã chuyển ở `projects/_giao-an/_sgk/`, rồi cắt đúng bài bằng `python tools\vi\giao_an.py trich-sgk <sgk.md> --bai "Bài <số>"` (ví dụ `--bai "Bài 5"`, không truyền cả tên bài), ghi ra `--ra projects/_giao-an/<tên_bài>/sgk-trich.md` để lần cắt sau không ghi đè lần trước.

### Khối thông tin đầu (meta)

Dòng đầu file là một dòng `---`. Sau đó mỗi dòng một cặp `khoá: giá trị`, đóng lại bằng một dòng `---` khác.

| Khoá | Bắt buộc | Ý nghĩa |
|---|---|---|
| `school` | có | Tên trường |
| `subject` | có | Tên môn |
| `grade` | có | Lớp, chỉ chữ số |
| `lesson` | có | Tên bài |
| `periods` | có | Số tiết, chỉ chữ số lớn hơn 0 |
| `department` | không | Tổ chuyên môn |
| `teacher` | không | Họ tên giáo viên |
| `track` | không | `chuyen` khi dạy hệ chuyên |
| `week` | không | Tuần dạy |
| `period_numbers` | không | Tiết thứ trong phân phối chương trình |
| `school_year` | không | Năm học |

Dùng khoá nào ngoài mười một khoá trên là lỗi cú pháp. `grade` hoặc `periods` không phải chữ số thì lỗi báo đúng dòng khai khoá đó. Thiếu một khoá bắt buộc thì lỗi báo ở dòng `---` đóng khối, kèm tên khoá còn thiếu. Chưa mở đầu bằng `---`, hoặc chưa đóng bằng `---`, cũng là lỗi.

### Các mục và mục con

Sau khối meta, các mục `## ` phải xuất hiện đúng thứ tự sau, không lặp lại: `## MUC TIEU`, `## THIET BI`, `## TIEN TRINH`, `## PHIEU HOC TAP` (tuỳ chọn), `## RUBRIC`, `## CAN SOAT` (tuỳ chọn, phải là mục cuối cùng nếu có). Thiếu một trong bốn mục bắt buộc (MUC TIEU, THIET BI, TIEN TRINH, RUBRIC) là lỗi.

`## MUC TIEU` có đúng sáu mục h3, viết không dấu, theo thứ tự: `### Kien thuc`, `### Nang luc chung`, `### Nang luc dac thu`, `### Nang luc so`, `### Nang luc AI`, `### Pham chat`. Thân mỗi mục là các dòng `- `, ít nhất một dòng.

Dòng trong `### Nang luc so` và (khi môn có khung AI) `### Nang luc AI` phải mở đầu bằng mã rồi dấu ` — `, ví dụ:

```
### Nang luc so
- 1.1.NC1a — Tìm kiếm, chọn lọc dữ liệu về ứng dụng của ammonia.
- 5.1.NC1a — Dùng thí nghiệm ảo PhET kiểm chứng chuyển dịch cân bằng.
```

Môn chưa có khung mã AI trong nang-luc-so-va-ai.md: dòng đầu của `### Nang luc AI` phải là đúng chuỗi `- (chưa có khung mã cho môn này)`; các dòng sau viết hoạt động AI bằng lời, không được mở đầu bằng mã kiểu `AI-… — `. Môn đã có khung thì tuyệt đối không được ghi chuỗi `(chưa có khung mã cho môn này)`. Cả hai trường hợp đều không tự đặt mã cho môn chưa có khung; mã chỉ báo là việc của Bộ và của tổ chuyên môn. Khung mã của môn/lớp không hợp chủ đề bài thì vẫn dùng mã của khung đó, không tự đặt mã khác; ghi một dòng vào `## CAN SOAT` để thầy cô cân nhắc.

`## THIET BI` có đúng hai mục h3: `### Giao vien`, `### Hoc sinh`; thân là các dòng `- `, ít nhất một dòng mỗi mục.

`## TIEN TRINH` có một mục h3 cho mỗi hoạt động; tên hoạt động là chữ tự do, viết không kèm số thứ tự (ví dụ `### Mở đầu`, không viết `### Hoạt động 1: Mở đầu`) vì công cụ tự đánh số theo thứ tự xuất hiện trong file Word. Thân mỗi hoạt động là các dòng `khoá: giá trị`, tám khoá sau bắt buộc, thiếu khoá nào báo lỗi kèm tên hoạt động:

| Khoá | Ý nghĩa |
|---|---|
| `thoi-luong:` | Số phút, chỉ chữ số lớn hơn 0 |
| `muc-tieu:` | Mục tiêu của hoạt động |
| `noi-dung:` | Nhiệm vụ giao cho học sinh |
| `san-pham:` | Kết quả cụ thể học sinh hoàn thành |
| `chuyen-giao:` | Bước 1: chuyển giao nhiệm vụ |
| `thuc-hien:` | Bước 2: học sinh thực hiện |
| `bao-cao:` | Bước 3: báo cáo, thảo luận |
| `ket-luan:` | Bước 4: kết luận, nhận định |

Hai khoá tuỳ chọn `nls:` và `ai:` ghi một hoặc nhiều mã, cách nhau bằng dấu phẩy. Một khoá lặp lại thành nhiều dòng thì mỗi dòng là một đoạn riêng trong file Word, theo thứ tự xuất hiện (ví dụ hai dòng `noi-dung:` cho ra hai đoạn) — khác với gói đề thi, nơi khoá lặp là lỗi. Toàn bộ tiến trình phải có ít nhất một hoạt động dùng `nls:` và ít nhất một hoạt động dùng `ai:`, kể cả khi môn chưa có khung: lúc đó hoạt động AI dùng đúng chuỗi `ai: (chưa có mã)`; môn đã có khung thì không được dùng `(chưa có mã)`. Mọi mã dùng trong `nls:`/`ai:` của một hoạt động phải đã khai ở `### Nang luc so`/`### Nang luc AI` tương ứng, không thì lỗi nêu rõ mã và tên hoạt động.

`## PHIEU HOC TAP` (tuỳ chọn) có một mục h3 cho mỗi phiếu; thân là các dòng `- `, mỗi dòng thành một đoạn riêng, căn lề trái trong file Word.

`## RUBRIC` có ít nhất hai mục h3, mỗi mục là một tiêu chí; thân đúng ba dòng, mở đầu lần lượt bằng `- Mức 1:`, `- Mức 2:`, `- Mức 3:`. Ít hơn hai tiêu chí, hoặc một tiêu chí không đủ ba mức hay sai thứ tự, là lỗi.

`## CAN SOAT` (tuỳ chọn) là các dòng `- `; nội dung đi vào can-soat.md, không vào file giáo án.

Trong mọi dòng văn bản chỉ dùng ba dấu đánh dấu: `~chỉ số dưới~`, `^chỉ số trên^`, `**in đậm**`.

Ngoài các lỗi trên, công cụ chỉ cảnh báo (vẫn xuất file) khi: tổng `thoi-luong` khác `periods × 45` phút; hoạt động nhắc một phiếu học tập không có trong `## PHIEU HOC TAP`; thiếu hẳn mục `## PHIEU HOC TAP`; đường dẫn thư mục dài hơn 200 ký tự; hoặc file Word đã có sẵn và sẽ bị ghi đè.

### Ví dụ đầy đủ

```
---
school: Trường THPT Nguyễn Trãi
subject: Hoá học
grade: 11
lesson: Bài 5. Ammonia và một số hợp chất ammonium
periods: 2
---
## MUC TIEU

### Kien thuc
- Trình bày được cấu tạo phân tử, tính chất vật lí và tính chất hoá học của ammonia.
- Giải thích được ứng dụng của ammonia và muối ammonium trong nông nghiệp, công nghiệp.

### Nang luc chung
- Tự học và tự chủ: chủ động tìm hiểu thông tin về ammonia từ nhiều nguồn.
- Giao tiếp và hợp tác: trình bày kết quả thảo luận nhóm rõ ràng, mạch lạc.

### Nang luc dac thu
- Nhận thức hoá học: phân tích cấu tạo phân tử ammonia theo thuyết VSEPR.
- Tìm hiểu thế giới tự nhiên dưới góc độ hoá học: đề xuất thí nghiệm kiểm chứng tính tan của ammonia.

### Nang luc so
- 1.1.NC1a — Tìm kiếm, chọn lọc dữ liệu về ứng dụng của ammonia trên các nguồn số tin cậy.
- 5.1.NC1a — Dùng thí nghiệm ảo PhET kiểm chứng chiều chuyển dịch cân bằng tổng hợp ammonia.

### Nang luc AI
- AI-H11.2 — Viết prompt để AI dự đoán chiều chuyển dịch cân bằng khi tăng áp suất theo Le Chatelier.
- AI-H11.4 — Đối chiếu kết quả dự đoán của AI với thí nghiệm thật, chỉ ra chỗ AI nói sai hoặc nói thiếu.

### Pham chat
- Trách nhiệm: cẩn thận khi thao tác với dung dịch ammonia trong phòng thí nghiệm.
- Trung thực: báo cáo đúng kết quả thí nghiệm, kể cả khi khác dự đoán ban đầu.

## THIET BI

### Giao vien
- Máy tính, máy chiếu, kết nối mạng để dùng công cụ AI và thí nghiệm ảo PhET.
- Hoá chất và dụng cụ thí nghiệm điều chế ammonia trong phòng thí nghiệm.

### Hoc sinh
- Điện thoại hoặc máy tính bảng có kết nối mạng, dùng theo nhóm.
- Sách giáo khoa Hoá học 11 và phiếu học tập.

## TIEN TRINH

### Mở đầu
thoi-luong: 10
muc-tieu: Huy động kiến thức nền về mùi và tính tan của ammonia trong đời sống.
noi-dung: Học sinh quan sát video thí nghiệm đài phun nước ammonia và nêu dự đoán.
san-pham: Câu trả lời miệng nêu dự đoán về tính tan của ammonia.
chuyen-giao: Giáo viên chiếu video và nêu câu hỏi dự đoán.
thuc-hien: Học sinh xem video, thảo luận cặp đôi trong 3 phút.
bao-cao: Đại diện hai cặp trình bày dự đoán trước lớp.
ket-luan: Giáo viên dẫn dắt vào bài học về ammonia.
nls: 1.1.NC1a

### Hình thành kiến thức
thoi-luong: 30
muc-tieu: Trình bày cấu tạo phân tử và tính chất hoá học của ammonia.
noi-dung: Học sinh dùng AI dự đoán chiều chuyển dịch cân bằng tổng hợp ammonia khi tăng áp suất, hoàn thành Phiếu học tập 1.
san-pham: Phiếu học tập 1 đã hoàn thành, có so sánh dự đoán AI với lý thuyết Le Chatelier.
chuyen-giao: Giáo viên giao Phiếu học tập 1 và hướng dẫn viết prompt cho AI.
thuc-hien: Học sinh viết prompt, đối chiếu kết quả AI với sách giáo khoa theo nhóm bốn người.
bao-cao: Mỗi nhóm trình bày một phần Phiếu học tập 1.
ket-luan: Giáo viên chốt kiến thức, chỉ ra chỗ AI trả lời chưa chính xác nếu có.
ai: AI-H11.2

### Luyện tập
thoi-luong: 30
muc-tieu: Vận dụng thí nghiệm ảo để kiểm chứng chiều chuyển dịch cân bằng.
noi-dung: Học sinh dùng PhET mô phỏng phản ứng tổng hợp ammonia và ghi lại kết quả.
san-pham: Bảng kết quả mô phỏng so với dự đoán ở Hoạt động 2.
chuyen-giao: Giáo viên hướng dẫn thao tác trên PhET.
thuc-hien: Học sinh thao tác mô phỏng theo nhóm, ghi số liệu.
bao-cao: Nhóm trình bày bảng số liệu và nhận xét.
ket-luan: Giáo viên nhận xét, củng cố nguyên lí Le Chatelier.
nls: 5.1.NC1a

### Vận dụng
thoi-luong: 20
muc-tieu: Đánh giá độ tin cậy của AI khi dự đoán phản ứng hoá học.
noi-dung: Học sinh so sánh kết quả AI ở Hoạt động 2 với kết quả thí nghiệm ảo ở Hoạt động 3, chỉ ra chỗ AI nói sai hoặc nói thiếu.
san-pham: Đoạn nhận xét ngắn nêu rõ chỗ AI đúng, chỗ AI sai.
chuyen-giao: Giáo viên nêu yêu cầu đối chiếu hai kết quả.
thuc-hien: Học sinh viết nhận xét cá nhân trong 5 phút.
bao-cao: Ba học sinh chia sẻ nhận xét trước lớp.
ket-luan: Giáo viên tổng kết, nhấn mạnh vai trò kiểm chứng của con người với AI.
ai: AI-H11.4

## PHIEU HOC TAP

### Phiếu học tập 1
- Viết một prompt yêu cầu AI dự đoán chiều chuyển dịch cân bằng khi tăng áp suất trong phản ứng tổng hợp ammonia.
- So sánh câu trả lời của AI với nguyên lí Le Chatelier trong sách giáo khoa.

## RUBRIC

### Mức độ hoàn thành Phiếu học tập
- Mức 1: Ghi được câu trả lời của AI nhưng chưa so sánh với lý thuyết.
- Mức 2: So sánh được câu trả lời của AI với lý thuyết nhưng chưa chỉ ra chỗ khác biệt.
- Mức 3: Chỉ ra rõ chỗ AI đúng, chỗ AI sai hoặc thiếu so với lý thuyết.

### Mức độ hợp tác nhóm
- Mức 1: Tham gia thảo luận khi được nhắc.
- Mức 2: Chủ động đóng góp ý kiến trong nhóm.
- Mức 3: Chủ động đóng góp và hỗ trợ bạn cùng nhóm hoàn thành nhiệm vụ.

## CAN SOAT

- Thầy cô kiểm tra lại số liệu áp suất dùng trong Hoạt động 3 có khớp với thiết bị PhET của trường không.
```

## Đầu ra

`projects/_giao-an/<tên_bài>/` chứa:
- `giao-an.md` — nguồn duy nhất.
- `giao-an.docx` — kế hoạch bài dạy để nộp cho trường.
- `can-soat.md` — ghi chú nội bộ (nội dung mục `## CAN SOAT` cộng các cảnh báo khi xuất), không nằm trong giáo án nộp cho trường.

## Ghi vào brief

- Loại việc: Soạn giáo án tích hợp năng lực số và năng lực AI, ghi đúng lời thầy cô.
- Brief lưu tại `projects/_giao-an/<tên_bài>/brief.md`, viết theo đúng bố cục của docs/vi/tro-ly/mau-brief.md.
- Loại việc này không dùng các bước dành cho PPTX: không kết thúc tin nhắn hỏi bằng dòng chốt cách xác nhận, không chạy `import-sources` hay `project_manager.py init`, và không có bước xác nhận của upstream — cũng không tự đặt mã cho môn chưa có khung khi ghi brief. Thầy cô trả lời xong lượt hỏi thì viết giao-an.md rồi chạy lệnh ngay.

Luồng A phải giữ nguyên nội dung chuyên môn của thầy cô: không rút gọn, không viết lại cho hay hơn, không đổi bài tập. Chỉ thêm phần năng lực số, năng lực AI và rubric. Không tự đặt mã chỉ báo cho môn chưa có khung. Chỗ nào không đọc được từ file gốc thì ghi vào mục CAN SOAT, không tự viết bù.
