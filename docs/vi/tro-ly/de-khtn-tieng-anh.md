# Loại việc: Soạn đề KHTN tiếng Anh

File dành cho AI. Luôn đọc docs/vi/tro-ly/quy-trinh-hoi.md trước file này, và đọc docs/vi/tro-ly/tieng-anh-khoa-hoc.md trước khi viết câu hỏi tiếng Anh.

## Khi nào dùng

Thầy cô cần đề kiểm tra môn khoa học tự nhiên viết bằng tiếng Anh, cho môn KHTN lớp 6–9 (chương trình tích hợp), hoặc Vật lí / Hoá học / Sinh học lớp 10–12 (chương trình tách môn).

Có hai trường hợp:
- Luồng A — thầy cô đã có đề tiếng Việt, cần chuyển sang tiếng Anh kèm bản song ngữ để tổ chuyên môn soát.
- Luồng B — thầy cô chưa có gì, cần AI soạn đề mới hoàn toàn bằng tiếng Anh: soạn ma trận đặc tả trước (chủ đề × mức độ × số câu) theo câu trả lời của thầy cô, viết câu hỏi trực tiếp bằng tiếng Anh, và ghi bản tiếng Việt của từng câu vào `vi:` để thầy cô soát.

Ví dụ câu lệnh:
- "Chuyển đề giữa kì Hoá 11 này sang tiếng Anh"
- "Soạn đề Vật lí 10 tiếng Anh 45 phút chương động lực học"

**Thầy cô đưa ảnh chụp đề thì không đọc được.** Xin bản PDF hoặc Word, không tự đoán nội dung đề.

Chuyển từ đề tiếng Việt thì không đổi số liệu, không đổi đáp án đúng, không đổi thứ tự câu. Câu gốc sai hoặc mơ hồ thì không tự sửa, ghi vào mục "Cần thầy cô soát".

## Câu hỏi bắt buộc

1. Môn và lớp (KHTN 6–9, hay Vật lí / Hoá học / Sinh học 10–12)?
   Gợi ý: lấy từ câu lệnh nếu đã có; chỉ hỏi phần còn thiếu.
2. Đề cho kỳ nào, làm trong bao nhiêu phút, tổng bao nhiêu điểm?
   Gợi ý: giữa kì hoặc cuối kì theo lịch nhà trường; 45 phút, thang điểm 10.
3. Nội dung thuộc chương hay chủ đề nào?
   Gợi ý: có đề tiếng Việt sẵn thì lấy theo đúng đề đó, không hỏi lại.
4. Số câu mỗi phần (Phần I trắc nghiệm, Phần II đúng/sai, Phần III trả lời ngắn)?
   Gợi ý: THPT dùng 18 – 4 – 6 theo đề tham khảo 2025. Môn KHTN ở THCS không có định dạng chung nên phải hỏi, không được lấy gợi ý làm mặc định.
5. Tỉ lệ mức độ biết – hiểu – vận dụng?
   Gợi ý: 4 – 4 – 2.
6. Đề dùng để làm gì (kiểm tra lớp song ngữ, đề luyện thêm, đề tham khảo cho tổ)?
   Gợi ý: kiểm tra lớp song ngữ thì dùng từ vựng đơn giản hơn đề tham khảo cho tổ chuyên môn. Câu này quyết định mức từ vựng tiếng Anh.
7. Có cần bản song ngữ để tổ soát không?
   Gợi ý: có, để tổ chuyên môn đối chiếu. Trả lời "không" thì chạy với `--phan de,dap-an`, không xuất `de-song-ngu.docx`; trường `vi:` trong `de.md` vẫn ghi để sau này xuất lại được.

Luồng A đã có đề sẵn thì câu 3, 4 và 5 đọc được từ chính đề, chỉ hỏi phần còn thiếu.

## Câu hỏi tuỳ chọn

- Có cần mã đề không? Chỉ hỏi khi thầy cô nhắc tới mã đề.

Bản này không có phần tự luận. Thầy cô hỏi hoặc muốn có câu tự luận thì nói rõ đề chỉ gồm Phần I trắc nghiệm, Phần II đúng/sai và Phần III trả lời ngắn, và gợi ý dùng Phần III trả lời ngắn thay cho câu tự luận.

## Tạo nhanh

1. Môn và lớp (câu hỏi bắt buộc 1).
2. Số câu mỗi phần (câu hỏi bắt buộc 4).

Hỏi xong hai câu trên thì viết `de.md` ngay và chạy lệnh; loại việc này không theo `quick-generate.md` của upstream (hồ sơ đó dành cho PPTX).

## Cấu trúc đề

Giữ cấu trúc đề Việt Nam đang dùng:
- Phần I — trắc nghiệm bốn lựa chọn (`A`–`D`).
- Phần II — đúng/sai bốn ý (`a`–`d`).
- Phần III — trả lời ngắn.

Thang điểm in trong file đáp án: Phần I mỗi câu 0,25 điểm; Phần II mỗi câu tối đa 1,0 điểm (chia theo số ý đúng); Phần III mỗi câu 0,25 điểm.

Ngữ pháp `de.md` dưới đây khớp đúng với `tools/vi/de_thi_parts/parse.py`. Sai một chỗ, công cụ báo lỗi kèm số dòng; viết đúng ngay từ đầu thay vì đoán.

### Khối thông tin đề (meta)

Dòng đầu file là một dòng `---`. Sau đó mỗi dòng một cặp `khoá: giá trị`, đóng lại bằng một dòng `---` khác.

| Khoá | Bắt buộc | Ý nghĩa |
|---|---|---|
| `school` | có | Tên trường |
| `title` | có | Tên kỳ kiểm tra |
| `subject` | có | Môn và lớp |
| `time` | có | Số phút làm bài, chỉ chữ số |
| `department` | không | Tổ chuyên môn |
| `code` | không | Mã đề |
| `points` | không | Tổng điểm, mặc định `10` |

Dùng khoá nào ngoài bảy khoá trên là lỗi cú pháp.

### Các phần và câu hỏi

Sau khối meta là các dòng tiêu đề phần, đúng thứ tự: `## PART I`, `## PART II`, `## PART III`. Phần nào không dùng thì bỏ hẳn dòng đó, không để trống. Không dùng tiêu đề `## ` nào khác, trừ mục `## CAN SOAT` không bắt buộc ở cuối file. Mỗi câu mở bằng `### <số>`; số câu đếm lại từ 1 trong mỗi phần, đúng theo thứ tự xuất hiện.

Khoá dùng trong một câu:

| Khoá | Ở phần | Bắt buộc | Ý nghĩa |
|---|---|---|---|
| `en` | I, II, III | có | Đề bài tiếng Anh |
| `vi` | I, II, III | không | Bản tiếng Việt; thiếu thì chỉ là cảnh báo, không phải lỗi |
| `level` | I, II, III | có | Một trong `biet`, `hieu`, `vandung` |
| `topic` | I, II, III | không | Chủ đề, dùng dựng ma trận |
| `why` | I, II, III | không | Giải thích ngắn, chỉ in trong file đáp án |
| `A`–`D` | I | có, đủ bốn | Bốn lựa chọn |
| `key` | I | có | Một chữ trong `A`–`D` |
| `a`–`d` | II | có, đủ bốn | Bốn ý, mỗi dòng kết bằng ` | T` hoặc ` | F`; phần nội dung trước dấu `|` không được để trống |
| `key` | III | có | Đáp số, phải là một số: `36`, `12.5`, `12,5`, `-0.25` — không phải chữ, không phải `12.` hay `.5` |
| `unit` | III | không | Đơn vị của đáp số |

Phần II **không** dùng khoá `key`; ghi khoá này ở Phần II là lỗi cú pháp. Mỗi khoá chỉ được ghi một lần trong một câu.

Trong `en:` và `vi:` chỉ dùng ba dấu đánh dấu: `~ ~` (chỉ số dưới, ví dụ `H~2~O`), `^ ^` (chỉ số trên, ví dụ `m/s^2^`), `**` (in đậm, ví dụ `**not**`). Không dùng Markdown nào khác trong hai trường này.

Sau phần cuối, file có thể có mục `## CAN SOAT`, phải là mục cuối cùng: mỗi dòng bắt đầu bằng `- ` là một việc AI muốn thầy cô soát lại. Mục này in ở cuối `dap-an.docx` thành "Cần thầy cô soát"; không có mục này thì file đáp án ghi "Không có mục nào cần soát".

### Ví dụ đầy đủ

```
---
school: TRƯỜNG THCS VÍ DỤ
title: ĐỀ KIỂM TRA GIỮA HỌC KÌ I
subject: KHTN — Grade 8
time: 45
points: 10
---
## PART I

### 1
en: Which of the following is **not** a state of matter?
vi: Chất nào sau đây không phải là một trạng thái của vật chất?
level: biet
topic: Trạng thái của vật chất
A: Solid
B: Liquid
C: Gas
D: Energy
key: D

## PART II

### 1
en: Consider the following statements about density.
vi: Xét các phát biểu sau về khối lượng riêng.
level: hieu
topic: Khối lượng riêng
a: Density is mass divided by volume. | T
b: Density has the unit of newton. | F
c: Two objects of the same volume always have the same density. | F
d: Density can be used to identify a material. | T

## PART III

### 1
en: A block has a mass of 2 kg and a volume of 0.001 m^3^. Calculate its density.
vi: Một khối vật chất có khối lượng 2 kg và thể tích 0,001 m^3^. Tính khối lượng riêng của nó.
level: vandung
topic: Khối lượng riêng
key: 2000
unit: kg/m^3^

## CAN SOAT

- Thuật ngữ "trạng thái vật chất" giữ nguyên state of matter; cần thầy cô soát lại cách diễn đạt.
```

## Đầu ra

`projects/_de-thi/<tên_đề>/` chứa:
- `de.md` — nguồn duy nhất, sinh ra cả ba file dưới.
- `de-en.docx` — đề tiếng Anh để in.
- `de-song-ngu.docx` — bản song ngữ Anh–Việt để tổ soát.
- `dap-an.docx` — đáp án, thang điểm, ma trận đặc tả và mục "Cần thầy cô soát".

## Ghi vào brief

- Loại việc: Soạn đề KHTN tiếng Anh.
- Brief lưu tại `projects/_de-thi/<tên_đề>/brief.md`, viết theo đúng bố cục của docs/vi/tro-ly/mau-brief.md.
- Thầy cô yêu cầu: ghi đúng lời thầy cô cho từng câu hỏi bắt buộc và tuỳ chọn.
- AI đề xuất (thầy cô đã đồng ý): các gợi ý thầy cô chấp nhận.
- Loại việc này không dùng các bước dành cho PPTX: không kết thúc tin nhắn hỏi bằng dòng chốt cách xác nhận, không chạy `project_manager.py import-sources`, không chạy `project_manager.py init`, và không có bước xác nhận của upstream. Thầy cô trả lời xong lượt hỏi thì viết `de.md` rồi chạy lệnh ngay.
