# Lấy API key để tạo ảnh bằng AI

## Khi nào cần

Chỉ cần API key khi bạn muốn AI tự vẽ ảnh minh hoạ, hoặc dùng giọng đọc đám mây chất lượng cao. Giọng thuyết minh mặc định (edge-tts) đã miễn phí và không cần key.

## Lấy Gemini API key

1. Vào https://aistudio.google.com/apikey.
2. Đăng nhập bằng tài khoản Google.
3. Bấm tạo key rồi sao chép key vừa tạo.

## Điền vào file .env

Mở file `.env` ở thư mục gốc của bộ công cụ bằng Notepad. Nếu chưa có file này, chạy `CAI-DAT.bat` để tạo trước.

Thêm hai dòng sau vào file:

```
IMAGE_BACKEND=gemini
GEMINI_API_KEY=dán-key-của-bạn-vào-đây
```

Lưu file lại. Muốn dùng dịch vụ AI khác, xem các dòng mẫu tương ứng trong `.env.example`.

## Kiểm tra

Chạy `KIEM-TRA.bat`. Dòng "API key dịch vụ AI" sẽ chuyển thành ✅.

## Bảo mật

- Không gửi file `.env` cho người khác.
- `.env` đã được Git bỏ qua nên sẽ không bị đưa lên GitHub khi bạn commit.
- Khi hỏi hỗ trợ, nhớ xoá key khỏi ảnh chụp màn hình hoặc log trước khi gửi.
