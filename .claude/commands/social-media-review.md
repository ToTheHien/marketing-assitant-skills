---
description: Review bài đăng mạng xã hội từ Google Drive link. Dùng khi cần review nội dung Facebook, LinkedIn, hoặc Website landing page cho Becamex VSIP.
---

Bạn vừa được gọi để review một bài đăng mạng xã hội. Đối số nhận được:

$ARGUMENTS

**Bước 1 — Phân tích đối số**

Đối số truyền vào theo thứ tự:
1. `google_drive_link` — link Google Drive chứa nội dung bài viết
2. `platform` — nền tảng đăng bài: `Facebook`, `LinkedIn`, hoặc `Website`
3. `objective` — mục tiêu bài đăng (ví dụ: "quảng bá khu công nghiệp", "thu hút nhà đầu tư FDI")

Nếu đối số được truyền theo định dạng: `<link> <platform> <objective>`, hãy tách ra tương ứng. Nếu thiếu đối số nào, hỏi lại người dùng trước khi tiếp tục.

**Bước 2 — Đọc nội dung từ Google Drive**

Sử dụng `gws-drive` skill để truy cập tài liệu Google Drive tại link trên và trích xuất toàn bộ nội dung văn bản.

**Bước 3 — Review bài viết**

Áp dụng skill `social-media-review` để đánh giá nội dung theo nền tảng và mục tiêu đã xác định.

Lưu ý khi review:
- Nền tảng đánh giá là `platform` từ đối số
- Mục tiêu bài đăng (`objective`) cần được cân nhắc khi đánh giá mức độ phù hợp nội dung

**Bước 4 — Lưu kết quả**

1. Lấy ngày hiện tại theo định dạng `YYYY-MM-DD`
2. Tạo folder `social_media/{YYYY-MM-DD}/` tại thư mục gốc của project (nếu chưa tồn tại)
3. Xác định tên file output:
   - Nếu `review.md` chưa tồn tại trong folder → dùng `review.md`
   - Nếu đã tồn tại → kiểm tra lần lượt `review(2).md`, `review(3).md`, ... cho đến khi tìm được tên chưa tồn tại
4. Lưu toàn bộ output review vào file vừa xác định bên trên
5. Thông báo đường dẫn đầy đủ của file đã lưu
