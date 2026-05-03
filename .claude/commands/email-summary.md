---
description: Tóm tắt email trong hộp thư đến. Dùng khi cần xem tổng quan email theo số ngày gần nhất và theo danh mục.
---

Bạn vừa được gọi để tóm tắt email. Đối số nhận được:

$ARGUMENTS

**Bước 1 — Phân tích đối số**

Đối số truyền vào theo thứ tự (tất cả đều tùy chọn):
1. `days` — số ngày gần nhất cần kiểm tra email. **Mặc định: 7**
2. `category` — danh mục email cần lọc (ví dụ: "work", "newsletter", "invoice"). **Mặc định: All**

Quy tắc xử lý đối số:
- Nếu không có đối số nào: dùng `days=7`, `category=All`
- Nếu chỉ có 1 đối số dạng số: đó là `days`
- Nếu chỉ có 1 đối số dạng chữ: đó là `category`
- Nếu có 2 đối số: đối số 1 là `days`, đối số 2 là `category`

**Bước 2 — Truy xuất email**

Sử dụng Gmail MCP tool để tìm email trong khoảng thời gian `days` ngày gần nhất.

Query tìm kiếm: `newer_than:{days}d` (kết hợp thêm bộ lọc category nếu `category` không phải `All`).

Lấy tối đa 50 email, bao gồm: subject, sender, date, và snippet nội dung.

**Bước 3 — Phân tích và tóm tắt**

Tóm tắt email theo cấu trúc sau:

## Tổng quan
- Tổng số email nhận được trong `{days}` ngày qua
- Phân bổ theo người gửi / tổ chức chính
- Danh mục nổi bật (nếu có nhiều loại)

## Email cần xử lý (Action Required)
Liệt kê email có yêu cầu hành động hoặc deadline cụ thể:
- **[Tiêu đề]** — *Từ: [Người gửi]* — [Tóm tắt 1 dòng] — **Deadline: [ngày nếu có]**

## Email thông tin quan trọng
Liệt kê email có thông tin đáng chú ý nhưng không cần phản hồi ngay:
- **[Tiêu đề]** — *Từ: [Người gửi]* — [Tóm tắt 1 dòng]

## Email tham khảo / Newsletter
Nhóm gọn các email ít quan trọng (newsletter, thông báo tự động, v.v.):
- [Tên newsletter/sender]: {số lượng} email

## Khuyến nghị
Đề xuất 2-3 hành động ưu tiên dựa trên nội dung email đã phân tích.

**Bước 4 — Lưu kết quả**

1. Lấy ngày hiện tại theo định dạng `YYYY-MM-DD`
2. Tạo folder `email_summary/{YYYY-MM-DD}/` tại thư mục gốc của project (nếu chưa tồn tại)
3. Xác định tên file output:
   - Nếu `summary.md` chưa tồn tại trong folder → dùng `summary.md`
   - Nếu đã tồn tại → kiểm tra lần lượt `summary(2).md`, `summary(3).md`, ... cho đến khi tìm được tên chưa tồn tại
4. Lưu toàn bộ output tóm tắt vào file vừa xác định bên trên
5. Thông báo đường dẫn đầy đủ của file đã lưu
