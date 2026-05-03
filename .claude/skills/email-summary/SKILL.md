---
name: email-summary
version: 1.0.0
description: >
  Kiểm tra, tổng hợp và phân loại email trong N ngày gần nhất từ Gmail. Đánh giá mức độ ưu tiên (Khẩn cấp/Cao/Trung bình/Thấp) cho từng email, sau đó đề xuất nội dung phản hồi dựa trên template có sẵn theo từng loại email. Sử dụng skill này khi người dùng muốn xem lại inbox, tổng hợp email tuần này/tháng này, phân loại thư đến, soạn thảo phản hồi nhanh, hoặc nói bất kỳ điều gì như "kiểm tra email", "xem mail gần đây", "tổng hợp inbox", "email mấy ngày nay", "soạn phản hồi email". Luôn sử dụng skill này khi người dùng đề cập đến email, inbox, thư điện tử, hay muốn trả lời email.
---

# Email Summary — Tổng hợp & Phân loại Email

## Mục tiêu

Giúp người dùng nắm nhanh tình hình inbox, hiểu rõ ưu tiên xử lý, và soạn phản hồi phù hợp — tất cả trong một lần chạy. Đây là công cụ tiết kiệm thời gian, không phải công cụ đọc email thủ công.

## Tham số đầu vào

| Tham số | Mô tả | Mặc định |
|---------|-------|---------|
| `n` | Số ngày cần kiểm tra | `7` |
| `max` | Số email tối đa xử lý | `30` |
| `category` | Lọc theo loại cụ thể (tùy chọn) | Tất cả |

**Ví dụ cách gọi:**
- "kiểm tra email 7 ngày gần nhất"
- "xem inbox 3 ngày qua, tối đa 20 email"
- "tổng hợp email tuần này về hợp tác"

## Quy trình thực hiện

### Bước 1 — Xác định khoảng thời gian

Tính ngày bắt đầu từ hôm nay trừ N ngày. Ví dụ nếu hôm nay là 2026-04-22 và N=7 thì after:2026/04/15.

Dùng `mcp__claude_ai_Gmail__search_threads` với query:
```
after:YYYY/MM/DD -label:sent -label:drafts
```

> Lấy tất cả thread trong hộp thư đến, bỏ qua thư đã gửi và nháp để tránh trùng lặp.

### Bước 2 — Đọc nội dung email

Với mỗi thread trả về, dùng `mcp__claude_ai_Gmail__get_thread` để lấy nội dung. Nếu thread có nhiều tin nhắn, chỉ cần đọc tin nhắn đầu tiên (để phân loại) và tin nhắn cuối cùng (để biết trạng thái mới nhất).

Giới hạn ở `max` email để tránh quá tải. Nếu số email > max, thông báo cho người dùng biết.

### Bước 3 — Phân loại email

Đọc file `reply-templates.md` để nắm rõ các loại email và cách phân biệt. Phân loại mỗi email vào một trong các nhóm sau:

| Loại | Mô tả ngắn |
|------|-----------|
| `hop-tac` | Đề xuất hợp tác, liên kết kinh doanh |
| `nha-dau-tu` | Hỏi thăm đầu tư, thuê đất/nhà xưởng |
| `truyen-thong` | Báo chí, truyền thông, phỏng vấn |
| `nha-cung-cap` | Chào hàng, dịch vụ, nhà cung cấp |
| `su-kien` | Mời tham dự sự kiện, hội thảo |
| `khieu-nai` | Phản ánh, khiếu nại, vấn đề phát sinh |
| `noi-bo` | Email từ nội bộ công ty |
| `hanh-chinh` | Thủ tục hành chính, hóa đơn, hợp đồng |
| `khac` | Không thuộc các loại trên |

### Bước 4 — Đánh giá mức độ ưu tiên

Ưu tiên dựa trên 3 yếu tố: **khẩn cấp** (deadline, từ cấp trên), **tầm quan trọng** (ảnh hưởng đến dự án lớn), **thời gian chờ** (email đã gửi lâu chưa được trả lời).

| Mức | Màu | Tiêu chí |
|-----|-----|---------|
| 🔴 Khẩn cấp | Đỏ | Có deadline trong 24h, từ ban lãnh đạo, khiếu nại cần xử lý ngay |
| 🟠 Cao | Cam | Cần phản hồi trong 2-3 ngày, nhà đầu tư tiềm năng, đối tác quan trọng |
| 🟡 Trung bình | Vàng | Cần phản hồi trong tuần, thông tin thông thường |
| 🟢 Thấp | Xanh | Không cần phản hồi gấp, thông tin tham khảo, FYI |

### Bước 5 — Xuất báo cáo tổng hợp

Xuất theo định dạng sau (không bỏ qua phần nào):

```
## 📊 Tổng quan Inbox — [N] ngày gần nhất ([ngày bắt đầu] → [hôm nay])

**Tổng số:** [X] email | **Chưa đọc:** [Y] | **Cần phản hồi:** [Z]

| # | Từ | Tiêu đề | Loại | Ưu tiên | Ngày |
|---|-----|---------|------|---------|------|
| 1 | ... | ... | ... | 🔴 Khẩn cấp | ... |
| 2 | ... | ... | ... | 🟠 Cao | ... |
...

---
## 📋 Phân loại theo nhóm
[Liệt kê số lượng theo từng loại]

---
## ✉️ Đề xuất phản hồi
[Xem Bước 6]
```

### Bước 6 — Đề xuất nội dung phản hồi

Sau khi xuất bảng tổng hợp, đề xuất phản hồi cho **các email có mức độ Khẩn cấp và Cao** (tối đa 5 email). Đọc `reply-templates.md` và lấy template phù hợp với loại email, sau đó điền thông tin cụ thể.

Với mỗi email cần phản hồi, xuất theo format:

```
### [#] Phản hồi: [Tiêu đề email gốc]
**Gửi đến:** [tên người gửi] <[email]>
**Loại:** [loại email]

---
[Nội dung phản hồi được điền sẵn từ template]
---

> 💡 Ghi chú: [Lý do chọn template này, điều cần lưu ý khi gửi]
```

Sau khi hiển thị tất cả phản hồi đề xuất, hỏi người dùng:
> "Bạn muốn tôi tạo draft cho email nào? (nêu số thứ tự hoặc 'tất cả')"

Nếu người dùng đồng ý, dùng `mcp__claude_ai_Gmail__create_draft` để tạo draft.

## Lưu ý khi xử lý

- Nếu email không có subject rõ ràng, dùng 5 từ đầu của nội dung làm tiêu đề
- Giữ bí mật thông tin nhạy cảm trong email — không log hay lưu nội dung email ra bên ngoài
- Nếu Gmail MCP chưa được xác thực, thông báo người dùng cần kết nối Gmail trước
- Ưu tiên hiển thị email chưa đọc (unread) lên đầu bảng
- Khi phân loại, nếu không chắc chắn thì chọn loại `khac` thay vì đoán sai

## File liên quan

- **`reply-templates.md`** — Đọc file này để lấy template phản hồi cho từng loại email. File chứa template mẫu cho tất cả 8 loại email. Đọc ngay khi bắt đầu Bước 3 để phân loại chính xác hơn.
