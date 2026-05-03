---
name: investment-progress-summary
description: "Tạo báo cáo tiến độ đầu tư (investment progress report) dưới dạng file .docx cho nhà đầu tư tại khu công nghiệp. Kích hoạt skill này khi người dùng muốn: tạo báo cáo tiến độ đầu tư tuần hoặc tháng, cập nhật tiến độ triển khai dự án cho nhà đầu tư, tổng hợp tình hình giải ngân vốn, hoặc khi có từ khóa như 'báo cáo tiến độ đầu tư', 'investment progress', 'investor update', 'weekly update', 'monthly investment summary', 'tiến độ nhà đầu tư', 'báo cáo tuần', 'báo cáo tháng đầu tư'. Also trigger when the user provides data files (spreadsheets, Drive links, documents) and asks to create a structured investor progress Word document."
---

# Investment Progress Summary

Skill này tạo báo cáo tiến độ đầu tư dạng `.docx` cho nhà đầu tư tại khu công nghiệp Becamex, sử dụng templates từ `templates/` folder.

## Input từ người dùng

| Tham số | Bắt buộc | Mô tả |
|---------|----------|-------|
| **Mode** | Có | `weekly` hoặc `monthly` |
| **Reference files** | Không | Google Drive links hoặc đường dẫn file local (xlsx, docx, pdf, csv) |
| **Context** | Nên có | Tổng quan nội dung cần báo cáo trong kỳ này |

Nếu người dùng không chỉ rõ, hỏi: "Báo cáo tuần hay tháng? Bạn có file dữ liệu tham chiếu không?"

## Templates

Templates lưu tại `templates/`:
- `weekly_template.docx` — Template báo cáo tuần
- `monthly_template.docx` — Template báo cáo tháng

**Nếu template chưa tồn tại:** Tạo dummy templates bằng cách chạy:
```bash
python .claude/skills/investment-progress-summary/scripts/create_dummy_templates.py
```

Đọc template để hiểu cấu trúc (styles, heading levels, tables, placeholders) trước khi tạo báo cáo.

## Quy trình tạo báo cáo

### Bước 1: Tải và đọc dữ liệu tham chiếu

**Google Drive files:**
- Dùng MCP Google Drive tool để tải file về
- Ưu tiên: spreadsheets (xlsx/Google Sheets) > documents > PDFs

**Local files:**
- `.xlsx/.csv`: Đọc bằng Python pandas hoặc `xlsx` skill
- `.docx`: Unpack XML rồi extract text, hoặc dùng `python-docx`
- `.pdf`: Dùng `pdf` skill

**Không có file tham chiếu:** Dùng nguyên context người dùng cung cấp làm nguồn dữ liệu chính.

### Bước 2: Trích xuất thông tin chính

Từ dữ liệu nguồn, tìm kiếm và tổng hợp:
- **Danh sách nhà đầu tư**: tên công ty, tên dự án, lô đất/nhà xưởng
- **Tiến độ**: % hoàn thành, milestone đã đạt, hạng mục đang thực hiện
- **Tài chính**: vốn cam kết, đã giải ngân, tỷ lệ giải ngân
- **Vướng mắc / rủi ro**: issues, blockers, cần hỗ trợ gì
- **Kế hoạch**: công việc tuần/tháng tới

Nếu thông tin không rõ ràng hoặc thiếu → đánh dấu `[Cần bổ sung]`, không được bịa đặt số liệu.

### Bước 3: Xác định kỳ báo cáo

- Weekly: "Tuần X, từ DD/MM đến DD/MM/YYYY"
- Monthly: "Tháng MM/YYYY"
- Nếu không rõ từ file, lấy ngày hiện tại làm mốc

### Bước 4: Tạo file .docx

Sử dụng `docx` skill hoặc `python-docx` để tạo báo cáo theo cấu trúc phù hợp.
Xem `references/report_structure.md` để biết chi tiết từng section.

**Đặt tên file output:**
- Weekly: `investment_progress_weekly_YYYY-WNN.docx` (ví dụ: `investment_progress_weekly_2026-W18.docx`)
- Monthly: `investment_progress_monthly_YYYY-MM.docx` (ví dụ: `investment_progress_monthly_2026-04.docx`)

Lưu file vào thư mục hiện tại hoặc theo yêu cầu người dùng.

### Bước 5: Xác nhận với người dùng

Sau khi tạo xong:
1. Thông báo đường dẫn file output
2. Tóm tắt nhanh nội dung đã điền (số lượng nhà đầu tư, các điểm nổi bật)
3. Hỏi xem có cần chỉnh sửa gì không

## Ngôn ngữ báo cáo

**Mặc định: Tiếng Việt.**
- Giữ nguyên tên công ty/nhà đầu tư theo tên chính thức (tiếng Anh/Hàn/Nhật nếu là tên gốc)
- Technical terms: viết tiếng Anh kèm giải thích nếu cần
- Nếu người dùng yêu cầu song ngữ (Việt-Anh): tạo bảng 2 cột hoặc thêm phần dịch

## Lưu ý quan trọng

- Không bịa đặt số liệu. Thiếu thông tin → ghi `[Cần bổ sung]`
- Giữ nguyên tên chính thức của nhà đầu tư, dự án, khu công nghiệp
- Nếu context mâu thuẫn với file tham chiếu → ưu tiên file tham chiếu và ghi chú
- Khi template có watermark "DRAFT" hoặc "DUMMY" → giữ nguyên trong dummy templates, xóa đi khi điền nội dung thật
