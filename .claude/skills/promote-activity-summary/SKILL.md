---
name: promote-activity-summary
description: "Tạo báo cáo hoạt động xúc tiến đầu tư (investment promotion activity report) dưới dạng file .docx. Kích hoạt skill này khi người dùng muốn: tổng hợp hoạt động xúc tiến đầu tư tuần hoặc tháng, cập nhật tình hình tiếp xúc nhà đầu tư tiềm năng, báo cáo kết quả gặp gỡ phòng thương mại hay hiệp hội ngành nghề, tổng kết hội nghị hoặc sự kiện xúc tiến đầu tư, khi có từ khóa như 'xúc tiến đầu tư', 'promote activity', 'investor outreach', 'tóm tắt hoạt động xúc tiến', 'báo cáo xúc tiến tuần', 'báo cáo xúc tiến tháng', 'kết quả gặp nhà đầu tư', 'chambers of commerce meeting', 'KCCI', 'KOCHAM', 'JETRO', 'EuroCham', 'roadshow', 'investment promotion event'. Also trigger when the user provides meeting notes, event summaries, or Drive files about investor outreach and wants a structured Word report."
---

# Promote Activity Summary

Skill này tạo báo cáo hoạt động xúc tiến đầu tư dạng `.docx` cho đội ngũ marketing/xúc tiến của Becamex IDC, theo templates từ `templates/` folder.

**Khác biệt với investment-progress-summary:** Skill này tập trung vào hoạt động tìm kiếm và tiếp cận nhà đầu tư MỚI (outreach, pipeline), không phải tiến độ xây dựng/giải ngân của nhà đầu tư đã ký kết.

## Input từ người dùng

| Tham số | Bắt buộc | Mô tả |
|---------|----------|-------|
| **Mode** | Có | `weekly` hoặc `monthly` |
| **Reference files** | Không | Google Drive links hoặc đường dẫn file local (xlsx, docx, pdf, csv) |
| **Context** | Nên có | Tổng quan hoạt động trong kỳ, thông tin chưa có trong file tham chiếu |

Nếu người dùng không chỉ rõ, hỏi: "Báo cáo tuần hay tháng? Bạn có file ghi chú cuộc họp hoặc dữ liệu tham chiếu không?"

## Templates

Templates lưu tại `templates/`:
- `weekly_template.docx` — Template báo cáo tuần
- `monthly_template.docx` — Template báo cáo tháng

**Nếu template chưa tồn tại:** Tạo dummy templates bằng cách chạy:
```bash
python .claude/skills/promote-activity-summary/scripts/create_dummy_templates.py
```

Đọc template để hiểu cấu trúc (styles, heading levels, tables, placeholders) trước khi tạo báo cáo.

## Quy trình tạo báo cáo

### Bước 1: Tải và đọc dữ liệu tham chiếu

**Google Drive files:**
- Dùng MCP Google Drive tool để tải file về
- Ưu tiên: spreadsheets > documents > PDFs

**Local files:**
- `.xlsx/.csv`: Đọc bằng Python pandas
- `.docx`: Unpack XML rồi extract text, hoặc dùng `python-docx`
- `.pdf`: Dùng `pdf` skill

**Không có file:** Dùng context người dùng cung cấp làm nguồn chính.

### Bước 2: Trích xuất thông tin xúc tiến đầu tư

Từ dữ liệu nguồn, tìm kiếm và tổng hợp:

**Gặp gỡ nhà đầu tư tiềm năng:**
- Tên công ty, quốc gia, ngành nghề
- Nội dung cuộc gặp (site visit, online meeting, trao đổi qua email/điện thoại)
- Quy mô dự án quan tâm (diện tích, vốn dự kiến)
- Kết quả và trạng thái pipeline (Mới tiếp cận / Đang quan tâm / Đang đàm phán / Chờ quyết định)
- Bước tiếp theo và người phụ trách

**Phòng thương mại & Hiệp hội:**
- Tên tổ chức (KCCI, KOCHAM, JETRO, EuroCham, AmCham, VCCI, JCCI, v.v.)
- Loại hoạt động (cuộc họp, hội thảo, trao đổi thông tin, kết nạp thành viên)
- Kết quả và follow-up

**Sự kiện xúc tiến:**
- Tên sự kiện, địa điểm, ngày tổ chức
- Becamex tổ chức hay tham dự
- Số lượng đại biểu/contacts thu được
- Leads tiềm năng từ sự kiện

**Thông tin còn thiếu:** Đánh dấu `[Cần bổ sung]`, không bịa đặt.

### Bước 3: Xác định kỳ báo cáo

- Weekly: "Tuần X, từ DD/MM đến DD/MM/YYYY"
- Monthly: "Tháng MM/YYYY"
- Nếu không rõ từ file → dùng ngày hiện tại làm mốc

### Bước 4: Tạo file .docx

Sử dụng `docx` skill (JavaScript docx-js) để tạo báo cáo theo cấu trúc trong `references/report_structure.md`.

**Đặt tên file output:**
- Weekly: `promote_activity_weekly_YYYY-WNN.docx`
- Monthly: `promote_activity_monthly_YYYY-MM.docx`

Lưu file vào thư mục làm việc hiện tại hoặc theo yêu cầu người dùng.

### Bước 5: Xác nhận với người dùng

1. Thông báo đường dẫn file output
2. Tóm tắt nhanh: số cuộc gặp, số sự kiện, điểm nổi bật
3. Hỏi xem có cần chỉnh sửa không

## Ngôn ngữ báo cáo

**Mặc định: Tiếng Việt.**
- Tên công ty/nhà đầu tư: giữ nguyên tên chính thức (tiếng Anh/Hàn/Nhật/Trung)
- Tên phòng thương mại: dùng tên viết tắt quen thuộc (KCCI, JETRO, v.v.)
- Nếu người dùng yêu cầu song ngữ → thêm cột/phần tiếng Anh

## Lưu ý quan trọng

- Không bịa đặt tên công ty, số liệu, hay kết quả cuộc gặp. Thiếu thông tin → ghi `[Cần bổ sung]`
- Nếu context mâu thuẫn với file tham chiếu → ưu tiên file, ghi chú sự khác biệt
- Pipeline status cần nhất quán: dùng đúng 4 trạng thái chuẩn (Mới tiếp cận / Đang quan tâm / Đang đàm phán / Chờ quyết định)
- Xem `references/report_structure.md` để biết chi tiết từng section và màu sắc chuẩn
