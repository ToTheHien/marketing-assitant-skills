---
description: Tạo báo cáo hoạt động xúc tiến đầu tư (weekly/monthly) dưới dạng .docx. Đọc file tham chiếu từ Google Drive hoặc local, tổng hợp nội dung, lưu output vào promote-activity/{mode}/{YYYY-MM-DD HH:MM}/.
---

Bạn vừa được gọi để tạo báo cáo hoạt động xúc tiến đầu tư. Đối số nhận được:

$ARGUMENTS

**Bước 1 — Phân tích đối số**

Đối số truyền vào theo cú pháp:

```
<mode> [references...] [--context <context_text>]
```

Quy tắc phân tích:
1. `mode` — đối số đầu tiên: `weekly` hoặc `monthly` (bắt buộc)
2. `references` — danh sách các đường dẫn file tham chiếu, phân cách bởi dấu cách. Mỗi reference là:
   - Google Drive link (bắt đầu bằng `https://drive.google.com/` hoặc `https://docs.google.com/`)
   - Đường dẫn file local (kết thúc bằng `.xlsx`, `.csv`, `.docx`, `.pdf`, hoặc bắt đầu bằng `/` hoặc `./`)
3. `--context` — phần nội dung sau cờ `--context` cho đến cuối dòng là context bổ sung (không bắt buộc)

Ví dụ hợp lệ:
- `weekly` — chỉ có mode, không có file
- `monthly https://drive.google.com/file/abc123` — mode + 1 Drive file
- `weekly ./data/tuan18.xlsx --context Tuần này có 2 đoàn site visit, gặp KCCI ngày thứ 4`
- `monthly https://drive.google.com/file/abc https://drive.google.com/file/xyz --context Tháng 4, tổ chức 1 roadshow tại Nhật Bản`

Nếu `mode` không phải `weekly` hoặc `monthly`, dừng lại và hỏi lại người dùng.

**Bước 2 — Tải và đọc dữ liệu tham chiếu**

Xử lý từng reference theo loại:

**Google Drive links:**
- Dùng `gws-drive` skill hoặc MCP Google Drive tool để tải nội dung
- Với Google Sheets/Spreadsheet: trích xuất dữ liệu dạng bảng
- Với Google Docs: trích xuất toàn bộ văn bản
- Với file binary (xlsx, pdf): tải về local trước

**Local files:**
- `.xlsx` hoặc `.csv`: Đọc bằng Python pandas hoặc `xlsx` skill, lấy toàn bộ dữ liệu
- `.docx`: Dùng `python-docx` để extract text và tables
- `.pdf`: Dùng `pdf` skill để extract text

**Không có references:** Dùng context người dùng cung cấp làm nguồn dữ liệu duy nhất.

**Bước 3 — Trích xuất thông tin xúc tiến đầu tư**

Từ dữ liệu đọc được và context, tổng hợp:

**Gặp gỡ nhà đầu tư tiềm năng:**
- Tên công ty, quốc gia, ngành nghề
- Hình thức (site visit, online meeting, trao đổi qua email/điện thoại)
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
- Số lượng đại biểu / contacts thu được
- Leads tiềm năng từ sự kiện

Thiếu thông tin → đánh dấu `[Cần bổ sung]`, không bịa đặt số liệu.

Context người dùng cung cấp qua `--context` có thể chứa thông tin nằm ngoài phạm vi file tham chiếu (sự kiện đặc biệt, quyết định nội bộ, điểm nhấn quan trọng) — hãy tích hợp vào báo cáo.

Nếu context mâu thuẫn với file tham chiếu → ưu tiên file tham chiếu, ghi chú sự khác biệt.

**Bước 4 — Xác định kỳ báo cáo**

- `weekly`: Xác định tuần từ ngày trong file tham chiếu hoặc ngày hiện tại → format "Tuần XX, từ DD/MM đến DD/MM/YYYY"
- `monthly`: Xác định tháng → format "Tháng MM/YYYY"

**Bước 5 — Tạo file báo cáo .docx**

Kích hoạt skill `promote-activity-summary` để tạo báo cáo với toàn bộ dữ liệu đã tổng hợp ở trên.

Đặt tên file output:
- Weekly: `promote_activity_weekly_YYYY-WNN.docx` (ví dụ: `promote_activity_weekly_2026-W18.docx`)
- Monthly: `promote_activity_monthly_YYYY-MM.docx` (ví dụ: `promote_activity_monthly_2026-04.docx`)

**Bước 6 — Lưu output**

1. Lấy ngày hiện tại theo định dạng `YYYY-MM-DD`
2. Xác định thư mục lưu:
   - Weekly: `promote-activity/weekly/{YYYY-MM-DD}/`
   - Monthly: `promote-activity/monthly/{YYYY-MM-DD}/`
3. Tạo folder nếu chưa tồn tại (kể cả toàn bộ path)
4. Xác định tên file output:
   - Nếu tên file chưa tồn tại trong folder → dùng tên gốc (ví dụ: `promote_activity_weekly_2026-W18.docx`)
   - Nếu đã tồn tại → thêm suffix `(n)` vào trước phần mở rộng, tăng dần cho đến khi tìm được tên chưa tồn tại (ví dụ: `promote_activity_weekly_2026-W18(2).docx`)
5. Lưu file `.docx` vào folder vừa xác định
6. Nếu có notes hoặc dữ liệu thô trung gian đáng giữ lại, lưu thêm file `notes.md` trong cùng folder (nếu đã tồn tại thì dùng `notes(2).md`, v.v.)

**Bước 7 — Xác nhận với người dùng**

Sau khi hoàn thành:
1. Thông báo đường dẫn đầy đủ của file output
2. Tóm tắt nhanh: số cuộc gặp, số hoạt động phòng thương mại, số sự kiện, các điểm nổi bật trong kỳ
3. Liệt kê các mục `[Cần bổ sung]` nếu có, đề nghị người dùng cung cấp thêm thông tin
