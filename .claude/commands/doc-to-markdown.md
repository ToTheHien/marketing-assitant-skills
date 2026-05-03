---
description: Chuyển đổi file Word (.doc/.docx) sang Markdown. Nhận đường dẫn file local hoặc URL online và thư mục đầu ra.
---

Bạn vừa được gọi để chuyển đổi tài liệu Word sang Markdown. Đối số nhận được:

$ARGUMENTS

**Bước 1 — Phân tích đối số**

Đối số truyền vào theo thứ tự:
1. `input` — đường dẫn file local (`.doc`/`.docx`) hoặc URL online trỏ đến file Word. **Bắt buộc.**
2. `output` — đường dẫn thư mục local để lưu kết quả. **Bắt buộc.**

Quy tắc xác định `input`:
- Nếu bắt đầu bằng `http://` hoặc `https://`: đây là URL online → cần tải về trước
- Ngược lại: đây là đường dẫn file local

Nếu thiếu một trong hai đối số, dừng lại và yêu cầu người dùng cung cấp.

**Bước 2 — Chuẩn bị file đầu vào**

Nếu `input` là URL online:
1. Tạo thư mục tạm: `/tmp/doc_to_md_download/`
2. Tải file về thư mục tạm bằng `curl -L -o /tmp/doc_to_md_download/<tên_file> "<URL>"`
3. Gán lại `input` = đường dẫn file vừa tải về
4. Xác nhận file đã tải thành công (kiểm tra size > 0)

Nếu `input` là file local:
- Xác nhận file tồn tại và có đuôi `.doc` hoặc `.docx`
- Nếu không tìm thấy file, báo lỗi rõ ràng và dừng

**Bước 3 — Chuẩn bị thư mục đầu ra**

1. Tạo thư mục `output` nếu chưa tồn tại: `mkdir -p "<output>"`
2. Xác định đường dẫn tuyệt đối của `output`

**Bước 4 — Chuyển đổi tài liệu**

Sử dụng skill `doc-to-markdown` để thực hiện chuyển đổi.

Chạy lệnh:
```bash
SKILL_DIR="/home/thehien/Projects/becamex/marketing_assitance_w_claude/.claude/skills/doc-to-markdown"
python3 "$SKILL_DIR/convert_word_to_markdown.py" "<input>"
```

Sau khi chạy xong, script sẽ tạo ra file `.md` và thư mục `*_images/` (nếu có hình ảnh) **cùng thư mục với file input**.

**Bước 5 — Di chuyển kết quả sang thư mục output**

1. Xác định tên file gốc (không có đuôi): `basename "<input>" | sed 's/\.[^.]*$//'`
2. Di chuyển file `.md` sang thư mục output:
   ```bash
   mv "<thư_mục_input>/<tên_file>.md" "<output>/"
   ```
3. Nếu tồn tại thư mục `<tên_file>_images/`, di chuyển luôn:
   ```bash
   mv "<thư_mục_input>/<tên_file>_images/" "<output>/" 2>/dev/null || true
   ```

**Bước 6 — Báo cáo kết quả**

Thông báo:
- Đường dẫn đầy đủ file Markdown đã tạo
- Số hình ảnh được trích xuất (nếu có)
- Kích thước file Markdown
- Bất kỳ cảnh báo hoặc lỗi nào phát sinh trong quá trình chuyển đổi
