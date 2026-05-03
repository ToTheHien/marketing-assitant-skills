from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

OUTPUT_PATH = "/home/thehien/Projects/becamex/marketing_assitance_skills/.claude/skills/promote-activity-summary-workspace/iteration-1/weekly-w18-context-only/without_skill/outputs/promote_activity_weekly_2026-W18.docx"

doc = Document()

# ── Page margins ──────────────────────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.0)

# ── Helper: set paragraph shading ────────────────────────────────────────────
def shade_paragraph(para, hex_color):
    pPr = para._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    pPr.append(shd)

# ── Helper: set cell background ───────────────────────────────────────────────
def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

# ── Helper: add table border ──────────────────────────────────────────────────
def set_table_borders(table, color="CCCCCC", size="4"):
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), size)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), color)
        tblBorders.append(border)
    tblPr.append(tblBorders)

# ─────────────────────────────────────────────────────────────────────────────
# HEADER BLOCK
# ─────────────────────────────────────────────────────────────────────────────
p_org = doc.add_paragraph()
p_org.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p_org.add_run("BAN QUẢN LÝ CÁC KHU CÔNG NGHIỆP BÌNH DƯƠNG (BIZA)")
run.bold = True
run.font.size = Pt(11)
run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

p_dept = doc.add_paragraph()
p_dept.alignment = WD_ALIGN_PARAGRAPH.CENTER
run2 = p_dept.add_run("PHÒNG XÚC TIẾN ĐẦU TƯ")
run2.bold = True
run2.font.size = Pt(10)
run2.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

doc.add_paragraph()  # spacer

# Title
p_title = doc.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
shade_paragraph(p_title, "1F497D")
run_title = p_title.add_run("BÁO CÁO HOẠT ĐỘNG XÚC TIẾN ĐẦU TƯ HÀNG TUẦN")
run_title.bold = True
run_title.font.size = Pt(14)
run_title.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

p_week = doc.add_paragraph()
p_week.alignment = WD_ALIGN_PARAGRAPH.CENTER
shade_paragraph(p_week, "2E75B6")
run_week = p_week.add_run("Tuần 18 | 28/04/2026 – 02/05/2026")
run_week.bold = True
run_week.font.size = Pt(12)
run_week.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

doc.add_paragraph()

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 – Tóm tắt tổng quan
# ─────────────────────────────────────────────────────────────────────────────
def add_section_heading(doc, number, text):
    p = doc.add_paragraph()
    shade_paragraph(p, "D6E4F0")
    run = p.add_run(f"  {number}. {text}")
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
    return p

add_section_heading(doc, "I", "TÓM TẮT TỔNG QUAN")

summary_data = [
    ("Tổng số cuộc gặp / liên hệ nhà đầu tư", "03"),
    ("Số buổi họp / sự kiện đối ngoại",         "01 (Họp KOCHAM)"),
    ("Số nhà đầu tư mới tiếp cận",              "01 (Công ty phụ tùng ô tô Nhật Bản)"),
    ("Khu công nghiệp liên quan",               "VSIP II, MTPII"),
    ("Tổng diện tích đang được quan tâm",       "≥ 7 ha (Samsung C&T: 5 ha, Nhật Bản: 2 ha)"),
]

tbl1 = doc.add_table(rows=1, cols=2)
tbl1.style = 'Table Grid'
set_table_borders(tbl1, color="9DC3E6", size="4")
tbl1.columns[0].width = Cm(8.5)
tbl1.columns[1].width = Cm(6.5)

# header row
hdr = tbl1.rows[0].cells
set_cell_bg(hdr[0], "2E75B6")
set_cell_bg(hdr[1], "2E75B6")
for i, txt in enumerate(["Chỉ tiêu", "Kết quả"]):
    p = hdr[i].paragraphs[0]
    r = p.add_run(txt)
    r.bold = True
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    r.font.size = Pt(10)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

for label, value in summary_data:
    row = tbl1.add_row().cells
    row[0].text = label
    row[0].paragraphs[0].runs[0].font.size = Pt(10)
    row[1].text = value
    row[1].paragraphs[0].runs[0].font.bold = True
    row[1].paragraphs[0].runs[0].font.size = Pt(10)
    row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_paragraph()

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 – Chi tiết các cuộc gặp nhà đầu tư
# ─────────────────────────────────────────────────────────────────────────────
add_section_heading(doc, "II", "CHI TIẾT CÁC CUỘC GẶP / TIẾP XÚC NHÀ ĐẦU TƯ")

meetings = [
    {
        "stt": "1",
        "company": "Samsung C&T Corporation",
        "country": "Hàn Quốc",
        "sector": "Logistics / Kho vận",
        "form": "Thăm thực địa (Site Visit)",
        "date": "29/04/2026",
        "location": "VSIP II, Bình Dương",
        "area": "5 ha",
        "status": "Quan tâm – chờ đề xuất chính thức",
        "note": "Đoàn đã tham quan khu đất tại VSIP II. Nhà đầu tư đánh giá cao vị trí và hạ tầng; yêu cầu nhận đề xuất thương mại trong tuần tới.",
    },
    {
        "stt": "2",
        "company": "Foxconn Technology Group",
        "country": "Đài Loan",
        "sector": "Điện tử / Công nghệ",
        "form": "Họp trực tuyến (Online Meeting)",
        "date": "30/04/2026",
        "location": "Khu công nghiệp MTPII",
        "area": "Đang thảo luận",
        "status": "Đang đàm phán",
        "note": "Trao đổi các điều kiện đầu tư tại MTPII, tập trung vào tiến độ hạ tầng và chính sách ưu đãi. Hai bên sẽ tiếp tục làm việc kỹ thuật.",
    },
    {
        "stt": "3",
        "company": "Công ty phụ tùng ô tô (Nhật Bản – chưa tiết lộ tên)",
        "country": "Nhật Bản",
        "sector": "Phụ tùng ô tô / Sản xuất",
        "form": "Liên hệ qua email",
        "date": "Tuần 18/2026",
        "location": "Đang khảo sát",
        "area": "2 ha",
        "status": "Liên hệ mới – chờ phản hồi chi tiết",
        "note": "Doanh nghiệp gửi email hỏi về quỹ đất 2 ha tại khu công nghiệp. Đây là đầu mối mới, cần thu thập thêm thông tin nhu cầu cụ thể.",
    },
]

col_labels = ["STT", "Công ty", "Quốc gia", "Lĩnh vực", "Hình thức", "Ngày", "KCN / Khu vực", "Diện tích", "Trạng thái", "Ghi chú"]
col_keys   = ["stt", "company", "country", "sector", "form", "date", "location", "area", "status", "note"]
col_widths = [Cm(0.9), Cm(4.0), Cm(1.8), Cm(2.8), Cm(2.5), Cm(1.8), Cm(2.6), Cm(1.8), Cm(2.8), Cm(4.5)]

tbl2 = doc.add_table(rows=1, cols=len(col_labels))
tbl2.style = 'Table Grid'
set_table_borders(tbl2, color="9DC3E6", size="4")

for i, (label, width) in enumerate(zip(col_labels, col_widths)):
    cell = tbl2.rows[0].cells[i]
    cell.width = width
    set_cell_bg(cell, "1F497D")
    p = cell.paragraphs[0]
    r = p.add_run(label)
    r.bold = True
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

row_colors = ["FFFFFF", "EBF3FB"]
for idx, meeting in enumerate(meetings):
    row_cells = tbl2.add_row().cells
    bg = row_colors[idx % 2]
    for i, key in enumerate(col_keys):
        c = row_cells[i]
        set_cell_bg(c, bg)
        p = c.paragraphs[0]
        r = p.add_run(meeting[key])
        r.font.size = Pt(9)
        if key == "status":
            if "Quan tâm" in meeting[key]:
                r.font.color.rgb = RGBColor(0x1F, 0x7A, 0x1F)
                r.bold = True
            elif "đàm phán" in meeting[key].lower():
                r.font.color.rgb = RGBColor(0xC5, 0x5A, 0x11)
                r.bold = True
            else:
                r.font.color.rgb = RGBColor(0x00, 0x70, 0xC0)
        if key in ("stt",):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_paragraph()

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 – Sự kiện đối ngoại
# ─────────────────────────────────────────────────────────────────────────────
add_section_heading(doc, "III", "SỰ KIỆN / HOẠT ĐỘNG ĐỐI NGOẠI")

p = doc.add_paragraph()
r = p.add_run("Họp giao lưu KOCHAM (Hiệp hội Doanh nghiệp Hàn Quốc tại Việt Nam)")
r.bold = True
r.font.size = Pt(10)
r.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

kocham_info = [
    ("Thời gian",         "02/05/2026"),
    ("Địa điểm",          "TP. Hồ Chí Minh / Bình Dương"),
    ("Đơn vị tham dự",    "Phòng Xúc tiến đầu tư – BIZA"),
    ("Nội dung chính",    "Tham dự buổi gặp gỡ doanh nghiệp Hàn Quốc, giới thiệu tiềm năng các khu công nghiệp tại Bình Dương."),
    ("Kết quả nổi bật",   "Tiếp nhận và ghi nhận thông tin 02 doanh nghiệp Hàn Quốc có nhu cầu tìm hiểu đầu tư vào Bình Dương. Đây là đầu mối tiềm năng, sẽ được theo dõi và liên hệ trong các tuần tiếp theo."),
]

tbl3 = doc.add_table(rows=len(kocham_info), cols=2)
tbl3.style = 'Table Grid'
set_table_borders(tbl3, color="9DC3E6", size="4")
tbl3.columns[0].width = Cm(4.5)
tbl3.columns[1].width = Cm(11.0)

for i, (label, value) in enumerate(kocham_info):
    row = tbl3.rows[i].cells
    bg = "D6E4F0" if i % 2 == 0 else "EBF3FB"
    set_cell_bg(row[0], bg)
    set_cell_bg(row[1], "FFFFFF")
    p0 = row[0].paragraphs[0]
    r0 = p0.add_run(label)
    r0.bold = True
    r0.font.size = Pt(10)
    row[1].text = value
    row[1].paragraphs[0].runs[0].font.size = Pt(10)

doc.add_paragraph()

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4 – Kế hoạch tuần tới
# ─────────────────────────────────────────────────────────────────────────────
add_section_heading(doc, "IV", "KẾ HOẠCH TUẦN TỚI (TUẦN 19 – 05/05 – 09/05/2026)")

next_week_items = [
    ("Gửi đề xuất thương mại cho Samsung C&T", "Trước 05/05/2026", "Phòng XTĐT", "Ưu tiên cao"),
    ("Họp với JETRO (Tổ chức Xúc tiến Thương mại Nhật Bản)", "07/05/2026", "Phòng XTĐT + Lãnh đạo BQL", "Ưu tiên cao"),
    ("Tiếp tục theo dõi và liên hệ 02 đầu mối từ KOCHAM", "Trong tuần 19", "Phòng XTĐT", "Bình thường"),
    ("Thu thập thêm thông tin từ công ty phụ tùng ô tô Nhật Bản", "Trong tuần 19", "Phòng XTĐT", "Bình thường"),
    ("Chuẩn bị tài liệu kỹ thuật cho Foxconn (MTPII)", "Trong tuần 19", "Phòng XTĐT + Phòng Kỹ thuật", "Bình thường"),
]

nw_headers = ["Nội dung công việc", "Thời hạn", "Đơn vị thực hiện", "Mức độ ưu tiên"]
nw_widths  = [Cm(7.0), Cm(3.5), Cm(5.0), Cm(3.5)]

tbl4 = doc.add_table(rows=1, cols=4)
tbl4.style = 'Table Grid'
set_table_borders(tbl4, color="9DC3E6", size="4")

for i, (label, width) in enumerate(zip(nw_headers, nw_widths)):
    cell = tbl4.rows[0].cells[i]
    cell.width = width
    set_cell_bg(cell, "2E75B6")
    p = cell.paragraphs[0]
    r = p.add_run(label)
    r.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

priority_colors = {"Ưu tiên cao": RGBColor(0xC0, 0x00, 0x00), "Bình thường": RGBColor(0x00, 0x70, 0xC0)}

for idx, item in enumerate(next_week_items):
    row = tbl4.add_row().cells
    bg = "FFFFFF" if idx % 2 == 0 else "EBF3FB"
    for i, val in enumerate(item):
        set_cell_bg(row[i], bg)
        p = row[i].paragraphs[0]
        r = p.add_run(val)
        r.font.size = Pt(10)
        if i == 3:  # priority column
            r.bold = True
            r.font.color.rgb = priority_colors.get(val, RGBColor(0, 0, 0))
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_paragraph()

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 5 – Nhận xét & Kiến nghị
# ─────────────────────────────────────────────────────────────────────────────
add_section_heading(doc, "V", "NHẬN XÉT VÀ KIẾN NGHỊ")

remarks = [
    ("Nhận xét chung",
     "Tuần 18 ghi nhận hoạt động xúc tiến đầu tư sôi nổi với 03 tiếp xúc nhà đầu tư từ 03 quốc gia/vùng lãnh thổ khác nhau (Hàn Quốc, Đài Loan, Nhật Bản). "
     "Chất lượng các cuộc gặp được đánh giá tích cực, đặc biệt là buổi thăm thực địa của Samsung C&T tại VSIP II."),
    ("Điểm tích cực",
     "• Samsung C&T thể hiện sự quan tâm rõ ràng sau site visit – cần nhanh chóng gửi đề xuất để duy trì động lực.\n"
     "• Foxconn đang trong giai đoạn đàm phán tích cực tại MTPII.\n"
     "• Họp KOCHAM mang lại 02 đầu mối doanh nghiệp Hàn Quốc mới."),
    ("Kiến nghị",
     "• Ưu tiên hoàn thiện và gửi đề xuất thương mại cho Samsung C&T trước ngày 05/05/2026.\n"
     "• Phối hợp chặt chẽ giữa Phòng XTĐT và Phòng Kỹ thuật để cung cấp thông tin đầy đủ cho Foxconn.\n"
     "• Lập danh sách theo dõi (pipeline) riêng cho các đầu mối từ KOCHAM và đầu mối Nhật Bản mới.\n"
     "• Chuẩn bị kỹ nội dung họp JETRO ngày 07/05 nhằm tiếp cận thêm nhà đầu tư Nhật Bản vào lĩnh vực phụ tùng ô tô."),
]

tbl5 = doc.add_table(rows=len(remarks), cols=2)
tbl5.style = 'Table Grid'
set_table_borders(tbl5, color="9DC3E6", size="4")
tbl5.columns[0].width = Cm(4.0)
tbl5.columns[1].width = Cm(11.5)

for i, (label, value) in enumerate(remarks):
    row = tbl5.rows[i].cells
    bg = "D6E4F0" if i % 2 == 0 else "EBF3FB"
    set_cell_bg(row[0], bg)
    set_cell_bg(row[1], "FFFFFF")
    p0 = row[0].paragraphs[0]
    r0 = p0.add_run(label)
    r0.bold = True
    r0.font.size = Pt(10)
    r0.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
    p1 = row[1].paragraphs[0]
    r1 = p1.add_run(value)
    r1.font.size = Pt(10)

doc.add_paragraph()

# ─────────────────────────────────────────────────────────────────────────────
# FOOTER – Signature block
# ─────────────────────────────────────────────────────────────────────────────
p_date = doc.add_paragraph()
p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
r_date = p_date.add_run("Bình Dương, ngày 02 tháng 05 năm 2026")
r_date.italic = True
r_date.font.size = Pt(10)

doc.add_paragraph()

# Signature table (left = receiver, right = preparer)
sig_tbl = doc.add_table(rows=1, cols=2)
sig_tbl.style = 'Table Grid'
# remove borders
for row in sig_tbl.rows:
    for cell in row.cells:
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcBorders = OxmlElement('w:tcBorders')
        for b in ('top', 'left', 'bottom', 'right'):
            border = OxmlElement(f'w:{b}')
            border.set(qn('w:val'), 'none')
            tcBorders.append(border)
        tcPr.append(tcBorders)

left_cell  = sig_tbl.rows[0].cells[0]
right_cell = sig_tbl.rows[0].cells[1]

for cell, title, name in [
    (left_cell,  "TRƯỞNG PHÒNG XÚC TIẾN ĐẦU TƯ", ""),
    (right_cell, "NGƯỜI LẬP BÁO CÁO", ""),
]:
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title)
    r.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run("(Ký và ghi rõ họ tên)")
    r2.italic = True
    r2.font.size = Pt(9)
    for _ in range(4):
        cell.add_paragraph()

# ─────────────────────────────────────────────────────────────────────────────
doc.save(OUTPUT_PATH)
print(f"Saved: {OUTPUT_PATH}")
