"""
Eval 3 with-skill: Weekly report Tuần 19 (05/05–09/05/2026)
Following promote-activity-summary SKILL.md and report_structure.md
"""
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

OUT_DIR = os.path.dirname(os.path.abspath(__file__))

BECAMEX_BLUE = (0x1F, 0x4E, 0x79)
ACCENT_BLUE  = (0x2E, 0x75, 0xB6)
LIGHT_BLUE   = (0xD6, 0xE4, 0xF0)
WHITE        = (0xFF, 0xFF, 0xFF)
RED          = (0xFF, 0x00, 0x00)

def rgb(*t): return RGBColor(*t)
def hex_str(t): return "".join(f"{c:02X}" for c in t)

def set_cell_bg(cell, color_tuple):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_str(color_tuple))
    tcPr.append(shd)

def add_heading(doc, text, color=BECAMEX_BLUE, size=13):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(size)
    run.font.color.rgb = rgb(*color)
    return p

def add_bullet(doc, text, size=11):
    p = doc.add_paragraph(style="List Bullet")
    run = p.add_run(text)
    run.font.size = Pt(size)
    return p

def header_cell(cell, text, bg=BECAMEX_BLUE):
    cell.text = ""
    set_cell_bg(cell, bg)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(9)
    run.font.color.rgb = rgb(*WHITE)

def data_cell(cell, text, bg=None, size=9):
    cell.text = ""
    if bg:
        set_cell_bg(cell, bg)
    p = cell.paragraphs[0]
    run = p.add_run(text)
    run.font.size = Pt(size)

doc = Document()
for sec in doc.sections:
    sec.top_margin = Cm(2)
    sec.bottom_margin = Cm(2)
    sec.left_margin = Cm(2.5)
    sec.right_margin = Cm(2)

# ── HEADER ──────────────────────────────────────────────
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("BECAMEX IDC — KHU CÔNG NGHIỆP")
r.bold = True; r.font.size = Pt(14); r.font.color.rgb = rgb(*BECAMEX_BLUE)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("BÁO CÁO HOẠT ĐỘNG XÚC TIẾN ĐẦU TƯ TUẦN")
r.bold = True; r.font.size = Pt(16); r.font.color.rgb = rgb(*BECAMEX_BLUE)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run("Kỳ báo cáo: ").font.size = Pt(11)
r2 = p.add_run("Tuần 19, 05/05 – 09/05/2026")
r2.bold = True; r2.font.size = Pt(11); r2.font.color.rgb = rgb(*ACCENT_BLUE)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run("Ngày lập: 09/05/2026").font.size = Pt(10)

doc.add_paragraph()

# ── SECTION 1: TÓM TẮT ĐIỀU HÀNH ───────────────────────
add_heading(doc, "I. TÓM TẮT ĐIỀU HÀNH")
p = doc.add_paragraph(); p.add_run("Số liệu nhanh trong tuần:").bold = True
add_bullet(doc, "Số cuộc gặp nhà đầu tư tiềm năng: 1 (Hyundai Engineering — site visit)")
add_bullet(doc, "Số hoạt động phòng thương mại/hiệp hội: 0 (Không có trong tuần)")
add_bullet(doc, "Số sự kiện tham dự/tổ chức: 1 (Vietnam Investment Forum 2026 — vai trò diễn giả)")
add_bullet(doc, "Follow-up proposals gửi đi: 4 công ty")
p = doc.add_paragraph(); p.add_run("Điểm nổi bật:").bold = True
add_bullet(doc, "Hyundai Engineering (Hàn Quốc) site visit kết quả tích cực — xác nhận sẽ làm LOI. Trạng thái: Chờ quyết định.")
add_bullet(doc, "Tham dự Vietnam Investment Forum 2026 tại TP.HCM với vai trò diễn giả — thu 8 business cards từ nhà đầu tư tiềm năng.")
add_bullet(doc, "Gửi proposal đến 4 công ty đang trong pipeline.")

# ── SECTION 2: GẶP GỠ NHÀ ĐẦU TƯ TIỀM NĂNG ────────────
add_heading(doc, "II. GẶP GỠ NHÀ ĐẦU TƯ TIỀM NĂNG")
t = doc.add_table(rows=1, cols=7)
t.style = "Table Grid"
hdrs = ["STT", "Tên công ty", "Quốc gia", "Ngành", "Hình thức", "Trạng thái pipeline", "Người phụ trách"]
for i, h in enumerate(hdrs):
    header_cell(t.rows[0].cells[i], h)

row = t.add_row()
cells = row.cells
data_cell(cells[0], "1", LIGHT_BLUE)
data_cell(cells[1], "Hyundai Engineering", LIGHT_BLUE)
data_cell(cells[2], "Hàn Quốc", LIGHT_BLUE)
data_cell(cells[3], "Xây dựng / Hạ tầng", LIGHT_BLUE)
data_cell(cells[4], "Site visit", LIGHT_BLUE)
data_cell(cells[5], "Chờ quyết định", LIGHT_BLUE)
data_cell(cells[6], "[Cần bổ sung]", LIGHT_BLUE)

doc.add_paragraph()
p = doc.add_paragraph()
r = p.add_run("Chi tiết — Hyundai Engineering")
r.bold = True; r.font.size = Pt(11); r.font.color.rgb = rgb(*ACCENT_BLUE)
add_bullet(doc, "Ngành / Quy mô: Xây dựng khu nhà ở công nhân tại khu công nghiệp Becamex")
add_bullet(doc, "Nội dung: Site visit tại khu vực đề xuất — phía Hyundai Engineering đánh giá tích cực về hạ tầng và vị trí")
add_bullet(doc, "Kết quả: Xác nhận sẽ ký LOI (Letter of Intent) — chờ phê duyệt nội bộ")
add_bullet(doc, "Bước tiếp theo: Soạn thảo LOI và gửi cho Hyundai Engineering để xem xét — Deadline: [Cần bổ sung]")

# ── SECTION 3: PHÒNG THƯƠNG MẠI & HIỆP HỘI ─────────────
add_heading(doc, "III. HOẠT ĐỘNG PHÒNG THƯƠNG MẠI & HIỆP HỘI")
p = doc.add_paragraph()
r = p.add_run("Không có hoạt động phòng thương mại hoặc hiệp hội ngành nghề trong Tuần 19.")
r.italic = True; r.font.size = Pt(11); r.font.color.rgb = rgb(0x59, 0x59, 0x59)
p2 = doc.add_paragraph()
p2.add_run("Lịch dự kiến gần nhất: ").font.size = Pt(11)
p2.add_run("[Cần bổ sung — xem kế hoạch tuần 20]").font.size = Pt(11)

# ── SECTION 4: SỰ KIỆN XÚC TIẾN ĐẦU TƯ ────────────────
add_heading(doc, "IV. SỰ KIỆN XÚC TIẾN ĐẦU TƯ")
t2 = doc.add_table(rows=1, cols=5)
t2.style = "Table Grid"
for i, h in enumerate(["Tên sự kiện", "Địa điểm", "Vai trò Becamex", "Số contacts", "Leads tiềm năng"]):
    header_cell(t2.rows[0].cells[i], h)
row2 = t2.add_row()
data_cell(row2.cells[0], "Vietnam Investment Forum 2026", LIGHT_BLUE)
data_cell(row2.cells[1], "TP. Hồ Chí Minh", LIGHT_BLUE)
data_cell(row2.cells[2], "Tham dự (Diễn giả)", LIGHT_BLUE)
data_cell(row2.cells[3], "8 business cards", LIGHT_BLUE)
data_cell(row2.cells[4], "[Cần theo dõi — 8 contacts từ sự kiện]", LIGHT_BLUE)

doc.add_paragraph()
add_bullet(doc, "Ngày: 07/05/2026 | Sự kiện có quy mô quốc gia, nhiều nhà đầu tư nước ngoài tham dự")
add_bullet(doc, "Becamex đóng vai trò diễn giả — cơ hội tốt để quảng bá khu công nghiệp")
add_bullet(doc, "8 business cards thu được — cần follow-up trong tuần 20")

# ── SECTION 5: KẾ HOẠCH TUẦN TỚI ───────────────────────
add_heading(doc, "V. KẾ HOẠCH TUẦN TỚI (Tuần 20, 12/05–16/05/2026)")
p = doc.add_paragraph(); p.add_run("Ưu tiên:").bold = True
add_bullet(doc, "Soạn thảo và gửi LOI cho Hyundai Engineering")
add_bullet(doc, "Follow-up 8 contacts từ Vietnam Investment Forum 2026")
add_bullet(doc, "Theo dõi phản hồi proposals đã gửi cho 4 công ty trong tuần 19")
p2 = doc.add_paragraph(); p2.add_run("Sự kiện:").bold = True
add_bullet(doc, "[Cần bổ sung — kiểm tra lịch sự kiện tháng 5]")
p3 = doc.add_paragraph(); p3.add_run("Hoạt động phòng thương mại:").bold = True
add_bullet(doc, "[Cần bổ sung — xem lịch KCCI, KOCHAM, JETRO tháng 5]")

# Footer
doc.add_paragraph()
p_footer = doc.add_paragraph()
p_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_f = p_footer.add_run("Người lập: _____________      Ngày lập: 09/05/2026      Phân phối: Ban lãnh đạo")
r_f.font.size = Pt(9); r_f.font.color.rgb = rgb(0x59, 0x59, 0x59)

out_path = os.path.join(OUT_DIR, "promote_activity_weekly_2026-W19.docx")
doc.save(out_path)
print(f"Saved: {out_path}")
