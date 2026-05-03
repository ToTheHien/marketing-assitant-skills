#!/usr/bin/env python3
"""Generate promote_activity_weekly_2026-W18.docx"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

# ── colour palette ────────────────────────────────────────────────────────────
BECAMEX_DARK   = RGBColor(0x1F, 0x4E, 0x79)   # #1F4E79
BECAMEX_ACCENT = RGBColor(0x2E, 0x75, 0xB6)   # #2E75B6
WHITE          = RGBColor(0xFF, 0xFF, 0xFF)
ZEBRA          = RGBColor(0xD6, 0xE4, 0xF0)   # row even

PIPELINE_COLORS = {
    "Mới tiếp cận":   RGBColor(0xD6, 0xE4, 0xF0),  # light blue
    "Đang quan tâm":  RGBColor(0x92, 0xD0, 0x50),  # green
    "Đang đàm phán":  RGBColor(0xFF, 0xFF, 0x00),  # yellow
    "Chờ quyết định": RGBColor(0xFF, 0xC0, 0x00),  # orange
}

# ── helpers ───────────────────────────────────────────────────────────────────
def set_cell_bg(cell, rgb: RGBColor):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    hex_color = f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def set_cell_border(cell, border_color="2E75B6"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        bd = OxmlElement(f"w:{side}")
        bd.set(qn("w:val"), "single")
        bd.set(qn("w:sz"), "4")
        bd.set(qn("w:space"), "0")
        bd.set(qn("w:color"), border_color)
        tcBorders.append(bd)
    tcPr.append(tcBorders)

def add_header_row(table, headers, col_widths=None):
    row = table.rows[0]
    for i, h in enumerate(headers):
        cell = row.cells[i]
        set_cell_bg(cell, BECAMEX_DARK)
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.bold = True
        run.font.color.rgb = WHITE
        run.font.size = Pt(9)
        set_cell_border(cell)

def add_data_row(table, values, row_idx, pipeline_col=None):
    row = table.add_row()
    bg = ZEBRA if row_idx % 2 == 0 else RGBColor(0xFF, 0xFF, 0xFF)
    for i, val in enumerate(values):
        cell = row.cells[i]
        if pipeline_col is not None and i == pipeline_col and val in PIPELINE_COLORS:
            set_cell_bg(cell, PIPELINE_COLORS[val])
        else:
            set_cell_bg(cell, bg)
        cell.text = str(val)
        cell.paragraphs[0].runs[0].font.size = Pt(9) if cell.paragraphs[0].runs else None
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for run in para.runs:
                run.font.size = Pt(9)
        set_cell_border(cell)

def add_section_heading(doc, title: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = BECAMEX_DARK

def add_sub_heading(doc, title: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = BECAMEX_ACCENT

def add_bullet(doc, text: str, level=0):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Inches(0.25 + level * 0.25)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    run.font.size = Pt(10)

def add_kv_para(doc, key: str, value: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Inches(0.25)
    r1 = p.add_run(f"{key}: ")
    r1.bold = True
    r1.font.size = Pt(10)
    r2 = p.add_run(value)
    r2.font.size = Pt(10)

# ── main ──────────────────────────────────────────────────────────────────────
def build_report():
    doc = Document()

    # page margins
    for section in doc.sections:
        section.top_margin    = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    # default font
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(11)

    # ── HEADER BLOCK ──────────────────────────────────────────────────────────
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run("BÁO CÁO HOẠT ĐỘNG XÚC TIẾN ĐẦU TƯ TUẦN")
    title_run.bold = True
    title_run.font.size = Pt(16)
    title_run.font.color.rgb = BECAMEX_DARK

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_p.add_run("Tuần 18 | 28/04 – 02/05/2026")
    sub_run.bold = True
    sub_run.font.size = Pt(13)
    sub_run.font.color.rgb = BECAMEX_ACCENT

    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_p.add_run("Ngày lập: 02/05/2026").font.size = Pt(10)

    doc.add_paragraph()  # spacer

    # ── SECTION 1: TÓM TẮT ĐIỀU HÀNH ─────────────────────────────────────────
    add_section_heading(doc, "I. TÓM TẮT ĐIỀU HÀNH")

    add_sub_heading(doc, "Số liệu nhanh trong tuần:")
    add_bullet(doc, "Số cuộc gặp nhà đầu tư tiềm năng: 3")
    add_bullet(doc, "Số hoạt động phòng thương mại/hiệp hội: 1 (KOCHAM – 02/05/2026)")
    add_bullet(doc, "Số sự kiện tham dự/tổ chức: 0")

    add_sub_heading(doc, "Điểm nổi bật:")
    add_bullet(doc, "Samsung C&T (Hàn Quốc, logistics) thực hiện site visit ngày 29/04 tại VSIP II — quan tâm thuê 5 ha, dự kiến gửi yêu cầu cụ thể trong thời gian tới.")
    add_bullet(doc, "Foxconn (Đài Loan, điện tử) tổ chức online meeting ngày 30/04 hỏi về khu công nghiệp MTPII — đang trong giai đoạn đàm phán giá thuê.")
    add_bullet(doc, "Một công ty Nhật (chưa tiết lộ tên, sản xuất linh kiện ô tô) liên hệ qua email hỏi về diện tích 2 ha — mới bước đầu tiếp cận.")
    add_bullet(doc, "Họp KOCHAM ngày 02/05: 2 công ty Hàn Quốc tiềm năng được giới thiệu để mở rộng pipeline.")

    add_sub_heading(doc, "Leads mới trong tuần:")
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.25)
    r = p.add_run("3 công ty (Hàn Quốc – logistics, Đài Loan – điện tử, Nhật Bản – linh kiện ô tô) + 2 giới thiệu qua KOCHAM")
    r.font.size = Pt(10)

    # ── SECTION 2: GẶP GỠ NHÀ ĐẦU TƯ TIỀM NĂNG ──────────────────────────────
    add_section_heading(doc, "II. GẶP GỠ NHÀ ĐẦU TƯ TIỀM NĂNG")

    add_sub_heading(doc, "Bảng tổng hợp:")
    headers_s2 = ["STT", "Tên công ty", "Quốc gia", "Ngành", "Hình thức", "Trạng thái pipeline", "Người phụ trách"]
    table2 = doc.add_table(rows=1, cols=len(headers_s2))
    table2.style = "Table Grid"
    table2.alignment = WD_TABLE_ALIGNMENT.CENTER
    add_header_row(table2, headers_s2)

    meetings = [
        ("1", "Samsung C&T", "Hàn Quốc", "Logistics", "Site visit", "Đang quan tâm", "[Cần bổ sung]"),
        ("2", "Foxconn", "Đài Loan", "Điện tử", "Online meeting", "Đang đàm phán", "[Cần bổ sung]"),
        ("3", "Công ty Nhật (chưa tiết lộ tên)", "Nhật Bản", "Linh kiện ô tô", "Email", "Mới tiếp cận", "[Cần bổ sung]"),
    ]
    for idx, row_data in enumerate(meetings):
        add_data_row(table2, row_data, idx, pipeline_col=5)

    # set column widths
    col_widths_s2 = [Cm(0.8), Cm(3.8), Cm(2.2), Cm(2.8), Cm(2.5), Cm(3.2), Cm(2.8)]
    for i, w in enumerate(col_widths_s2):
        for row in table2.rows:
            row.cells[i].width = w

    doc.add_paragraph()  # spacer

    # ── Section 2 detail cards ──────────────────────────────────────────────
    add_sub_heading(doc, "Chi tiết từng cuộc gặp:")

    # --- Samsung C&T ---
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    run = p.add_run("1. Samsung C&T")
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = BECAMEX_DARK

    add_kv_para(doc, "Ngành / Quy mô quan tâm", "Logistics — 5 ha tại VSIP II")
    add_kv_para(doc, "Ngày gặp", "29/04/2026 (Site visit)")
    add_kv_para(doc, "Nội dung", "Đoàn Samsung C&T thực hiện khảo sát thực địa tại VSIP II. Họ thể hiện mức độ quan tâm cao đến lô đất 5 ha phục vụ mục đích logistics.")
    add_kv_para(doc, "Kết quả", "Phía Samsung C&T đang quan tâm; dự kiến sẽ gửi yêu cầu thuê đất cụ thể trong thời gian tới.")
    add_kv_para(doc, "Bước tiếp theo", "Chuẩn bị và gửi proposal chi tiết cho Samsung C&T (dự kiến tuần 19)")
    add_kv_para(doc, "Deadline follow-up", "07/05/2026")

    # --- Foxconn ---
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    run = p.add_run("2. Foxconn")
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = BECAMEX_DARK

    add_kv_para(doc, "Ngành / Quy mô quan tâm", "Điện tử — khu công nghiệp MTPII (diện tích cụ thể chưa xác định)")
    add_kv_para(doc, "Ngày gặp", "30/04/2026 (Online meeting)")
    add_kv_para(doc, "Nội dung", "Cuộc họp trực tuyến trao đổi thông tin về khu công nghiệp MTPII. Foxconn đặt câu hỏi về giá thuê đất, hạ tầng và các điều kiện ưu đãi.")
    add_kv_para(doc, "Kết quả", "Đang trong giai đoạn đàm phán giá; hai bên chưa đạt được thỏa thuận.")
    add_kv_para(doc, "Bước tiếp theo", "Tiếp tục đàm phán giá và điều kiện thuê đất; cung cấp thêm tài liệu ưu đãi đầu tư")
    add_kv_para(doc, "Deadline follow-up", "[Cần bổ sung]")

    # --- Công ty Nhật ---
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    run = p.add_run("3. Công ty Nhật Bản (chưa tiết lộ tên) — sản xuất linh kiện ô tô")
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = BECAMEX_DARK

    add_kv_para(doc, "Ngành / Quy mô quan tâm", "Linh kiện ô tô — 2 ha")
    add_kv_para(doc, "Ngày liên hệ", "Tuần 18/2026 (Qua email)")
    add_kv_para(doc, "Nội dung", "Công ty liên hệ qua email để hỏi thông tin về khả năng thuê đất tại khu công nghiệp của Becamex IDC, diện tích yêu cầu khoảng 2 ha.")
    add_kv_para(doc, "Kết quả", "Mới tiếp cận; chưa có phản hồi chi tiết từ phía Becamex.")
    add_kv_para(doc, "Bước tiếp theo", "Phản hồi email, gửi brochure và thông tin khu công nghiệp phù hợp; xác nhận tên công ty")
    add_kv_para(doc, "Deadline follow-up", "[Cần bổ sung]")

    # ── SECTION 3: HOẠT ĐỘNG PHÒNG THƯƠNG MẠI & HIỆP HỘI ─────────────────────
    add_section_heading(doc, "III. HOẠT ĐỘNG PHÒNG THƯƠNG MẠI & HIỆP HỘI")

    headers_s3 = ["Tổ chức", "Loại hoạt động", "Ngày", "Kết quả", "Follow-up"]
    table3 = doc.add_table(rows=1, cols=len(headers_s3))
    table3.style = "Table Grid"
    table3.alignment = WD_TABLE_ALIGNMENT.CENTER
    add_header_row(table3, headers_s3)

    chambers = [
        ("KOCHAM", "Cuộc họp – kết nạp thành viên mới", "02/05/2026", "2 công ty Hàn Quốc tiềm năng được giới thiệu vào pipeline", "Liên hệ trực tiếp 2 công ty để đặt lịch gặp"),
    ]
    for idx, row_data in enumerate(chambers):
        add_data_row(table3, row_data, idx)

    col_widths_s3 = [Cm(2.5), Cm(4.5), Cm(2.5), Cm(5.0), Cm(4.5)]
    for i, w in enumerate(col_widths_s3):
        for row in table3.rows:
            row.cells[i].width = w

    # ── SECTION 4: SỰ KIỆN XÚC TIẾN ĐẦU TƯ ──────────────────────────────────
    add_section_heading(doc, "IV. SỰ KIỆN XÚC TIẾN ĐẦU TƯ")

    headers_s4 = ["Tên sự kiện", "Địa điểm", "Vai trò Becamex", "Số contacts", "Leads tiềm năng"]
    table4 = doc.add_table(rows=1, cols=len(headers_s4))
    table4.style = "Table Grid"
    table4.alignment = WD_TABLE_ALIGNMENT.CENTER
    add_header_row(table4, headers_s4)
    add_data_row(table4, ["(Không có sự kiện trong tuần 18)", "—", "—", "—", "—"], 0)

    col_widths_s4 = [Cm(4.5), Cm(3.0), Cm(3.0), Cm(2.5), Cm(4.0)]
    for i, w in enumerate(col_widths_s4):
        for row in table4.rows:
            row.cells[i].width = w

    # ── SECTION 5: KẾ HOẠCH TUẦN TỚI ─────────────────────────────────────────
    add_section_heading(doc, "V. KẾ HOẠCH TUẦN TỚI (Tuần 19 — 05/05 – 09/05/2026)")

    add_sub_heading(doc, "Các cuộc gặp đã lên lịch:")
    add_bullet(doc, "07/05/2026: Họp với JETRO — hình thức và địa điểm [Cần bổ sung] — Người phụ trách: [Cần bổ sung]")

    add_sub_heading(doc, "Công việc cần hoàn thành:")
    add_bullet(doc, "Gửi proposal chi tiết cho Samsung C&T (ưu tiên cao — deadline 07/05/2026)")
    add_bullet(doc, "Tiếp tục đàm phán giá với Foxconn — chuẩn bị thêm tài liệu ưu đãi đầu tư")
    add_bullet(doc, "Phản hồi email và xác nhận thông tin công ty Nhật Bản (linh kiện ô tô, 2 ha)")
    add_bullet(doc, "Liên hệ 2 công ty Hàn Quốc được KOCHAM giới thiệu để đặt lịch gặp sơ bộ")

    add_sub_heading(doc, "Sự kiện:")
    add_bullet(doc, "(Chưa có sự kiện xúc tiến đầu tư được lên lịch trong tuần 19)")

    # ── PIPELINE STATUS LEGEND ─────────────────────────────────────────────────
    doc.add_paragraph()
    add_section_heading(doc, "PHỤ LỤC: Chú giải màu trạng thái Pipeline")

    legend_table = doc.add_table(rows=1, cols=3)
    legend_table.style = "Table Grid"
    add_header_row(legend_table, ["Trạng thái", "Màu nền", "Ý nghĩa"])

    legend_data = [
        ("Mới tiếp cận",   "Xanh nhạt (#D6E4F0)", "Lần đầu liên hệ, chưa có phản hồi rõ ràng"),
        ("Đang quan tâm",  "Xanh lá (#92D050)",    "Đã phản hồi tích cực, đang tìm hiểu thêm"),
        ("Đang đàm phán",  "Vàng (#FFFF00)",        "Đang thảo luận điều kiện cụ thể"),
        ("Chờ quyết định", "Cam (#FFC000)",          "Đã nhận đủ thông tin, đang ra quyết định nội bộ"),
    ]
    for idx, (status, color_name, meaning) in enumerate(legend_data):
        row = legend_table.add_row()
        bg = ZEBRA if idx % 2 == 0 else RGBColor(0xFF, 0xFF, 0xFF)
        # status cell gets pipeline colour
        set_cell_bg(row.cells[0], PIPELINE_COLORS[status])
        row.cells[0].text = status
        row.cells[0].paragraphs[0].runs[0].font.size = Pt(9)
        for j, val in enumerate([color_name, meaning]):
            set_cell_bg(row.cells[j+1], bg)
            row.cells[j+1].text = val
            row.cells[j+1].paragraphs[0].runs[0].font.size = Pt(9)
        for cell in row.cells:
            set_cell_border(cell)

    for i, w in enumerate([Cm(3.5), Cm(4.0), Cm(9.5)]):
        for row in legend_table.rows:
            row.cells[i].width = w

    # ── FOOTER ────────────────────────────────────────────────────────────────
    doc.add_paragraph()
    hr_p = doc.add_paragraph()
    hr_run = hr_p.add_run("─" * 80)
    hr_run.font.size = Pt(8)
    hr_run.font.color.rgb = BECAMEX_ACCENT

    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r1 = footer_p.add_run("Người lập: _____________     |     Ngày: 02/05/2026     |     Phân phối: Ban lãnh đạo")
    r1.font.size = Pt(9)
    r1.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

    # ── SAVE ──────────────────────────────────────────────────────────────────
    out_path = "/home/thehien/Projects/becamex/marketing_assitance_skills/.claude/skills/promote-activity-summary-workspace/iteration-1/weekly-w18-context-only/with_skill/outputs/promote_activity_weekly_2026-W18.docx"
    doc.save(out_path)
    print(f"Saved: {out_path}")


if __name__ == "__main__":
    build_report()
