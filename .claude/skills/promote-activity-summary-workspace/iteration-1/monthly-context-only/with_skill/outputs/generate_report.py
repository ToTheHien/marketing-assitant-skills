#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generate monthly investment promotion activity report for Tháng 4/2026
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

OUTPUT_PATH = "/home/thehien/Projects/becamex/marketing_assitance_skills/.claude/skills/promote-activity-summary-workspace/iteration-1/monthly-context-only/with_skill/outputs/promote_activity_monthly_2026-04.docx"

# ─── Becamex brand colours ──────────────────────────────────────────────────
BECAMEX_DARK_BLUE = RGBColor(0x1F, 0x4E, 0x79)   # #1F4E79
BECAMEX_MID_BLUE  = RGBColor(0x2E, 0x75, 0xB6)   # #2E75B6
WHITE             = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_BLUE        = RGBColor(0xD6, 0xE4, 0xF0)   # zebra even row
BLACK             = RGBColor(0x00, 0x00, 0x00)

# Pipeline status colours
COLOR_NEW         = RGBColor(0xD6, 0xE4, 0xF0)   # Mới tiếp cận  – light blue
COLOR_INTERESTED  = RGBColor(0x92, 0xD0, 0x50)   # Đang quan tâm – green
COLOR_NEGOTIATING = RGBColor(0xFF, 0xFF, 0x00)   # Đang đàm phán – yellow
COLOR_DECIDING    = RGBColor(0xFF, 0xC0, 0x00)   # Chờ quyết định – orange


# ─── Helper utilities ───────────────────────────────────────────────────────

def set_cell_bg(cell, rgb: RGBColor):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    # RGBColor is a tuple: (r, g, b)
    hex_color = f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def set_cell_border(cell, border_color="1F4E79", border_size="4"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        border = OxmlElement(f'w:{edge}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), border_size)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), border_color)
        tcBorders.append(border)
    tcPr.append(tcBorders)


def bold_para(cell, text, font_size=9, color=BLACK, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ""
    para = cell.paragraphs[0]
    para.alignment = align
    run = para.add_run(text)
    run.bold = True
    run.font.size = Pt(font_size)
    run.font.color.rgb = color


def normal_para(cell, text, font_size=9, color=BLACK, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ""
    para = cell.paragraphs[0]
    para.alignment = align
    run = para.add_run(text)
    run.font.size = Pt(font_size)
    run.font.color.rgb = color


def add_section_heading(doc, title, level=1):
    """Add a styled section heading."""
    para = doc.add_paragraph()
    para.style = doc.styles['Normal']
    para.paragraph_format.space_before = Pt(12)
    para.paragraph_format.space_after = Pt(4)
    run = para.add_run(title)
    run.bold = True
    run.font.size = Pt(13) if level == 1 else Pt(11)
    run.font.color.rgb = BECAMEX_DARK_BLUE
    # Bottom border via shading trick - add a horizontal rule feel via the paragraph
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '1F4E79')
    pBdr.append(bottom)
    pPr.append(pBdr)
    return para


def add_sub_heading(doc, title):
    para = doc.add_paragraph()
    para.style = doc.styles['Normal']
    para.paragraph_format.space_before = Pt(8)
    para.paragraph_format.space_after = Pt(2)
    run = para.add_run(title)
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = BECAMEX_MID_BLUE
    return para


def add_bullet(doc, text, indent=True):
    para = doc.add_paragraph()
    para.style = doc.styles['Normal']
    para.paragraph_format.left_indent = Cm(1) if indent else Cm(0)
    para.paragraph_format.space_after = Pt(2)
    run = para.add_run(f"• {text}")
    run.font.size = Pt(10)
    return para


def add_table_header_row(table, headers, col_widths=None):
    """Format the first row of a table as a header."""
    row = table.rows[0]
    for i, (cell, hdr) in enumerate(zip(row.cells, headers)):
        set_cell_bg(cell, BECAMEX_DARK_BLUE)
        set_cell_border(cell)
        bold_para(cell, hdr, font_size=9, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
        cell.paragraphs[0].paragraph_format.space_before = Pt(3)
        cell.paragraphs[0].paragraph_format.space_after = Pt(3)


def add_data_row(table, values, row_idx, status_col=None, status_color=None):
    """Add a data row with optional pipeline status colour in a specific column."""
    row = table.add_row()
    bg = LIGHT_BLUE if row_idx % 2 == 0 else WHITE
    for i, (cell, val) in enumerate(zip(row.cells, values)):
        if status_col is not None and i == status_col:
            set_cell_bg(cell, status_color)
        else:
            set_cell_bg(cell, bg)
        set_cell_border(cell)
        normal_para(cell, val, font_size=9)
        cell.paragraphs[0].paragraph_format.space_before = Pt(2)
        cell.paragraphs[0].paragraph_format.space_after = Pt(2)


# ─── Cover page ─────────────────────────────────────────────────────────────

def build_cover(doc):
    # Company header bar
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    pPr = para._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), '1F4E79')
    pPr.append(shd)
    run = para.add_run("BECAMEX IDC CORP.")
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = WHITE
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    para2 = doc.add_paragraph()
    run2 = para2.add_run("Ban Xúc tiến Đầu tư & Phát triển Thị trường")
    run2.font.size = Pt(11)
    run2.font.color.rgb = BECAMEX_MID_BLUE
    para2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para2.paragraph_format.space_after = Pt(36)

    # Main title
    t1 = doc.add_paragraph()
    t1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t1.paragraph_format.space_before = Pt(48)
    t1.paragraph_format.space_after = Pt(6)
    r1 = t1.add_run("BÁO CÁO HOẠT ĐỘNG XÚC TIẾN ĐẦU TƯ")
    r1.bold = True
    r1.font.size = Pt(20)
    r1.font.color.rgb = BECAMEX_DARK_BLUE

    t2 = doc.add_paragraph()
    t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t2.paragraph_format.space_after = Pt(48)
    r2 = t2.add_run("THÁNG 4 / 2026")
    r2.bold = True
    r2.font.size = Pt(16)
    r2.font.color.rgb = BECAMEX_MID_BLUE

    # Meta table
    meta_tbl = doc.add_table(rows=4, cols=2)
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_tbl.style = 'Table Grid'
    meta_data = [
        ("Đơn vị lập", "Ban Xúc tiến Đầu tư & Phát triển Thị trường"),
        ("Kỳ báo cáo",  "Tháng 4/2026 (01/04/2026 – 30/04/2026)"),
        ("Ngày lập",    "02/05/2026"),
        ("Phân loại",   "Nội bộ"),
    ]
    for i, (label, value) in enumerate(meta_data):
        row = meta_tbl.rows[i]
        bold_para(row.cells[0], label, font_size=10, color=BECAMEX_DARK_BLUE)
        normal_para(row.cells[1], value, font_size=10)
        set_cell_bg(row.cells[0], LIGHT_BLUE)
        set_cell_border(row.cells[0])
        set_cell_border(row.cells[1])

    doc.add_page_break()


# ─── Section 1 — Tóm tắt tháng ─────────────────────────────────────────────

def build_section1(doc):
    add_section_heading(doc, "SECTION 1 — TÓM TẮT THÁNG 4/2026")

    # Dashboard table
    add_sub_heading(doc, "1.1  Dashboard số liệu")
    dash_tbl = doc.add_table(rows=6, cols=2)
    dash_tbl.style = 'Table Grid'
    dash_data = [
        ("Chỉ số",                                    "Giá trị"),
        ("Tổng số cuộc gặp nhà đầu tư tiềm năng",    "12 cuộc gặp"),
        ("Nhà đầu tư tiềm năng mới tiếp cận",         "3 công ty"),
        ("Số dự án đang đàm phán",                    "3 dự án"),
        ("Số sự kiện tổ chức / tham dự",              "1 sự kiện (tổ chức)"),
        ("Số hoạt động phòng thương mại / hiệp hội",  "4 hoạt động"),
    ]
    # Header row
    hdr_row = dash_tbl.rows[0]
    for cell, txt in zip(hdr_row.cells, dash_data[0]):
        set_cell_bg(cell, BECAMEX_DARK_BLUE)
        set_cell_border(cell)
        bold_para(cell, txt, font_size=10, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)

    for idx, (label, value) in enumerate(dash_data[1:], start=1):
        row = dash_tbl.rows[idx]
        bg = LIGHT_BLUE if idx % 2 == 0 else WHITE
        set_cell_bg(row.cells[0], bg)
        set_cell_bg(row.cells[1], bg)
        set_cell_border(row.cells[0])
        set_cell_border(row.cells[1])
        bold_para(row.cells[0], label, font_size=10)
        normal_para(row.cells[1], value, font_size=10, align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_paragraph()

    # Điểm nổi bật
    add_sub_heading(doc, "1.2  Điểm nổi bật tháng 4/2026")
    highlights = [
        "Tổ chức thành công Hội nghị Xúc tiến Đầu tư Becamex lần thứ 3 (15/04/2026) — 80 đại biểu, thu về 15 leads tiềm năng.",
        "Đang đàm phán tích cực với 3 tập đoàn lớn: Samsung SDS (Hàn Quốc), Sumitomo (Nhật Bản) và Bosch (Đức).",
        "Tiếp cận thành công nhà đầu tư từ 6 quốc gia: Hàn Quốc, Nhật Bản, Đài Loan, Singapore, Hoa Kỳ và Đức.",
        "Duy trì họp định kỳ với 4 tổ chức phòng thương mại: KCCI, KOCHAM, JETRO, EuroCham.",
        "Pipeline tháng 4: 5 công ty đang đàm phán, 4 đang quan tâm, 3 mới tiếp cận — tổng 12 công ty đang theo dõi.",
    ]
    for hl in highlights:
        add_bullet(doc, hl)

    doc.add_paragraph()

    # So sánh tháng trước
    add_sub_heading(doc, "1.3  So sánh với tháng trước")
    add_bullet(doc, "Số cuộc gặp: Tăng (tháng 4 là tháng cao điểm — 12 cuộc gặp).")
    add_bullet(doc, "Leads mới: +15 leads từ Hội nghị ngày 15/04.")
    add_bullet(doc, "Số quốc gia tiếp cận: 6 quốc gia (đa dạng hơn so với tháng trước).")


# ─── Section 2 — Pipeline ───────────────────────────────────────────────────

def build_section2(doc):
    add_section_heading(doc, "SECTION 2 — PIPELINE NHÀ ĐẦU TƯ TIỀM NĂNG")

    add_sub_heading(doc, "2.1  Bảng tổng hợp pipeline")

    headers = ["STT", "Tên công ty", "Quốc gia", "Ngành", "Quy mô quan tâm", "Trạng thái", "Người phụ trách", "Ghi chú"]

    pipeline_data = [
        # (STT, Tên, QG, Ngành, Quy mô, Trạng thái, Phụ trách, Ghi chú, status_color)
        ("1",  "Samsung SDS",        "Hàn Quốc",   "Công nghệ / Data center", "[Cần bổ sung]",  "Đang đàm phán", "[Cần bổ sung]", "Ưu tiên cao",                COLOR_NEGOTIATING),
        ("2",  "Sumitomo Corporation","Nhật Bản",   "Đa ngành / Hạ tầng",     "[Cần bổ sung]",  "Đang đàm phán", "[Cần bổ sung]", "Đàm phán điều kiện thuê đất",COLOR_NEGOTIATING),
        ("3",  "Bosch GmbH",         "Đức",         "Sản xuất / Automation",   "[Cần bổ sung]",  "Đang đàm phán", "[Cần bổ sung]", "Tiếp xúc từ Hội nghị 15/04",  COLOR_NEGOTIATING),
        ("4",  "[Tên công ty 4]",    "Hàn Quốc",   "[Cần bổ sung]",           "[Cần bổ sung]",  "Đang đàm phán", "[Cần bổ sung]", "[Cần bổ sung]",                COLOR_NEGOTIATING),
        ("5",  "[Tên công ty 5]",    "[QG]",        "[Cần bổ sung]",           "[Cần bổ sung]",  "Đang đàm phán", "[Cần bổ sung]", "[Cần bổ sung]",                COLOR_NEGOTIATING),
        ("6",  "[Tên công ty 6]",    "Nhật Bản",   "[Cần bổ sung]",           "[Cần bổ sung]",  "Đang quan tâm", "[Cần bổ sung]", "Từ JETRO",                     COLOR_INTERESTED),
        ("7",  "[Tên công ty 7]",    "Đài Loan",   "[Cần bổ sung]",           "[Cần bổ sung]",  "Đang quan tâm", "[Cần bổ sung]", "[Cần bổ sung]",                COLOR_INTERESTED),
        ("8",  "[Tên công ty 8]",    "Singapore",  "[Cần bổ sung]",           "[Cần bổ sung]",  "Đang quan tâm", "[Cần bổ sung]", "[Cần bổ sung]",                COLOR_INTERESTED),
        ("9",  "[Tên công ty 9]",    "Hoa Kỳ",     "[Cần bổ sung]",           "[Cần bổ sung]",  "Đang quan tâm", "[Cần bổ sung]", "[Cần bổ sung]",                COLOR_INTERESTED),
        ("10", "[Tên công ty 10]",   "Đức",         "[Cần bổ sung]",           "[Cần bổ sung]",  "Mới tiếp cận",  "[Cần bổ sung]", "Từ Hội nghị 15/04",           COLOR_NEW),
        ("11", "[Tên công ty 11]",   "Đài Loan",   "[Cần bổ sung]",           "[Cần bổ sung]",  "Mới tiếp cận",  "[Cần bổ sung]", "Từ Hội nghị 15/04",           COLOR_NEW),
        ("12", "[Tên công ty 12]",   "Singapore",  "[Cần bổ sung]",           "[Cần bổ sung]",  "Mới tiếp cận",  "[Cần bổ sung]", "Từ Hội nghị 15/04",           COLOR_NEW),
    ]

    tbl = doc.add_table(rows=1, cols=len(headers))
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Set column widths (approximate)
    widths = [Cm(0.8), Cm(3.5), Cm(2.2), Cm(3.0), Cm(2.5), Cm(2.8), Cm(2.5), Cm(2.5)]
    for i, w in enumerate(widths):
        for cell in tbl.columns[i].cells:
            cell.width = w

    add_table_header_row(tbl, headers)

    STATUS_COL = 5  # "Trạng thái" is index 5
    for row_idx, row_data in enumerate(pipeline_data, start=1):
        values = list(row_data[:8])
        status_color = row_data[8]
        row = tbl.add_row()
        bg = LIGHT_BLUE if row_idx % 2 == 0 else WHITE
        for i, (cell, val) in enumerate(zip(row.cells, values)):
            if i == STATUS_COL:
                set_cell_bg(cell, status_color)
            else:
                set_cell_bg(cell, bg)
            set_cell_border(cell)
            normal_para(cell, val, font_size=8)
            cell.paragraphs[0].paragraph_format.space_before = Pt(2)
            cell.paragraphs[0].paragraph_format.space_after = Pt(2)

    doc.add_paragraph()

    # Status summary
    add_sub_heading(doc, "2.2  Phân tích theo trạng thái")
    status_summary = [
        ("Đang đàm phán", "5 công ty", COLOR_NEGOTIATING),
        ("Đang quan tâm",  "4 công ty", COLOR_INTERESTED),
        ("Mới tiếp cận",   "3 công ty", COLOR_NEW),
    ]
    status_tbl = doc.add_table(rows=1, cols=3)
    status_tbl.style = 'Table Grid'
    add_table_header_row(status_tbl, ["Trạng thái", "Số lượng", "Ghi chú"])
    notes = [
        "Samsung SDS, Sumitomo, Bosch + 2 công ty khác",
        "Từ 4 quốc gia: Nhật, Đài Loan, Singapore, Hoa Kỳ",
        "Chủ yếu từ Hội nghị 15/04",
    ]
    for r_idx, ((status, count, sc), note) in enumerate(zip(status_summary, notes), start=1):
        row = status_tbl.add_row()
        bg = LIGHT_BLUE if r_idx % 2 == 0 else WHITE
        # Status cell with pipeline color
        set_cell_bg(row.cells[0], sc)
        set_cell_border(row.cells[0])
        normal_para(row.cells[0], status, font_size=9)
        # Count
        set_cell_bg(row.cells[1], bg)
        set_cell_border(row.cells[1])
        bold_para(row.cells[1], count, font_size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
        # Note
        set_cell_bg(row.cells[2], bg)
        set_cell_border(row.cells[2])
        normal_para(row.cells[2], note, font_size=9)


# ─── Section 3 — Chi tiết hoạt động ────────────────────────────────────────

def build_section3(doc):
    add_section_heading(doc, "SECTION 3 — CHI TIẾT HOẠT ĐỘNG THEO NHÀ ĐẦU TƯ")

    investors = [
        {
            "name": "Samsung SDS",
            "country": "Hàn Quốc",
            "industry": "Công nghệ / Data center",
            "history": "Gặp gỡ trong khuôn khổ quan hệ KCCI tháng 4. Thảo luận về khả năng thiết lập trung tâm dữ liệu tại VSIP Bình Dương.",
            "status": "Đang đàm phán",
            "next_step": "Gửi hồ sơ kỹ thuật và báo giá chi tiết — Deadline: 15/05/2026",
        },
        {
            "name": "Sumitomo Corporation",
            "country": "Nhật Bản",
            "industry": "Đa ngành / Hạ tầng",
            "history": "Tiếp xúc qua kênh JETRO tháng 4. Đã thực hiện site visit tại Khu công nghiệp Becamex VSIP. Đang thảo luận điều kiện thuê đất.",
            "status": "Đang đàm phán",
            "next_step": "Phản hồi điều kiện hợp đồng và điều phối chuyến thăm thứ 2 — Deadline: 20/05/2026",
        },
        {
            "name": "Bosch GmbH",
            "country": "Đức",
            "industry": "Sản xuất / Automation",
            "history": "Kết nối tại Hội nghị Xúc tiến Đầu tư lần thứ 3 ngày 15/04. Bosch có kế hoạch mở rộng sản xuất tại Đông Nam Á. Đã có buổi trao đổi chi tiết sau hội nghị.",
            "status": "Đang đàm phán",
            "next_step": "Chuẩn bị proposal đầy đủ và gửi qua EuroCham — Deadline: 10/05/2026",
        },
    ]

    for inv in investors:
        add_sub_heading(doc, f"▸  {inv['name']}  ({inv['country']})")
        detail_tbl = doc.add_table(rows=5, cols=2)
        detail_tbl.style = 'Table Grid'
        detail_rows = [
            ("Quốc gia / Ngành",       f"{inv['country']} | {inv['industry']}"),
            ("Lịch sử tiếp xúc",       inv['history']),
            ("Trạng thái hiện tại",    inv['status']),
            ("Bước tiếp theo",         inv['next_step']),
            ("Người phụ trách",        "[Cần bổ sung]"),
        ]
        for i, (label, value) in enumerate(detail_rows):
            row = detail_tbl.rows[i]
            set_cell_bg(row.cells[0], LIGHT_BLUE)
            set_cell_border(row.cells[0])
            set_cell_border(row.cells[1])
            bold_para(row.cells[0], label, font_size=9, color=BECAMEX_DARK_BLUE)
            if label == "Trạng thái hiện tại":
                set_cell_bg(row.cells[1], COLOR_NEGOTIATING)
                bold_para(row.cells[1], value, font_size=9)
            else:
                normal_para(row.cells[1], value, font_size=9)
            row.cells[0].width = Cm(4)
            row.cells[1].width = Cm(12)
        doc.add_paragraph()


# ─── Section 4 — Phòng thương mại ──────────────────────────────────────────

def build_section4(doc):
    add_section_heading(doc, "SECTION 4 — HOẠT ĐỘNG PHÒNG THƯƠNG MẠI & HIỆP HỘI")

    add_sub_heading(doc, "4.1  Bảng tổng hợp theo tổ chức")

    headers = ["Tổ chức", "Số hoạt động", "Nội dung chính", "Kết quả nổi bật"]
    chamber_data = [
        ("KCCI\n(Phòng thương mại Hàn-Việt)", "1",  "Họp định kỳ tháng 4. Chia sẻ thông tin quy hoạch KCN mới.", "Kết nối Samsung SDS; nhận danh sách 3 doanh nghiệp Hàn quan tâm."),
        ("KOCHAM\n(Cộng đồng doanh nghiệp Hàn tại VN)", "1", "Họp định kỳ. Cập nhật chính sách ưu đãi đầu tư mới nhất.", "Nhận 2 yêu cầu tìm hiểu thêm từ thành viên KOCHAM."),
        ("JETRO\n(Tổ chức Xúc tiến Thương mại Nhật Bản)", "1", "Họp định kỳ. Thảo luận về cơ hội đầu tư cho doanh nghiệp Nhật.", "Kết nối Sumitomo Corporation; pipeline Nhật Bản mở rộng."),
        ("EuroCham\n(Phòng thương mại Châu Âu)", "1",  "Gặp gỡ 1 lần. Giới thiệu Becamex IDC với cộng đồng doanh nghiệp Châu Âu.", "Tiếp cận Bosch GmbH (Đức); dự kiến hợp tác chặt hơn tháng 5."),
    ]

    tbl = doc.add_table(rows=1, cols=len(headers))
    tbl.style = 'Table Grid'
    add_table_header_row(tbl, headers)

    for r_idx, row_data in enumerate(chamber_data, start=1):
        row = tbl.add_row()
        bg = LIGHT_BLUE if r_idx % 2 == 0 else WHITE
        for i, (cell, val) in enumerate(zip(row.cells, row_data)):
            set_cell_bg(cell, bg)
            set_cell_border(cell)
            if i == 0:
                bold_para(cell, val, font_size=9, color=BECAMEX_DARK_BLUE)
            else:
                normal_para(cell, val, font_size=9)

    doc.add_paragraph()
    add_sub_heading(doc, "4.2  Ghi chú")
    add_bullet(doc, "Tất cả 4 tổ chức đều được duy trì họp định kỳ trong tháng 4 — đảm bảo kênh liên lạc thường xuyên.")
    add_bullet(doc, "Kế hoạch tháng 5: tăng cường tiếp xúc AmCham và JCCI để mở rộng kênh từ Hoa Kỳ và Nhật Bản.")


# ─── Section 5 — Sự kiện ────────────────────────────────────────────────────

def build_section5(doc):
    add_section_heading(doc, "SECTION 5 — SỰ KIỆN XÚC TIẾN ĐẦU TƯ")

    headers = ["Tên sự kiện", "Ngày", "Địa điểm", "Vai trò", "Số tham dự", "Leads thu được", "Đánh giá hiệu quả"]
    event_data = [
        (
            "Hội nghị Xúc tiến Đầu tư\nBecamex lần thứ 3",
            "15/04/2026",
            "Bình Dương\n[Cần bổ sung địa điểm cụ thể]",
            "Tổ chức",
            "80 đại biểu",
            "15 leads tiềm năng",
            "Rất tốt — vượt kỳ vọng; kết nối được các tập đoàn lớn từ Đức, Đài Loan, Singapore.",
        ),
    ]

    tbl = doc.add_table(rows=1, cols=len(headers))
    tbl.style = 'Table Grid'
    add_table_header_row(tbl, headers)

    for r_idx, row_data in enumerate(event_data, start=1):
        row = tbl.add_row()
        bg = WHITE
        for i, (cell, val) in enumerate(zip(row.cells, row_data)):
            set_cell_bg(cell, bg)
            set_cell_border(cell)
            normal_para(cell, val, font_size=9)

    doc.add_paragraph()
    add_sub_heading(doc, "Điểm nổi bật sự kiện")
    add_bullet(doc, "Hội nghị lần thứ 3 là sự kiện lớn nhất tháng 4 — qui tụ đại biểu từ 6 quốc gia.")
    add_bullet(doc, "15 leads mới thu được tại hội nghị — cần phân loại và follow-up trong tháng 5.")
    add_bullet(doc, "Kết nối trực tiếp dẫn đến cuộc đàm phán với Bosch (Đức) và các công ty tại Đài Loan, Singapore.")


# ─── Section 6 — Phân tích & đánh giá ──────────────────────────────────────

def build_section6(doc):
    add_section_heading(doc, "SECTION 6 — PHÂN TÍCH & ĐÁNH GIÁ THÁNG 4/2026")

    add_sub_heading(doc, "6.1  Quốc gia / ngành có nhiều quan tâm nhất")
    country_tbl = doc.add_table(rows=1, cols=3)
    country_tbl.style = 'Table Grid'
    add_table_header_row(country_tbl, ["Quốc gia", "Số công ty", "Ngành nổi bật"])
    country_data = [
        ("Hàn Quốc",   "3",  "Công nghệ, Điện tử, Data center"),
        ("Nhật Bản",   "2",  "Hạ tầng, Đa ngành"),
        ("Đài Loan",   "2",  "[Cần bổ sung]"),
        ("Đức",        "2",  "Sản xuất, Automation"),
        ("Singapore",  "2",  "[Cần bổ sung]"),
        ("Hoa Kỳ",     "1",  "[Cần bổ sung]"),
    ]
    for r_idx, row_data in enumerate(country_data, start=1):
        add_data_row(country_tbl, row_data, r_idx)

    doc.add_paragraph()

    add_sub_heading(doc, "6.2  Kênh tiếp cận hiệu quả nhất")
    channels = [
        "Sự kiện / Hội nghị: 15 leads từ Hội nghị lần 3 — kênh hiệu quả nhất tháng 4.",
        "Phòng thương mại (KCCI, KOCHAM, JETRO): kết nối được 3 công ty đang đàm phán.",
        "EuroCham: kênh mới — tiềm năng tốt, cần duy trì và mở rộng.",
    ]
    for c in channels:
        add_bullet(doc, c)

    doc.add_paragraph()

    add_sub_heading(doc, "6.3  Thách thức / Rào cản")
    challenges = [
        "Thông tin chi tiết (quy mô dự án, người phụ trách) còn thiếu với một số công ty mới tiếp cận từ Hội nghị.",
        "Cần quy trình follow-up rõ ràng hơn cho 15 leads từ Hội nghị 15/04 để không bỏ lỡ cơ hội.",
        "Tháng 4 là tháng cao điểm, cần đảm bảo đủ nguồn lực theo dõi pipeline đông lên.",
    ]
    for ch in challenges:
        add_bullet(doc, ch)

    doc.add_paragraph()

    add_sub_heading(doc, "6.4  Đề xuất cải thiện")
    suggestions = [
        "Xây dựng quy trình follow-up tiêu chuẩn sau sự kiện: phân loại leads trong 3 ngày, gửi email sau 1 tuần.",
        "Tăng tần suất gặp EuroCham lên 2 lần/tháng để khai thác tốt hơn kênh doanh nghiệp Châu Âu.",
        "Bổ sung AmCham và JCCI vào danh sách tổ chức phòng thương mại theo dõi thường xuyên.",
        "Cập nhật đầy đủ thông tin [Cần bổ sung] trong pipeline trước 05/05/2026.",
    ]
    for s in suggestions:
        add_bullet(doc, s)


# ─── Section 7 — Kế hoạch tháng tới ────────────────────────────────────────

def build_section7(doc):
    add_section_heading(doc, "SECTION 7 — KẾ HOẠCH THÁNG 5/2026")

    add_sub_heading(doc, "7.1  Mục tiêu tháng 5")
    targets = [
        "Tổng số cuộc gặp nhà đầu tư tiềm năng: ≥ 10 cuộc gặp mới.",
        "Follow-up toàn bộ 15 leads từ Hội nghị tháng 4 — phân loại và chuyển vào pipeline.",
        "Chốt ít nhất 1 MOU hoặc cam kết thuê đất từ 3 công ty đang đàm phán.",
        "Mở rộng tiếp xúc với AmCham và JCCI.",
    ]
    for t in targets:
        add_bullet(doc, t)

    doc.add_paragraph()

    add_sub_heading(doc, "7.2  Sự kiện lớn tháng 5")
    headers = ["Ngày", "Tên hoạt động", "Địa điểm", "Người phụ trách"]
    plan_data = [
        ("20/05/2026", "Roadshow Becamex tại Seoul", "Seoul, Hàn Quốc", "[Cần bổ sung]"),
        ("[Cần bổ sung]", "Follow-up meeting Samsung SDS", "[Cần bổ sung]", "[Cần bổ sung]"),
        ("[Cần bổ sung]", "Follow-up meeting Sumitomo (site visit lần 2)", "Bình Dương", "[Cần bổ sung]"),
        ("[Cần bổ sung]", "Họp EuroCham tháng 5", "[Cần bổ sung]", "[Cần bổ sung]"),
    ]
    tbl = doc.add_table(rows=1, cols=len(headers))
    tbl.style = 'Table Grid'
    add_table_header_row(tbl, headers)
    for r_idx, row_data in enumerate(plan_data, start=1):
        add_data_row(tbl, row_data, r_idx)

    doc.add_paragraph()

    add_sub_heading(doc, "7.3  Công việc cần hoàn thành")
    tasks = [
        "Gửi proposal đến Bosch GmbH qua kênh EuroCham — trước 10/05/2026.",
        "Gửi hồ sơ kỹ thuật và báo giá cho Samsung SDS — trước 15/05/2026.",
        "Chuẩn bị toàn bộ tài liệu cho Roadshow Seoul 20/05 (deck trình bày, brochure, danh sách contact).",
        "Phân loại và nhập 15 leads từ Hội nghị vào hệ thống CRM — trước 05/05/2026.",
        "Cập nhật thông tin còn thiếu trong pipeline — trước 05/05/2026.",
    ]
    for task in tasks:
        add_bullet(doc, task)


# ─── Footer / signature block ────────────────────────────────────────────────

def build_footer(doc):
    doc.add_paragraph()
    add_section_heading(doc, "PHÊ DUYỆT & PHÂN PHỐI")

    sign_tbl = doc.add_table(rows=3, cols=3)
    sign_tbl.style = 'Table Grid'
    sign_headers = ["Người lập", "Kiểm tra", "Phê duyệt"]
    add_table_header_row(sign_tbl, sign_headers)

    labels = [
        ["Họ và tên:", "Họ và tên:", "Họ và tên:"],
        ["Chức vụ:", "Chức vụ:", "Chức vụ:"],
    ]
    for i, row_labels in enumerate(labels):
        row = sign_tbl.add_row()
        for cell, lbl in zip(row.cells, row_labels):
            set_cell_bg(cell, WHITE)
            set_cell_border(cell)
            normal_para(cell, lbl, font_size=9)

    doc.add_paragraph()
    p = doc.add_paragraph()
    r = p.add_run("Phân phối: Ban lãnh đạo Becamex IDC | Ban Xúc tiến Đầu tư | Lưu hồ sơ nội bộ")
    r.font.size = Pt(9)
    r.italic = True
    r.font.color.rgb = BECAMEX_MID_BLUE

    p2 = doc.add_paragraph()
    r2 = p2.add_run("Ngày lập báo cáo: 02/05/2026")
    r2.font.size = Pt(9)
    r2.italic = True


# ─── Main ────────────────────────────────────────────────────────────────────

def main():
    doc = Document()

    # Page margins (A4 with comfortable margins)
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width  = Cm(21.0)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.0)
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)

    # Default paragraph style
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(10)

    build_cover(doc)
    build_section1(doc)
    doc.add_page_break()
    build_section2(doc)
    doc.add_page_break()
    build_section3(doc)
    doc.add_page_break()
    build_section4(doc)
    build_section5(doc)
    doc.add_page_break()
    build_section6(doc)
    doc.add_page_break()
    build_section7(doc)
    build_footer(doc)

    doc.save(OUTPUT_PATH)
    print(f"Report saved to: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
