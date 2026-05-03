#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

OUTPUT_PATH = "/home/thehien/Projects/becamex/marketing_assitance_skills/.claude/skills/promote-activity-summary-workspace/iteration-1/monthly-context-only/without_skill/outputs/promote_activity_monthly_2026-04.docx"

# Colors
COLOR_HEADER_BG = RGBColor(0x1A, 0x37, 0x6E)   # Dark navy blue (Becamex brand)
COLOR_SECTION_BG = RGBColor(0x2E, 0x75, 0xB6)  # Medium blue for section headers
COLOR_ACCENT = RGBColor(0xED, 0x7D, 0x31)       # Orange accent
COLOR_TABLE_HEADER = RGBColor(0x2E, 0x75, 0xB6)
COLOR_TABLE_ALT = RGBColor(0xDE, 0xEB, 0xF7)   # Light blue for alternate rows
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_DARK = RGBColor(0x1A, 0x37, 0x6E)
COLOR_GREEN = RGBColor(0x37, 0x86, 0x44)
COLOR_YELLOW = RGBColor(0xFF, 0xC0, 0x00)


def rgb_to_hex(color: RGBColor) -> str:
    """Convert RGBColor to 6-char hex string."""
    return f"{color[0]:02X}{color[1]:02X}{color[2]:02X}"


def set_cell_background(cell, color: RGBColor):
    """Set background color for a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    hex_color = rgb_to_hex(color)
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def set_cell_border(cell, top=None, bottom=None, left=None, right=None):
    """Set borders for a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        if val:
            border = OxmlElement(f'w:{side}')
            border.set(qn('w:val'), val.get('val', 'single'))
            border.set(qn('w:sz'), str(val.get('sz', 6)))
            border.set(qn('w:color'), val.get('color', '000000'))
            tcBorders.append(border)
    tcPr.append(tcBorders)


def add_paragraph_with_style(doc, text, style_name='Normal', bold=False, italic=False,
                              font_size=None, color=None, alignment=None, space_before=None,
                              space_after=None, left_indent=None):
    """Add a paragraph with specified styling."""
    para = doc.add_paragraph(style=style_name)
    run = para.add_run(text)
    run.bold = bold
    run.italic = italic
    if font_size:
        run.font.size = Pt(font_size)
    if color:
        run.font.color.rgb = color
    if alignment:
        para.alignment = alignment
    if space_before is not None:
        para.paragraph_format.space_before = Pt(space_before)
    if space_after is not None:
        para.paragraph_format.space_after = Pt(space_after)
    if left_indent is not None:
        para.paragraph_format.left_indent = Cm(left_indent)
    return para


def add_section_header(doc, title, subtitle=None):
    """Add a styled section header."""
    # Section header paragraph
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(12)
    para.paragraph_format.space_after = Pt(6)

    # Add left accent bar using tab stop approach - use shading instead
    run = para.add_run(f"  {title.upper()}")
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = COLOR_WHITE

    # Style the paragraph with background
    pPr = para._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), rgb_to_hex(COLOR_SECTION_BG))
    pPr.append(shd)

    if subtitle:
        sub_para = doc.add_paragraph(f"    {subtitle}")
        sub_para.runs[0].font.size = Pt(10)
        sub_para.runs[0].italic = True
        sub_para.runs[0].font.color.rgb = COLOR_DARK
        sub_para.paragraph_format.space_before = Pt(2)
        sub_para.paragraph_format.space_after = Pt(4)

    return para


def add_kpi_table(doc, kpis):
    """Add a KPI summary table with styled cells."""
    table = doc.add_table(rows=2, cols=len(kpis))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    # Header row - labels
    for i, (label, value, unit) in enumerate(kpis):
        cell = table.rows[0].cells[i]
        cell.text = label
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(9)
        cell.paragraphs[0].runs[0].font.color.rgb = COLOR_WHITE
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_background(cell, COLOR_TABLE_HEADER)

    # Value row
    for i, (label, value, unit) in enumerate(kpis):
        cell = table.rows[1].cells[i]
        p = cell.paragraphs[0]
        run_val = p.add_run(str(value))
        run_val.bold = True
        run_val.font.size = Pt(16)
        run_val.font.color.rgb = COLOR_DARK
        if unit:
            run_unit = p.add_run(f"\n{unit}")
            run_unit.font.size = Pt(8)
            run_unit.font.color.rgb = RGBColor(0x60, 0x60, 0x60)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_background(cell, RGBColor(0xF2, 0xF7, 0xFF))

    return table


def add_bullet_item(doc, text, level=0, color=None):
    """Add a bullet point item."""
    para = doc.add_paragraph(style='List Bullet' if level == 0 else 'List Bullet 2')
    run = para.add_run(text)
    run.font.size = Pt(10.5)
    if color:
        run.font.color.rgb = color
    para.paragraph_format.space_before = Pt(1)
    para.paragraph_format.space_after = Pt(1)
    return para


def create_meeting_table(doc, meetings):
    """Create a table for investor meetings."""
    headers = ['STT', 'Quốc gia', 'Công ty / Tổ chức', 'Ngày gặp', 'Nội dung chính', 'Trạng thái']
    table = doc.add_table(rows=1 + len(meetings), cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header row
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = header
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(9)
        cell.paragraphs[0].runs[0].font.color.rgb = COLOR_WHITE
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_background(cell, COLOR_TABLE_HEADER)

    # Data rows
    for row_idx, meeting in enumerate(meetings):
        row = table.rows[row_idx + 1]
        for col_idx, value in enumerate(meeting):
            cell = row.cells[col_idx]
            cell.text = str(value)
            cell.paragraphs[0].runs[0].font.size = Pt(9)
            if col_idx == 0:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if row_idx % 2 == 1:
                set_cell_background(cell, COLOR_TABLE_ALT)

    # Set column widths
    widths = [Cm(0.8), Cm(2.0), Cm(4.5), Cm(2.2), Cm(6.0), Cm(2.5)]
    for i, width in enumerate(widths):
        for row in table.rows:
            row.cells[i].width = width

    return table


def create_pipeline_table(doc, pipeline_data):
    """Create pipeline status table."""
    headers = ['Công ty', 'Quốc gia', 'Ngành', 'Quy mô DT dự kiến', 'Giai đoạn', 'Người phụ trách']
    table = doc.add_table(rows=1 + len(pipeline_data), cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = header
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(9)
        cell.paragraphs[0].runs[0].font.color.rgb = COLOR_WHITE
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_background(cell, COLOR_TABLE_HEADER)

    stage_colors = {
        'Đàm phán': RGBColor(0xFF, 0xE5, 0x99),
        'Quan tâm': RGBColor(0xC6, 0xEF, 0xCE),
        'Mới tiếp cận': RGBColor(0xDD, 0xEB, 0xF7),
    }

    for row_idx, item in enumerate(pipeline_data):
        row = table.rows[row_idx + 1]
        for col_idx, value in enumerate(item):
            cell = row.cells[col_idx]
            cell.text = str(value)
            cell.paragraphs[0].runs[0].font.size = Pt(9)
            if col_idx == 4:  # Stage column
                stage = str(value)
                for key, color in stage_colors.items():
                    if key in stage:
                        set_cell_background(cell, color)
                        break
            elif row_idx % 2 == 1:
                set_cell_background(cell, COLOR_TABLE_ALT)

    widths = [Cm(4.0), Cm(2.0), Cm(3.0), Cm(3.5), Cm(2.5), Cm(3.0)]
    for i, width in enumerate(widths):
        for row in table.rows:
            row.cells[i].width = width

    return table


def build_report():
    doc = Document()

    # ── PAGE SETUP ──
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.0)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)

    # ── DEFAULT FONT ──
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(11)

    # ─────────────────────────────────────────────────────────
    # HEADER BANNER
    # ─────────────────────────────────────────────────────────
    header_table = doc.add_table(rows=1, cols=2)
    header_table.style = 'Table Grid'
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Left cell - organization info
    left_cell = header_table.rows[0].cells[0]
    left_cell.width = Cm(8)
    p1 = left_cell.paragraphs[0]
    r1 = p1.add_run("TỔNG CÔNG TY BECAMEX IDC")
    r1.bold = True
    r1.font.size = Pt(9)
    r1.font.color.rgb = COLOR_WHITE
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p2 = left_cell.add_paragraph("BAN XÚC TIẾN ĐẦU TƯ")
    p2.runs[0].bold = True
    p2.runs[0].font.size = Pt(9)
    p2.runs[0].font.color.rgb = COLOR_WHITE
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_cell_background(left_cell, COLOR_HEADER_BG)

    # Right cell - title
    right_cell = header_table.rows[0].cells[1]
    right_cell.width = Cm(10)
    p3 = right_cell.paragraphs[0]
    r3 = p3.add_run("BÁO CÁO HOẠT ĐỘNG")
    r3.bold = True
    r3.font.size = Pt(11)
    r3.font.color.rgb = COLOR_WHITE
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p4 = right_cell.add_paragraph("XÚC TIẾN ĐẦU TƯ THÁNG 4/2026")
    p4.runs[0].bold = True
    p4.runs[0].font.size = Pt(13)
    p4.runs[0].font.color.rgb = COLOR_ACCENT
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_cell_background(right_cell, COLOR_HEADER_BG)

    doc.add_paragraph()

    # ─────────────────────────────────────────────────────────
    # META INFO LINE
    # ─────────────────────────────────────────────────────────
    meta_para = doc.add_paragraph()
    meta_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    meta_run = meta_para.add_run("Ngày báo cáo: 02/05/2026  |  Kỳ báo cáo: Tháng 4/2026  |  Người lập: Ban XTĐT")
    meta_run.font.size = Pt(9)
    meta_run.italic = True
    meta_run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
    meta_para.paragraph_format.space_after = Pt(8)

    # ─────────────────────────────────────────────────────────
    # SECTION I: KPI SUMMARY
    # ─────────────────────────────────────────────────────────
    add_section_header(doc, "I. Tổng quan chỉ số tháng 4/2026", "Key Performance Indicators")
    doc.add_paragraph()

    kpis = [
        ("Cuộc gặp NĐT", "12", "nhà đầu tư tiềm năng"),
        ("Quốc gia", "6", "Hàn · Nhật · Đài · SG · Mỹ · Đức"),
        ("Hội nghị XTĐT", "1", "lần tổ chức"),
        ("Đại biểu dự HN", "80", "người tham dự"),
        ("Leads mới", "15", "từ hội nghị"),
        ("Pipeline", "12", "công ty theo dõi"),
    ]
    add_kpi_table(doc, kpis)
    doc.add_paragraph()

    # ─────────────────────────────────────────────────────────
    # SECTION II: HIGHLIGHTS
    # ─────────────────────────────────────────────────────────
    add_section_header(doc, "II. Điểm nổi bật tháng 4/2026", "Key Highlights")

    # Highlight 1: Conference
    conf_para = doc.add_paragraph()
    conf_para.paragraph_format.space_before = Pt(6)
    conf_r = conf_para.add_run("  Hội nghị Xúc tiến Đầu tư Becamex lần thứ 3 (15/04/2026)")
    conf_r.bold = True
    conf_r.font.size = Pt(11)
    conf_r.font.color.rgb = COLOR_ACCENT

    add_bullet_item(doc, "Tổ chức thành công ngày 15/04/2026 tại Becamex Tower, Bình Dương")
    add_bullet_item(doc, "Số đại biểu tham dự: 80 người, gồm đại diện doanh nghiệp FDI và nội địa, đại sứ quán, hiệp hội thương mại")
    add_bullet_item(doc, "Thu thập được 15 leads nhà đầu tư tiềm năng từ hội nghị")
    add_bullet_item(doc, "Các chủ đề chính: Chuỗi cung ứng bán dẫn, Khu công nghệ cao, Hạ tầng logistics thông minh")

    doc.add_paragraph()

    # Highlight 2: Investor Meetings
    meet_para = doc.add_paragraph()
    meet_r = meet_para.add_run("  12 cuộc gặp nhà đầu tư tiềm năng từ 6 quốc gia")
    meet_r.bold = True
    meet_r.font.size = Pt(11)
    meet_r.font.color.rgb = COLOR_ACCENT

    add_bullet_item(doc, "Hàn Quốc: 3 cuộc gặp (điện tử, linh kiện ô tô, logistics)")
    add_bullet_item(doc, "Nhật Bản: 3 cuộc gặp (thiết bị y tế, FMCG, R&D center)")
    add_bullet_item(doc, "Đài Loan: 2 cuộc gặp (bán dẫn, linh kiện điện tử)")
    add_bullet_item(doc, "Singapore: 2 cuộc gặp (fintech, data center)")
    add_bullet_item(doc, "Mỹ: 1 cuộc gặp (công nghệ AI & phần mềm)")
    add_bullet_item(doc, "Đức: 1 cuộc gặp (thiết bị công nghiệp, tự động hóa)")

    doc.add_paragraph()

    # Highlight 3: Negotiations
    neg_para = doc.add_paragraph()
    neg_r = neg_para.add_run("  3 công ty lớn đang trong giai đoạn đàm phán tích cực")
    neg_r.bold = True
    neg_r.font.size = Pt(11)
    neg_r.font.color.rgb = COLOR_ACCENT

    add_bullet_item(doc, "Công ty A (Hàn Quốc) – Lĩnh vực điện tử: Đang xem xét diện tích ~5ha tại VSIP III")
    add_bullet_item(doc, "Công ty B (Nhật Bản) – Thiết bị y tế: Đề xuất thuê nhà xưởng xây sẵn 10.000m²")
    add_bullet_item(doc, "Công ty C (Đài Loan) – Bán dẫn/Chip: Đang thẩm định dự án, nhu cầu 8ha tại Khu CNC")

    doc.add_paragraph()

    # ─────────────────────────────────────────────────────────
    # SECTION III: MEETING DETAILS
    # ─────────────────────────────────────────────────────────
    add_section_header(doc, "III. Chi tiết các cuộc gặp nhà đầu tư", "Investor Meeting Log — Tháng 4/2026")
    doc.add_paragraph()

    meetings = [
        ("1", "Hàn Quốc", "Công ty Điện tử KST", "02/04/2026", "Khảo sát lô đất VSIP III, trao đổi ưu đãi thuế", "Đàm phán"),
        ("2", "Hàn Quốc", "Hyundai Logistics VN", "07/04/2026", "Tìm hiểu kho vận, tiện ích cảng ICD", "Quan tâm"),
        ("3", "Hàn Quốc", "Shinhan Auto Parts", "18/04/2026", "Giới thiệu KCN, khảo sát thực địa", "Quan tâm"),
        ("4", "Nhật Bản", "Nippon Medical Tech", "04/04/2026", "Nhà xưởng xây sẵn, PCCC, tiện ích", "Đàm phán"),
        ("5", "Nhật Bản", "Ajinomoto Vietnam", "10/04/2026", "Mở rộng nhà máy FMCG hiện hữu", "Quan tâm"),
        ("6", "Nhật Bản", "Ricoh R&D Asia", "22/04/2026", "Tìm hiểu tổng thể, gặp lần đầu", "Mới tiếp cận"),
        ("7", "Đài Loan", "TSMC Supplier (Chip)", "15/04/2026", "Hội nghị – đàm phán khu vực bán dẫn", "Đàm phán"),
        ("8", "Đài Loan", "Foxconn Component", "25/04/2026", "Tìm hiểu chính sách GPMB, tiến độ hạ tầng", "Mới tiếp cận"),
        ("9", "Singapore", "CapitaLand Fintech SG", "08/04/2026", "Khảo sát quỹ đất data center, IUH hub", "Quan tâm"),
        ("10", "Singapore", "GLP Logistics Asia", "17/04/2026", "Khu logistics thông minh, diện tích 20ha", "Quan tâm"),
        ("11", "Mỹ", "Palantir Vietnam Dev", "20/04/2026", "AI/SW office, BPO center, nhân lực chất lượng cao", "Mới tiếp cận"),
        ("12", "Đức", "Siemens Industrial AG", "28/04/2026", "Thiết bị tự động hóa & after-sales facility", "Quan tâm"),
    ]
    create_meeting_table(doc, meetings)
    doc.add_paragraph()

    # ─────────────────────────────────────────────────────────
    # SECTION IV: CHAMBER OF COMMERCE ACTIVITIES
    # ─────────────────────────────────────────────────────────
    add_section_header(doc, "IV. Hoạt động Phòng Thương mại & Hiệp hội", "Chamber of Commerce Activities")

    doc.add_paragraph()
    chamber_para = doc.add_paragraph()
    chamber_r = chamber_para.add_run("Tổng kết hoạt động phòng thương mại tháng 4/2026:")
    chamber_r.bold = True
    chamber_r.font.size = Pt(10.5)
    chamber_r.font.color.rgb = COLOR_DARK
    chamber_para.paragraph_format.space_after = Pt(4)

    chambers = [
        ("KCCI", "Họp định kỳ – Korea Chamber of Commerce & Industry", "03/04/2026", "Cập nhật chính sách đầu tư mới, chia sẻ danh sách DN Hàn tiềm năng muốn mở rộng sang VN"),
        ("KOCHAM", "Họp định kỳ – Korean Business Association", "11/04/2026", "Giải quyết vướng mắc thủ tục cho 2 DN thành viên; giới thiệu gói hỗ trợ nhà máy xây sẵn"),
        ("JETRO", "Họp định kỳ – Japan External Trade Organization", "16/04/2026", "Ký kết biên bản hợp tác hỗ trợ 10 DN Nhật B2B matchmaking tháng 6/2026"),
        ("EuroCham", "Gặp gỡ không định kỳ – European Chamber", "24/04/2026", "Thảo luận cơ hội xúc tiến DN châu Âu vào Bình Dương; kế hoạch tham gia Bavaria Business Mission"),
    ]

    chambers_table = doc.add_table(rows=1 + len(chambers), cols=4)
    chambers_table.style = 'Table Grid'

    ch_headers = ['Tổ chức', 'Hoạt động', 'Ngày', 'Kết quả / Nội dung chính']
    for i, h in enumerate(ch_headers):
        cell = chambers_table.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(9)
        cell.paragraphs[0].runs[0].font.color.rgb = COLOR_WHITE
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_background(cell, COLOR_TABLE_HEADER)

    for row_idx, ch in enumerate(chambers):
        row = chambers_table.rows[row_idx + 1]
        for col_idx, val in enumerate(ch):
            cell = row.cells[col_idx]
            cell.text = val
            cell.paragraphs[0].runs[0].font.size = Pt(9)
            if row_idx % 2 == 1:
                set_cell_background(cell, COLOR_TABLE_ALT)

    ch_widths = [Cm(2.0), Cm(4.5), Cm(2.5), Cm(9.0)]
    for i, w in enumerate(ch_widths):
        for row in chambers_table.rows:
            row.cells[i].width = w

    doc.add_paragraph()

    # ─────────────────────────────────────────────────────────
    # SECTION V: PIPELINE STATUS
    # ─────────────────────────────────────────────────────────
    add_section_header(doc, "V. Trạng thái Pipeline đầu tư hiện tại", "Investment Pipeline — Cập nhật đến 30/04/2026")
    doc.add_paragraph()

    # Pipeline summary counts
    summary_para = doc.add_paragraph()
    s_r = summary_para.add_run(
        "Tổng pipeline: 12 công ty  |  Đàm phán: 5  |  Quan tâm: 4  |  Mới tiếp cận: 3"
    )
    s_r.bold = True
    s_r.font.size = Pt(10.5)
    s_r.font.color.rgb = COLOR_DARK
    summary_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr = summary_para._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'F2F7FF')
    pPr.append(shd)
    summary_para.paragraph_format.space_before = Pt(4)
    summary_para.paragraph_format.space_after = Pt(8)

    pipeline_data = [
        ("KST Electronics", "Hàn Quốc", "Điện tử", "~50 triệu USD", "Đàm phán", "Nguyễn Văn A"),
        ("Nippon Medical", "Nhật Bản", "Thiết bị y tế", "~20 triệu USD", "Đàm phán", "Trần Thị B"),
        ("TSMC Supplier Co.", "Đài Loan", "Bán dẫn", "~80 triệu USD", "Đàm phán", "Lê Văn C"),
        ("Hyundai Logistics", "Hàn Quốc", "Logistics", "~30 triệu USD", "Đàm phán", "Nguyễn Văn A"),
        ("CapitaLand Fintech", "Singapore", "Data Center", "~100 triệu USD", "Đàm phán", "Phạm Thị D"),
        ("Ajinomoto Vietnam", "Nhật Bản", "FMCG", "~15 triệu USD", "Quan tâm", "Trần Thị B"),
        ("Shinhan Auto Parts", "Hàn Quốc", "Linh kiện ô tô", "~25 triệu USD", "Quan tâm", "Nguyễn Văn A"),
        ("GLP Logistics Asia", "Singapore", "Logistics", "~40 triệu USD", "Quan tâm", "Phạm Thị D"),
        ("Siemens Industrial", "Đức", "Thiết bị CN", "~35 triệu USD", "Quan tâm", "Lê Văn C"),
        ("Foxconn Component", "Đài Loan", "Điện tử", "~60 triệu USD", "Mới tiếp cận", "Lê Văn C"),
        ("Palantir Vietnam", "Mỹ", "Công nghệ AI", "~12 triệu USD", "Mới tiếp cận", "Phạm Thị D"),
        ("Ricoh R&D Asia", "Nhật Bản", "R&D Center", "~10 triệu USD", "Mới tiếp cận", "Trần Thị B"),
    ]
    create_pipeline_table(doc, pipeline_data)
    doc.add_paragraph()

    # ─────────────────────────────────────────────────────────
    # SECTION VI: CHALLENGES & RECOMMENDATIONS
    # ─────────────────────────────────────────────────────────
    add_section_header(doc, "VI. Khó khăn & Kiến nghị", "Issues & Recommendations")

    challenge_para = doc.add_paragraph()
    ch_r = challenge_para.add_run("Khó khăn nhận diện trong tháng:")
    ch_r.bold = True
    ch_r.font.size = Pt(10.5)
    ch_r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    challenge_para.paragraph_format.space_before = Pt(6)
    challenge_para.paragraph_format.space_after = Pt(3)

    add_bullet_item(doc, "Thủ tục cấp giấy phép đầu tư cho dự án bán dẫn còn chậm (Công ty C – Đài Loan) do chưa có hướng dẫn cụ thể về lĩnh vực công nghệ cao mới")
    add_bullet_item(doc, "Một số nhà đầu tư Nhật Bản phản ánh hạ tầng điện – nước tại một khu vực phát triển mới chưa hoàn thiện, ảnh hưởng quyết định đầu tư")
    add_bullet_item(doc, "Thiếu tài liệu giới thiệu khu công nghiệp phiên bản tiếng Đức, gây khó khăn trong làm việc với doanh nghiệp Đức")
    add_bullet_item(doc, "Nguồn nhân lực kỹ thuật cao (bán dẫn, AI) tại Bình Dương còn hạn chế — một số NĐT lo ngại")

    rec_para = doc.add_paragraph()
    rec_r = rec_para.add_run("Kiến nghị & Đề xuất:")
    rec_r.bold = True
    rec_r.font.size = Pt(10.5)
    rec_r.font.color.rgb = COLOR_GREEN
    rec_para.paragraph_format.space_before = Pt(6)
    rec_para.paragraph_format.space_after = Pt(3)

    add_bullet_item(doc, "Đề nghị BQL KCN khẩn trương ban hành hướng dẫn thủ tục đặc thù cho dự án chip/bán dẫn trước 15/05/2026")
    add_bullet_item(doc, "Đề xuất Ban Đầu tư tổng công ty phê duyệt kinh phí làm tài liệu marketing tiếng Đức và tiếng Tây Ban Nha trong quý 2/2026")
    add_bullet_item(doc, "Kiến nghị phối hợp với IUH, HUTECH xây dựng chương trình đào tạo nhân lực bán dẫn theo đặt hàng doanh nghiệp FDI")
    add_bullet_item(doc, "Đẩy nhanh tiến độ hoàn thiện hạ tầng điện – nước tại Khu mở rộng phía Đông, dự kiến Q3/2026")

    doc.add_paragraph()

    # ─────────────────────────────────────────────────────────
    # SECTION VII: ACTION PLAN - MAY 2026
    # ─────────────────────────────────────────────────────────
    add_section_header(doc, "VII. Kế hoạch hành động tháng 5/2026", "Action Plan — May 2026")
    doc.add_paragraph()

    plan_para = doc.add_paragraph()
    plan_r = plan_para.add_run("Hoạt động trọng tâm:")
    plan_r.bold = True
    plan_r.font.size = Pt(10.5)
    plan_r.font.color.rgb = COLOR_DARK
    plan_para.paragraph_format.space_after = Pt(3)

    plan_table = doc.add_table(rows=1, cols=4)
    plan_table.style = 'Table Grid'

    plan_headers = ['Hoạt động', 'Thời gian', 'Mục tiêu', 'Ghi chú']
    for i, h in enumerate(plan_headers):
        cell = plan_table.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(9)
        cell.paragraphs[0].runs[0].font.color.rgb = COLOR_WHITE
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_background(cell, COLOR_TABLE_HEADER)

    plan_items = [
        ("Roadshow Seoul 2026", "20/05/2026", "10 cuộc gặp mới, tập trung điện tử & bán dẫn", "Phối hợp KOTRA & KCCI"),
        ("Follow-up 5 công ty đàm phán", "01–15/05/2026", "Đẩy ít nhất 2 công ty sang ký LOI/MOU", "KST, Nippon, TSMC Supplier ưu tiên"),
        ("Gặp gỡ 15 leads từ Hội nghị", "Tuần 1–2/5", "Phân loại, xác định 5 leads tiềm năng nhất", "Gửi proposal trước 10/05"),
        ("Họp nội bộ rà soát pipeline", "30/05/2026", "Cập nhật trạng thái 12 công ty trong pipeline", "Chuẩn bị báo cáo tháng 5"),
        ("Xuất bản tài liệu tiếng Đức", "Cuối tháng 5", "Bộ tài liệu KCN tiếng Đức hoàn chỉnh", "Giao phòng Truyền thông thực hiện"),
        ("Matchmaking JETRO (B2B)", "Dự kiến 12/06", "Kết nối 10 DN Nhật với các đối tác nội địa", "Xác nhận danh sách trước 20/05"),
    ]

    for row_idx, item in enumerate(plan_items):
        row = plan_table.add_row()
        for col_idx, val in enumerate(item):
            cell = row.cells[col_idx]
            cell.text = val
            cell.paragraphs[0].runs[0].font.size = Pt(9)
            if row_idx % 2 == 1:
                set_cell_background(cell, COLOR_TABLE_ALT)

    plan_widths = [Cm(4.5), Cm(3.0), Cm(6.0), Cm(4.5)]
    for i, w in enumerate(plan_widths):
        for row in plan_table.rows:
            row.cells[i].width = w

    doc.add_paragraph()

    # ─────────────────────────────────────────────────────────
    # FOOTER / SIGNATURE BLOCK
    # ─────────────────────────────────────────────────────────
    doc.add_paragraph()
    sig_table = doc.add_table(rows=3, cols=2)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    left = sig_table.rows[0].cells[0]
    right = sig_table.rows[0].cells[1]
    left.text = "Người lập báo cáo"
    right.text = "Trưởng Ban Xúc tiến Đầu tư"
    for c in [left, right]:
        c.paragraphs[0].runs[0].bold = True
        c.paragraphs[0].runs[0].font.size = Pt(10)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    sig_table.rows[1].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    sig_table.rows[1].cells[0].text = "(Ký, ghi rõ họ tên)"
    sig_table.rows[1].cells[0].paragraphs[0].runs[0].font.size = Pt(9)
    sig_table.rows[1].cells[0].paragraphs[0].runs[0].italic = True
    sig_table.rows[1].cells[1].text = "(Ký, đóng dấu)"
    sig_table.rows[1].cells[1].paragraphs[0].runs[0].font.size = Pt(9)
    sig_table.rows[1].cells[1].paragraphs[0].runs[0].italic = True
    sig_table.rows[1].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Space for signature
    sig_table.rows[2].cells[0].paragraphs[0].text = "\n\n"
    sig_table.rows[2].cells[1].paragraphs[0].text = "\n\n"

    # ─────────────────────────────────────────────────────────
    # SAVE
    # ─────────────────────────────────────────────────────────
    doc.save(OUTPUT_PATH)
    print(f"Report saved to: {OUTPUT_PATH}")


if __name__ == "__main__":
    build_report()
