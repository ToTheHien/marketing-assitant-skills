"""
Eval 3 baseline (no skill): Weekly report Tuần 19 (05/05–09/05/2026)
Generic approach without skill guidance.
"""
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os

OUT_DIR = os.path.dirname(os.path.abspath(__file__))

doc = Document()
for sec in doc.sections:
    sec.top_margin = Cm(2); sec.bottom_margin = Cm(2)
    sec.left_margin = Cm(2.5); sec.right_margin = Cm(2)

def heading(text, size=13, bold=True):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = bold; r.font.size = Pt(size)
    r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    p.paragraph_format.space_before = Pt(10)
    return p

def bullet(text):
    p = doc.add_paragraph(style="List Bullet")
    p.add_run(text).font.size = Pt(11)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("BÁO CÁO HOẠT ĐỘNG XÚC TIẾN ĐẦU TƯ TUẦN")
r.bold = True; r.font.size = Pt(16)
r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

p2 = doc.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p2.add_run("Tuần 19: 05/05/2026 – 09/05/2026").font.size = Pt(12)

doc.add_paragraph()

heading("1. Tóm Tắt Tuần")
bullet("1 cuộc gặp nhà đầu tư (Hyundai Engineering – site visit)")
bullet("0 cuộc họp phòng thương mại")
bullet("1 sự kiện: Vietnam Investment Forum 2026 (TP.HCM, 07/05)")
bullet("Gửi proposal cho 4 công ty")

heading("2. Cuộc Gặp Nhà Đầu Tư")
p3 = doc.add_paragraph()
p3.add_run("Hyundai Engineering (Hàn Quốc)\n").bold = True
p3.add_run("Ngành: Xây dựng khu nhà ở công nhân\n")
p3.add_run("Hình thức: Site visit\n")
p3.add_run("Kết quả: Tích cực — phía Hyundai Engineering sẽ chuẩn bị LOI\n")
p3.add_run("Bước tiếp theo: Soạn thảo LOI\n")
for run in p3.runs:
    run.font.size = Pt(11)

heading("3. Phòng Thương Mại & Hiệp Hội")
doc.add_paragraph("Không có hoạt động trong tuần 19.").runs[0].font.size = Pt(11)

heading("4. Sự Kiện")
p4 = doc.add_paragraph()
p4.add_run("Vietnam Investment Forum 2026\n").bold = True
p4.add_run("Địa điểm: TP. Hồ Chí Minh | Ngày: 07/05/2026\n")
p4.add_run("Vai trò: Diễn giả\n")
p4.add_run("Kết quả: Thu 8 business cards từ nhà đầu tư tiềm năng\n")
for run in p4.runs:
    run.font.size = Pt(11)

heading("5. Kế Hoạch Tuần 20 (12/05–16/05)")
bullet("Soạn thảo và gửi LOI cho Hyundai Engineering")
bullet("Follow-up 8 contacts từ Vietnam Investment Forum")
bullet("Theo dõi phản hồi proposals đã gửi (4 công ty)")

out_path = os.path.join(OUT_DIR, "promote_activity_weekly_2026-W19.docx")
doc.save(out_path)
print(f"Saved: {out_path}")
