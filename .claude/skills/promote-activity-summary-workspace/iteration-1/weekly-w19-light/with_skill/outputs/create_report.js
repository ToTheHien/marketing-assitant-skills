const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, LevelFormat, TabStopType,
  TabStopPosition
} = require('docx');
const fs = require('fs');

// ─── Becamex color palette ───────────────────────────────────────────────────
const BECAMEX_BLUE     = "1F4E79";
const BECAMEX_ACCENT   = "2E75B6";
const WHITE            = "FFFFFF";
const ZEBRA            = "D6E4F0";
const STATUS_WAIT      = "FFC000";  // Chờ quyết định — Cam
const STATUS_INTEREST  = "92D050";  // Đang quan tâm — Xanh lá
const STATUS_NEW       = "D6E4F0";  // Mới tiếp cận — Xanh nhạt

// ─── Page setup (A4) ─────────────────────────────────────────────────────────
const PAGE_W    = 11906;   // A4 width  in DXA
const PAGE_H    = 16838;   // A4 height in DXA
const MARGIN    = 1134;    // ~2 cm margins
const CONTENT_W = PAGE_W - MARGIN * 2;  // 9638 DXA

// ─── Border helper ───────────────────────────────────────────────────────────
const b = (color = "CCCCCC", size = 4) => ({
  top:    { style: BorderStyle.SINGLE, size, color },
  bottom: { style: BorderStyle.SINGLE, size, color },
  left:   { style: BorderStyle.SINGLE, size, color },
  right:  { style: BorderStyle.SINGLE, size, color },
});

// ─── Cell helper ─────────────────────────────────────────────────────────────
function cell(text, w, { fill = null, bold = false, color = "000000", center = false } = {}) {
  return new TableCell({
    borders: b(),
    width: { size: w, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: [new TextRun({ text, bold, color, font: "Arial", size: 20 })]
    })]
  });
}

// ─── Header row helper ───────────────────────────────────────────────────────
function headerRow(labels, widths) {
  return new TableRow({
    tableHeader: true,
    children: labels.map((lbl, i) => cell(lbl, widths[i], { fill: BECAMEX_BLUE, bold: true, color: WHITE, center: true }))
  });
}

// ─── Section heading ─────────────────────────────────────────────────────────
function sectionHeading(num, title) {
  return new Paragraph({
    spacing: { before: 300, after: 120 },
    children: [
      new TextRun({ text: `${num}. ${title}`, bold: true, size: 26, font: "Arial", color: BECAMEX_BLUE }),
    ],
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BECAMEX_ACCENT, space: 2 } }
  });
}

// ─── Bullet paragraph ────────────────────────────────────────────────────────
function bullet(text, bold = false) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    children: [new TextRun({ text, font: "Arial", size: 22, bold })]
  });
}

// ─── Normal paragraph ────────────────────────────────────────────────────────
function para(text, { bold = false, size = 22, spacing = {} } = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60, ...spacing },
    children: [new TextRun({ text, font: "Arial", size, bold })]
  });
}

// ─── Sub-label (bold label + value) ─────────────────────────────────────────
function kv(label, value) {
  return new Paragraph({
    spacing: { before: 40, after: 40 },
    children: [
      new TextRun({ text: label + ": ", font: "Arial", size: 22, bold: true }),
      new TextRun({ text: value, font: "Arial", size: 22 }),
    ]
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
//  DOCUMENT
// ═══════════════════════════════════════════════════════════════════════════════
const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "•",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 480, hanging: 240 } } }
      }]
    }]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: BECAMEX_BLUE },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: BECAMEX_ACCENT },
        paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 1 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: PAGE_H },
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }
      }
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            shading: { fill: BECAMEX_BLUE, type: ShadingType.CLEAR },
            spacing: { before: 80, after: 80 },
            children: [
              new TextRun({ text: "BECAMEX IDC", bold: true, size: 24, font: "Arial", color: WHITE }),
              new TextRun({ text: "  |  BÁO CÁO HOẠT ĐỘNG XÚC TIẾN ĐẦU TƯ", size: 22, font: "Arial", color: WHITE }),
            ]
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: BECAMEX_ACCENT, space: 2 } },
            children: [
              new TextRun({ text: "Nội bộ — Becamex IDC\t", font: "Arial", size: 18, color: "666666" }),
              new TextRun({ text: "Trang ", font: "Arial", size: 18, color: "666666" }),
              new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: "666666" }),
            ]
          })
        ]
      })
    },
    children: [

      // ── TITLE BLOCK ──────────────────────────────────────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 60 },
        children: [new TextRun({
          text: "BÁO CÁO HOẠT ĐỘNG XÚC TIẾN ĐẦU TƯ TUẦN",
          bold: true, font: "Arial", size: 36, color: BECAMEX_BLUE
        })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({
          text: "Tuần 19 (05/05 – 09/05/2026)",
          bold: true, font: "Arial", size: 28, color: BECAMEX_ACCENT
        })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 300 },
        children: [new TextRun({ text: "Ngày lập: 09/05/2026", font: "Arial", size: 22, color: "666666" })]
      }),

      // ── SECTION 1: TÓM TẮT ĐIỀU HÀNH ────────────────────────────────────
      sectionHeading("I", "TÓM TẮT ĐIỀU HÀNH"),
      para("Số liệu nhanh trong tuần:", { bold: true, spacing: { before: 120, after: 40 } }),

      new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: [Math.floor(CONTENT_W * 0.6), Math.ceil(CONTENT_W * 0.4)],
        rows: [
          new TableRow({ children: [
            cell("Chỉ số", Math.floor(CONTENT_W * 0.6), { fill: BECAMEX_BLUE, bold: true, color: WHITE }),
            cell("Tuần 19", Math.ceil(CONTENT_W * 0.4), { fill: BECAMEX_BLUE, bold: true, color: WHITE, center: true }),
          ]}),
          new TableRow({ children: [
            cell("Số cuộc gặp / site visit nhà đầu tư tiềm năng", Math.floor(CONTENT_W * 0.6), { fill: WHITE }),
            cell("1", Math.ceil(CONTENT_W * 0.4), { fill: ZEBRA, center: true }),
          ]}),
          new TableRow({ children: [
            cell("Số proposal gửi đi", Math.floor(CONTENT_W * 0.6), { fill: ZEBRA }),
            cell("4", Math.ceil(CONTENT_W * 0.4), { fill: WHITE, center: true }),
          ]}),
          new TableRow({ children: [
            cell("Số hoạt động phòng thương mại / hiệp hội", Math.floor(CONTENT_W * 0.6), { fill: WHITE }),
            cell("0", Math.ceil(CONTENT_W * 0.4), { fill: ZEBRA, center: true }),
          ]}),
          new TableRow({ children: [
            cell("Số sự kiện tham dự / tổ chức", Math.floor(CONTENT_W * 0.6), { fill: ZEBRA }),
            cell("1", Math.ceil(CONTENT_W * 0.4), { fill: WHITE, center: true }),
          ]}),
          new TableRow({ children: [
            cell("Business cards thu được", Math.floor(CONTENT_W * 0.6), { fill: WHITE }),
            cell("8", Math.ceil(CONTENT_W * 0.4), { fill: ZEBRA, center: true }),
          ]}),
        ]
      }),

      para("Điểm nổi bật tuần 19:", { bold: true, spacing: { before: 160, after: 60 } }),
      bullet("Site visit với Hyundai Engineering (Hàn Quốc) diễn ra thuận lợi — nhà đầu tư dự kiến phát hành LOI.", true),
      bullet("Gửi proposal cho 4 công ty tiềm năng trong tuần (follow-up pipeline hiện tại)."),
      bullet("Tham dự Vietnam Investment Forum 2026 tại TP.HCM ngày 07/05 với vai trò diễn giả; thu được 8 business cards."),
      bullet("Không có hoạt động phòng thương mại hoặc hiệp hội trong tuần này."),

      // ── SECTION 2: GẶP GỠ NHÀ ĐẦU TƯ TIỀM NĂNG ─────────────────────────
      sectionHeading("II", "GẶP GỠ NHÀ ĐẦU TƯ TIỀM NĂNG"),
      para("Bảng tổng hợp:", { bold: true, spacing: { before: 120, after: 60 } }),

      (() => {
        const cols = [400, 1700, 1000, 1400, 1200, 1838, 1100];
        const total = cols.reduce((a, b) => a + b, 0); // = 8638, close enough to CONTENT_W
        return new Table({
          width: { size: CONTENT_W, type: WidthType.DXA },
          columnWidths: cols,
          rows: [
            headerRow(["STT", "Tên công ty", "Quốc gia", "Ngành", "Hình thức", "Trạng thái pipeline", "Người phụ trách"], cols),
            new TableRow({ children: [
              cell("1", cols[0], { fill: ZEBRA, center: true }),
              cell("Hyundai Engineering", cols[1], { fill: ZEBRA }),
              cell("Hàn Quốc", cols[2], { fill: ZEBRA }),
              cell("Xây dựng (khu nhà ở công nhân)", cols[3], { fill: ZEBRA }),
              cell("Site visit", cols[4], { fill: ZEBRA }),
              cell("Chờ quyết định", cols[5], { fill: STATUS_WAIT, center: true }),
              cell("[Cần bổ sung]", cols[6], { fill: ZEBRA }),
            ]}),
            new TableRow({ children: [
              cell("2", cols[0], { fill: WHITE, center: true }),
              cell("[Công ty A — đã gửi proposal]", cols[1], { fill: WHITE }),
              cell("[Cần bổ sung]", cols[2], { fill: WHITE }),
              cell("[Cần bổ sung]", cols[3], { fill: WHITE }),
              cell("Gửi proposal", cols[4], { fill: WHITE }),
              cell("Đang quan tâm", cols[5], { fill: STATUS_INTEREST, center: true }),
              cell("[Cần bổ sung]", cols[6], { fill: WHITE }),
            ]}),
            new TableRow({ children: [
              cell("3", cols[0], { fill: ZEBRA, center: true }),
              cell("[Công ty B — đã gửi proposal]", cols[1], { fill: ZEBRA }),
              cell("[Cần bổ sung]", cols[2], { fill: ZEBRA }),
              cell("[Cần bổ sung]", cols[3], { fill: ZEBRA }),
              cell("Gửi proposal", cols[4], { fill: ZEBRA }),
              cell("Đang quan tâm", cols[5], { fill: STATUS_INTEREST, center: true }),
              cell("[Cần bổ sung]", cols[6], { fill: ZEBRA }),
            ]}),
            new TableRow({ children: [
              cell("4", cols[0], { fill: WHITE, center: true }),
              cell("[Công ty C — đã gửi proposal]", cols[1], { fill: WHITE }),
              cell("[Cần bổ sung]", cols[2], { fill: WHITE }),
              cell("[Cần bổ sung]", cols[3], { fill: WHITE }),
              cell("Gửi proposal", cols[4], { fill: WHITE }),
              cell("Đang quan tâm", cols[5], { fill: STATUS_INTEREST, center: true }),
              cell("[Cần bổ sung]", cols[6], { fill: WHITE }),
            ]}),
            new TableRow({ children: [
              cell("5", cols[0], { fill: ZEBRA, center: true }),
              cell("[Công ty D — đã gửi proposal]", cols[1], { fill: ZEBRA }),
              cell("[Cần bổ sung]", cols[2], { fill: ZEBRA }),
              cell("[Cần bổ sung]", cols[3], { fill: ZEBRA }),
              cell("Gửi proposal", cols[4], { fill: ZEBRA }),
              cell("Đang quan tâm", cols[5], { fill: STATUS_INTEREST, center: true }),
              cell("[Cần bổ sung]", cols[6], { fill: ZEBRA }),
            ]}),
          ]
        });
      })(),

      // Detail block: Hyundai Engineering
      para("Chi tiết cuộc gặp đáng chú ý:", { bold: true, spacing: { before: 200, after: 80 } }),
      new Paragraph({
        spacing: { before: 60, after: 40 },
        shading: { fill: ZEBRA, type: ShadingType.CLEAR },
        border: { left: { style: BorderStyle.SINGLE, size: 12, color: STATUS_WAIT, space: 4 } },
        indent: { left: 240, right: 240 },
        children: [new TextRun({ text: "Hyundai Engineering — Site Visit", bold: true, font: "Arial", size: 24, color: BECAMEX_BLUE })]
      }),
      kv("Quốc gia / Ngành", "Hàn Quốc — Xây dựng (khu nhà ở công nhân)"),
      kv("Hình thức", "Site visit tại khu công nghiệp Becamex IDC"),
      kv("Kết quả", "Buổi thăm địa điểm diễn ra thuận lợi; phía Hyundai Engineering đánh giá tích cực. Công ty dự kiến phát hành LOI (Letter of Intent) trong thời gian sắp tới."),
      kv("Trạng thái pipeline", "Chờ quyết định"),
      kv("Bước tiếp theo", "Chờ Hyundai Engineering phát hành LOI; chuẩn bị hồ sơ tiếp theo khi nhận được LOI."),
      kv("Deadline follow-up", "[Cần bổ sung]"),

      // ── SECTION 3: HOẠT ĐỘNG PHÒNG THƯƠNG MẠI & HIỆP HỘI ────────────────
      sectionHeading("III", "HOẠT ĐỘNG PHÒNG THƯƠNG MẠI & HIỆP HỘI"),

      new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: [1400, 1800, 1200, 2200, 2038],
        rows: [
          headerRow(["Tổ chức", "Loại hoạt động", "Ngày", "Kết quả", "Follow-up"],
                    [1400, 1800, 1200, 2200, 2038]),
          new TableRow({ children: [
            cell("—", 1400, { fill: ZEBRA, center: true }),
            cell("Không có hoạt động trong tuần 19", 1800, { fill: ZEBRA }),
            cell("—", 1200, { fill: ZEBRA, center: true }),
            cell("—", 2200, { fill: ZEBRA }),
            cell("—", 2038, { fill: ZEBRA }),
          ]}),
        ]
      }),

      para("Ghi chú: Tuần 19 (05/05 – 09/05/2026) không phát sinh hoạt động với các phòng thương mại hoặc hiệp hội ngành nghề (KCCI, KOCHAM, JETRO, EuroCham, AmCham, VCCI, v.v.).", { spacing: { before: 80, after: 60 } }),

      // ── SECTION 4: SỰ KIỆN XÚC TIẾN ĐẦU TƯ ─────────────────────────────
      sectionHeading("IV", "SỰ KIỆN XÚC TIẾN ĐẦU TƯ"),

      (() => {
        const cols = [2200, 1400, 1200, 1400, 1200, 2238];
        return new Table({
          width: { size: CONTENT_W, type: WidthType.DXA },
          columnWidths: cols,
          rows: [
            headerRow(["Tên sự kiện", "Địa điểm", "Ngày", "Vai trò Becamex", "Số contacts", "Leads tiềm năng"], cols),
            new TableRow({ children: [
              cell("Vietnam Investment Forum 2026", cols[0], { fill: ZEBRA }),
              cell("TP. Hồ Chí Minh", cols[1], { fill: ZEBRA }),
              cell("07/05/2026", cols[2], { fill: ZEBRA, center: true }),
              cell("Diễn giả (Tham dự)", cols[3], { fill: ZEBRA }),
              cell("8 business cards", cols[4], { fill: ZEBRA, center: true }),
              cell("[Cần bổ sung sau khi phân loại contacts]", cols[5], { fill: ZEBRA }),
            ]}),
          ]
        });
      })(),

      para("Chi tiết: Becamex IDC tham dự Vietnam Investment Forum 2026 tại TP. Hồ Chí Minh với vai trò diễn giả vào ngày 07/05/2026. Đây là cơ hội nâng cao nhận diện thương hiệu và thu hút nhà đầu tư tiềm năng. Thu được 8 business cards trong sự kiện; cần phân loại và theo dõi các contacts này trong tuần tới.", { spacing: { before: 100, after: 60 } }),

      // ── SECTION 5: KẾ HOẠCH TUẦN TỚI ────────────────────────────────────
      sectionHeading("V", "KẾ HOẠCH TUẦN TỚI (12/05 – 16/05/2026)"),
      para("Các cuộc gặp / follow-up đã dự kiến:", { bold: true, spacing: { before: 120, after: 60 } }),
      bullet("Theo dõi và nhận LOI từ Hyundai Engineering — người phụ trách: [Cần bổ sung]."),
      bullet("Phân loại 8 contacts thu được từ Vietnam Investment Forum 2026; lên kế hoạch tiếp cận."),
      bullet("Follow-up phản hồi từ 4 công ty đã nhận proposal trong tuần 19."),
      para("Sự kiện:", { bold: true, spacing: { before: 120, after: 60 } }),
      bullet("[Cần bổ sung nếu có sự kiện tuần 20]"),
      para("Công việc cần hoàn thành:", { bold: true, spacing: { before: 120, after: 60 } }),
      bullet("Chuẩn bị hồ sơ pháp lý / đề xuất chi tiết sẵn sàng cho bước tiếp theo với Hyundai Engineering."),
      bullet("Tổng hợp và lưu trữ business cards từ Vietnam Investment Forum 2026."),

      // ── FOOTER BLOCK ─────────────────────────────────────────────────────
      new Paragraph({
        spacing: { before: 400, after: 60 },
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: BECAMEX_ACCENT, space: 4 } },
        children: []
      }),
      new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: [Math.floor(CONTENT_W / 2), Math.ceil(CONTENT_W / 2)],
        rows: [
          new TableRow({ children: [
            new TableCell({
              borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
              width: { size: Math.floor(CONTENT_W / 2), type: WidthType.DXA },
              children: [
                para("Người lập:", { bold: true }),
                para("_______________________________"),
                para("Ngày: 09/05/2026"),
              ]
            }),
            new TableCell({
              borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
              width: { size: Math.ceil(CONTENT_W / 2), type: WidthType.DXA },
              children: [
                para("Phân phối:", { bold: true }),
                para("Ban lãnh đạo / [Danh sách nhận]"),
                para("Phân loại: Nội bộ"),
              ]
            }),
          ]})
        ]
      }),

    ] // end children
  }]
});

// ─── Write output ─────────────────────────────────────────────────────────────
const OUTPUT = "/home/thehien/Projects/becamex/marketing_assitance_skills/.claude/skills/promote-activity-summary-workspace/iteration-1/weekly-w19-light/with_skill/outputs/promote_activity_weekly_2026-W19.docx";

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUTPUT, buf);
  console.log("Created: " + OUTPUT);
}).catch(err => {
  console.error("Error:", err);
  process.exit(1);
});
