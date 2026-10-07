const express = require("express");
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, ShadingType, Header, Footer,
  PageNumber, TabStopType, VerticalAlign, HeightRule,
} = require("docx");

const router = express.Router();
const LOGO_DIR = path.join(__dirname, "assets", "eia");
const FONT = "Outfit"; // Word substitutes it if not installed; use "Calibri" or "Arial" to be safe
const SUPERVISOR = "Mario Natan Jr.";

// A4, margins in twips (1px = 15 twips): left 76px, right 32px
const CONTENT_W = 11906 - 1140 - 480; // 10286

const COMPANIES = {
  sgc: {
    name: "SILVERGRAPH CORPORATION", logo: "EIALogoSGC.png", w: 74, h: 70,
    address: ["1124 Triumph Square", "1618 Quezon Avenue", "Brgy. South Triangle",
      "District 4 1104 Quezon City", "info@silvergraph.ai"],
    formCode: "SDB-FRM-IT-000001 Rev 0", rev: "04-April-2023",
  },
  sdb: {
    name: "SILVERDAB CORPORATION", logo: "EIALogoSDB.png", w: 70, h: 74,
    address: ["7th Floor Unit 3, Hexagon Corporate Center,", "1471 Quezon Ave., West Triangle,",
      "Quezon City, Philippines 1104", "info@silverdab.com | (2) 5322-1900"],
    formCode: "SDB-FRM-IT-000001 Rev 0", rev: "04-April-2023",
  },
  ocg: {
    name: "Oriental Consultants Global Co. Ltd", logo: "EIALogoOCG.png", w: 130, h: 67,
    address: ["Unit 12 - 14th Flr Triumph Square, 1618,", "Quezon Ave, Brgy. South Triangle,", "Quezon City"],
    formCode: "OCG-FRM-IT-000001 Rev 0", rev: "04-April-2023",
  },
};

// ── date helpers (same as the frontend) ─────────────────────────────────────
const MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const isoOk = (v) => /^\d{4}-\d{2}-\d{2}$/.test(v || "");
const fmtLong = (v) => { if (!isoOk(v)) return v || ""; const [y, m, d] = v.split("-").map(Number); return `${MONTHS[m - 1]} ${d}, ${y}`; };
const fmtShort = (v) => { if (!isoOk(v)) return v || ""; const [y, m, d] = v.split("-").map(Number); return `${m}-${d}-${String(y).slice(2)}`; };

// ── building blocks ─────────────────────────────────────────────────────────
const NONE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const LINE = { style: BorderStyle.SINGLE, size: 6, color: "000000" };
const noBorders = { top: NONE, bottom: NONE, left: NONE, right: NONE };
const allBorders = { top: LINE, bottom: LINE, left: LINE, right: LINE };
const underline = { top: NONE, left: NONE, right: NONE, bottom: LINE };
const TBL_NONE = { ...noBorders, insideHorizontal: NONE, insideVertical: NONE };
const DXA = WidthType.DXA;
const CENTER = AlignmentType.CENTER;
const RIGHT = AlignmentType.RIGHT;

const run = (text, o = {}) => new TextRun({ text: text || "", font: FONT, size: 22, ...o });
const para = (text, o = {}, p = {}) => new Paragraph({ children: [run(text, o)], ...p });
const lines = (text, o = {}, p = {}) => String(text || "").split("\n").map((l) => para(l, o, p));
const spacer = (after = 160) => new Paragraph({ children: [], spacing: { after } });

const cell = (children, width, o = {}) =>
  new TableCell({
    width: { size: width, type: DXA },
    children: Array.isArray(children) ? children : [children],
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    ...o,
  });

const buildHeader = (c) => {
  const logo = new ImageRun({
    type: "png",
    data: fs.readFileSync(path.join(LOGO_DIR, c.logo)),
    transformation: { width: c.w, height: c.h },
  });
  const right = [
    para(c.name, { bold: true, size: 24 }, { alignment: RIGHT }),
    ...c.address.map((a) => para(a, { size: 20 }, { alignment: RIGHT })),
  ];
  return new Header({
    children: [
      new Table({
        width: { size: CONTENT_W, type: DXA },
        columnWidths: [3000, CONTENT_W - 3000],
        borders: TBL_NONE,
        rows: [new TableRow({ children: [
          cell(new Paragraph({ children: [logo] }), 3000, { borders: noBorders }),
          cell(right, CONTENT_W - 3000, { borders: noBorders, verticalAlign: VerticalAlign.TOP }),
        ] })],
      }),
      new Paragraph({ children: [], border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: "000000", space: 4 } } }),
    ],
  });
};

const buildFooter = (c) =>
  new Footer({
    children: [
      new Paragraph({
        tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
        children: [
          run(c.formCode, { size: 16 }),
          run("\t", { size: 16 }),
          new TextRun({ children: [PageNumber.CURRENT, " | ", PageNumber.TOTAL_PAGES], font: FONT, size: 18, color: "888888" }),
        ],
      }),
      para(c.rev, { size: 16 }),
    ],
  });

const sigBlock = (label, name, date, cap, w) => {
  const nameW = w - 1200 - 200;
  return new Table({
    width: { size: w, type: DXA },
    columnWidths: [nameW, 200, 1200],
    borders: TBL_NONE,
    rows: [
      new TableRow({ children: [new TableCell({ columnSpan: 3, width: { size: w, type: DXA }, borders: noBorders, children: [para(label)] })] }),
      new TableRow({
        height: { value: 800, rule: HeightRule.ATLEAST },
        children: [
          cell(para(name, {}, { alignment: CENTER }), nameW, { borders: underline, verticalAlign: VerticalAlign.BOTTOM }),
          cell(para(""), 200, { borders: noBorders }),
          cell(para(date, {}, { alignment: CENTER }), 1200, { borders: underline, verticalAlign: VerticalAlign.BOTTOM }),
        ],
      }),
      new TableRow({ children: [
        cell(para(cap, { size: 20 }, { alignment: CENTER }), nameW, { borders: noBorders }),
        cell(para(""), 200, { borders: noBorders }),
        cell(para("Date", { size: 20 }, { alignment: CENTER }), 1200, { borders: noBorders }),
      ] }),
    ],
  });
};

const HALF = 4900;
const GAP = CONTENT_W - HALF * 2; // 486
const flush = { top: 0, bottom: 0, left: 0, right: 0 };
const sigCell = (block, w) => cell([block, new Paragraph({ children: [] })], w, { borders: noBorders, margins: flush, verticalAlign: VerticalAlign.TOP });
const emptyCell = (w) => cell(new Paragraph({ children: [] }), w, { borders: noBorders });

// ── document ────────────────────────────────────────────────────────────────
const buildDoc = (d) => {
  const c = COMPANIES[d.company];
  const items = Array.isArray(d.items) ? d.items : [];
  const pad = Math.max(0, 5 - items.length);

  const title = new Table({
    width: { size: CONTENT_W, type: DXA }, columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [cell(
      para("EQUIPMENT ISSUANCE AGREEMENT SIGN SHEET", { bold: true, size: 26 }, { alignment: CENTER }),
      CONTENT_W, { borders: allBorders, shading: { type: ShadingType.CLEAR, fill: "D9D9D9", color: "auto" } })] })],
  });

  const metaW = [1500, 4300, 1100, 3386];
  const lab = (t, w) => cell(para(t), w, { borders: noBorders });
  const val = (t, w) => cell(para(t), w, { borders: underline });
  const blank = (w) => cell(para(""), w, { borders: noBorders });
  const meta = new Table({
    width: { size: CONTENT_W, type: DXA }, columnWidths: metaW, borders: TBL_NONE,
    rows: [
      new TableRow({ children: [lab("Name:", metaW[0]), val(d.name, metaW[1]), lab("Date:", metaW[2]), val(fmtLong(d.date), metaW[3])] }),
      new TableRow({ children: [lab("Department:", metaW[0]), val(d.department, metaW[1]), lab("Ref. No.:", metaW[2]), val(d.refNo, metaW[3])] }),
      new TableRow({ children: [lab("Issued No.:", metaW[0]), val(d.issuedNo, metaW[1]), blank(metaW[2]), blank(metaW[3])] }),
    ],
  });

  const cols = [780, 2250, 3836, 1650, 1770];
  const head = new TableRow({
    tableHeader: true, cantSplit: true,
    children: ["Qty.", "Brand / Model", "Description", "Warranty", "Date Purchase"].map((h, i) =>
      cell(para(h, { bold: true, size: 22 }, { alignment: CENTER }), cols[i],
        { borders: allBorders, shading: { type: ShadingType.CLEAR, fill: "D9D9D9", color: "auto" } })),
  });
  const rowOf = (vals) => new TableRow({
    cantSplit: true, height: { value: 450, rule: HeightRule.ATLEAST },
    children: vals.map((v, i) => cell(lines(v, { size: 22 }, { alignment: CENTER }), cols[i], { borders: allBorders })),
  });
  const table = new Table({
    width: { size: CONTENT_W, type: DXA }, columnWidths: cols,
    rows: [
      head,
      ...items.map((r) => rowOf([r.qty, r.brand, r.desc, r.warranty, r.purchased])),
      ...Array.from({ length: pad }, () => rowOf(["", "", "", "", ""])),
    ],
  });

  const remarks = new Table({
    width: { size: CONTENT_W, type: DXA }, columnWidths: [CONTENT_W],
    rows: [new TableRow({ cantSplit: true, children: [cell(
      [para("Remarks:", { bold: true }, { spacing: { after: 100 } }), ...lines(d.remarks)],
      CONTENT_W, { borders: allBorders, margins: { top: 100, bottom: 120, left: 160, right: 160 } })] })],
  });

  const issuedName = (d.issuedTo && d.issuedTo.name) || d.name;
  const sigRow1 = new Table({
    width: { size: CONTENT_W, type: DXA }, columnWidths: [HALF, GAP, HALF], borders: TBL_NONE,
    rows: [new TableRow({ cantSplit: true, children: [
      sigCell(sigBlock("Issued to:", issuedName, fmtShort(d.issuedTo && d.issuedTo.date), "Signature Over Printed Name", HALF), HALF),
      emptyCell(GAP),
      sigCell(sigBlock("Delivered by:", d.deliveredBy && d.deliveredBy.name, fmtShort(d.deliveredBy && d.deliveredBy.date), "IT Support", HALF), HALF),
    ] })],
  });
  const sigRow2 = new Table({
    width: { size: CONTENT_W, type: DXA }, columnWidths: [HALF, CONTENT_W - HALF], borders: TBL_NONE,
    rows: [new TableRow({ cantSplit: true, children: [
      sigCell(sigBlock("Approved by:", SUPERVISOR, fmtShort(d.approvedBy && d.approvedBy.date), "IT Supervisor", HALF), HALF),
      emptyCell(CONTENT_W - HALF),
    ] })],
  });

  return new Document({
    sections: [{
      properties: { page: {
        size: { width: 11906, height: 16838 },
        // bottom margin = blank space above the footer; raise it for more room
        margin: { top: 720, bottom: 1200, left: 1140, right: 480, header: 500, footer: 400 },
      } },
      headers: { default: buildHeader(c) },
      footers: { default: buildFooter(c) },
      children: [
        title,
        new Paragraph({ children: [run(d.copyLabel || "", { italics: true, size: 21 })], spacing: { before: 80, after: 120 } }),
        meta, spacer(240),
        table, spacer(320),
        remarks, spacer(400),
        sigRow1, spacer(300),
        sigRow2,
      ],
    }],
  });
};

// ── route ───────────────────────────────────────────────────────────────────
router.post("/eia/export", async (req, res) => {
  try {
    const d = req.body || {};
    if (!COMPANIES[d.company]) return res.status(400).json({ error: "Unknown company" });

    const buf = await Packer.toBuffer(buildDoc(d));
    const safe = (s) => String(s || "").replace(/[^\w.-]+/g, "_");
    const filename = `${safe(d.refNo) || "EIA"}_${safe(d.name) || "Employee"}.docx`;

    res.set({
      "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "Content-Disposition": `attachment; filename="${filename}"`,
    });
    res.send(buf);
  } catch (err) {
    console.error("EIA export failed:", err);
    res.status(500).json({ error: "Export failed" });
  }
});

module.exports = router;