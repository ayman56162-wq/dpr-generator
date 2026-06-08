const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign
} = require('docx');
const XLSX = require('xlsx');
const fs   = require('fs');
const path = require('path');
const os   = require('os');

// ─── Parse multipart manually (Vercel edge) ───────────────────────────────
const { IncomingForm } = require('formidable');

// ─── Colours / constants ──────────────────────────────────────────────────
const GOLD      = "C9A84C";
const DARK_BLUE = "1F3864";
const LIGHT_BLUE= "D6E4F0";
const WHITE     = "FFFFFF";
const LGRAY     = "F2F2F2";
const TEXT_DARK = "1A1A1A";

const PAGE_W = 11906, PAGE_H = 16838, MARGIN = 576;
const CW = PAGE_W - MARGIN * 2; // 10754

const BD   = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const BDS  = { top: BD, bottom: BD, left: BD, right: BD };
const NB   = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const NBDS = { top: NB, bottom: NB, left: NB, right: NB };

function mc(text, o = {}) {
  const { bold=false, color=TEXT_DARK, bg=WHITE, align=AlignmentType.LEFT,
    size=15, colspan=1, rowspan=1, valign=VerticalAlign.CENTER,
    w=null, bds=BDS, italic=false } = o;
  return new TableCell({
    columnSpan: colspan, rowSpan: rowspan, verticalAlign: valign,
    width: w ? { size: w, type: WidthType.DXA } : undefined,
    shading: { fill: bg, type: ShadingType.CLEAR },
    margins: { top: 50, bottom: 50, left: 80, right: 80 },
    borders: bds,
    children: [new Paragraph({
      alignment: align, spacing: { before: 0, after: 0 },
      children: [new TextRun({ text: String(text ?? ''), bold, color, size, font: "Arial", italics: italic })]
    })]
  });
}

function hc(text, o = {}) {
  return mc(text, { bold: true, color: WHITE, bg: DARK_BLUE, align: AlignmentType.CENTER, size: 15, ...o });
}

function banner(label) {
  return new Table({
    width: { size: CW, type: WidthType.DXA }, columnWidths: [CW],
    rows: [new TableRow({ children: [new TableCell({
      shading: { fill: GOLD, type: ShadingType.CLEAR }, borders: BDS,
      margins: { top: 80, bottom: 80, left: 140, right: 140 },
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: label, bold: true, color: WHITE, size: 18, font: "Arial" })]
      })]
    })] })]
  });
}

const sp = (pts = 70) => new Paragraph({ spacing: { before: pts, after: 0 }, children: [] });

function groupActs(acts) {
  const out = []; let cur = null;
  for (const r of acts) {
    if (!cur || r.system !== cur.system) { cur = { system: r.system, rows: [] }; out.push(cur); }
    cur.rows.push(r);
  }
  return out;
}

// Col widths: System=1500, Zone=1400, Level=700, Discipline=1800, Work=3754, Qty=700, Unit=900 → 10754
const AC = [1500, 1400, 700, 1800, 3754, 700, 900];

function buildActTable(rows) {
  const groups = groupActs(rows);
  const tRows = [new TableRow({ tableHeader: true, children: [
    hc("SYSTEM",       { w: AC[0] }),
    hc("ZONE",         { w: AC[1] }),
    hc("LEVEL",        { w: AC[2] }),
    hc("DISCIPLINE",   { w: AC[3] }),
    hc("TYPE OF WORK", { w: AC[4] }),
    hc("QTY",          { w: AC[5], align: AlignmentType.CENTER }),
    hc("UNIT",         { w: AC[6], align: AlignmentType.CENTER }),
  ]})];

  for (const grp of groups) {
    const cnt = grp.rows.length;
    grp.rows.forEach((row, i) => {
      const cells = [];
      if (i === 0) {
        cells.push(new TableCell({
          rowSpan: cnt, verticalAlign: VerticalAlign.CENTER,
          shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR }, borders: BDS,
          width: { size: AC[0], type: WidthType.DXA },
          margins: { top: 50, bottom: 50, left: 80, right: 80 },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: grp.system, bold: true, size: 15, font: "Arial", color: DARK_BLUE })]
          })]
        }));
      }
      cells.push(mc(row.zone,       { size: 15, w: AC[1] }));
      cells.push(mc(row.level,      { size: 15, align: AlignmentType.CENTER, w: AC[2] }));
      cells.push(mc(row.discipline, { size: 15, w: AC[3] }));
      cells.push(mc(row.work,       { size: 15, w: AC[4] }));
      cells.push(mc(row.qty,        { size: 15, align: AlignmentType.CENTER, w: AC[5] }));
      cells.push(mc(row.unit,       { size: 15, align: AlignmentType.CENTER, w: AC[6] }));
      tRows.push(new TableRow({ children: cells }));
    });
  }
  return new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: AC, rows: tRows });
}

// ─── Parse Excel files ────────────────────────────────────────────────────
function parseExcel(filePath) {
  const wb  = XLSX.readFile(filePath);
  const ws  = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  const activities = [];
  const nextDay    = [];
  let section = null;
  let headerRow = null;

  for (let i = 0; i < raw.length; i++) {
    const row = raw[i];
    const c0  = String(row[0] || '').trim();

    if (c0 === "Today's Activities") { section = 'today'; headerRow = null; continue; }
    if (c0 === "Next Day Activities")  { section = 'next';  headerRow = null; continue; }

    if (!section) continue;

    // Detect header row (contains "Zone" or "Level")
    const rowStr = row.map(v => String(v).toLowerCase()).join('|');
    if (!headerRow && (rowStr.includes('zone') || rowStr.includes('level'))) {
      headerRow = row.map(v => String(v).trim().toLowerCase());
      continue;
    }
    if (!headerRow) continue;

    // Skip empty rows
    const num = String(row[0]).trim();
    if (!num || isNaN(parseFloat(num))) continue;

    // Map columns dynamically
    const get = (key) => {
      const idx = headerRow.findIndex(h => h.includes(key));
      return idx >= 0 ? String(row[idx] || '').trim() : '';
    };

    const zone       = get('zone');
    const level      = get('level');
    const discipline = get('discipline') || get('disc');
    const system     = get('system');
    let   work       = get('type') || get('work');
    const qty        = get('qty') || get('quantity') || get('count');
    const unit       = get('unit');
    const remark     = get('remark') || get('remarks');

    // Use remark if work is empty
    if (!work && remark) work = remark;
    if (!work) continue;

    const entry = { zone, level, discipline, system: system || 'General', work, qty, unit };

    if (section === 'today') activities.push(entry);
    if (section === 'next')  nextDay.push(entry);
  }

  return { activities, nextDay };
}

// ─── Build Word doc ───────────────────────────────────────────────────────
async function buildDoc({ date, reportNo, supervision, labors, pkg, activities, nextDay, photos, logoPath }) {
  const ch = [];

  // Format date
  const d = new Date(date);
  const dateStr = `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;

  // LOGO
  const logo = fs.readFileSync(logoPath);
  ch.push(new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 130 },
    children: [new ImageRun({ data: logo, type: "png", transformation: { width: 370, height: 70 } })]
  }));

  // GENERAL INFO
  const gi = [2100, 3900, 1754, 3000];
  ch.push(new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: gi, rows: [
    new TableRow({ children: [
      mc("Contractor / SubCont.", { bold: true, bg: LGRAY, size: 14, w: gi[0] }),
      mc("MBL / First Fix",       { size: 14, w: gi[1] }),
      mc("Package",               { bold: true, bg: LGRAY, size: 14, w: gi[2] }),
      mc(pkg,                     { size: 14, w: gi[3] }),
    ]}),
    new TableRow({ children: [
      mc("Client",                { bold: true, bg: LGRAY, size: 14, w: gi[0] }),
      mc("Jeddah Central Development Co.", { size: 14, w: gi[1] }),
      mc("Subcontract Ref.",      { bold: true, bg: LGRAY, size: 14, w: gi[2] }),
      mc("P1020096-OCN-014",      { size: 14, w: gi[3] }),
    ]}),
  ]}));

  ch.push(sp(90));

  // DATE / REPORT / MANPOWER + TITLE
  const leftW = 4800, titleW = CW - 4800;
  ch.push(new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [leftW, titleW],
    rows: [new TableRow({ children: [
      new TableCell({
        verticalAlign: VerticalAlign.CENTER,
        shading: { fill: WHITE, type: ShadingType.CLEAR }, borders: BDS,
        width: { size: leftW, type: WidthType.DXA }, margins: { top: 0, bottom: 0, left: 0, right: 0 },
        children: [new Table({ width: { size: leftW, type: WidthType.DXA }, columnWidths: [1700, 3100], rows: [
          new TableRow({ children: [mc("Date",      { bold: true, bg: LGRAY, size: 16, w: 1700 }), mc(dateStr,     { size: 16, w: 3100 })] }),
          new TableRow({ children: [mc("Report No", { bold: true, bg: LGRAY, size: 16, w: 1700 }), mc(reportNo,   { size: 16, w: 3100 })] }),
          new TableRow({ children: [mc("Manpower",  { bold: true, bg: LGRAY, align: AlignmentType.CENTER, size: 16, colspan: 2, w: 4800 })] }),
          new TableRow({ children: [mc("Supervision",{ bold: true, bg: LGRAY, align: AlignmentType.CENTER, size: 15, w: 1700 }), mc(String(supervision), { align: AlignmentType.CENTER, size: 16, bold: true, w: 3100 })] }),
          new TableRow({ children: [mc("Labors",    { bold: true, bg: LGRAY, align: AlignmentType.CENTER, size: 15, w: 1700 }), mc(String(labors),      { align: AlignmentType.CENTER, size: 16, bold: true, w: 3100 })] }),
        ]})]
      }),
      new TableCell({
        verticalAlign: VerticalAlign.CENTER,
        shading: { fill: WHITE, type: ShadingType.CLEAR }, borders: BDS,
        width: { size: titleW, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 100, right: 100 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Daily Site Work Report", bold: true, size: 28, color: DARK_BLUE, font: "Arial" })]
        })]
      }),
    ]})]
  }));

  ch.push(sp(100));
  ch.push(banner("DAILY SITE ACTIVITIES"));
  ch.push(sp(50));
  ch.push(new Paragraph({ spacing: { before: 30, after: 50 },
    children: [new TextRun({ text: "MECHANICAL PROGRESS", bold: true, size: 16, color: DARK_BLUE, font: "Arial" })]
  }));

  ch.push(buildActTable(activities));
  ch.push(sp(100));

  ch.push(banner("next day activities"));
  ch.push(sp(50));
  if (nextDay.length > 0) {
    ch.push(buildActTable(nextDay));
  } else {
    ch.push(new Paragraph({ spacing: { before: 40, after: 40 },
      children: [new TextRun({ text: "To be updated", size: 15, color: "999999", italics: true, font: "Arial" })]
    }));
  }
  ch.push(sp(100));

  // PHOTOS
  ch.push(banner("PHOTOS"));
  ch.push(sp(70));
  const photoW = Math.floor(CW / 2);
  for (let i = 0; i < photos.length; i += 2) {
    const cells = [];
    for (let j = i; j < Math.min(i + 2, photos.length); j++) {
      const buf = fs.readFileSync(photos[j]);
      const ext = path.extname(photos[j]).toLowerCase().replace('.', '');
      const type = ext === 'jpg' ? 'jpeg' : ext;
      cells.push(new TableCell({
        borders: NBDS, width: { size: photoW, type: WidthType.DXA },
        margins: { top: 50, bottom: 50, left: 50, right: 50 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new ImageRun({ data: buf, type, transformation: { width: 250, height: 188 } })]
        })]
      }));
    }
    if (cells.length === 1) cells.push(new TableCell({ borders: NBDS, width: { size: photoW, type: WidthType.DXA }, children: [new Paragraph({ children: [] })] }));
    ch.push(new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [photoW, photoW], rows: [new TableRow({ children: cells })] }));
    ch.push(sp(40));
  }

  const doc = new Document({
    styles: { default: { document: { run: { font: "Arial", size: 15, color: TEXT_DARK } } } },
    sections: [{ properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN } } }, children: ch }]
  });

  return Packer.toBuffer(doc);
}

// ─── Vercel handler ───────────────────────────────────────────────────────
module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'dpr-'));

  try {
    // Parse form
    const form = new IncomingForm({ uploadDir: tmpDir, keepExtensions: true, multiples: true });
    const { fields, files } = await new Promise((ok, err) => form.parse(req, (e, f, fi) => e ? err(e) : ok({ fields: f, files: fi })));

    const date       = Array.isArray(fields.date)       ? fields.date[0]       : fields.date;
    const reportNo   = Array.isArray(fields.reportNo)   ? fields.reportNo[0]   : fields.reportNo;
    const supervision= Array.isArray(fields.supervision)? fields.supervision[0]: fields.supervision;
    const labors     = Array.isArray(fields.labors)     ? fields.labors[0]     : fields.labors;
    const pkg        = Array.isArray(fields.package)    ? fields.package[0]    : (fields.package || 'MEP');

    // Excel files
    const excelArr = Array.isArray(files.excel) ? files.excel : (files.excel ? [files.excel] : []);
    const photoArr = Array.isArray(files.photos) ? files.photos : (files.photos ? [files.photos] : []);

    // Parse all Excel files
    let activities = [], nextDay = [];
    for (const f of excelArr) {
      const fp = f.filepath || f.path;
      const parsed = parseExcel(fp);
      activities = activities.concat(parsed.activities);
      nextDay    = nextDay.concat(parsed.nextDay);
    }

    // Logo path (bundled with app)
    const logoPath = path.join(process.cwd(), 'public', 'logo.png');

    const photoPaths = photoArr.map(f => f.filepath || f.path);

    // Build doc
    const docBuffer = await buildDoc({
      date, reportNo, supervision, labors, pkg,
      activities, nextDay,
      photos: photoPaths,
      logoPath
    });

    // Return as docx
    const filename = `DPR_${reportNo}.docx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(docBuffer);

  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  } finally {
    // Cleanup tmp
    try { fs.rmSync(tmpDir, { recursive: true }); } catch {}
  }
};
