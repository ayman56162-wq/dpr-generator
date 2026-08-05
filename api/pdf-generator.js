// pdf-generator.js — generates a PDF version of the DPR using pdf-lib (pure JS, works on Vercel)
const { PDFDocument, StandardFonts, rgb, PageSizes } = require('pdf-lib');

// Colours (as 0-1 RGB for pdf-lib)
const GOLD       = rgb(0.788, 0.659, 0.298); // C9A84C
const DARK_BLUE  = rgb(0.122, 0.220, 0.392); // 1F3864
const LIGHT_BLUE = rgb(0.839, 0.894, 0.941); // D6E4F0
const WHITE      = rgb(1, 1, 1);
const LGRAY      = rgb(0.949, 0.949, 0.949); // F2F2F2
const TEXT_DARK  = rgb(0.102, 0.102, 0.102); // 1A1A1A
const BORDER_GRAY= rgb(0.8, 0.8, 0.8);

const PAGE_W = 595.28;  // A4 width in points
const PAGE_H = 841.89;  // A4 height in points
const MARGIN = 30;
const CW = PAGE_W - MARGIN * 2; // content width

// ── Text wrapping helper ──
function wrapText(text, font, fontSize, maxWidth) {
  text = String(text ?? '');
  if (text === '') return [''];
  const words = text.split(/\s+/);
  const lines = [];
  let current = '';
  for (const word of words) {
    const test = current ? current + ' ' + word : word;
    const width = font.widthOfTextAtSize(test, fontSize);
    if (width > maxWidth && current) {
      lines.push(current);
      current = word;
    } else {
      current = test;
    }
  }
  if (current) lines.push(current);
  return lines.length ? lines : [''];
}

// ── Main builder ──
async function buildPdf({ date, reportNo, supervision, labors, pkg, activities, nextDay, photos, logoBuffer }) {
  const pdfDoc = await PDFDocument.create();
  const font     = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

  let page = pdfDoc.addPage([PAGE_W, PAGE_H]);
  let y = PAGE_H - MARGIN;

  function newPageIfNeeded(neededHeight) {
    if (y - neededHeight < MARGIN + 20) {
      page = pdfDoc.addPage([PAGE_W, PAGE_H]);
      y = PAGE_H - MARGIN;
    }
  }

  // ── Draw a single cell (with wrapping) ──
  function drawCell(x, cellY, w, h, text, opts = {}) {
    const {
      bold = false, size = 8, align = 'left', fill = null,
      color = TEXT_DARK, valign = 'middle'
    } = opts;
    if (fill) {
      page.drawRectangle({ x, y: cellY - h, width: w, height: h, color: fill });
    }
    page.drawRectangle({ x, y: cellY - h, width: w, height: h, borderColor: BORDER_GRAY, borderWidth: 0.5 });

    const f = bold ? fontBold : font;
    const pad = 4;
    const maxW = w - pad * 2;
    const lines = wrapText(text, f, size, maxW);
    const lineHeight = size * 1.25;
    const totalTextH = lines.length * lineHeight;
    let startY;
    if (valign === 'middle') startY = cellY - (h - totalTextH) / 2 - size;
    else startY = cellY - pad - size;

    lines.forEach((line, i) => {
      const lineW = f.widthOfTextAtSize(line, size);
      let tx = x + pad;
      if (align === 'center') tx = x + (w - lineW) / 2;
      if (align === 'right') tx = x + w - pad - lineW;
      page.drawText(line, { x: tx, y: startY - i * lineHeight, size, font: f, color });
    });
  }

  // ── Compute row height needed for a row of cells ──
  function computeRowHeight(cells, colWidths, sizeDefault = 8) {
    let maxLines = 1;
    cells.forEach((c, i) => {
      const size = c.size || sizeDefault;
      const f = c.bold ? fontBold : font;
      const w = colWidths[i] - 8;
      const lines = wrapText(c.text, f, size, w);
      if (lines.length > maxLines) maxLines = lines.length;
    });
    const lineHeight = sizeDefault * 1.25;
    return Math.max(20, maxLines * lineHeight + 8);
  }

  // ── Section banner (gold bar) ──
  function drawBanner(label) {
    newPageIfNeeded(28);
    const h = 24;
    page.drawRectangle({ x: MARGIN, y: y - h, width: CW, height: h, color: GOLD, borderColor: BORDER_GRAY, borderWidth: 0.5 });
    const size = 10;
    const tw = fontBold.widthOfTextAtSize(label, size);
    page.drawText(label, { x: MARGIN + (CW - tw) / 2, y: y - h / 2 - size / 2 + 2, size, font: fontBold, color: WHITE });
    y -= h + 8;
  }

  // ── Activities table (System | Zone | Level | Discipline | Type of Work | Qty | Unit) ──
  function drawActivitiesTable(rows) {
    const colW = [85, 80, 40, 100, 170, 40, 50]; // sums to 565... adjust to CW
    const scale = CW / colW.reduce((a, b) => a + b, 0);
    const AC = colW.map(w => w * scale);

    // Header row
    newPageIfNeeded(24);
    let x = MARGIN;
    const headers = ['SYSTEM', 'ZONE', 'LEVEL', 'DISCIPLINE', 'TYPE OF WORK', 'QTY', 'UNIT'];
    const headerH = 22;
    headers.forEach((h, i) => {
      drawCell(x, y, AC[i], headerH, h, { bold: true, size: 8, align: 'center', fill: DARK_BLUE, color: WHITE });
      x += AC[i];
    });
    y -= headerH;

    // Group rows by system
    const groups = [];
    let cur = null;
    for (const r of rows) {
      if (!cur || r.system !== cur.system) { cur = { system: r.system, rows: [] }; groups.push(cur); }
      cur.rows.push(r);
    }

    for (const grp of groups) {
      // For simplicity, draw system cell per-row but merge visually by leaving blank after first.
      grp.rows.forEach((row, i) => {
        const cells = [
          { text: i === 0 ? grp.system : '', bold: true, size: 8 },
          { text: row.zone || '', size: 8 },
          { text: row.level || '', size: 8 },
          { text: row.discipline || '', size: 8 },
          { text: row.work || '', size: 8 },
          { text: row.qty || '', size: 8 },
          { text: row.unit || '', size: 8 },
        ];
        const rowH = computeRowHeight(cells, AC, 8);
        newPageIfNeeded(rowH);
        // re-draw header if new page started mid-group (simple approach: skip re-header for brevity)
        let cx = MARGIN;
        cells.forEach((c, idx) => {
          const isSystemCol = idx === 0;
          const align = (idx === 2 || idx === 5 || idx === 6) ? 'center' : 'left';
          drawCell(cx, y, AC[idx], rowH, c.text, {
            bold: c.bold || false, size: c.size, align,
            fill: isSystemCol ? LIGHT_BLUE : WHITE,
            color: isSystemCol ? DARK_BLUE : TEXT_DARK
          });
          cx += AC[idx];
        });
        y -= rowH;
      });
    }
  }

  // ════════════════════════════════════════
  // LOGO
  // ════════════════════════════════════════
  if (logoBuffer) {
    try {
      const logoImg = await pdfDoc.embedPng(logoBuffer);
      const logoW = 200, logoH = 38;
      page.drawImage(logoImg, { x: (PAGE_W - logoW) / 2, y: y - logoH, width: logoW, height: logoH });
      y -= logoH + 10;
    } catch (e) {
      y -= 10;
    }
  }

  // ════════════════════════════════════════
  // GENERAL INFO TABLE
  // ════════════════════════════════════════
  {
    const colW = [110, 215, 95, CW - 110 - 215 - 95];
    const rows = [
      [
        { text: 'Contractor / SubCont.', bold: true, fill: LGRAY },
        { text: 'MBL / First Fix' },
        { text: 'Package', bold: true, fill: LGRAY },
        { text: pkg },
      ],
      [
        { text: 'Client', bold: true, fill: LGRAY },
        { text: 'Jeddah Central Development Co.' },
        { text: 'Subcontract Ref.', bold: true, fill: LGRAY },
        { text: 'P1020096-OCN-014' },
      ],
    ];
    rows.forEach(rowCells => {
      const rowH = computeRowHeight(rowCells.map(c => ({ text: c.text, size: 8, bold: c.bold })), colW, 8);
      newPageIfNeeded(rowH);
      let x = MARGIN;
      rowCells.forEach((c, i) => {
        drawCell(x, y, colW[i], rowH, c.text, { bold: c.bold || false, size: 8, fill: c.fill || WHITE });
        x += colW[i];
      });
      y -= rowH;
    });
    y -= 8;
  }

  // ════════════════════════════════════════
  // DATE / REPORT NO / MANPOWER + TITLE
  // ════════════════════════════════════════
  {
    const leftW = CW * 0.45;
    const titleW = CW - leftW;
    const labelW = 80;
    const valueW = leftW - labelW;

    const leftRows = [
      [{ text: 'Date', bold: true, fill: LGRAY, w: labelW }, { text: date, w: valueW }],
      [{ text: 'Report No', bold: true, fill: LGRAY, w: labelW }, { text: String(reportNo), w: valueW }],
      [{ text: 'Manpower', bold: true, fill: LGRAY, w: leftW, align: 'center', span: true }],
      [{ text: 'Supervision', bold: true, fill: LGRAY, w: labelW, align: 'center' }, { text: String(supervision), w: valueW, align: 'center', bold: true }],
      [{ text: 'Labors', bold: true, fill: LGRAY, w: labelW, align: 'center' }, { text: String(labors), w: valueW, align: 'center', bold: true }],
    ];

    const rowHeights = leftRows.map(r => computeRowHeight(r.map(c => ({ text: c.text, size: 8, bold: c.bold })), r.map(c => c.w), 8));
    const totalLeftH = rowHeights.reduce((a, b) => a + b, 0);

    newPageIfNeeded(totalLeftH);
    const startY = y;
    let cy = y;
    leftRows.forEach((rowCells, ri) => {
      let x = MARGIN;
      const rh = rowHeights[ri];
      rowCells.forEach(c => {
        drawCell(x, cy, c.w, rh, c.text, { bold: c.bold || false, size: 8, fill: c.fill || WHITE, align: c.align || 'left' });
        x += c.w;
      });
      cy -= rh;
    });

    // Title box on the right
    drawCell(MARGIN + leftW, startY, titleW, totalLeftH, 'Daily Site Work Report', {
      bold: true, size: 14, align: 'center', fill: WHITE, color: DARK_BLUE
    });

    y -= totalLeftH;
    y -= 10;
  }

  // ════════════════════════════════════════
  // DAILY SITE ACTIVITIES
  // ════════════════════════════════════════
  drawBanner('DAILY SITE ACTIVITIES');
  newPageIfNeeded(16);
  page.drawText('MECHANICAL PROGRESS', { x: MARGIN, y: y - 10, size: 9, font: fontBold, color: DARK_BLUE });
  y -= 18;
  drawActivitiesTable(activities);
  y -= 10;

  // ════════════════════════════════════════
  // NEXT DAY ACTIVITIES
  // ════════════════════════════════════════
  drawBanner('next day activities');
  if (nextDay && nextDay.length > 0) {
    drawActivitiesTable(nextDay);
  } else {
    newPageIfNeeded(20);
    page.drawText('To be updated', { x: MARGIN, y: y - 10, size: 9, font, color: rgb(0.6, 0.6, 0.6) });
    y -= 20;
  }
  y -= 10;

  // ════════════════════════════════════════
  // PHOTOS
  // ════════════════════════════════════════
  drawBanner('PHOTOS');
  if (photos && photos.length > 0) {
    const photoW = (CW - 10) / 2;
    const photoH = 140;
    for (let i = 0; i < photos.length; i += 2) {
      newPageIfNeeded(photoH + 10);
      for (let j = 0; j < 2; j++) {
        const idx = i + j;
        if (idx >= photos.length) break;
        const { buffer, ext } = photos[idx];
        try {
          let img;
          if (ext === 'png') img = await pdfDoc.embedPng(buffer);
          else img = await pdfDoc.embedJpg(buffer);

          // Fit image within photoW x photoH preserving aspect ratio
          const scale = Math.min(photoW / img.width, photoH / img.height);
          const w = img.width * scale;
          const h = img.height * scale;
          const x = MARGIN + j * (photoW + 10) + (photoW - w) / 2;
          page.drawImage(img, { x, y: y - photoH + (photoH - h) / 2, width: w, height: h });
        } catch (e) {
          // skip broken image
        }
      }
      y -= photoH + 10;
    }
  }

  return pdfDoc.save();
}

module.exports = { buildPdf };
