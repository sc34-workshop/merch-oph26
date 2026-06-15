// ═══════════════════════════════════════════════════════════
//  Merch Order Form · Google Apps Script Backend
//  OPH 2026 × MWIT 36 ปี
//
//  SETUP:
//  1. Set SHEET_ID and DRIVE_FOLDER_ID in CONFIG below
//  2. Deploy → New deployment → Web app (Execute as: Me, Anyone)
//  3. Paste the URL into index.html and admin.html as GAS_URL
//
//  ONE-TIME SETUP (run manually from GAS editor):
//  - setupCheckboxes()   → adds ✅/🚚 checkbox columns + onEdit auto-update
//  - formatOrdersSheet() → applies colour formatting to all data rows
//  - setupSummarySheet() → (re)builds the Summary tab with SUMPRODUCT formulas
// ═══════════════════════════════════════════════════════════

const CONFIG = {
  SHEET_ID:        'PASTE_YOUR_GOOGLE_SHEET_ID_HERE',
  SHEET_NAME:      'Orders',
  SUMMARY_SHEET:   'Summary',
  DRIVE_FOLDER_ID: 'PASTE_YOUR_DRIVE_FOLDER_ID_HERE'
};

const HEADERS = [
  'Timestamp', 'OrderRef', 'ชื่อ', 'เบอร์โทร', 'LINE_IG',
  'วิธีรับ', 'ที่อยู่',
  'พวง_มังกรMWIT', 'พวง_โลโก้36ปี', 'พวง_มาสคอต',
  'สติ๊ก_ScienceSeries', 'สติ๊ก_SchoolLifeSeries',
  'กระเป๋า',
  'เสื้อ_S_ขาว', 'เสื้อ_M_ขาว', 'เสื้อ_L_ขาว', 'เสื้อ_XL_ขาว',
  'เสื้อ_S_กรมท่า', 'เสื้อ_M_กรมท่า', 'เสื้อ_L_กรมท่า', 'เสื้อ_XL_กรมท่า',
  'SetA_qty', 'SetA_ลายพวง', 'SetA_ลายสติ๊ก',
  'SetB_qty', 'SetB_ลายพวง', 'SetB_ลายสติ๊ก', 'SetB_ไซส์', 'SetB_สี',
  'ราคาสินค้า', 'ค่าส่ง', 'รวมทั้งหมด',
  'URL_สลิป', 'เลขอ้างอิง', 'เวลาโอน', 'หมายเหตุ', 'สถานะ'
];

// Checkbox columns sit immediately after สถานะ (last HEADERS column)
const STATUS_COL  = HEADERS.length;      // col 37 = AK  (สถานะ text)
const CONFIRM_COL = HEADERS.length + 1;  // col 38 = AL  ✅ ยืนยันชำระเงิน
const DELIVER_COL = HEADERS.length + 2;  // col 39 = AM  🚚 จัดส่งแล้ว

// Column groups: [startCol, endCol, headerBg, label]
// Each group gets its own header colour; lane dividers go on data rows too.
const GROUPS = [
  [1,   5,  '#0d1b4b', 'ข้อมูลออเดอร์'],
  [6,   7,  '#065f46', 'การจัดส่ง'],
  [8,  10,  '#4c1d95', 'พวงกุญแจ 🔑'],
  [11, 12,  '#831843', 'สติ๊กเกอร์ 📋'],
  [13, 13,  '#064e3b', 'กระเป๋า 👜'],
  [14, 17,  '#92400e', 'เสื้อขาว 👕'],
  [18, 21,  '#1e3a5f', 'เสื้อกรมท่า 👕'],
  [22, 24,  '#3730a3', 'เซ็ต A 🎁'],
  [25, 29,  '#1d4ed8', 'เซ็ต B 🎁'],
  [30, 32,  '#14532d', '💰 ราคา'],
  [33, 36,  '#334155', 'การชำระเงิน'],
  [37, 37,  '#7f1d1d', 'สถานะ'],
];

// ── ONEDIT TRIGGER (simple trigger — runs automatically) ──
// When staff tick/untick a checkbox, สถานะ updates instantly.
function onEdit(e) {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== CONFIG.SHEET_NAME) return;
    const col = e.range.getColumn();
    const row = e.range.getRow();
    if (row <= 1) return;                         // skip header row
    if (col !== CONFIRM_COL && col !== DELIVER_COL) return;

    const confirmed = sheet.getRange(row, CONFIRM_COL).getValue() === true;
    const delivered = sheet.getRange(row, DELIVER_COL).getValue() === true;

    if (delivered) {
      sheet.getRange(row, STATUS_COL).setValue('จัดส่งแล้ว');
      if (!confirmed) sheet.getRange(row, CONFIRM_COL).setValue(true); // auto-tick confirm too
    } else if (confirmed) {
      sheet.getRange(row, STATUS_COL).setValue('ยืนยันแล้ว');
    } else {
      sheet.getRange(row, STATUS_COL).setValue('รอยืนยัน');
    }
  } catch (_) {}  // never block the user's edit
}

// ── SETUP CHECKBOXES (run once manually from GAS editor) ──
// Adds ✅/🚚 columns, checkbox validation, initial values from existing
// สถานะ data, and conditional formatting on the whole row.
function setupCheckboxes() {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) { ss.toast('ไม่พบ Orders sheet', '⚠ ข้อผิดพลาด', 5); return; }
  ensureHeaders(sheet);

  const lastRow = sheet.getLastRow();
  const maxData = 10000;  // validation covers rows 2 – 10000

  // ── Column headers ─────────────────────────────────────
  sheet.getRange(1, CONFIRM_COL).setValue('✅ ยืนยันชำระเงิน');
  sheet.getRange(1, DELIVER_COL).setValue('🚚 จัดส่งแล้ว');
  sheet.getRange(1, CONFIRM_COL, 1, 2)
    .setBackground('#0d1b4b').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  sheet.setColumnWidth(CONFIRM_COL, 140);
  sheet.setColumnWidth(DELIVER_COL, 115);

  // ── Checkbox data validation ────────────────────────────
  const cbValidation = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange(2, CONFIRM_COL, maxData - 1).setDataValidation(cbValidation).setHorizontalAlignment('center');
  sheet.getRange(2, DELIVER_COL, maxData - 1).setDataValidation(cbValidation).setHorizontalAlignment('center');

  // ── Seed checkbox state from existing สถานะ values ─────
  if (lastRow >= 2) {
    const statuses    = sheet.getRange(2, STATUS_COL, lastRow - 1).getValues();
    const confirmVals = statuses.map(r => [r[0] === 'ยืนยันแล้ว' || r[0] === 'จัดส่งแล้ว']);
    const deliverVals = statuses.map(r => [r[0] === 'จัดส่งแล้ว']);
    sheet.getRange(2, CONFIRM_COL, lastRow - 1).setValues(confirmVals);
    sheet.getRange(2, DELIVER_COL, lastRow - 1).setValues(deliverVals);
  }

  // ── Conditional formatting ──────────────────────────────
  // Highlight the entire row green when confirmed, blue when delivered
  const allCols   = sheet.getRange(2, 1, maxData - 1, DELIVER_COL);
  const cfConfirm = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${colLetter(CONFIRM_COL)}2=TRUE`)
    .setBackground('#ecfdf5').setFontColor('#065f46')
    .setRanges([allCols]).build();
  const cfDeliver = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${colLetter(DELIVER_COL)}2=TRUE`)
    .setBackground('#eff6ff').setFontColor('#1e40af')
    .setRanges([allCols]).build();

  // Keep existing rules, add ours at the top (highest priority)
  const existing = sheet.getConditionalFormatRules().filter(r => {
    // Drop old full-row rules from previous setupCheckboxes calls
    const src = r.getBooleanCondition();
    if (!src) return true;
    const f = src.getCriteriaValues()[0] || '';
    return !f.toString().includes(colLetter(CONFIRM_COL)) && !f.toString().includes(colLetter(DELIVER_COL));
  });
  sheet.setConditionalFormatRules([cfDeliver, cfConfirm, ...existing]);

  ss.toast('✅ ยืนยัน → ยืนยันแล้ว  |  🚚 จัดส่ง → จัดส่งแล้ว  |  แถวเปลี่ยนสีอัตโนมัติ', '✅ ตั้งค่า Checkbox เรียบร้อย!', 8);
}

// Convert column number to letter(s): 1→A, 26→Z, 27→AA, 38→AL
function colLetter(n) {
  let s = '';
  while (n > 0) { s = String.fromCharCode(64 + (n - 1) % 26 + 1) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// ── POST: รับ order จาก frontend ─────────────────────────
function doPost(e) {
  try {
    const data  = JSON.parse(e.postData.contents);
    let slipUrl = '';
    if (data.payment && data.payment.slipBase64 && data.payment.slipBase64.length > 100) {
      slipUrl = uploadSlip(data.payment.slipBase64, data.orderRef || 'unknown');
    }
    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
    sheet.appendRow(buildRow(data, slipUrl));
    return ok({ orderRef: data.orderRef });
  } catch (err) {
    return fail(err.message);
  }
}

// ── GET: ส่งข้อมูลทั้งหมดให้ admin page ─────────────────
function doGet(e) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

    if (!sheet || sheet.getLastRow() === 0) {
      return ok({ orders: [], summary: {} });
    }

    // Auto-fix: insert header row if data exists but headers are missing
    ensureHeaders(sheet);

    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return ok({ orders: [], summary: {} });

    const headers = rows[0];
    // Filter out empty rows — setupCheckboxes() adds validation to 10 000 rows,
    // so getDataRange() returns them all; skip any row with no OrderRef.
    const orders  = rows.slice(1)
      .filter(row => row[1] !== '' && row[1] !== null && row[1] !== undefined)
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = row[i]; });
        return obj;
      });

    return ok({ orders, summary: buildSummary(orders) });
  } catch (err) {
    return fail(err.message);
  }
}

// ── BUILD ROW ────────────────────────────────────────────
function buildRow(d, slipUrl) {
  const items = d.items    || {};
  const kc    = items.keychain || {};
  const st    = items.sticker  || {};
  const ts    = items.tshirt   || {};
  const sA    = items.setA     || {};
  const sB    = items.setB     || {};
  const pay   = d.payment  || {};
  const price = d.pricing  || {};
  const ship  = d.shipping || {};
  const cust  = d.customer || {};

  return [
    new Date(),
    d.orderRef         || '',
    cust.name          || '',
    cust.phone         || '',
    cust.lineig        || '',
    ship.method        || '',
    ship.address       || '',
    n(kc['มังกร MWIT']),
    n(kc['โลโก้ 36 ปี']),
    n(kc['มาสคอต']),
    n(st['Science Series']),
    n(st['School Life Series']),
    n(items.tote),
    n(ts['S_ขาว']),  n(ts['M_ขาว']),  n(ts['L_ขาว']),  n(ts['XL_ขาว']),
    n(ts['S_กรมท่า']), n(ts['M_กรมท่า']), n(ts['L_กรมท่า']), n(ts['XL_กรมท่า']),
    n(sA.qty), sA.keychain || '', sA.sticker || '',
    n(sB.qty), sB.keychain || '', sB.sticker || '', sB.size || '', sB.color || '',
    n(price.subtotal),
    n(price.shippingCost),
    n(price.total),
    slipUrl,
    pay.refNumber  || '',
    pay.datetime   || '',
    d.notes        || '',
    'รอยืนยัน'
  ];
}

// ── SHEET INIT & HEADER FIX ──────────────────────────────
function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  ensureHeaders(sheet);
  return sheet;
}

// Inserts header row if missing (detects by checking A1 = 'Timestamp')
function ensureHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    styleHeaderRow(sheet);
    return;
  }
  const a1 = sheet.getRange(1, 1).getValue();
  if (a1 !== 'Timestamp') {
    // Data row at row 1 with no header — push data down and insert headers
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    styleHeaderRow(sheet);
  }
}

function styleHeaderRow(sheet) {
  sheet.setFrozenRows(1);

  // Per-group header colours
  GROUPS.forEach(([start, end, bg]) => {
    sheet.getRange(1, start, 1, end - start + 1)
      .setBackground(bg)
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  });
  sheet.setRowHeight(1, 32);
  setColumnWidths(sheet);
}

function setColumnWidths(sheet) {
  const widths = [
    160, 190, 140, 110, 120, // A-E  order basics
     90, 260,                 // F-G  shipping
     65,  80,  70,            // H-J  keychains
     80,  95,                 // K-L  stickers
     65,                      // M    bag
     62,  62,  62,  65,       // N-Q  shirts white
     65,  65,  65,  70,       // R-U  shirts dark
     62, 120, 120,            // V-X  set A
     62, 120, 120,  62,  70,  // Y-AC set B
     90,  70, 115,            // AD-AF pricing
    200, 100, 130, 200,  95,  // AG-AK payment + status
  ];
  widths.forEach((w, i) => { try { sheet.setColumnWidth(i + 1, w); } catch(_) {} });
}

// Adds medium-weight lane dividers between column groups on data rows
function applyGroupBorders(sheet, fromRow, rows) {
  const divColor = '#8fa8c8';
  const style    = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  GROUPS.forEach(([start]) => {
    if (start <= 1) return;
    sheet.getRange(fromRow, start, rows, 1)
      .setBorder(null, true, null, null, null, null, divColor, style);
  });
}

// ── FORMAT ORDERS SHEET (run manually from GAS editor) ───
function formatOrdersSheet() {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) { ss.toast('ไม่พบ Orders sheet', '⚠ ข้อผิดพลาด', 5); return; }

  ensureHeaders(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ss.toast('ยังไม่มีข้อมูลใน Orders', '⚠ ยังว่างเปล่า', 5); return; }
  const dataRows  = lastRow - 1;
  const totalCols = DELIVER_COL; // include checkbox columns in formatting

  // ── Base font ───────────────────────────────────────────
  sheet.getRange(1, 1, lastRow, totalCols)
    .setFontFamily('Noto Sans Thai').setFontSize(10);

  // ── Header row ──────────────────────────────────────────
  styleHeaderRow(sheet);

  // ── Alternating data rows (navy blue family) ────────────
  // Bright row: very light icy blue
  // Dark row:   clear periwinkle — distinctly different, still easy on eyes
  const ROW_BRIGHT = '#edf3ff';
  const ROW_DARK   = '#cddcf8';
  for (let r = 2; r <= lastRow; r++) {
    sheet.getRange(r, 1, 1, totalCols)
      .setBackground(r % 2 === 0 ? ROW_BRIGHT : ROW_DARK)
      .setVerticalAlignment('middle');
    sheet.setRowHeight(r, 26);
  }

  // ── Column group lane dividers ──────────────────────────
  applyGroupBorders(sheet, 2, dataRows);

  // ── Name column — bold, easy to scan ───────────────────
  const nameCol = HEADERS.indexOf('ชื่อ') + 1;
  sheet.getRange(2, nameCol, dataRows).setFontWeight('bold');

  // ── Address — wrap & top-align ─────────────────────────
  const addrCol = HEADERS.indexOf('ที่อยู่') + 1;
  sheet.getRange(2, addrCol, dataRows).setWrap(true).setVerticalAlignment('top');

  // ── Pricing columns ─────────────────────────────────────
  ['ราคาสินค้า', 'ค่าส่ง'].forEach(h => {
    const ci = HEADERS.indexOf(h) + 1;
    sheet.getRange(2, ci, dataRows)
      .setNumberFormat('#,##0 "฿"')
      .setHorizontalAlignment('right');
  });
  const totalCol = HEADERS.indexOf('รวมทั้งหมด') + 1;
  sheet.getRange(2, totalCol, dataRows)
    .setNumberFormat('#,##0 "฿"')
    .setFontWeight('bold')
    .setFontColor('#db2777')
    .setBackground('#fff0f7') // permanent pink tint — stands out in both row colours
    .setHorizontalAlignment('right');

  // ── Date/time format ────────────────────────────────────
  sheet.getRange(2, 1, dataRows).setNumberFormat('d/M/yyyy  HH:mm');

  // ── Numeric item columns — center ──────────────────────
  // Cols 8-29 are all qty numbers — center them
  sheet.getRange(2, 8, dataRows, 22).setHorizontalAlignment('center');

  // ── Conditional formatting ──────────────────────────────
  sheet.clearConditionalFormatRules();

  const statusCol = HEADERS.indexOf('สถานะ') + 1;
  const shipCol   = HEADERS.indexOf('วิธีรับ') + 1;
  const statusR   = sheet.getRange(2, statusCol, dataRows);
  const shipR     = sheet.getRange(2, shipCol,   dataRows);

  const cfRules = [
    // Status cell colours (badge-style)
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('รอยืนยัน')
      .setBackground('#fef3c7').setFontColor('#92400e').setBold(true)
      .setRanges([statusR]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('ยืนยันแล้ว')
      .setBackground('#d1fae5').setFontColor('#065f46').setBold(true)
      .setRanges([statusR]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('จัดส่งแล้ว')
      .setBackground('#dbeafe').setFontColor('#1e40af').setBold(true)
      .setRanges([statusR]).build(),
    // Shipping cell colours
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('postal')
      .setBackground('#fce7f3').setFontColor('#9d174d').setBold(true)
      .setRanges([shipR]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('pickup')
      .setBackground('#e0f2fe').setFontColor('#0c4a6e').setBold(true)
      .setRanges([shipR]).build(),
  ];
  sheet.setConditionalFormatRules(cfRules);

  ss.toast('หากยังไม่มี checkbox ให้ run setupCheckboxes() ด้วย', '✅ จัดรูปแบบ Orders sheet เรียบร้อย!', 8);
}

// ── SUMMARY SHEET SETUP (run manually from GAS editor) ───
function setupSummarySheet() {
  const ss     = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sumSheet = ss.getSheetByName(CONFIG.SUMMARY_SHEET);
  if (sumSheet) ss.deleteSheet(sumSheet);
  sumSheet = ss.insertSheet(CONFIG.SUMMARY_SHEET, 0);

  const data = [
    ['📦 สรุปสินค้าที่ระลึก OPH 2026 × MWIT 36 ปี', '', '', ''],
    ['(อัปเดตอัตโนมัติเมื่อมีออเดอร์ใหม่)', '', '', ''],
    ['', '', '', ''],
    ['สินค้า / ตัวเลือก', 'รับเอง', 'ส่งไปรษณีย์', 'รวม'],
    ['พวงกุญแจ มังกร MWIT',     '=SUMPRODUCT((Orders!F2:F10000="pickup")*Orders!H2:H10000)', '=SUMPRODUCT((Orders!F2:F10000="postal")*Orders!H2:H10000)', '=B5+C5'],
    ['พวงกุญแจ โลโก้ 36 ปี',    '=SUMPRODUCT((Orders!F2:F10000="pickup")*Orders!I2:I10000)', '=SUMPRODUCT((Orders!F2:F10000="postal")*Orders!I2:I10000)', '=B6+C6'],
    ['พวงกุญแจ มาสคอต',         '=SUMPRODUCT((Orders!F2:F10000="pickup")*Orders!J2:J10000)', '=SUMPRODUCT((Orders!F2:F10000="postal")*Orders!J2:J10000)', '=B7+C7'],
    ['สติ๊กเกอร์ Science Series',     '=SUMPRODUCT((Orders!F2:F10000="pickup")*Orders!K2:K10000)', '=SUMPRODUCT((Orders!F2:F10000="postal")*Orders!K2:K10000)', '=B8+C8'],
    ['สติ๊กเกอร์ School Life Series', '=SUMPRODUCT((Orders!F2:F10000="pickup")*Orders!L2:L10000)', '=SUMPRODUCT((Orders!F2:F10000="postal")*Orders!L2:L10000)', '=B9+C9'],
    ['กระเป๋า Canvas Tote',     '=SUMPRODUCT((Orders!F2:F10000="pickup")*Orders!M2:M10000)', '=SUMPRODUCT((Orders!F2:F10000="postal")*Orders!M2:M10000)', '=B10+C10'],
    ['เสื้อ S ขาว',    '=SUMPRODUCT((Orders!F2:F10000="pickup")*Orders!N2:N10000)', '=SUMPRODUCT((Orders!F2:F10000="postal")*Orders!N2:N10000)', '=B11+C11'],
    ['เสื้อ M ขาว',    '=SUMPRODUCT((Orders!F2:F10000="pickup")*Orders!O2:O10000)', '=SUMPRODUCT((Orders!F2:F10000="postal")*Orders!O2:O10000)', '=B12+C12'],
    ['เสื้อ L ขาว',    '=SUMPRODUCT((Orders!F2:F10000="pickup")*Orders!P2:P10000)', '=SUMPRODUCT((Orders!F2:F10000="postal")*Orders!P2:P10000)', '=B13+C13'],
    ['เสื้อ XL ขาว',   '=SUMPRODUCT((Orders!F2:F10000="pickup")*Orders!Q2:Q10000)', '=SUMPRODUCT((Orders!F2:F10000="postal")*Orders!Q2:Q10000)', '=B14+C14'],
    ['เสื้อ S กรมท่า',  '=SUMPRODUCT((Orders!F2:F10000="pickup")*Orders!R2:R10000)', '=SUMPRODUCT((Orders!F2:F10000="postal")*Orders!R2:R10000)', '=B15+C15'],
    ['เสื้อ M กรมท่า',  '=SUMPRODUCT((Orders!F2:F10000="pickup")*Orders!S2:S10000)', '=SUMPRODUCT((Orders!F2:F10000="postal")*Orders!S2:S10000)', '=B16+C16'],
    ['เสื้อ L กรมท่า',  '=SUMPRODUCT((Orders!F2:F10000="pickup")*Orders!T2:T10000)', '=SUMPRODUCT((Orders!F2:F10000="postal")*Orders!T2:T10000)', '=B17+C17'],
    ['เสื้อ XL กรมท่า', '=SUMPRODUCT((Orders!F2:F10000="pickup")*Orders!U2:U10000)', '=SUMPRODUCT((Orders!F2:F10000="postal")*Orders!U2:U10000)', '=B18+C18'],
    ['เซ็ต A', '=SUMPRODUCT((Orders!F2:F10000="pickup")*Orders!V2:V10000)', '=SUMPRODUCT((Orders!F2:F10000="postal")*Orders!V2:V10000)', '=B19+C19'],
    ['เซ็ต B', '=SUMPRODUCT((Orders!F2:F10000="pickup")*Orders!Y2:Y10000)', '=SUMPRODUCT((Orders!F2:F10000="postal")*Orders!Y2:Y10000)', '=B20+C20'],
    ['', '', '', ''],
    ['💰 การเงิน', '', '', ''],
    ['รายได้สินค้ารวม', '', '', '=SUM(Orders!AD2:AD10000)'],
    ['รายได้ค่าส่งรวม', '', '', '=SUM(Orders!AE2:AE10000)'],
    ['รายได้ทั้งหมด',   '', '', '=SUM(Orders!AF2:AF10000)'],
    ['', '', '', ''],
    ['📊 ออเดอร์', '', '', ''],
    ['ออเดอร์ทั้งหมด',  '', '', '=COUNTA(Orders!B2:B10000)'],
    ['รับที่โรงเรียน',   '', '', '=COUNTIF(Orders!F2:F10000,"pickup")'],
    ['ส่งไปรษณีย์',      '', '', '=COUNTIF(Orders!F2:F10000,"postal")'],
    ['รอยืนยัน',         '', '', '=COUNTIF(Orders!AJ2:AJ10000,"รอยืนยัน")'],
    ['ยืนยันแล้ว',        '', '', '=COUNTIF(Orders!AJ2:AJ10000,"ยืนยันแล้ว")'],
    ['จัดส่งแล้ว',        '', '', '=COUNTIF(Orders!AJ2:AJ10000,"จัดส่งแล้ว")'],
  ];

  sumSheet.getRange(1, 1, data.length, 4).setValues(data);

  // Styling
  sumSheet.getRange(1, 1, 1, 4).setBackground('#0d1b4b').setFontColor('#ffffff').setFontWeight('bold').setFontSize(13);
  sumSheet.getRange(2, 1, 1, 4).setBackground('#1e2d6b').setFontColor('#94a3b8').setFontSize(10).setFontStyle('italic');
  sumSheet.getRange(4, 1, 1, 4).setBackground('#db2777').setFontColor('#ffffff').setFontWeight('bold');
  sumSheet.getRange(22, 1, 1, 4).setBackground('#1e3a8a').setFontColor('#ffffff').setFontWeight('bold');
  sumSheet.getRange(27, 1, 1, 4).setBackground('#1e3a8a').setFontColor('#ffffff').setFontWeight('bold');

  // Alternating rows for product list (rows 5-20)
  for (let r = 5; r <= 20; r++) {
    sumSheet.getRange(r, 1, 1, 4).setBackground(r % 2 === 0 ? '#f0f4ff' : '#ffffff');
    sumSheet.getRange(r, 4).setFontWeight('bold').setFontColor('#0d1b4b');
  }

  // Money rows
  ['D23','D24','D25'].forEach((cell, i) => {
    sumSheet.getRange(23 + i, 4).setNumberFormat('#,##0 "฿"').setFontWeight('bold').setFontColor('#db2777');
  });

  // Borders on data area
  sumSheet.getRange(4, 1, 17, 4).setBorder(true, true, true, true, true, false, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);
  sumSheet.getRange(22, 1, 10, 4).setBorder(true, true, true, true, true, false, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);

  sumSheet.setFrozenRows(4);
  sumSheet.setColumnWidth(1, 220);
  sumSheet.setColumnWidth(2, 110);
  sumSheet.setColumnWidth(3, 130);
  sumSheet.setColumnWidth(4, 110);

  ss.toast('เปิดแท็บ Summary ดูได้เลย', '✅ สร้าง Summary sheet เรียบร้อย!', 6);
}

// ── SUMMARY for admin ────────────────────────────────────
function buildSummary(orders) {
  const s = {
    totalOrders:  orders.length,
    totalRevenue: orders.reduce((t, o) => t + n(o['รวมทั้งหมด']), 0),
    shipping: {
      pickup: orders.filter(o => o['วิธีรับ'] === 'pickup').length,
      postal: orders.filter(o => o['วิธีรับ'] === 'postal').length
    },
    items: {}
  };
  const numCols = HEADERS.filter((_, i) => i >= 7 && i <= 28);
  numCols.forEach(col => {
    s.items[col] = orders.reduce((t, o) => t + n(o[col]), 0);
  });
  return s;
}

// ── DRIVE: อัปโหลดสลิป ───────────────────────────────────
function uploadSlip(base64Data, orderRef) {
  try {
    const raw    = base64Data.includes(',') ? base64Data.split(',')[1] : base64Data;
    const blob   = Utilities.newBlob(Utilities.base64Decode(raw), 'image/jpeg', `slip_${orderRef}.jpg`);
    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    const file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (_) {
    return '';
  }
}

// ── HELPERS ──────────────────────────────────────────────
function n(v)    { return parseInt(v) || 0; }
function ok(d)   { return resp({ success: true,  ...d }); }
function fail(m) { return resp({ success: false, error: m }); }
function resp(d) {
  return ContentService
    .createTextOutput(JSON.stringify(d))
    .setMimeType(ContentService.MimeType.JSON);
}
