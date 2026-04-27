import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { restoreChartsFromOriginal } from './xlsxChartPreserver.js';

// ===== Interfaces removed (TypeScript only) =====

function totalSale(s) {
  return s.originsSale + s.svSale + (s.aeSale || 0);
}

// ===== Date Parsing =====

const DATE_REGEXES = [
  {
    // yyyy-MM-dd HH:mm:ss +ZZZZ  or  yyyy-MM-ddTHH:mm:ssXXX  or  yyyy-MM-ddTHH:mm:ss  or  yyyy-MM-dd
    regex: /^(\d{4})-(\d{2})-(\d{2})/,
    parse: (m) => new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3])),
  },
  {
    // dd/mm/yyyy (with optional time)
    regex: /^(\d{1,2})\/(\d{1,2})\/(\d{4})/,
    parse: (m) => new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1])),
  },
];

function parseFulfilledDate(raw) {
  if (raw == null) return null;
  const trimmed = String(raw).trim();
  if (!trimmed) return null;

  if (/^\d+(\.\d+)?$/.test(trimmed)) {
    const serial = Number(trimmed);
    if (!Number.isNaN(serial) && serial > 0) {
      const utcDays = Math.floor(serial - 25569);
      const utcValue = utcDays * 86400;
      const dateInfo = new Date(utcValue * 1000);
      if (!isNaN(dateInfo.getTime())) {
        return new Date(dateInfo.getUTCFullYear(), dateInfo.getUTCMonth(), dateInfo.getUTCDate());
      }
    }
  }

  for (const { regex, parse } of DATE_REGEXES) {
    const m = trimmed.match(regex);
    if (m) {
      const d = parse(m);
      if (d && !isNaN(d.getTime())) return d;
    }
  }

  // Fallback for formats like "Feb 21, 2026, 10:30 AM" that appear in some exports.
  const native = new Date(trimmed);
  if (!isNaN(native.getTime())) {
    return new Date(native.getFullYear(), native.getMonth(), native.getDate());
  }

  return null;
}

function dateOnly(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
}

function parseLocalDate(s) {
  if (!s) return null;
  const d = new Date(s);
  if (isNaN(d.getTime())) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

// ===== String Helpers =====

function normalizeCode(s) {
  if (!s) return '';
  let t = s.trim();
  t = t.replace(/^=+/, '');
  t = t.replace(/^"|"$/g, '');
  return t.trim().toLowerCase();
}

// Like normalizeCode but also strips hyphens — for AE Trading memo matching (MER-110 → mer110)
function normalizeAECode(s) {
  if (!s) return '';
  let t = s.trim();
  t = t.replace(/^=+/, '');
  t = t.replace(/^"|"$/g, '');
  t = t.trim().toLowerCase();
  t = t.replace(/-/g, '');
  return t;
}

// Strip spaces, dots, hyphens, slashes and lowercase for fuzzy name comparison
function normalizeName(s) {
  return (s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function normalizeHeader(s) {
  if (!s) return '';
  return s
    .trim()
    .toLowerCase()
    .replace(/[\n\r\t ]+/g, '')
    .replace(/[_\-]/g, '');
}

function parseDouble(s) {
  if (!s) return 0;
  const cleaned = s.replace(/[^\d.\-]/g, '');
  if (!cleaned) return 0;
  const v = parseFloat(cleaned);
  return isNaN(v) ? 0 : v;
}

function isExactTarget(headerLower) {
  if (!headerLower || !headerLower.trim()) return false;
  const h = headerLower.toLowerCase();
  if (!h.includes('target')) return false;
  if (
    h.includes('march') ||
    h.includes('per day') ||
    h.includes('forecast') ||
    h.includes('achievement') ||
    h.includes('day')
  )
    return false;
  return true;
}

/** Convert 0-based column index to Excel column letter (0=A, 1=B, 25=Z, 26=AA) */
function colLetter(col) {
  let sb = '';
  let c = col + 1; // 1-based
  while (c > 0) {
    c--;
    sb = String.fromCharCode(65 + (c % 26)) + sb;
    c = Math.floor(c / 26);
  }
  return sb;
}

/**
 * Read an XLSX sheet into a grid keyed by actual Excel row/col (0-based).
 * Returns { grid, startRow, startCol, endRow, endCol }.
 * grid[row][col] = cell string value (using actual Excel positions).
 */
function readSheetGrid(sheet) {
  if (!sheet || !sheet['!ref']) return { grid: {}, startRow: 0, startCol: 0, endRow: 0, endCol: 0 };
  const range = XLSX.utils.decode_range(sheet['!ref']);
  const grid = {};
  for (let R = range.s.r; R <= range.e.r; R++) {
    grid[R] = {};
    for (let C = range.s.c; C <= range.e.c; C++) {
      const addr = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = sheet[addr];
      grid[R][C] = cell ? String(cell.v ?? '') : '';
    }
  }
  return { grid, startRow: range.s.r, startCol: range.s.c, endRow: range.e.r, endCol: range.e.c };
}

// ===== OrderRow helper =====

function getVal(row, col) {
  if (!col) return '';
  // 1. Exact case-insensitive match
  for (const key of Object.keys(row.data)) {
    if (key.toLowerCase() === col.toLowerCase()) return row.data[key];
  }
  const normCol = normalizeHeader(col);
  // 2. normKey === normCol (after header normalization)
  for (const key of Object.keys(row.data)) {
    if (normalizeHeader(key) === normCol) return row.data[key];
  }
  // 3. normKey starts with normCol (e.g. "total" matches "total sale" but NOT "subtotal")
  for (const key of Object.keys(row.data)) {
    const normKey = normalizeHeader(key);
    if (normKey.startsWith(normCol) || normCol.startsWith(normKey)) return row.data[key];
  }
  // 4. Last resort: normKey contains normCol (original fuzzy fallback)
  for (const key of Object.keys(row.data)) {
    const normKey = normalizeHeader(key);
    if (normKey.includes(normCol)) return row.data[key];
  }
  return '';
}

function createOrderRow(company, data) {
  const row = { company, data, total: 0, financialStatus: '', discountCode: '', createdAt: '' };
  row.financialStatus = getVal(row, 'Financial Status');
  row.discountCode = normalizeCode(getVal(row, 'Discount Code'));
  row.total = parseDouble(getVal(row, 'Total'));
  row.createdAt = getVal(row, 'Created at');
  return row;
}

// ===== File Readers (using xlsx library) =====

function cellToString(v) {
  if (v == null) return '';
  return String(v).trim();
}

function readExcelOrCsv(file) {
  const wb = XLSX.read(file.buffer, { type: 'buffer' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  if (!sheet) return { rows: [], headers: [] };
  const ref = sheet['!ref'];
  if (!ref) return { rows: [], headers: [] };
  const range = XLSX.utils.decode_range(ref);

  // Read a cell as its displayed text (cell.w) to preserve dates exactly as shown in the file
  function getCellText(r, c) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = sheet[addr];
    if (!cell) return '';
    if (cell.w != null) return String(cell.w).trim();
    if (cell.v == null) return '';
    return String(cell.v).trim();
  }

  const headers = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    headers.push(getCellText(range.s.r, c));
  }

  const rows = [];
  for (let r = range.s.r + 1; r <= range.e.r; r++) {
    const map = {};
    let hasData = false;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const val = getCellText(r, c);
      if (val) hasData = true;
      map[headers[c - range.s.c]] = val;
    }
    if (hasData) rows.push(map);
  }
  return { rows, headers };
}

function readCouponFile(file) {
  const coupons = [];
  const wb = XLSX.read(file.buffer, { type: 'buffer' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  if (!sheet) return coupons;

  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
  if (raw.length === 0) return coupons;

  // Find header row (scan rows 0-5 for "coupon", "code", "owner", "mer")
  let headerRowIdx = -1;
  for (let r = 0; r < Math.min(raw.length, 6); r++) {
    for (const cell of raw[r]) {
      const val = String(cell ?? '').trim().toLowerCase();
      if (val.includes('coupon') || val.includes('owner') || val.includes('mer')) {
        headerRowIdx = r;
        break;
      }
    }
    if (headerRowIdx >= 0) break;
  }

  if (headerRowIdx < 0) return coupons;

  const headerArr = raw[headerRowIdx];

  // Detect columns
  let colMerchant = -1;
  let colCode = -1;
  let colType = -1;

  for (let c = 0; c < headerArr.length; c++) {
    const h = String(headerArr[c] ?? '').trim();
    const hLower = h.toLowerCase();

    if (colMerchant < 0 && hLower.includes('owner')) {
      colMerchant = c;
    } else if (colCode < 0 && hLower.includes('code') && !hLower.includes('owner')) {
      colCode = c;
    } else if (colType < 0 && hLower.includes('type')) {
      colType = c;
    }
  }

  // Fallback for code column
  if (colCode < 0) {
    for (let c = 0; c < headerArr.length; c++) {
      if (c === colMerchant || c === colType) continue;
      const hLower = String(headerArr[c] ?? '').trim().toLowerCase();
      if (hLower.includes('coupon') || hLower.includes('discount') || hLower.includes('code')) {
        colCode = c;
        break;
      }
    }
  }

  const dataStartRow = headerRowIdx + 1;
  for (let r = dataStartRow; r < raw.length; r++) {
    const rowArr = raw[r];
    const merchantName = colMerchant >= 0 ? String(rowArr[colMerchant] ?? '').trim() : '';
    const code = colCode >= 0 ? String(rowArr[colCode] ?? '').trim() : '';
    const type = colType >= 0 ? String(rowArr[colType] ?? '').trim() : '';

    if (!merchantName && !code) continue;
    coupons.push({ merchantName, discountCode: code, type });
  }

  return coupons;
}

/**
 * Read an AE Trading export file.
 * Columns: Type (filter: Invoice only), Memo (coupon/merchant code), Amount (sale total).
 */
function readAETradingFile(file) {
  const entries = [];
  const wb = XLSX.read(file.buffer, { type: 'buffer' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  if (!sheet) return entries;

  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
  if (raw.length === 0) return entries;

  // Find header row — look for a row containing "Memo" and ("Amount" or "Type")
  let headerRowIdx = -1;
  for (let r = 0; r < Math.min(raw.length, 15); r++) {
    const rowVals = raw[r].map((v) => String(v ?? '').trim().toLowerCase());
    if (rowVals.some((v) => v === 'memo') && rowVals.some((v) => v.includes('amount') || v === 'type')) {
      headerRowIdx = r;
      break;
    }
  }

  if (headerRowIdx < 0) return entries;

  const headers = raw[headerRowIdx];
  let colMemo = -1;
  let colAmount = -1;
  let colType = -1;
  let colNo = -1;

  for (let c = 0; c < headers.length; c++) {
    const h = String(headers[c] ?? '').trim().toLowerCase();
    if ((h === 'memo' || h.includes('memo')) && colMemo < 0) colMemo = c;
    else if ((h === 'amount' || h.includes('amount')) && colAmount < 0) colAmount = c;
    else if ((h === 'type' || h.includes('type')) && colType < 0) colType = c;
    else if ((h === 'no.' || h === 'no' || h === 'number') && colNo < 0) colNo = c;
  }

  if (colMemo < 0 || colAmount < 0) return entries;

  for (let r = headerRowIdx + 1; r < raw.length; r++) {
    const rowArr = raw[r];
    // Only include Invoice rows (case-insensitive)
    if (colType >= 0) {
      const type = String(rowArr[colType] ?? '').trim().toLowerCase();
      if (type && type !== 'invoice') continue;
    }
    const memo = String(rowArr[colMemo] ?? '').trim();
    const amount = parseDouble(String(rowArr[colAmount] ?? ''));
    const no = colNo >= 0 ? String(rowArr[colNo] ?? '').trim() : '';
    if (!memo && amount === 0) continue;
    entries.push({ memo, amount, no });
  }

  return entries;
}

/** Find the header row (actual Excel 0-based row number) that contains "merchant name" */
function findHeaderRow(grid, startRow, endRow, startCol, endCol) {
  for (let r = startRow; r <= Math.min(endRow, startRow + 10); r++) {
    const rowData = grid[r];
    if (!rowData) continue;
    for (let c = startCol; c <= endCol; c++) {
      const val = String(rowData[c] ?? '').toLowerCase();
      if (val.includes('merchant name') || val.includes('merchant')) return r;
    }
  }
  return -1;
}

function readTargetTable(file) {
  const result = [];
  const wb = XLSX.read(file.buffer, { type: 'buffer' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  if (!sheet) return result;

  const { grid, startRow, startCol, endRow, endCol } = readSheetGrid(sheet);

  const headerRow = findHeaderRow(grid, startRow, endRow, startCol, endCol);
  if (headerRow < 0) return result;

  const hRow = grid[headerRow] || {};

  // Check for sub-header row
  let dataStartRow = headerRow + 1;
  if (dataStartRow <= endRow) {
    const nextRow = grid[dataStartRow] || {};
    let isSubHeader = false;
    for (let c = startCol; c <= endCol; c++) {
      const v = String(nextRow[c] ?? '').toLowerCase();
      if (v.includes('origin') || v.includes('sv.lk') || v.includes('total sale')) {
        isSubHeader = true;
        break;
      }
    }
    if (isSubHeader) dataStartRow++;
  }

  const tRowAbove = headerRow > startRow ? grid[headerRow - 1] || {} : {};
  const tRowBelow = headerRow + 1 <= endRow ? grid[headerRow + 1] || {} : {};

  let colMerchant = -1;
  let colTarget = -1;
  let colOutlet = -1;

  for (let c = startCol; c <= endCol; c++) {
    const above = String(tRowAbove[c] ?? '').trim();
    const main = String(hRow[c] ?? '').trim();
    const below = String(tRowBelow[c] ?? '').trim();
    const val = main.toLowerCase().replace(/[^a-z0-9% /().,]/g, '').trim();
    const allRows = (above + ' ' + main + ' ' + below)
      .toLowerCase()
      .replace(/[^a-z0-9% /().,]/g, '')
      .trim();

    if (colMerchant < 0 && (val.includes('merchant name') || allRows.includes('merchant name')))
      colMerchant = c;
    else if (colTarget < 0 && isExactTarget(val)) colTarget = c;
    else if (colOutlet < 0 && (val === 'outlet' || allRows.includes('outlet'))) colOutlet = c;
  }

  // Fallback for Target column from above/below
  if (colTarget < 0) {
    for (let c = startCol; c <= endCol; c++) {
      const above = String(tRowAbove[c] ?? '')
        .toLowerCase()
        .replace(/[^a-z0-9 ]/g, '')
        .trim();
      const below = String(tRowBelow[c] ?? '')
        .toLowerCase()
        .replace(/[^a-z0-9 ]/g, '')
        .trim();
      if (isExactTarget(above) || isExactTarget(below)) {
        colTarget = c;
        break;
      }
    }
  }

  if (colMerchant < 0) return result;

  let startedRows = false;
  for (let r = dataStartRow; r <= endRow; r++) {
    const rowData = grid[r] || {};
    const merchant = String(rowData[colMerchant] ?? '').trim();
    if (!merchant) {
      if (startedRows) break;
      continue;
    }
    startedRows = true;

    const merchantLower = merchant.toLowerCase();
    if (merchantLower === 'total' || merchantLower === 'grand total') break;
    if (merchantLower.includes('merchant total')) continue;

    const tr = {
      rowIndex: r, // actual Excel 0-based row number
      merchantName: merchant,
      target: colTarget >= 0 ? parseDouble(String(rowData[colTarget] ?? '')) : 0,
      outlet: colOutlet >= 0 ? String(rowData[colOutlet] ?? '').trim() : '',
    };
    result.push(tr);
  }

  return result;
}

// ===== Merchant Sales Matching =====

function findSalesForMerchant(
  merchantKey,
  salesMap
) {
  // Exact match
  const exact = salesMap.get(merchantKey);
  if (exact) return exact;
  // Case-insensitive match
  for (const [key, val] of salesMap) {
    if (key.toLowerCase() === merchantKey.toLowerCase()) return val;
  }
  // Normalized partial match (strips spaces, dots, etc.)
  const keyNorm = normalizeName(merchantKey);
  for (const [key, val] of salesMap) {
    const kNorm = normalizeName(key);
    if (kNorm.includes(keyNorm) || keyNorm.includes(kNorm)) return val;
  }
  return null;
}

function getMerchantType(merchantName, typeMap) {
  if (merchantName.toLowerCase() === 'dm general/sandali') return 'Outlet';
  const type = typeMap.get(merchantName.toLowerCase());
  if (type) return type;
  // Partial match
  for (const [key, val] of typeMap) {
    if (key.includes(merchantName.toLowerCase()) || merchantName.toLowerCase().includes(key)) {
      return val;
    }
  }
  return 'Online'; // default
}

// ===== Sheet Writers (using ExcelJS) =====

function writeAllSheet(
  wb,
  allOrders,
  headers
) {
  const ws = wb.addWorksheet('ALL');
  if (!headers || allOrders.length === 0) return;

  const allHeaders = ['Company', ...headers];

  // Header row
  const headerRow = ws.addRow(allHeaders);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true };
  });

  // Data rows
  for (const or of allOrders) {
    const vals = [or.company];
    for (const h of headers) {
      const v = or.data[h] ?? '';
      const num = parseFloat(v);
      vals.push(!isNaN(num) && v.trim() !== '' ? num : v);
    }
    ws.addRow(vals);
  }
}

function writeSalesSheet(wb, salesMap) {
  const ws = wb.addWorksheet('Sales');

  const cols = ['Merchant Name', 'Origins Sale', 'SupplementVault Sale', 'AE Trading Sale', 'Total Sale'];
  const headerRow = ws.addRow(cols);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true };
  });

  for (const [name, sales] of salesMap) {
    const row = ws.addRow([name, sales.originsSale, sales.svSale, sales.aeSale || 0, totalSale(sales)]);
    for (let c = 2; c <= 5; c++) {
      row.getCell(c).numFmt = '#,##0.00';
    }
  }

  // Total row
  const dataStart = 2; // first data row (row 1 is header)
  const dataEnd = dataStart + salesMap.size - 1;
  const totalRow = ws.addRow([
    'Total',
    { formula: `SUM(B${dataStart}:B${dataEnd})` },
    { formula: `SUM(C${dataStart}:C${dataEnd})` },
    { formula: `SUM(D${dataStart}:D${dataEnd})` },
    { formula: `SUM(E${dataStart}:E${dataEnd})` },
  ]);
  totalRow.getCell(1).font = { bold: true };
  for (let c = 2; c <= 5; c++) {
    totalRow.getCell(c).font = { bold: true };
    totalRow.getCell(c).numFmt = '#,##0.00';
  }
}

function writeExtraMerchantsSheet(
  wb,
  extraMerchants
) {
  const ws = wb.addWorksheet('Other Merchants');

  const cols = ['Merchant Name', 'Origins Sale', 'SupplementVault Sale', 'AE Trading Sale', 'Total Sale'];
  const headerRow = ws.addRow(cols);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true };
  });

  for (const [name, sales] of extraMerchants) {
    const row = ws.addRow([name, sales.originsSale, sales.svSale, sales.aeSale || 0, totalSale(sales)]);
    for (let c = 2; c <= 5; c++) {
      row.getCell(c).numFmt = '#,##0.00';
    }
  }
}

function writeMissingCouponSheet(wb, missingEntries) {
  const ws = wb.addWorksheet('Missing Coupon (AE)');
  const headerRow = ws.addRow(['No.', 'Memo', 'Amount']);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true };
  });
  for (const entry of missingEntries) {
    const row = ws.addRow([entry.no || '', entry.memo, entry.amount]);
    row.getCell(3).numFmt = '#,##0.00';
  }
}

// ===== Report Sheet: Modify target workbook sheet in-place =====

async function fillReportSheet(
  reportWb,
  targetBuffer,
  targetRows,
  merchantSalesMap,
  merchantTypeMap,
  daysRemainingOnline,
  daysRemainingOutlet,
  totalDays,
  reportDay
) {
  const reportSheet = reportWb.worksheets[0];
  if (!reportSheet) return;

  // Rename to "Report"
  reportSheet.name = 'Report';

  // Read actual cell positions using xlsx library for column detection
  const rawWb = XLSX.read(targetBuffer, { type: 'buffer' });
  const rawSheet = rawWb.Sheets[rawWb.SheetNames[0]];
  if (!rawSheet) return;
  const { grid, startRow, startCol, endRow, endCol } = readSheetGrid(rawSheet);

  const headerRowIdx = findHeaderRow(grid, startRow, endRow, startCol, endCol);
  if (headerRowIdx < 0) return;

  // Detect columns from grid (using actual Excel column numbers)
  const hRow = grid[headerRowIdx] || {};
  const rowAbove = headerRowIdx > startRow ? grid[headerRowIdx - 1] || {} : {};
  const rowBelow = headerRowIdx + 1 <= endRow ? grid[headerRowIdx + 1] || {} : {};

  let colMerchant = -1;
  let colTarget = -1;
  let colOriginSale = -1;
  let colSvSale = -1;
  let colAeSale = -1;
  let colTotalSale = -1;
  let colBalance = -1;
  let colPerDayTarget = -1;
  let colForecast = -1;
  const achievementPctCols = [];

  for (let c = startCol; c <= endCol; c++) {
    let above = String(rowAbove[c] ?? '')
      .toLowerCase()
      .trim()
      .replace(/[^a-z0-9% /().,]/g, '')
      .trim();
    let main = String(hRow[c] ?? '')
      .toLowerCase()
      .trim()
      .replace(/[^a-z0-9% /().,]/g, '')
      .trim();
    let below = String(rowBelow[c] ?? '')
      .toLowerCase()
      .trim()
      .replace(/[^a-z0-9% /().,]/g, '')
      .trim();
    const all = above + ' ' + main + ' ' + below;

    if (colMerchant < 0 && all.includes('merchant')) colMerchant = c;
    if (colTarget < 0 && (isExactTarget(main) || isExactTarget(above) || isExactTarget(below)))
      colTarget = c;
    if (
      colOriginSale < 0 &&
      (below.includes('origin') || (main.includes('origin') && !main.includes('achievement')))
    )
      colOriginSale = c;
    if (
      colSvSale < 0 &&
      (below.includes('sv.lk') ||
        below.includes('sv ') ||
        main.includes('sv.lk') ||
        main.includes('sv '))
    )
      colSvSale = c;
    if (
      colAeSale < 0 &&
      (below.includes('ae trading') || main.includes('ae trading') ||
        below === 'ae' || main === 'ae')
    )
      colAeSale = c;
    if (
      colTotalSale < 0 &&
      ((below.includes('total') && below.includes('sale')) ||
        (main.includes('total') && main.includes('sale')))
    )
      colTotalSale = c;
    if (colBalance < 0 && all.includes('balance')) colBalance = c;
    if (colPerDayTarget < 0 && all.includes('per day')) colPerDayTarget = c;
    if (
      colForecast < 0 &&
      all.includes('forecast') &&
      (all.includes('month') ||
        main === 'forecast' ||
        below === 'forecast' ||
        above === 'forecast') &&
      !all.includes('achievement %')
    )
      colForecast = c;

    if (all.includes('achievement') && all.includes('%')) {
      achievementPctCols.push(c);
    }
  }

  achievementPctCols.sort((a, b) => a - b);

  const colAchievement = achievementPctCols.length >= 1 ? achievementPctCols[0] : -1;
  const colForecastPct =
    achievementPctCols.length >= 2 ? achievementPctCols[achievementPctCols.length - 1] : -1;

  // Fallback for forecast value column
  if (colForecast < 0) {
    for (let c = startCol; c <= endCol; c++) {
      for (let rr = Math.max(startRow, headerRowIdx - 1); rr <= headerRowIdx + 1 && rr <= endRow; rr++) {
        const scanRow = grid[rr] || {};
        const v = String(scanRow[c] ?? '')
          .toLowerCase()
          .trim();
        if (v.includes('forecast') && (v.includes('month') || v === 'forecast')) {
          colForecast = c;
          break;
        }
      }
      if (colForecast >= 0) break;
    }
  }

  // Fill in data for each target row
  // Note: tr.rowIndex is the actual Excel 0-based row number; ExcelJS uses 1-based
  for (const tr of targetRows) {
    const excelRowNum = tr.rowIndex + 1; // ExcelJS 1-based row
    const row = reportSheet.getRow(excelRowNum);

    const merchantKey = tr.merchantName;
    const sales = findSalesForMerchant(merchantKey, merchantSalesMap);

    const originSale = sales ? sales.originsSale : 0;
    const svSale = sales ? sales.svSale : 0;

    const type = getMerchantType(merchantKey, merchantTypeMap);
    const daysRemaining = type.toLowerCase().includes('outlet')
      ? daysRemainingOutlet
      : daysRemainingOnline;

    // Helper: set cell value + numFmt while keeping existing style intact
    function writeCell(cell, value, numFmt) {
      cell.value = value;
      if (numFmt) cell.numFmt = numFmt;
    }

    // ExcelJS getCell uses 1-based column; colLetter() takes 0-based

    // Number format that shows "-" for zero values
    const NUM_FMT = '#,##0.00;-#,##0.00;"-"';
    const originColRef = colOriginSale >= 0 ? colLetter(colOriginSale) : '';
    const svColRef = colSvSale >= 0 ? colLetter(colSvSale) : '';
    const aeColRef = colAeSale >= 0 ? colLetter(colAeSale) : '';
    const totalColRef = colTotalSale >= 0 ? colLetter(colTotalSale) : '';
    const targetColRef = colTarget >= 0 ? colLetter(colTarget) : '';
    const balanceColRef = colBalance >= 0 ? colLetter(colBalance) : '';
    const forecastColRef = colForecast >= 0 ? colLetter(colForecast) : '';

    // Write sale values — only these two; all other columns keep their original formulas
    if (colOriginSale >= 0) {
      writeCell(row.getCell(colOriginSale + 1), originSale, NUM_FMT);
    }
    if (colSvSale >= 0) {
      writeCell(row.getCell(colSvSale + 1), svSale, NUM_FMT);
    }
    if (colAeSale >= 0) {
      writeCell(row.getCell(colAeSale + 1), sales ? (sales.aeSale || 0) : 0, NUM_FMT);
    }
    if (colTotalSale >= 0 && originColRef && svColRef) {
      const aeFormulaPart = aeColRef ? `+${aeColRef}${excelRowNum}` : '';
      writeCell(
        row.getCell(colTotalSale + 1),
        { formula: `(${originColRef}${excelRowNum}+${svColRef}${excelRowNum}${aeFormulaPart})` },
        NUM_FMT
      );
    }
    if (colAchievement >= 0 && totalColRef && targetColRef) {
      writeCell(
        row.getCell(colAchievement + 1),
        { formula: `IF(${targetColRef}${excelRowNum}=0,0,${totalColRef}${excelRowNum}/${targetColRef}${excelRowNum})` },
        '0%'
      );
    }
    if (colBalance >= 0 && targetColRef && totalColRef) {
      writeCell(
        row.getCell(colBalance + 1),
        { formula: `MAX(${targetColRef}${excelRowNum}-${totalColRef}${excelRowNum},0)` },
        NUM_FMT
      );
    }
    if (colPerDayTarget >= 0 && balanceColRef) {
      writeCell(
        row.getCell(colPerDayTarget + 1),
        { formula: `IF(${balanceColRef}${excelRowNum}=0,0,${balanceColRef}${excelRowNum}/${Math.max(daysRemaining, 1)})` },
        NUM_FMT
      );
    }
    if (colForecast >= 0 && totalColRef) {
      writeCell(
        row.getCell(colForecast + 1),
        { formula: `(${totalColRef}${excelRowNum}/${Math.max(reportDay, 1)})*${Math.max(totalDays, 1)}` },
        NUM_FMT
      );
    }
    if (colForecastPct >= 0 && forecastColRef && targetColRef) {
      writeCell(
        row.getCell(colForecastPct + 1),
        { formula: `IF(${targetColRef}${excelRowNum}=0,0,${forecastColRef}${excelRowNum}/${targetColRef}${excelRowNum})` },
        '0%'
      );
    }
    // Leave Total Sale, Achievement %, Balance, Per Day Target, Forecast, Forecast Achievement %
    // untouched — the original formulas in the target file will recalculate via fullCalcOnLoad
  }
}

function detectSheetColumns(grid, startRow, endRow, startCol, endCol) {
  let headerRowIdx = -1;
  for (let r = startRow; r <= Math.min(endRow, startRow + 10); r++) {
    const rowData = grid[r];
    if (!rowData) continue;
    for (let c = startCol; c <= endCol; c++) {
      const val = String(rowData[c] ?? '').toLowerCase();
      if (val.includes('merchant')) {
        headerRowIdx = r;
        break;
      }
    }
    if (headerRowIdx >= 0) break;
  }
  if (headerRowIdx < 0) return null;

  const hRow = grid[headerRowIdx] || {};
  const rowAbove = headerRowIdx > startRow ? grid[headerRowIdx - 1] || {} : {};
  const rowBelow = headerRowIdx + 1 <= endRow ? grid[headerRowIdx + 1] || {} : {};

  let colMerchant = -1;
  let colTarget = -1;
  let colInvoiceSale = -1;
  let colCompleteSale = -1;
  const achievementPctCols = [];
  let colInvoicePct = -1;
  let colCompletePct = -1;

  for (let c = startCol; c <= endCol; c++) {
    const above = String(rowAbove[c] ?? '').toLowerCase().trim();
    const main = String(hRow[c] ?? '').toLowerCase().trim();
    const below = String(rowBelow[c] ?? '').toLowerCase().trim();
    const combined = above + ' ' + main + ' ' + below;

    if (colMerchant < 0 && combined.includes('merchant')) colMerchant = c;
    if (colTarget < 0 && isExactTarget(main)) colTarget = c;
    if (colTarget < 0 && isExactTarget(above)) colTarget = c;
    if (colTarget < 0 && isExactTarget(below)) colTarget = c;
    if (colInvoiceSale < 0 && combined.includes('total sale') && combined.includes('invoice')) colInvoiceSale = c;
    if (colCompleteSale < 0 && combined.includes('total sale') && combined.includes('complete')) colCompleteSale = c;
    if (combined.includes('achievement') && combined.includes('%')) achievementPctCols.push(c);
  }

  if (colInvoiceSale < 0 || colCompleteSale < 0) {
    for (let c = startCol; c <= endCol; c++) {
      const main = String(hRow[c] ?? '').toLowerCase().trim();
      const above = String(rowAbove[c] ?? '').toLowerCase().trim();
      if ((main.includes('invoice date') || above.includes('invoice date')) &&
          (main.includes('total sale') || above.includes('total sale')) &&
          colInvoiceSale < 0) {
        colInvoiceSale = c;
      }
      if ((main.includes('complete date') || above.includes('complete date')) &&
          (main.includes('total sale') || above.includes('total sale')) &&
          colCompleteSale < 0) {
        colCompleteSale = c;
      }
    }
  }

  achievementPctCols.sort((a, b) => a - b);
  if (achievementPctCols.length >= 1) colInvoicePct = achievementPctCols[0];
  if (achievementPctCols.length >= 2) colCompletePct = achievementPctCols[1];

  let dataStartRow = headerRowIdx + 1;
  while (dataStartRow <= endRow) {
    const dr = grid[dataStartRow] || {};
    let isSub = false;
    for (let c = startCol; c <= endCol; c++) {
      const v = String(dr[c] ?? '').toLowerCase();
      if (v.includes('origin') || v.includes('sv.lk') || v.includes('total sale')) {
        isSub = true;
        break;
      }
    }
    if (!isSub) break;
    dataStartRow++;
  }

  const rows = [];
  for (let r = dataStartRow; r <= endRow; r++) {
    const rowData = grid[r] || {};
    const merchant = colMerchant >= 0 ? String(rowData[colMerchant] ?? '').trim() : '';
    if (!merchant) continue;
    const mLower = merchant.toLowerCase();
    if (mLower === 'total') break;
    if (mLower.includes('merchant total')) continue;
    rows.push({
      rowIndex: r,
      merchantName: merchant,
      target: colTarget >= 0 ? parseDouble(String(rowData[colTarget] ?? '')) : 0,
    });
  }

  return {
    headerRowIdx,
    colMerchant,
    colTarget,
    colInvoiceSale,
    colCompleteSale,
    colInvoicePct,
    colCompletePct,
    dataStartRow,
    rows,
  };
}

function detectSideTables(grid, startRow, endRow, startCol, endCol, mainTableEndCol) {
  const tables = [];
  const scanStart = mainTableEndCol + 1;
  const usedHeaders = new Set();

  for (let c = scanStart; c <= endCol; c++) {
    for (let r = startRow; r <= endRow; r++) {
      const val = String(grid[r]?.[c] ?? '').trim().toLowerCase();
      if (val.includes('row label') && !usedHeaders.has(`${r},${c}`)) {
        usedHeaders.add(`${r},${c}`);
        const table = { headerRow: r, colRowLabel: c, dataCols: [], merchantRows: [] };

        for (let cc = c + 1; cc <= endCol; cc++) {
          const hVal = String(grid[r]?.[cc] ?? '').trim();
          if (!hVal) break;
          table.dataCols.push({ col: cc, headerName: hVal });
        }

        for (let rr = r + 1; rr <= endRow; rr++) {
          const mName = String(grid[rr]?.[c] ?? '').trim();
          if (!mName) break;
          table.merchantRows.push({ row: rr, merchantName: mName });
        }

        if (table.dataCols.length > 0 && table.merchantRows.length > 0) {
          tables.push(table);
        }
      }
    }
  }

  return tables;
}

function fillSheet2or3(reportSheet, rawSheet, invoiceSalesMap, completeSalesMap = new Map()) {
  if (!reportSheet) return { invoiceTotals: new Map() };
  const { grid, startRow, startCol, endRow, endCol } = readSheetGrid(rawSheet);
  const cols = detectSheetColumns(grid, startRow, endRow, startCol, endCol);
  if (!cols) return { invoiceTotals: new Map() };

  const NUM_FMT = '#,##0.00;-#,##0.00;"-"';
  const invoiceTotals = new Map();

  for (const mr of cols.rows) {
    const excelRowNum = mr.rowIndex + 1;
    const row = reportSheet.getRow(excelRowNum);
    const invoiceSales = findSalesForMerchant(mr.merchantName, invoiceSalesMap);
    const completeSales = findSalesForMerchant(mr.merchantName, completeSalesMap);
    const invoiceTotal = invoiceSales ? totalSale(invoiceSales) : 0;
    const completeTotal = completeSales ? totalSale(completeSales) : 0;

    invoiceTotals.set(mr.merchantName, invoiceTotal);

    if (cols.colInvoiceSale >= 0) {
      row.getCell(cols.colInvoiceSale + 1).value = invoiceTotal;
      row.getCell(cols.colInvoiceSale + 1).numFmt = NUM_FMT;
    }
    if (cols.colCompleteSale >= 0) {
      row.getCell(cols.colCompleteSale + 1).value = completeTotal;
      row.getCell(cols.colCompleteSale + 1).numFmt = NUM_FMT;
    }
  }

  let mainEndCol = 0;
  for (const c of [cols.colMerchant, cols.colTarget, cols.colInvoiceSale, cols.colCompleteSale, cols.colInvoicePct, cols.colCompletePct]) {
    if (c > mainEndCol) mainEndCol = c;
  }

  const sideTables = detectSideTables(grid, startRow, endRow, startCol, endCol, mainEndCol);
  for (const st of sideTables) {
    for (const mr of st.merchantRows) {
      const excelRowNum = mr.row + 1;
      const row = reportSheet.getRow(excelRowNum);
      const sales = findSalesForMerchant(mr.merchantName, invoiceSalesMap);
      const invoiceVal = sales ? totalSale(sales) : 0;

      for (const dc of st.dataCols) {
        const existingRaw = String(grid[mr.row]?.[dc.col] ?? '').trim();
        const existingNum = parseDouble(existingRaw);
        if (!existingRaw || existingNum === 0) {
          row.getCell(dc.col + 1).value = invoiceVal;
          row.getCell(dc.col + 1).numFmt = '#,##0';
          break;
        }
      }
    }
  }

  return { invoiceTotals };
}

function fillSheet4(reportSheet, rawSheet, totalInvoiceSales) {
  if (!reportSheet) return;
  const { grid, startRow, startCol, endRow, endCol } = readSheetGrid(rawSheet);

  let headerRow = -1;
  let colTarget = -1;
  let colAchievement = -1;

  for (let r = startRow; r <= Math.min(endRow, startRow + 10); r++) {
    for (let c = startCol; c <= endCol; c++) {
      const val = String(grid[r]?.[c] ?? '').trim().toLowerCase();
      if (val.includes('target')) {
        headerRow = r;
        colTarget = c;
        break;
      }
    }
    if (headerRow >= 0) break;
  }

  if (headerRow < 0) return;

  const hRowData = grid[headerRow] || {};
  for (let c = startCol; c <= endCol; c++) {
    const val = String(hRowData[c] ?? '').trim().toLowerCase();
    if (val.includes('achievement')) colAchievement = c;
    if (val.includes('target')) colTarget = c;
  }

  if (colAchievement < 0) return;

  for (let r = headerRow + 1; r <= endRow; r++) {
    const achievementVal = String(grid[r]?.[colAchievement] ?? '').trim();
    if (!achievementVal || parseDouble(achievementVal) === 0) {
      const excelRowNum = r + 1;
      const row = reportSheet.getRow(excelRowNum);
      row.getCell(colAchievement + 1).value = totalInvoiceSales;
      row.getCell(colAchievement + 1).numFmt = '#,##0';
      break;
    }
  }
}

// ===== Main Entry Point =====

export async function generateSupplementVaultReport(
  params
) {
  const {
    orderFiles,
    couponFile,
    aeFile,
    targetFile,
    daysRemainingOnline,
    daysRemainingOutlet,
    totalDays,
    reportDay,
    startDate,
    endDate,
  } = params;

  // 1. Read all order reports, tag with company, merge
  const allOrders = [];
  let orderHeaders = null;

  for (const f of orderFiles) {
    const fileName = f.name.toLowerCase();
    const company = fileName.includes('origin') ? 'Origins' : 'SupplementVault';

    const { rows, headers } = readExcelOrCsv(f);
    if (rows.length > 0 && orderHeaders === null) {
      orderHeaders = headers;
    }
    for (const row of rows) {
      allOrders.push(createOrderRow(company, row));
    }
  }

  // 2. Filter: only keep paid or pending
  let filtered = allOrders.filter((r) => {
    const fs = r.financialStatus.trim().toLowerCase();
    return fs === 'paid' || fs === 'pending';
  });

  // 2b. Created date filtering
  const dateFrom = startDate ? parseLocalDate(startDate) : null;
  const dateTo = endDate ? parseLocalDate(endDate) : null;

  if (dateFrom && dateTo) {
    const fromTime = dateOnly(dateFrom);
    const toTime = dateOnly(dateTo);
    filtered = filtered.filter((r) => {
      const rawDate = r.createdAt || getVal(r, 'Created at');
      const d = parseFulfilledDate(rawDate);
      if (!d) return false;
      const t = dateOnly(d);
      return t >= fromTime && t <= toTime;
    });
  }

  // 3. Read the target table (read early so we can resolve coupon → target mapping)
  const targetRows = readTargetTable(targetFile);
  const rawWb = XLSX.read(targetFile.buffer, { type: 'buffer' });

  // Build set of target merchant names (lowercase → original-cased)
  const targetMerchantNames = new Map();
  for (const tr of targetRows) {
    if (tr.merchantName) {
      targetMerchantNames.set(tr.merchantName.toLowerCase(), tr.merchantName);
    }
  }

  for (let si = 1; si < rawWb.SheetNames.length && si <= 2; si++) {
    const sh = rawWb.Sheets[rawWb.SheetNames[si]];
    if (!sh) continue;
    const { grid: g, startRow: sr, startCol: sc, endRow: er, endCol: ec } = readSheetGrid(sh);
    const detected = detectSheetColumns(g, sr, er, sc, ec);
    if (!detected) continue;

    for (const mr of detected.rows) {
      if (mr.merchantName && !targetMerchantNames.has(mr.merchantName.toLowerCase())) {
        targetMerchantNames.set(mr.merchantName.toLowerCase(), mr.merchantName);
      }
    }
  }

  // 4. Read Merchant Coupon Code file
  const coupons = readCouponFile(couponFile);

  // For each coupon, resolve which name matches a target merchant row.
  // Owner (merchantName) is checked first; if that's not in the target, try Code (discountCode).
  for (const mc of coupons) {
    if (mc.merchantName && targetMerchantNames.has(mc.merchantName.toLowerCase())) {
      mc.targetName = targetMerchantNames.get(mc.merchantName.toLowerCase());
    } else if (mc.discountCode && targetMerchantNames.has(mc.discountCode.toLowerCase())) {
      mc.targetName = targetMerchantNames.get(mc.discountCode.toLowerCase());
    } else {
      // Fuzzy match: normalize both sides (strip spaces, dots, special chars)
      let found = false;
      const ownerNorm = normalizeName(mc.merchantName);
      const codeNorm = normalizeName(mc.discountCode);
      for (const [tnLower, tnOriginal] of targetMerchantNames) {
        const tnNorm = normalizeName(tnLower);
        if (ownerNorm && (tnNorm.includes(ownerNorm) || ownerNorm.includes(tnNorm))) {
          mc.targetName = tnOriginal;
          found = true;
          break;
        }
        if (codeNorm && (tnNorm.includes(codeNorm) || codeNorm.includes(tnNorm))) {
          mc.targetName = tnOriginal;
          found = true;
          break;
        }
      }
      if (!found) {
        mc.targetName = mc.merchantName; // fallback to owner
      }
    }
  }

  // Build discount code -> coupon mapping. Only map actual discount codes, NOT owner names.
  const codeToMerchant = new Map();
  for (const mc of coupons) {
    if (mc.discountCode && mc.discountCode.trim()) {
      const normCode = normalizeCode(mc.discountCode);
      if (!codeToMerchant.has(normCode)) codeToMerchant.set(normCode, mc);
    }
  }

  // Build merchant name -> type mapping (index by both owner and code, and targetName)
  const merchantTypeMap = new Map();
  for (const mc of coupons) {
    if (mc.merchantName) merchantTypeMap.set(mc.merchantName.toLowerCase(), mc.type);
    if (mc.discountCode) merchantTypeMap.set(mc.discountCode.toLowerCase(), mc.type);
    if (mc.targetName) merchantTypeMap.set(mc.targetName.toLowerCase(), mc.type);
  }

  // 5. Calculate sales per merchant per company
  const merchantSalesMap = new Map();
  let unmatchedOrigins = 0;
  let unmatchedSV = 0;

  for (const row of filtered) {
    const code = normalizeCode(row.discountCode);
    const mc = codeToMerchant.get(code);

    if (mc && mc.targetName) {
      const merchantKey = mc.targetName;
      let sales = merchantSalesMap.get(merchantKey);
      if (!sales) {
        sales = { originsSale: 0, svSale: 0 };
        merchantSalesMap.set(merchantKey, sales);
      }
      if (row.company === 'Origins') {
        sales.originsSale += row.total;
      } else {
        sales.svSale += row.total;
      }
    } else {
      if (row.company === 'Origins') {
        unmatchedOrigins += row.total;
      } else {
        unmatchedSV += row.total;
      }
    }
  }

  // Add unmatched to DM General/Sandali
  if (unmatchedOrigins !== 0 || unmatchedSV !== 0) {
    let dmSales = merchantSalesMap.get('DM General/Sandali');
    if (!dmSales) {
      dmSales = { originsSale: 0, svSale: 0 };
      merchantSalesMap.set('DM General/Sandali', dmSales);
    }
    dmSales.originsSale += unmatchedOrigins;
    dmSales.svSale += unmatchedSV;
  }

  // 5b. Process AE Trading file
  // Build AE-specific code map (hyphens stripped for broader matching: MER-110 = MER110)
  const codeToMerchantAE = new Map();
  for (const mc of coupons) {
    if (mc.discountCode && mc.discountCode.trim()) {
      const normCode = normalizeAECode(mc.discountCode);
      if (!codeToMerchantAE.has(normCode)) codeToMerchantAE.set(normCode, mc);
    }
  }

  const missingCouponEntries = [];
  if (aeFile) {
    const aeEntries = readAETradingFile(aeFile);
    for (const entry of aeEntries) {
      // Normalize memo: strip hyphens and lowercase
      let normMemo = normalizeAECode(entry.memo);
      let mc = codeToMerchantAE.get(normMemo);

      // Fallback: try to extract a MER code pattern from within the memo
      if (!mc) {
        const merMatch = entry.memo.match(/\bmer-?\d+\b/i);
        if (merMatch) {
          normMemo = normalizeAECode(merMatch[0]);
          mc = codeToMerchantAE.get(normMemo);
        }
      }

      if (mc && mc.targetName) {
        const merchantKey = mc.targetName;
        let sales = merchantSalesMap.get(merchantKey);
        if (!sales) {
          sales = { originsSale: 0, svSale: 0, aeSale: 0 };
          merchantSalesMap.set(merchantKey, sales);
        }
        if (!sales.aeSale) sales.aeSale = 0;
        sales.aeSale += entry.amount;
      } else {
        missingCouponEntries.push(entry);
      }
    }
  }

  // Identify extra merchants (have sales but not in target)
  const extraMerchants = new Map();
  for (const [name, sales] of merchantSalesMap) {
    // Check exact match first, then normalized match
    if (targetMerchantNames.has(name.toLowerCase())) continue;
    const nameNorm = normalizeName(name);
    let matched = false;
    for (const tnLower of targetMerchantNames.keys()) {
      const tnNorm = normalizeName(tnLower);
      if (tnNorm.includes(nameNorm) || nameNorm.includes(tnNorm)) { matched = true; break; }
    }
    if (!matched) extraMerchants.set(name, sales);
  }

  // 6. Write the output workbook — load the target file as base to preserve all formatting
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(targetFile.buffer);

  // Force Excel to recalculate all formulas when the file is opened
  wb.calcProperties = { fullCalcOnLoad: true };

  // Remember original sheet count for chart restoration
  const originalSheetCount = wb.worksheets.length;

  // Fill the Report sheet (modifies the first sheet in-place)
  await fillReportSheet(
    wb,
    targetFile.buffer,
    targetRows,
    merchantSalesMap,
    merchantTypeMap,
    daysRemainingOnline,
    daysRemainingOutlet,
    totalDays,
    reportDay
  );

  // Created-date mode should only update the top table on sheet 1.

  // Add Sheet: ALL (merged order data)
  writeAllSheet(wb, allOrders, orderHeaders);

  // Add Sheet: Sales
  writeSalesSheet(wb, merchantSalesMap);

  // Add Sheet: Extra Merchants
  if (extraMerchants.size > 0) {
    writeExtraMerchantsSheet(wb, extraMerchants);
  }

  if (missingCouponEntries.length > 0) {
    writeMissingCouponSheet(wb, missingCouponEntries);
  }

  // Write to buffer
  const arrayBuffer = await wb.xlsx.writeBuffer();

  // Restore charts/drawings/media that ExcelJS drops during load/save
  const finalBuffer = await restoreChartsFromOriginal(
    targetFile.buffer,
    Buffer.from(arrayBuffer),
    originalSheetCount
  );

  return finalBuffer;
}
