import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { restoreChartsFromOriginal } from './xlsxChartPreserver.js';

// ===== Helpers =====

function totalSale(s) {
  return s.originsSale + s.svSale + (s.aeSale || 0);
}

const DATE_REGEXES = [
  {
    // yyyy-MM-dd (ISO format, with optional time)
    regex: /^(\d{4})-(\d{2})-(\d{2})/,
    parse: (m) => new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3])),
  },
  {
    // dd/mm/yyyy (with optional time)
    regex: /^(\d{1,2})\/(\d{1,2})\/(\d{4})/,
    parse: (m) => new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1])),
  },
];

function parseDate(raw) {
  if (raw == null) return null;
  const trimmed = String(raw).trim();
  if (!trimmed) return null;
  for (const { regex, parse } of DATE_REGEXES) {
    const m = trimmed.match(regex);
    if (m) {
      const d = parse(m);
      if (d && !isNaN(d.getTime())) return d;
    }
  }

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
  return s.trim().toLowerCase().replace(/[\n\r\t ]+/g, '').replace(/[_\-]/g, '');
}

function parseDouble(s) {
  if (!s) return 0;
  const cleaned = String(s).replace(/[^\d.\-]/g, '');
  if (!cleaned) return 0;
  const v = parseFloat(cleaned);
  return isNaN(v) ? 0 : v;
}

function isExactTarget(headerLower) {
  if (!headerLower || !headerLower.trim()) return false;
  const h = headerLower.toLowerCase();
  if (!h.includes('target')) return false;
  if (h.includes('march') || h.includes('per day') || h.includes('forecast') || h.includes('achievement') || h.includes('day'))
    return false;
  return true;
}

function colLetter(col) {
  let sb = '';
  let c = col + 1;
  while (c > 0) {
    c--;
    sb = String.fromCharCode(65 + (c % 26)) + sb;
    c = Math.floor(c / 26);
  }
  return sb;
}

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
  const row = { company, data, total: 0, financialStatus: '', discountCode: '', fulfilledAt: '', createdAt: '' };
  row.financialStatus = getVal(row, 'Financial Status');
  row.discountCode = normalizeCode(getVal(row, 'Discount Code'));
  row.total = parseDouble(getVal(row, 'Total'));
  row.fulfilledAt = getVal(row, 'Fulfilled at');
  row.createdAt = getVal(row, 'Created at');
  return row;
}

// ===== File Readers =====

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
  let colMerchant = -1, colCode = -1, colType = -1;
  for (let c = 0; c < headerArr.length; c++) {
    const hLower = String(headerArr[c] ?? '').trim().toLowerCase();
    if (colMerchant < 0 && hLower.includes('owner')) colMerchant = c;
    else if (colCode < 0 && hLower.includes('code') && !hLower.includes('owner')) colCode = c;
    else if (colType < 0 && hLower.includes('type')) colType = c;
  }
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

  for (let r = headerRowIdx + 1; r < raw.length; r++) {
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
 * Columns: Type (filter: Invoice only), Memo (coupon/merchant code), Amount (sale total), Date, No.
 */
function readAETradingFile(file, dateFrom, dateTo) {
  const entries = [];
  const wb = XLSX.read(file.buffer, { type: 'buffer' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  if (!sheet) return entries;

  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
  if (raw.length === 0) return entries;

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
  let colMemo = -1, colAmount = -1, colType = -1, colNo = -1, colDate = -1;
  for (let c = 0; c < headers.length; c++) {
    const h = String(headers[c] ?? '').trim().toLowerCase();
    if ((h === 'memo' || h.includes('memo')) && colMemo < 0) colMemo = c;
    else if ((h === 'amount' || h.includes('amount')) && colAmount < 0) colAmount = c;
    else if ((h === 'type' || h.includes('type')) && colType < 0) colType = c;
    else if ((h === 'no.' || h === 'no' || h === 'number') && colNo < 0) colNo = c;
    else if ((h === 'date' || h.includes('date')) && colDate < 0) colDate = c;
  }
  if (colMemo < 0 || colAmount < 0) return entries;

  const fromTime = dateFrom ? dateOnly(dateFrom) : null;
  const toTime = dateTo ? dateOnly(dateTo) : null;

  for (let r = headerRowIdx + 1; r < raw.length; r++) {
    const rowArr = raw[r];
    if (colType >= 0) {
      const type = String(rowArr[colType] ?? '').trim().toLowerCase();
      if (type && type !== 'invoice') continue;
    }
    if ((fromTime || toTime) && colDate >= 0) {
      const rawDate = String(rowArr[colDate] ?? '').trim();
      const d = parseDate(rawDate);
      if (!d) continue;
      const t = dateOnly(d);
      if (fromTime && t < fromTime) continue;
      if (toTime && t > toTime) continue;
    }
    const memo = String(rowArr[colMemo] ?? '').trim();
    const amount = parseDouble(String(rowArr[colAmount] ?? ''));
    const no = colNo >= 0 ? String(rowArr[colNo] ?? '').trim() : '';
    if (!memo && amount === 0) continue;
    entries.push({ memo, amount, no });
  }
  return entries;
}

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

function readTargetTableSheet1(file) {
  const result = [];
  const wb = XLSX.read(file.buffer, { type: 'buffer' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  if (!sheet) return result;
  const { grid, startRow, startCol, endRow, endCol } = readSheetGrid(sheet);
  const headerRow = findHeaderRow(grid, startRow, endRow, startCol, endCol);
  if (headerRow < 0) return result;

  const hRow = grid[headerRow] || {};
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

  let colMerchant = -1, colTarget = -1, colOutlet = -1;
  const tRowAbove = headerRow > startRow ? grid[headerRow - 1] || {} : {};
  const tRowBelow = headerRow + 1 <= endRow ? grid[headerRow + 1] || {} : {};

  for (let c = startCol; c <= endCol; c++) {
    const above = String(tRowAbove[c] ?? '').trim();
    const main = String(hRow[c] ?? '').trim();
    const below = String(tRowBelow[c] ?? '').trim();
    const val = main.toLowerCase().replace(/[^a-z0-9% /().,]/g, '').trim();
    const allRows = (above + ' ' + main + ' ' + below).toLowerCase().replace(/[^a-z0-9% /().,]/g, '').trim();

    if (colMerchant < 0 && (val.includes('merchant name') || allRows.includes('merchant name'))) colMerchant = c;
    else if (colTarget < 0 && isExactTarget(val)) colTarget = c;
    else if (colOutlet < 0 && (val === 'outlet' || allRows.includes('outlet'))) colOutlet = c;
  }

  if (colTarget < 0) {
    for (let c = startCol; c <= endCol; c++) {
      const above = String(tRowAbove[c] ?? '').toLowerCase().replace(/[^a-z0-9 ]/g, '').trim();
      const below = String(tRowBelow[c] ?? '').toLowerCase().replace(/[^a-z0-9 ]/g, '').trim();
      if (isExactTarget(above) || isExactTarget(below)) { colTarget = c; break; }
    }
  }

  if (colMerchant < 0) return result;

  for (let r = dataStartRow; r <= endRow; r++) {
    const rowData = grid[r] || {};
    const merchant = String(rowData[colMerchant] ?? '').trim();
    if (!merchant) continue;
    const merchantLower = merchant.toLowerCase();
    if (merchantLower.includes('total') || merchantLower.includes('grand total')) break;
    result.push({
      rowIndex: r,
      merchantName: merchant,
      target: colTarget >= 0 ? parseDouble(String(rowData[colTarget] ?? '')) : 0,
      outlet: colOutlet >= 0 ? String(rowData[colOutlet] ?? '').trim() : '',
    });
  }
  return result;
}

// ===== Merchant Sales Computation =====

function buildMerchantMappings(coupons, targetMerchantNames) {
  // Resolve targetName for each coupon: owner first, then code, then fuzzy match
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

  // Build code → coupon mapping. Only map actual discount codes, NOT owner names.
  const codeToMerchant = new Map();
  for (const mc of coupons) {
    if (mc.discountCode && mc.discountCode.trim()) {
      const normCode = normalizeCode(mc.discountCode);
      if (!codeToMerchant.has(normCode)) codeToMerchant.set(normCode, mc);
    }
  }

  const merchantTypeMap = new Map();
  for (const mc of coupons) {
    if (mc.merchantName) merchantTypeMap.set(mc.merchantName.toLowerCase(), mc.type);
    if (mc.discountCode) merchantTypeMap.set(mc.discountCode.toLowerCase(), mc.type);
    if (mc.targetName) merchantTypeMap.set(mc.targetName.toLowerCase(), mc.type);
  }

  return { codeToMerchant, merchantTypeMap };
}

// Words too generic to use for name-matching heuristic
const COMMON_WORDS = new Set(['sale', 'shop', 'order', 'pvt', 'ltd', 'the', 'for', 'and', 'trading', 'lk', 'web']);

function significantWords(str) {
  return str
    .toLowerCase()
    .split(/[^a-z0-9]+/)
    .filter((w) => w.length >= 4 && !COMMON_WORDS.has(w));
}

function findTargetByCode(codes, targetMerchantNames) {
  for (const code of codes) {
    if (targetMerchantNames.has(code)) return targetMerchantNames.get(code);
    const codeNorm = normalizeName(code);
    for (const [tnLower, tnOriginal] of targetMerchantNames) {
      const tnNorm = normalizeName(tnLower);
      if (codeNorm && tnNorm && (codeNorm === tnNorm || codeNorm.includes(tnNorm) || tnNorm.includes(codeNorm))) {
        return tnOriginal;
      }
    }
    const codeWords = significantWords(code);
    if (codeWords.length === 0) continue;
    for (const [tnLower, tnOriginal] of targetMerchantNames) {
      const tnWords = significantWords(tnLower);
      if (tnWords.length === 0) continue;
      const shared = codeWords.filter((w) => tnWords.includes(w));
      if (shared.length > 0 && shared.length >= Math.min(codeWords.length, tnWords.length) / 2) {
        return tnOriginal;
      }
    }
  }
  return null;
}

function computeSalesMap(filteredOrders, codeToMerchant, targetMerchantNames) {
  const merchantSalesMap = new Map();
  let unmatchedOrigins = 0;
  let unmatchedSV = 0;

  for (const row of filteredOrders) {
    const rawCodes = String(row.discountCode || '').split(',').map((s) => normalizeCode(s.trim())).filter(Boolean);
    let mc = null;
    for (const code of rawCodes) {
      const candidate = codeToMerchant.get(code);
      if (candidate) { mc = candidate; break; }
    }
    let merchantKey = mc?.targetName ?? null;
    if (!merchantKey) {
      merchantKey = findTargetByCode(rawCodes, targetMerchantNames);
    }
    if (merchantKey) {
      let sales = merchantSalesMap.get(merchantKey);
      if (!sales) { sales = { originsSale: 0, svSale: 0 }; merchantSalesMap.set(merchantKey, sales); }
      if (row.company === 'Origins') sales.originsSale += row.total;
      else sales.svSale += row.total;
    } else {
      if (row.company === 'Origins') unmatchedOrigins += row.total;
      else unmatchedSV += row.total;
    }
  }

  if (unmatchedOrigins !== 0 || unmatchedSV !== 0) {
    let dmSales = merchantSalesMap.get('DM General/Sandali');
    if (!dmSales) { dmSales = { originsSale: 0, svSale: 0 }; merchantSalesMap.set('DM General/Sandali', dmSales); }
    dmSales.originsSale += unmatchedOrigins;
    dmSales.svSale += unmatchedSV;
  }

  return merchantSalesMap;
}

function findSalesForMerchant(merchantKey, salesMap) {
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

// ===== Date range filtering =====

function filterByDateColumn(orders, colName, fromTime, toTime) {
  return orders.filter((r) => {
    const raw = getVal(r, colName);
    const d = parseDate(raw);
    if (!d) return false;
    const t = dateOnly(d);
    return t >= fromTime && t <= toTime;
  });
}

// ===== Sheet 1: same as created-date logic but using Fulfilled at date range =====

function fillSheet1(reportSheet, rawSheet, targetRows, merchantSalesMap) {
  if (!reportSheet) return;
  const { grid, startRow, startCol, endRow, endCol } = readSheetGrid(rawSheet);
  const headerRowIdx = findHeaderRow(grid, startRow, endRow, startCol, endCol);
  if (headerRowIdx < 0) return;

  const hRow = grid[headerRowIdx] || {};
  const rowAbove = headerRowIdx > startRow ? grid[headerRowIdx - 1] || {} : {};
  const rowBelow = headerRowIdx + 1 <= endRow ? grid[headerRowIdx + 1] || {} : {};

  let colMerchant = -1;
  // Collect ALL target columns, total sale columns, and achievement % columns
  const targetCols = [];
  const totalSaleCols = [];
  const achievementPctCols = [];

  for (let c = startCol; c <= endCol; c++) {
    const above = String(rowAbove[c] ?? '').toLowerCase().trim().replace(/[^a-z0-9% /().,]/g, '').trim();
    const main = String(hRow[c] ?? '').toLowerCase().trim().replace(/[^a-z0-9% /().,]/g, '').trim();
    const below = String(rowBelow[c] ?? '').toLowerCase().trim().replace(/[^a-z0-9% /().,]/g, '').trim();
    const all = above + ' ' + main + ' ' + below;

    if (colMerchant < 0 && all.includes('merchant')) colMerchant = c;
    if (isExactTarget(main) || isExactTarget(above) || isExactTarget(below)) targetCols.push(c);
    if ((below.includes('total') && below.includes('sale')) || (main.includes('total') && main.includes('sale'))) totalSaleCols.push(c);
    if (all.includes('achievement') && all.includes('%')) achievementPctCols.push(c);
  }

  // Find the first data row to check which columns are empty
  const firstDataRow = targetRows.length > 0 ? targetRows[0].rowIndex : -1;

  // Find the EMPTY Total Sale column (the one for the current month)
  let colTotalSale = -1;
  if (firstDataRow >= 0 && totalSaleCols.length > 1) {
    for (const c of totalSaleCols) {
      const val = String(grid[firstDataRow]?.[c] ?? '').trim();
      if (!val) { colTotalSale = c; break; }
    }
  }
  // Fallback: if only one Total Sale col or none were empty, use the last one
  if (colTotalSale < 0 && totalSaleCols.length > 0) {
    colTotalSale = totalSaleCols[totalSaleCols.length - 1];
  }

  // Find the matching Achievement % and Target columns for the chosen Total Sale column
  // The Achievement % column is the one immediately after the Total Sale column
  let colAchievement = -1;
  achievementPctCols.sort((a, b) => a - b);
  for (const ac of achievementPctCols) {
    if (ac > colTotalSale) { colAchievement = ac; break; }
  }
  // Fallback if none found after: use the last achievement col
  if (colAchievement < 0 && achievementPctCols.length > 0) {
    colAchievement = achievementPctCols[achievementPctCols.length - 1];
  }

  // Find the Target column that corresponds to the current month
  // It's the target column closest before the chosen Total Sale column
  let colTarget = -1;
  targetCols.sort((a, b) => a - b);
  for (let i = targetCols.length - 1; i >= 0; i--) {
    if (targetCols[i] < colTotalSale) { colTarget = targetCols[i]; break; }
  }
  // Fallback: use last target col
  if (colTarget < 0 && targetCols.length > 0) colTarget = targetCols[targetCols.length - 1];

  const NUM_FMT = '#,##0.00;-#,##0.00;"-"';

  for (const tr of targetRows) {
    const excelRowNum = tr.rowIndex + 1;
    const row = reportSheet.getRow(excelRowNum);
    const sales = findSalesForMerchant(tr.merchantName, merchantSalesMap);
    const total = sales ? totalSale(sales) : 0;

    if (colTotalSale >= 0) {
      row.getCell(colTotalSale + 1).value = total;
      row.getCell(colTotalSale + 1).numFmt = NUM_FMT;
    }
    // Leave Achievement % and Total rows untouched — original formulas will recalculate
  }
}

// ===== Sheets 2 & 3: Invoice Date + Complete Date columns + side tables =====

function detectSheetColumns(grid, startRow, endRow, startCol, endCol) {
  // Find header row with "Merchant" or "Merchant Name"
  let headerRowIdx = -1;
  for (let r = startRow; r <= Math.min(endRow, startRow + 10); r++) {
    const rowData = grid[r];
    if (!rowData) continue;
    for (let c = startCol; c <= endCol; c++) {
      const val = String(rowData[c] ?? '').toLowerCase();
      if (val.includes('merchant')) { headerRowIdx = r; break; }
    }
    if (headerRowIdx >= 0) break;
  }
  if (headerRowIdx < 0) return null;

  const hRow = grid[headerRowIdx] || {};
  const rowAbove = headerRowIdx > startRow ? grid[headerRowIdx - 1] || {} : {};
  const rowBelow = headerRowIdx + 1 <= endRow ? grid[headerRowIdx + 1] || {} : {};

  let colMerchant = -1, colTarget = -1, colInvoiceSale = -1, colCompleteSale = -1;
  const achievementPctCols = [];
  let colInvoicePct = -1, colCompletePct = -1;

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

  // If we couldn't find them via combined text, try direct header matching
  if (colInvoiceSale < 0 || colCompleteSale < 0) {
    for (let c = startCol; c <= endCol; c++) {
      const main = String(hRow[c] ?? '').toLowerCase().trim();
      const above = String(rowAbove[c] ?? '').toLowerCase().trim();
      if (main.includes('invoice date') || above.includes('invoice date')) {
        if (main.includes('total sale') || above.includes('total sale')) {
          if (colInvoiceSale < 0) colInvoiceSale = c;
        }
      }
      if (main.includes('complete date') || above.includes('complete date')) {
        if (main.includes('total sale') || above.includes('total sale')) {
          if (colCompleteSale < 0) colCompleteSale = c;
        }
      }
    }
  }

  // Achievement % columns — sort and assign: first is for invoice, second for complete
  achievementPctCols.sort((a, b) => a - b);
  if (achievementPctCols.length >= 1) colInvoicePct = achievementPctCols[0];
  if (achievementPctCols.length >= 2) colCompletePct = achievementPctCols[1];

  // Find data rows
  let dataStartRow = headerRowIdx + 1;
  // Skip sub-header rows
  while (dataStartRow <= endRow) {
    const dr = grid[dataStartRow] || {};
    let isSub = false;
    for (let c = startCol; c <= endCol; c++) {
      const v = String(dr[c] ?? '').toLowerCase();
      if (v.includes('origin') || v.includes('sv.lk') || v.includes('total sale')) { isSub = true; break; }
    }
    if (!isSub) break;
    dataStartRow++;
  }

  // Collect merchant rows (skip subtotal rows like "Merchant Total" but don't stop)
  const rows = [];
  for (let r = dataStartRow; r <= endRow; r++) {
    const rowData = grid[r] || {};
    const merchant = colMerchant >= 0 ? String(rowData[colMerchant] ?? '').trim() : '';
    if (!merchant) continue;
    const mLower = merchant.toLowerCase();
    // Stop only at the final "Total" row (exact match), not at subtotals
    if (mLower === 'total') break;
    // Skip subtotal rows like "Merchant Total" — don't process but continue scanning
    if (mLower.includes('merchant total')) continue;
    rows.push({ rowIndex: r, merchantName: merchant, target: colTarget >= 0 ? parseDouble(String(rowData[colTarget] ?? '')) : 0 });
  }

  return {
    headerRowIdx, colMerchant, colTarget, colInvoiceSale, colCompleteSale,
    colInvoicePct, colCompletePct, dataStartRow, rows,
  };
}

/**
 * Detect side tables in a sheet. Side tables are smaller tables to the right of the main table.
 * They have Row Labels / merchant names and numeric columns for months.
 * There can be multiple side tables at different row positions.
 * Returns array of { headerRow, colRowLabel, dataCols: [{col, headerName}], merchantRows: [{row, merchantName}] }
 */
function detectSideTables(grid, startRow, endRow, startCol, endCol, mainTableEndCol) {
  const tables = [];
  const scanStart = mainTableEndCol + 1;

  // Track which (row, col) combos we've already claimed as a table header
  const usedHeaders = new Set();

  // Scan ALL rows (not just top 10) for "Row Label" blocks to the right of main table
  for (let c = scanStart; c <= endCol; c++) {
    for (let r = startRow; r <= endRow; r++) {
      const val = String(grid[r]?.[c] ?? '').trim().toLowerCase();
      if (val.includes('row label') && !usedHeaders.has(`${r},${c}`)) {
        usedHeaders.add(`${r},${c}`);
        const table = { headerRow: r, colRowLabel: c, dataCols: [], merchantRows: [] };

        // Find data columns (month columns) to the right
        for (let cc = c + 1; cc <= endCol; cc++) {
          const hVal = String(grid[r]?.[cc] ?? '').trim();
          if (!hVal) break;
          table.dataCols.push({ col: cc, headerName: hVal });
        }

        // Find merchant rows below
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

function fillSheet2or3(reportSheet, rawSheet, invoiceSalesMap, completeSalesMap) {
  if (!reportSheet) return { invoiceTotals: new Map() };
  const { grid, startRow, startCol, endRow, endCol } = readSheetGrid(rawSheet);
  const cols = detectSheetColumns(grid, startRow, endRow, startCol, endCol);
  if (!cols) return { invoiceTotals: new Map() };

  const NUM_FMT = '#,##0.00;-#,##0.00;"-"';
  const invoiceTotals = new Map(); // merchant -> total invoice sale

  for (const mr of cols.rows) {
    const excelRowNum = mr.rowIndex + 1;
    const row = reportSheet.getRow(excelRowNum);

    const invoiceSales = findSalesForMerchant(mr.merchantName, invoiceSalesMap);
    const completeSales = findSalesForMerchant(mr.merchantName, completeSalesMap);

    const invoiceTotal = invoiceSales ? totalSale(invoiceSales) : 0;
    const completeTotal = completeSales ? totalSale(completeSales) : 0;

    invoiceTotals.set(mr.merchantName, invoiceTotal);

    // Invoice Date column — write raw value only
    if (cols.colInvoiceSale >= 0) {
      row.getCell(cols.colInvoiceSale + 1).value = invoiceTotal;
      row.getCell(cols.colInvoiceSale + 1).numFmt = NUM_FMT;
    }

    // Complete Date column — write raw value only
    if (cols.colCompleteSale >= 0) {
      row.getCell(cols.colCompleteSale + 1).value = completeTotal;
      row.getCell(cols.colCompleteSale + 1).numFmt = NUM_FMT;
    }
    // Leave Achievement %, Total rows, etc. untouched — original formulas will recalculate
  }

  // Fill side tables with invoice date totals
  let mainEndCol = 0;
  for (const c of [cols.colMerchant, cols.colTarget, cols.colInvoiceSale, cols.colCompleteSale, cols.colInvoicePct, cols.colCompletePct]) {
    if (c > mainEndCol) mainEndCol = c;
  }

  const sideTables = detectSideTables(grid, startRow, endRow, startCol, endCol, mainEndCol);

  for (const st of sideTables) {
    for (const mr of st.merchantRows) {
      const excelRowNum = mr.row + 1;
      const row = reportSheet.getRow(excelRowNum);

      let invoiceVal = 0;
      const sales = findSalesForMerchant(mr.merchantName, invoiceSalesMap);
      if (sales) invoiceVal = totalSale(sales);

      // Find the next available (empty) data cell for this merchant in the side table
      for (const dc of st.dataCols) {
        const existingVal = String(grid[mr.row]?.[dc.col] ?? '').trim();
        if (!existingVal) {
          row.getCell(dc.col + 1).value = invoiceVal;
          row.getCell(dc.col + 1).numFmt = '#,##0';
          break;
        }
      }
    }
  }

  return { invoiceTotals };
}

// ===== Sheet 4: Achievement table =====

function fillSheet4(reportSheet, rawSheet, totalInvoiceSales) {
  if (!reportSheet) return;
  const { grid, startRow, startCol, endRow, endCol } = readSheetGrid(rawSheet);

  // Find the header with "Target" and "Achievement"
  let headerRow = -1;
  let colLabel = -1, colTarget = -1, colAchievement = -1;

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

  // Find the label column (usually one to the left of target) and achievement column
  const hRowData = grid[headerRow] || {};
  for (let c = startCol; c <= endCol; c++) {
    const val = String(hRowData[c] ?? '').trim().toLowerCase();
    if (val.includes('achievement')) colAchievement = c;
    if (val.includes('target')) colTarget = c;
  }

  // Label column: first column before target with no typical header or the first column
  colLabel = colTarget > startCol ? colTarget - 1 : startCol;

  if (colAchievement < 0) return;

  // Find the next empty row in the achievement column
  for (let r = headerRow + 1; r <= endRow; r++) {
    const achievementVal = String(grid[r]?.[colAchievement] ?? '').trim();
    if (!achievementVal) {
      // This is the next empty cell — write the total invoice sales here
      const excelRowNum = r + 1;
      const row = reportSheet.getRow(excelRowNum);
      row.getCell(colAchievement + 1).value = totalInvoiceSales;
      row.getCell(colAchievement + 1).numFmt = '#,##0';
      break;
    }
  }
}

// ===== Sheet Writers (same as created-date logic) =====

function writeAllSheet(wb, allOrders, headers) {
  const ws = wb.addWorksheet('ALL');
  if (!headers || allOrders.length === 0) return;
  const allHeaders = ['Company', ...headers];
  const headerRow = ws.addRow(allHeaders);
  headerRow.eachCell((cell) => { cell.font = { bold: true }; });
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

function writeSalesSheet(wb, fulfilledMap, createdMap) {
  const ws = wb.addWorksheet('Sales');
  const cols = ['Merchant Name', 'Complete Sale (Fulfilled at)', 'Invoice Sale (Created at)'];
  const headerRow = ws.addRow(cols);
  headerRow.eachCell((cell) => { cell.font = { bold: true }; });

  const allNames = new Set([...fulfilledMap.keys(), ...createdMap.keys()]);
  for (const name of allNames) {
    const fs = fulfilledMap.get(name);
    const cs = createdMap.get(name);
    const row = ws.addRow([
      name,
      fs ? totalSale(fs) : 0,
      cs ? totalSale(cs) : 0,
    ]);
    for (let c = 2; c <= 3; c++) row.getCell(c).numFmt = '#,##0.00';
  }
  const dataStart = 2;
  const dataEnd = dataStart + allNames.size - 1;
  const totalRow = ws.addRow([
    'Total',
    { formula: `SUM(B${dataStart}:B${dataEnd})` },
    { formula: `SUM(C${dataStart}:C${dataEnd})` },
  ]);
  totalRow.getCell(1).font = { bold: true };
  for (let c = 2; c <= 3; c++) {
    totalRow.getCell(c).font = { bold: true };
    totalRow.getCell(c).numFmt = '#,##0.00';
  }
}

// Debug sheet: every individual order included in the Invoice (Created at) filter
function writeInvoiceOrdersDebugSheet(wb, createdFiltered) {
  const ws = wb.addWorksheet('Invoice Orders (Debug)');
  const headers = ['Name', 'Company', 'Financial Status', 'Created at', 'Fulfilled at', 'Total', 'Discount Code'];
  const headerRow = ws.addRow(headers);
  headerRow.eachCell((cell) => { cell.font = { bold: true }; });
  for (const r of createdFiltered) {
    ws.addRow([
      getVal(r, 'Name'),
      r.company,
      r.financialStatus,
      r.createdAt,
      r.fulfilledAt,
      r.total,
      getVal(r, 'Discount Code'),
    ]);
  }
  // Grand total of what the code is summing
  const dataEnd = createdFiltered.length + 1;
  const totalRow = ws.addRow(['TOTAL', '', '', '', '', { formula: `SUM(F2:F${dataEnd})` }, '']);
  totalRow.getCell(1).font = { bold: true };
  totalRow.getCell(6).font = { bold: true };
  totalRow.getCell(6).numFmt = '#,##0.00';
}

function writeExtraMerchantsSheet(wb, extraMerchants) {
  const ws = wb.addWorksheet('Other Merchants');
  const cols = ['Merchant Name', 'Origins Sale', 'SupplementVault Sale', 'Total Sale'];
  const headerRow = ws.addRow(cols);
  headerRow.eachCell((cell) => { cell.font = { bold: true }; });
  for (const [name, sales] of extraMerchants) {
    const row = ws.addRow([name, sales.originsSale, sales.svSale, totalSale(sales)]);
    for (let c = 2; c <= 4; c++) row.getCell(c).numFmt = '#,##0.00';
  }
}

function writeMissingCouponSheet(wb, missingEntries) {
  const ws = wb.addWorksheet('Missing Coupon (AE)');
  const headerRow = ws.addRow(['No.', 'Memo', 'Amount']);
  headerRow.eachCell((cell) => { cell.font = { bold: true }; });
  for (const entry of missingEntries) {
    const row = ws.addRow([entry.no || '', entry.memo, entry.amount]);
    row.getCell(3).numFmt = '#,##0.00';
  }
}

// ===== Main Entry Point =====

export async function generateFulfilledDateReport(params) {
  const { orderFiles, couponFile, aeFile, targetFile, startDate, endDate } = params;

  // 1. Read all order reports
  const allOrders = [];
  let orderHeaders = null;
  for (const f of orderFiles) {
    const fileName = f.name.toLowerCase();
    const company = fileName.includes('origin') ? 'Origins' : 'SupplementVault';
    const { rows, headers } = readExcelOrCsv(f);
    if (rows.length > 0 && orderHeaders === null) orderHeaders = headers;
    for (const row of rows) {
      allOrders.push(createOrderRow(company, row));
    }
  }

  // 2. Filter: only paid or pending
  const paidOrPending = allOrders.filter((r) => {
    const fs = r.financialStatus.trim().toLowerCase();
    return fs === 'paid' || fs === 'pending';
  });

  // 3. Date range
  const fromDate = parseLocalDate(startDate);
  const toDate = parseLocalDate(endDate);
  if (!fromDate || !toDate) throw new Error('Invalid date range');
  const fromTime = dateOnly(fromDate);
  const toTime = dateOnly(toDate);

  // Filter by Fulfilled at
  const fulfilledFiltered = filterByDateColumn(paidOrPending, 'Fulfilled at', fromTime, toTime);
  // Filter by Created at
  const createdFiltered = filterByDateColumn(paidOrPending, 'Created at', fromTime, toTime);

  // 4. Read coupon file and target (Sheet1)
  const coupons = readCouponFile(couponFile);
  const sheet1Targets = readTargetTableSheet1(targetFile);

  // Map lowercase name → original-cased name for all target merchants
  const targetMerchantNames = new Map();
  for (const tr of sheet1Targets) {
    if (tr.merchantName) targetMerchantNames.set(tr.merchantName.toLowerCase(), tr.merchantName);
  }

  // Also gather merchant names from sheets 2 & 3
  const rawWb = XLSX.read(targetFile.buffer, { type: 'buffer' });
  for (let si = 1; si < rawWb.SheetNames.length && si <= 2; si++) {
    const sh = rawWb.Sheets[rawWb.SheetNames[si]];
    if (!sh) continue;
    const { grid: g, startRow: sr, startCol: sc, endRow: er, endCol: ec } = readSheetGrid(sh);
    const c = detectSheetColumns(g, sr, er, sc, ec);
    if (c) {
      for (const mr of c.rows) {
        if (mr.merchantName && !targetMerchantNames.has(mr.merchantName.toLowerCase())) {
          targetMerchantNames.set(mr.merchantName.toLowerCase(), mr.merchantName);
        }
      }
    }
  }

  const { codeToMerchant, merchantTypeMap } = buildMerchantMappings(coupons, targetMerchantNames);

  // 5. Compute sales maps
  const fulfilledSalesMap = computeSalesMap(fulfilledFiltered, codeToMerchant, targetMerchantNames);
  const createdSalesMap = computeSalesMap(createdFiltered, codeToMerchant, targetMerchantNames);

  // 5b. Process AE Trading file — AE has no fulfilled/created distinction so add to both maps
  const codeToMerchantAE = new Map();
  for (const mc of coupons) {
    if (mc.discountCode && mc.discountCode.trim()) {
      const normCode = normalizeAECode(mc.discountCode);
      if (!codeToMerchantAE.has(normCode)) codeToMerchantAE.set(normCode, mc);
    }
  }

  const missingCouponEntries = [];
  if (aeFile) {
    const fromDate = parseLocalDate(startDate);
    const toDate = parseLocalDate(endDate);
    const aeEntries = readAETradingFile(aeFile, fromDate, toDate);
    for (const entry of aeEntries) {
      let normMemo = normalizeAECode(entry.memo);
      let mc = codeToMerchantAE.get(normMemo);
      if (!mc) {
        const merMatch = entry.memo.match(/\bmer-?\d+\b/i);
        if (merMatch) {
          normMemo = normalizeAECode(merMatch[0]);
          mc = codeToMerchantAE.get(normMemo);
        }
      }
      if (mc && mc.targetName) {
        const merchantKey = mc.targetName;
        // Add to both maps with same amount (AE has no date distinction)
        for (const sMap of [fulfilledSalesMap, createdSalesMap]) {
          let sales = sMap.get(merchantKey);
          if (!sales) {
            sales = { originsSale: 0, svSale: 0, aeSale: 0 };
            sMap.set(merchantKey, sales);
          }
          if (!sales.aeSale) sales.aeSale = 0;
          sales.aeSale += entry.amount;
        }
      } else {
        missingCouponEntries.push(entry);
      }
    }
  }

  // Identify extra merchants (have sales but not in target)
  const extraMerchants = new Map();
  for (const [name, sales] of fulfilledSalesMap) {
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

  // 6. Load the target file as base workbook to preserve formatting, graphs, etc.
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(targetFile.buffer);

  // Force Excel to recalculate all formulas when the file is opened
  wb.calcProperties = { fullCalcOnLoad: true };

  // Remember original sheet count for chart restoration
  const originalSheetCount = wb.worksheets.length;

  // Sheet 1: Fill Total Sale (using Fulfilled at filtered data)
  if (wb.worksheets.length >= 1) {
    const rawSheet1 = rawWb.Sheets[rawWb.SheetNames[0]];
    fillSheet1(wb.worksheets[0], rawSheet1, sheet1Targets, fulfilledSalesMap);
  }

  // Sheet 2: Invoice Date (Created at) + Complete Date (Fulfilled at)
  let sheet2InvoiceTotals = new Map();
  if (wb.worksheets.length >= 2 && rawWb.SheetNames.length >= 2) {
    const rawSheet2 = rawWb.Sheets[rawWb.SheetNames[1]];
    const result = fillSheet2or3(wb.worksheets[1], rawSheet2, createdSalesMap, fulfilledSalesMap);
    sheet2InvoiceTotals = result.invoiceTotals;
  }

  // Sheet 3: Same logic as sheet 2
  let sheet3InvoiceTotals = new Map();
  if (wb.worksheets.length >= 3 && rawWb.SheetNames.length >= 3) {
    const rawSheet3 = rawWb.Sheets[rawWb.SheetNames[2]];
    const result = fillSheet2or3(wb.worksheets[2], rawSheet3, createdSalesMap, fulfilledSalesMap);
    sheet3InvoiceTotals = result.invoiceTotals;
  }

  // Sheet 4: Sum invoice totals from sheets 2 & 3 and write to Achievement
  if (wb.worksheets.length >= 4 && rawWb.SheetNames.length >= 4) {
    let totalInvoiceSales = 0;
    for (const [, val] of sheet2InvoiceTotals) totalInvoiceSales += val;
    for (const [, val] of sheet3InvoiceTotals) totalInvoiceSales += val;

    const rawSheet4 = rawWb.Sheets[rawWb.SheetNames[3]];
    fillSheet4(wb.worksheets[3], rawSheet4, totalInvoiceSales);
  }

  // Add Sheet: ALL (merged order data)
  writeAllSheet(wb, allOrders, orderHeaders);

  // Add Sheet: Sales (fulfilled) + Invoice Sales (created) side by side for comparison
  writeSalesSheet(wb, fulfilledSalesMap, createdSalesMap);

  // Add debug sheet: every order row included in the Invoice (Created at) calculation
  writeInvoiceOrdersDebugSheet(wb, createdFiltered);

  // Add Sheet: Other Merchants
  if (extraMerchants.size > 0) {
    writeExtraMerchantsSheet(wb, extraMerchants);
  }

  if (missingCouponEntries.length > 0) {
    writeMissingCouponSheet(wb, missingCouponEntries);
  }

  const arrayBuffer = await wb.xlsx.writeBuffer();

  // Restore charts/drawings/media that ExcelJS drops during load/save
  const finalBuffer = await restoreChartsFromOriginal(
    targetFile.buffer,
    Buffer.from(arrayBuffer),
    originalSheetCount
  );

  return finalBuffer;
}
