import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';

// ===== Interfaces removed (TypeScript only) =====

function totalSale(s) {
  return s.originsSale + s.svSale;
}

// ===== Date Parsing =====

const DATE_REGEXES = [
  {
    // yyyy-MM-dd HH:mm:ss +ZZZZ  or  yyyy-MM-ddTHH:mm:ssXXX  or  yyyy-MM-ddTHH:mm:ss  or  yyyy-MM-dd
    regex: /^(\d{4})-(\d{2})-(\d{2})/,
    parse: (m) => new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3])),
  },
  {
    // M/d/yyyy or MM/dd/yyyy
    regex: /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/,
    parse: (m) => new Date(parseInt(m[3]), parseInt(m[1]) - 1, parseInt(m[2])),
  },
];

function parseFulfilledDate(raw) {
  if (!raw || !raw.trim()) return null;
  const trimmed = raw.trim();
  for (const { regex, parse } of DATE_REGEXES) {
    const m = trimmed.match(regex);
    if (m) {
      const d = parse(m);
      if (d && !isNaN(d.getTime())) return d;
    }
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
  return t.trim();
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
  // direct case-insensitive key match
  for (const key of Object.keys(row.data)) {
    if (key.toLowerCase() === col.toLowerCase()) return row.data[key];
  }
  // relaxed contains match
  const normCol = normalizeHeader(col);
  for (const key of Object.keys(row.data)) {
    const normKey = normalizeHeader(key);
    if (normKey.includes(normCol) || normCol.includes(normKey)) return row.data[key];
  }
  return '';
}

function createOrderRow(company, data) {
  const row = { company, data, total: 0, financialStatus: '', discountCode: '' };
  row.financialStatus = getVal(row, 'Financial Status');
  row.discountCode = normalizeCode(getVal(row, 'Discount Code'));
  row.total = parseDouble(getVal(row, 'Total'));
  return row;
}

// ===== File Readers (using xlsx library) =====

function readExcelOrCsv(file) {
  const wb = XLSX.read(file.buffer, { type: 'buffer' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  if (!sheet) return { rows: [], headers: [] };

  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
  if (raw.length === 0) return { rows: [], headers: [] };

  const headers = raw[0].map((h) => (h != null ? String(h).trim() : ''));
  const rows = [];

  for (let r = 1; r < raw.length; r++) {
    const rowArr = raw[r];
    const map = {};
    let hasData = false;
    for (let c = 0; c < headers.length; c++) {
      const val = rowArr[c] != null ? String(rowArr[c]).trim() : '';
      if (val) hasData = true;
      map[headers[c]] = val;
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

/** Find the header row (actual Excel 0-based row number) that contains "merchant name" */
function findHeaderRow(grid, startRow, endRow, startCol, endCol) {
  for (let r = startRow; r <= Math.min(endRow, startRow + 10); r++) {
    const rowData = grid[r];
    if (!rowData) continue;
    for (let c = startCol; c <= endCol; c++) {
      const val = String(rowData[c] ?? '').toLowerCase();
      if (val.includes('merchant name')) return r;
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

  for (let r = dataStartRow; r <= endRow; r++) {
    const rowData = grid[r] || {};
    const merchant = String(rowData[colMerchant] ?? '').trim();
    if (!merchant) continue;

    const merchantLower = merchant.toLowerCase();
    if (merchantLower.includes('total') || merchantLower.includes('grand total')) break;

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

  const cols = ['Merchant Name', 'Origins Sale', 'SupplementVault Sale', 'Total Sale'];
  const headerRow = ws.addRow(cols);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true };
  });

  for (const [name, sales] of salesMap) {
    const row = ws.addRow([name, sales.originsSale, sales.svSale, totalSale(sales)]);
    for (let c = 2; c <= 4; c++) {
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
  ]);
  totalRow.getCell(1).font = { bold: true };
  for (let c = 2; c <= 4; c++) {
    totalRow.getCell(c).font = { bold: true };
    totalRow.getCell(c).numFmt = '#,##0.00';
  }
}

function writeExtraMerchantsSheet(
  wb,
  extraMerchants
) {
  const ws = wb.addWorksheet('Other Merchants');

  const cols = ['Merchant Name', 'Origins Sale', 'SupplementVault Sale', 'Total Sale'];
  const headerRow = ws.addRow(cols);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true };
  });

  for (const [name, sales] of extraMerchants) {
    const row = ws.addRow([name, sales.originsSale, sales.svSale, totalSale(sales)]);
    for (let c = 2; c <= 4; c++) {
      row.getCell(c).numFmt = '#,##0.00';
    }
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

    // Write sale values
    if (colOriginSale >= 0) {
      writeCell(row.getCell(colOriginSale + 1), originSale, NUM_FMT);
    }
    if (colSvSale >= 0) {
      writeCell(row.getCell(colSvSale + 1), svSale, NUM_FMT);
    }

    // Column letters for formulas
    const colLetterOrigin = colOriginSale >= 0 ? colLetter(colOriginSale) : '';
    const colLetterSv = colSvSale >= 0 ? colLetter(colSvSale) : '';
    const colLetterTotal = colTotalSale >= 0 ? colLetter(colTotalSale) : '';
    const colLetterTarget = colTarget >= 0 ? colLetter(colTarget) : '';
    const colLetterBalance = colBalance >= 0 ? colLetter(colBalance) : '';
    const colLetterForecast = colForecast >= 0 ? colLetter(colForecast) : '';

    // Total Sale = ORIGIN Sale + SV.LK Sale
    if (colTotalSale >= 0 && colOriginSale >= 0 && colSvSale >= 0) {
      const formula = `${colLetterOrigin}${excelRowNum}+${colLetterSv}${excelRowNum}`;
      writeCell(row.getCell(colTotalSale + 1), { formula, result: originSale + svSale }, NUM_FMT);
    } else if (colTotalSale >= 0) {
      writeCell(row.getCell(colTotalSale + 1), originSale + svSale, NUM_FMT);
    }

    // Achievement % = IF(Target=0, 0, Total Sale / Target)
    if (colAchievement >= 0 && colTotalSale >= 0 && colTarget >= 0) {
      const totalRef = `${colLetterTotal}${excelRowNum}`;
      const targetRef = `${colLetterTarget}${excelRowNum}`;
      const formula = `IF(${targetRef}=0,0,${totalRef}/${targetRef})`;
      writeCell(row.getCell(colAchievement + 1), { formula, result: 0 }, '0%');
    }

    // Balance = MAX(Target - Total Sale, 0)
    if (colBalance >= 0 && colTarget >= 0 && colTotalSale >= 0) {
      const totalRef = `${colLetterTotal}${excelRowNum}`;
      const targetRef = `${colLetterTarget}${excelRowNum}`;
      const formula = `MAX(${targetRef}-${totalRef},0)`;
      writeCell(row.getCell(colBalance + 1), { formula, result: 0 }, NUM_FMT);
    }

    // Per Day Target = IF(Balance=0, 0, Balance / daysRemaining)
    if (colPerDayTarget >= 0 && colBalance >= 0 && daysRemaining > 0) {
      const balanceRef = `${colLetterBalance}${excelRowNum}`;
      const formula = `IF(${balanceRef}=0,0,${balanceRef}/${daysRemaining})`;
      writeCell(row.getCell(colPerDayTarget + 1), { formula, result: 0 }, NUM_FMT);
    }

    // Forecast Month End Achievement = (Total Sale / reportDay) * totalDays
    if (colForecast >= 0 && colTotalSale >= 0 && reportDay > 0) {
      const totalRef = `${colLetterTotal}${excelRowNum}`;
      const formula = `(${totalRef}/${reportDay})*${totalDays}`;
      writeCell(row.getCell(colForecast + 1), { formula, result: 0 }, NUM_FMT);
    }

    // Forecast Achievement % = IF(Target=0, 0, Forecast / Target)
    if (colForecastPct >= 0 && colForecast >= 0 && colTarget >= 0) {
      const forecastRef = `${colLetterForecast}${excelRowNum}`;
      const targetRef = `${colLetterTarget}${excelRowNum}`;
      const formula = `IF(${targetRef}=0,0,${forecastRef}/${targetRef})`;
      writeCell(row.getCell(colForecastPct + 1), { formula, result: 0 }, '0%');
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
    targetFile,
    daysRemainingOnline,
    daysRemainingOutlet,
    totalDays,
    reportDay,
    fulfilledFrom,
    fulfilledTo,
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

  // 2b. Fulfilled date filtering
  const dateFrom = fulfilledFrom ? parseLocalDate(fulfilledFrom) : null;
  const dateTo = fulfilledTo ? parseLocalDate(fulfilledTo) : null;

  if (dateFrom && dateTo) {
    const fromTime = dateOnly(dateFrom);
    const toTime = dateOnly(dateTo);
    filtered = filtered.filter((r) => {
      const rawDate = getVal(r, 'Fulfilled at');
      const d = parseFulfilledDate(rawDate);
      if (!d) return false;
      const t = dateOnly(d);
      return t >= fromTime && t <= toTime;
    });
  }

  // 3. Read the target table (read early so we can resolve coupon → target mapping)
  const targetRows = readTargetTable(targetFile);

  // Build set of target merchant names
  const targetMerchantNames = new Set();
  for (const tr of targetRows) {
    if (tr.merchantName) {
      targetMerchantNames.add(tr.merchantName.toLowerCase());
    }
  }

  // 4. Read Merchant Coupon Code file
  const coupons = readCouponFile(couponFile);

  // For each coupon, resolve which name matches a target merchant row.
  // Owner (merchantName) is checked first; if that's not in the target, try Code (discountCode).
  for (const mc of coupons) {
    if (mc.merchantName && targetMerchantNames.has(mc.merchantName.toLowerCase())) {
      mc.targetName = mc.merchantName;
    } else if (mc.discountCode && targetMerchantNames.has(mc.discountCode.toLowerCase())) {
      mc.targetName = mc.discountCode;
    } else {
      mc.targetName = mc.merchantName; // fallback to owner
    }
  }

  // Build discount code -> coupon mapping (orders may use either the code or owner as discount code)
  const codeToMerchant = new Map();
  for (const mc of coupons) {
    if (mc.discountCode && mc.discountCode.trim()) {
      const normCode = normalizeCode(mc.discountCode);
      if (!codeToMerchant.has(normCode)) codeToMerchant.set(normCode, mc);
    }
    if (mc.merchantName && mc.merchantName.trim()) {
      const normOwner = normalizeCode(mc.merchantName);
      if (!codeToMerchant.has(normOwner)) codeToMerchant.set(normOwner, mc);
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

  // Identify extra merchants (have sales but not in target)
  const extraMerchants = new Map();
  for (const [name, sales] of merchantSalesMap) {
    if (!targetMerchantNames.has(name.toLowerCase())) {
      extraMerchants.set(name, sales);
    }
  }

  // 6. Write the output workbook — load the target file as base to preserve all formatting
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(targetFile.buffer);

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

  // Add Sheet: ALL (merged order data)
  writeAllSheet(wb, allOrders, orderHeaders);

  // Add Sheet: Sales
  writeSalesSheet(wb, merchantSalesMap);

  // Add Sheet: Extra Merchants
  if (extraMerchants.size > 0) {
    writeExtraMerchantsSheet(wb, extraMerchants);
  }

  // Write to buffer
  const arrayBuffer = await wb.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}
