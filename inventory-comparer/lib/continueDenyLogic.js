import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';

// ── Cell value helper ──────────────────────────────────────────────────────────

function getCellValue(sheet, row, col) {
  const addr = XLSX.utils.encode_cell({ r: row, c: col });
  const cell = sheet[addr];
  if (!cell) return '';
  if (cell.t === 'n') return cell.v;
  if (cell.t === 'b') return cell.v;
  if (cell.v !== undefined) return String(cell.v).trim();
  return '';
}

function getCellString(sheet, row, col) {
  const val = getCellValue(sheet, row, col);
  return val === null || val === undefined ? '' : String(val).trim();
}

// ── SKU normalisation: replace underscores with dashes ────────────────────────

function normaliseSku(sku) {
  let s = sku.replace(/_/g, '-');
  if (s.toUpperCase().startsWith('OGF-')) s = s.slice(4);
  return s;
}

// ── Find column index by header name ──────────────────────────────────────────

function findColumnIndex(sheet, headerRow, headerName) {
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cellVal = getCellString(sheet, headerRow, c);
    if (cellVal.toLowerCase() === headerName.toLowerCase()) {
      return c;
    }
  }
  return -1;
}

// ── Workbook reader (supports .xlsx, .xls, .csv) ────────────────────────────

function readWorkbook(buffer, fileName) {
  const ext = (fileName || '').split('.').pop().toLowerCase();
  if (ext === 'csv') {
    return XLSX.read(buffer.toString('utf8'), { type: 'string' });
  }
  return XLSX.read(buffer, { type: 'buffer' });
}

// ── Parse Cosmetics file ───────────────────────────────────────────────────────
// Required columns: SKU, Inventory Policy, Shop Name, Product_Status
// Filters applied at this stage:
//   - Shop Name is exactly "Cosmetics.lk" (does NOT contain a dash)
// Product_Status is stored but NOT filtered here — both Active and Draft are
// included so that Draft SKUs are not incorrectly reported as missing.
// SKUs are normalised: underscores replaced with dashes.

function parseCosmeticsFile(buffer, fileName = '') {
  const workbook = readWorkbook(buffer, fileName);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  if (!sheet || !sheet['!ref']) {
    throw new Error('Cosmetics file is empty or unreadable');
  }

  const range = XLSX.utils.decode_range(sheet['!ref']);
  const headerRow = range.s.r;

  const skuCol    = findColumnIndex(sheet, headerRow, 'SKU');
  const policyCol = findColumnIndex(sheet, headerRow, 'Inventory Policy');
  const shopCol   = findColumnIndex(sheet, headerRow, 'Shop Name');
  const statusCol = findColumnIndex(sheet, headerRow, 'Product_Status');

  if (skuCol    === -1) throw new Error('Cosmetics file is missing a "SKU" column');
  if (policyCol === -1) throw new Error('Cosmetics file is missing an "Inventory Policy" column');
  if (shopCol   === -1) throw new Error('Cosmetics file is missing a "Shop Name" column');
  if (statusCol === -1) throw new Error('Cosmetics file is missing a "Product_Status" column');

  // Map: normalised SKU (lowercase key) -> { sku, policy, productStatus }
  const map = new Map();
  for (let r = headerRow + 1; r <= range.e.r; r++) {
    const rawSku      = getCellString(sheet, r, skuCol);
    const policy      = getCellString(sheet, r, policyCol);
    const shopName    = getCellString(sheet, r, shopCol);
    const productStatus = getCellString(sheet, r, statusCol);

    if (!rawSku) continue;

    // Only rows where Shop Name is exactly "Cosmetics.lk" (no dash suffix)
    if (shopName.includes('-')) continue;
    if (shopName.toLowerCase() !== 'cosmetics.lk') continue;

    const sku = normaliseSku(rawSku);
    // If the same SKU appears as both Draft and Active, Active wins
    const existing = map.get(sku.toLowerCase());
    if (existing && existing.productStatus.toLowerCase() === 'active') continue;

    map.set(sku.toLowerCase(), { sku, policy, productStatus });
  }

  return map;
}

// ── Parse Supplement file ──────────────────────────────────────────────────────
// Expected columns: SKU, Available

function parseSupplementFile(buffer, fileName = '') {
  const workbook = readWorkbook(buffer, fileName);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  if (!sheet || !sheet['!ref']) {
    throw new Error('Supplement file is empty or unreadable');
  }

  const range = XLSX.utils.decode_range(sheet['!ref']);
  const headerRow = range.s.r;

  const skuCol = findColumnIndex(sheet, headerRow, 'SKU');
  const availableCol = findColumnIndex(sheet, headerRow, 'Available');

  if (skuCol === -1) throw new Error('Supplement file is missing a "SKU" column');
  if (availableCol === -1) throw new Error('Supplement file is missing an "Available" column');

  const rows = [];
  for (let r = headerRow + 1; r <= range.e.r; r++) {
    const rawSku = getCellString(sheet, r, skuCol);
    if (!rawSku) continue;
    const sku = normaliseSku(rawSku);
    const rawAvail = getCellValue(sheet, r, availableCol);
    const available = typeof rawAvail === 'number' ? rawAvail : parseFloat(String(rawAvail).replace(/[^\d.\-]/g, '')) || 0;
    rows.push({ sku, available });
  }

  return rows;
}

// ── Main report generator ──────────────────────────────────────────────────────

export async function generateContinueDenyReport(cosmeticsFile, supplementFile) {
  const cosmeticsMap = parseCosmeticsFile(cosmeticsFile.buffer, cosmeticsFile.name);
  const supplementRows = parseSupplementFile(supplementFile.buffer, supplementFile.name);

  const matched = [];      // Active rows
  const draftMatched = []; // Draft rows — present in Cosmetics but not Active
  const missing = [];      // Truly absent from Cosmetics

  for (const { sku, available } of supplementRows) {
    const cosmeticsEntry = cosmeticsMap.get(sku.toLowerCase());
    if (!cosmeticsEntry) {
      missing.push(sku);
      continue;
    }

    const { policy, productStatus } = cosmeticsEntry;
    const isDraft = productStatus.toLowerCase() === 'draft';
    const flagged =
      (available > 0 && policy.toLowerCase() === 'deny') ||
      (available <= 0 && policy.toLowerCase() === 'continue');

    const row = {
      sku: cosmeticsEntry.sku,
      available,
      inventoryPolicy: policy,
      productStatus,
      flagged,
    };

    if (isDraft) {
      draftMatched.push(row);
    } else {
      matched.push(row);
    }
  }

  // ── Build Excel workbook ───────────────────────────────────────────────────

  const workbook = new ExcelJS.Workbook();

  // ── Helper: style a header row ─────────────────────────────────────────────

  function styleHeader(row, bgArgb = 'FF1E1E1E') {
    row.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = { bottom: { style: 'thin', color: { argb: 'FF444444' } } };
    });
    row.height = 20;
  }

  // ── Sheet 1: Active Matched SKUs ──────────────────────────────────────────

  const matchedSheet = workbook.addWorksheet('Matched SKUs (Active)');
  matchedSheet.columns = [
    { header: 'SKU', key: 'sku', width: 30 },
    { header: 'Available (Supplement)', key: 'available', width: 24 },
    { header: 'Inventory Policy (Cosmetics)', key: 'inventoryPolicy', width: 28 },
    { header: 'Flagged', key: 'flagged', width: 12 },
  ];
  styleHeader(matchedSheet.getRow(1));

  for (const row of matched) {
    const dataRow = matchedSheet.addRow({
      sku: row.sku,
      available: row.available,
      inventoryPolicy: row.inventoryPolicy,
      flagged: row.flagged ? 'YES' : 'NO',
    });

    if (row.flagged) {
      dataRow.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFCCCC' } };
        cell.font = { bold: true, color: { argb: 'FFCC0000' } };
      });
    }

    dataRow.getCell('flagged').alignment = { horizontal: 'center' };
    dataRow.getCell('available').alignment = { horizontal: 'center' };
    dataRow.getCell('inventoryPolicy').alignment = { horizontal: 'center' };
  }

  matchedSheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 4 } };

  // ── Sheet 2: Draft SKUs ───────────────────────────────────────────────────
  // These exist in Cosmetics but have Product_Status = Draft.
  // Available > 0 AND Inventory Policy = Deny is flagged as MAJOR ISSUE.

  const draftSheet = workbook.addWorksheet('Draft SKUs');
  draftSheet.columns = [
    { header: 'SKU', key: 'sku', width: 30 },
    { header: 'Available (Supplement)', key: 'available', width: 24 },
    { header: 'Inventory Policy (Cosmetics)', key: 'inventoryPolicy', width: 28 },
    { header: 'Product Status', key: 'productStatus', width: 16 },
    { header: 'Flag', key: 'flag', width: 18 },
  ];
  styleHeader(draftSheet.getRow(1), 'FF424242');

  for (const row of draftMatched) {
    const flagLabel = row.flagged ? 'MAJOR ISSUE' : '-';
    const dataRow = draftSheet.addRow({
      sku: row.sku,
      available: row.available,
      inventoryPolicy: row.inventoryPolicy,
      productStatus: row.productStatus,
      flag: flagLabel,
    });

    if (row.flagged) {
      dataRow.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE0B2' } };
        cell.font = { bold: true, color: { argb: 'FFE65100' } };
      });
    }

    dataRow.getCell('flag').alignment = { horizontal: 'center' };
    dataRow.getCell('available').alignment = { horizontal: 'center' };
    dataRow.getCell('inventoryPolicy').alignment = { horizontal: 'center' };
    dataRow.getCell('productStatus').alignment = { horizontal: 'center' };
  }

  draftSheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 5 } };

  // ── Sheet 3: Missing SKUs ─────────────────────────────────────────────────

  const missingSheet = workbook.addWorksheet('Missing in Cosmetics');
  missingSheet.columns = [
    { header: 'SKU (not found in Cosmetics)', key: 'sku', width: 36 },
  ];
  styleHeader(missingSheet.getRow(1), 'FF424242');

  for (const sku of missing) {
    const dataRow = missingSheet.addRow({ sku });
    dataRow.getCell('sku').fill = {
      type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9C4' },
    };
    dataRow.getCell('sku').font = { color: { argb: 'FF7B5E00' } };
  }

  // ── Serialize ──────────────────────────────────────────────────────────────

  const buffer = await workbook.xlsx.writeBuffer();
  return buffer;
}
