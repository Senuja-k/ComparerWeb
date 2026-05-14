import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';

// ── Priority shop definitions ──────────────────────────────────────────────────
// Priority 1 (index 0): Pevi Trading - Web and SPK Trading Lanka Pvt Ltd - Web
// Priority 2 (index 1): Pepiliyana Shop
// Priority 3 (index 2): any other shop

const PRIORITY_1_SHOPS = [
  'cosmetics.lk - pevi trading - web',
  'cosmetics.lk - spk trading lanka pvt ltd - web',
];
const PRIORITY_2_SHOPS = [
  'cosmetics.lk - pepiliyana shop',
];

function getShopPriority(shopNameLower) {
  if (PRIORITY_1_SHOPS.includes(shopNameLower)) return 1;
  if (PRIORITY_2_SHOPS.includes(shopNameLower)) return 2;
  return 3;
}

// ── Short display names for the report ────────────────────────────────────────

const SHOP_DISPLAY_NAMES = {
  'cosmetics.lk - pevi trading - web':              'Pevi',
  'cosmetics.lk - chami trading - web':             'Chami',
  'cosmetics.lk - spk trading lanka pvt ltd - web': 'SPK',
  'cosmetics.lk - kiribathgoda shop':               'Kiribathgoda',
  'cosmetics.lk - pepiliyana shop':                 'Pepiliyana',
  'cosmetics.lk - maharagama shop':                 'Maharagama',
  'cosmetics.lk - one galle face outlet':           'OGF',
  'cosmetics.lk - cool planet outlet':              'Cool Planet',
};

function displayShopName(shopName) {
  return SHOP_DISPLAY_NAMES[shopName.toLowerCase()] ?? shopName;
}

// ── Cell helpers ───────────────────────────────────────────────────────────────

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

// ── SKU normalisation ─────────────────────────────────────────────────────────

function normaliseSku(sku) {
  let s = sku.replace(/_/g, '-');
  if (s.toUpperCase().startsWith('OGF-')) s = s.slice(4);
  return s;
}

// ── Parse inventory file ───────────────────────────────────────────────────────
// Returns:
//   mainRows  — SKU→{ sku, qty, productTitle } for Cosmetics.lk Active rows with qty ≤ 0
//   otherRows — Map<normalisedSku, Array<{ shopName, qty }>> for all other shops

function parseInventoryFile(buffer, threshold = 0, fileName = '') {
  const workbook = readWorkbook(buffer, fileName);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  if (!sheet || !sheet['!ref']) throw new Error('Inventory file is empty or unreadable');

  const range = XLSX.utils.decode_range(sheet['!ref']);
  const headerRow = range.s.r;

  const skuCol    = findColumnIndex(sheet, headerRow, 'SKU');
  const shopCol   = findColumnIndex(sheet, headerRow, 'Shop Name');
  const statusCol = findColumnIndex(sheet, headerRow, 'Product_Status');
  const qtyCol    = findColumnIndex(sheet, headerRow, 'Inventory Quantity');
  const titleCol  = findColumnIndex(sheet, headerRow, 'Product Title');

  if (skuCol  === -1) throw new Error('File is missing a "SKU" column');
  if (shopCol === -1) throw new Error('File is missing a "Shop Name" column');
  if (qtyCol  === -1) throw new Error('File is missing an "Inventory Quantity" column');
  if (statusCol === -1) throw new Error('File is missing a "Product_Status" column');

  // mainMap: normalisedSku (lower) → { sku, qty, productTitle }
  // Only Cosmetics.lk (exact, no dash), Active, qty ≤ 0
  const mainMap = new Map();

  // otherMap: normalisedSku (lower) → Array<{ shopName, qty }>
  const otherMap = new Map();

  for (let r = headerRow + 1; r <= range.e.r; r++) {
    const rawSku    = getCellString(sheet, r, skuCol);
    if (!rawSku) continue;

    const sku       = normaliseSku(rawSku);
    const shopName  = getCellString(sheet, r, shopCol);
    const status    = getCellString(sheet, r, statusCol);
    const rawQty    = getCellValue(sheet, r, qtyCol);
    const qty       = typeof rawQty === 'number'
      ? rawQty
      : parseFloat(String(rawQty).replace(/[^\d.\-]/g, '')) || 0;
    const title     = titleCol !== -1 ? getCellString(sheet, r, titleCol) : '';

    const shopLower = shopName.toLowerCase();
    const isMainShop = !shopName.includes('-') && shopLower === 'cosmetics.lk';

    if (isMainShop) {
      // Only Active, only out-of-stock (qty ≤ threshold)
      if (status.toLowerCase() !== 'active') continue;
      if (qty > threshold) continue;
      // Last-write wins if duplicate SKU
      mainMap.set(sku.toLowerCase(), { sku, qty, productTitle: title });
    } else {
      // All other shops — collect for lookup regardless of status
      if (!otherMap.has(sku.toLowerCase())) {
        otherMap.set(sku.toLowerCase(), []);
      }
      otherMap.get(sku.toLowerCase()).push({ shopName, qty });
    }
  }

  return { mainMap, otherMap };
}

// ── Main report generator ─────────────────────────────────────────────────────

export async function generateCosmeticsStockReport(inventoryFile, threshold = 0) {
  const { mainMap, otherMap } = parseInventoryFile(inventoryFile.buffer, threshold, inventoryFile.name);

  // Build result rows
  const results = [];

  for (const [skuLower, mainEntry] of mainMap) {
    const otherShops = otherMap.get(skuLower) || [];

    const p1 = otherShops
      .filter((s) => getShopPriority(s.shopName.toLowerCase()) === 1 && s.qty > 0)
      .sort((a, b) => a.shopName.localeCompare(b.shopName));

    const p2 = otherShops
      .filter((s) => getShopPriority(s.shopName.toLowerCase()) === 2 && s.qty > 0)
      .sort((a, b) => a.shopName.localeCompare(b.shopName));

    const p3 = otherShops
      .filter((s) => getShopPriority(s.shopName.toLowerCase()) === 3 && s.qty > 0)
      .sort((a, b) => a.shopName.localeCompare(b.shopName));

    const hasAnyStock = p1.length > 0 || p2.length > 0 || p3.length > 0;

    results.push({
      sku: mainEntry.sku,
      productTitle: mainEntry.productTitle,
      cosmeticsQty: mainEntry.qty,
      hasAnyStock,
      priority1: p1,
      priority2: p2,
      priority3: p3,
    });
  }

  // Sort: rows that have stock elsewhere first, then alphabetically by SKU
  results.sort((a, b) => {
    if (a.hasAnyStock !== b.hasAnyStock) return a.hasAnyStock ? -1 : 1;
    return a.sku.localeCompare(b.sku);
  });

  // ── Build Excel workbook ─────────────────────────────────────────────────

  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet('Stock Report');

  // Col indices (1-based):
  // 1=SKU, 2=Title, 3=CosmeticsQty, 4=P1Shop, 5=P1Qty, 6=P2Shop, 7=P2Qty, 8=P3Shop, 9=P3Qty, 10=Available
  ws.columns = [
    { header: 'SKU',                       key: 'sku',         width: 28 },
    { header: 'Product Title',             key: 'title',       width: 36 },
    { header: 'Cosmetics.lk Qty',          key: 'cosmeticsQty',width: 18 },
    { header: 'Priority 1 Shop(s)',        key: 'p1shops',     width: 26 },
    { header: 'Priority 1 Qty',            key: 'p1qty',       width: 14 },
    { header: 'Priority 2 Shop(s)',        key: 'p2shops',     width: 26 },
    { header: 'Priority 2 Qty',            key: 'p2qty',       width: 14 },
    { header: 'Priority 3 Shop(s)',        key: 'p3shops',     width: 26 },
    { header: 'Priority 3 Qty',            key: 'p3qty',       width: 14 },
    { header: 'Stock Available Elsewhere', key: 'available',   width: 26 },
  ];

  // Header styling
  const headerRow = ws.getRow(1);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00695C' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FF004D40' } } };
  });
  headerRow.height = 22;

  let currentExcelRow = 2;

  for (const row of results) {
    const numRows = Math.max(row.priority1.length, row.priority2.length, row.priority3.length, 1);
    const startRow = currentExcelRow;
    const endRow   = currentExcelRow + numRows - 1;

    for (let i = 0; i < numRows; i++) {
      const excelRow = ws.getRow(currentExcelRow);

      // Shared (merged) columns — only write on first sub-row
      if (i === 0) {
        excelRow.getCell(1).value  = row.sku;
        excelRow.getCell(2).value  = row.productTitle;
        excelRow.getCell(3).value  = row.cosmeticsQty;
        excelRow.getCell(10).value = row.hasAnyStock ? 'YES' : 'NO';
      }

      // Priority 1
      const p1 = row.priority1[i];
      excelRow.getCell(4).value = p1 ? displayShopName(p1.shopName) : null;
      excelRow.getCell(5).value = p1 ? p1.qty : null;

      // Priority 2
      const p2 = row.priority2[i];
      excelRow.getCell(6).value = p2 ? displayShopName(p2.shopName) : null;
      excelRow.getCell(7).value = p2 ? p2.qty : null;

      // Priority 3
      const p3 = row.priority3[i];
      excelRow.getCell(8).value = p3 ? displayShopName(p3.shopName) : null;
      excelRow.getCell(9).value = p3 ? p3.qty : null;

      [4, 6, 8].forEach((c) => {
        excelRow.getCell(c).alignment = { horizontal: 'left', vertical: 'middle' };
      });
      [5, 7, 9].forEach((c) => {
        excelRow.getCell(c).alignment = { horizontal: 'center', vertical: 'middle' };
      });

      excelRow.commit();
      currentExcelRow++;
    }

    // Merge shared columns across all sub-rows for this SKU
    if (numRows > 1) {
      ws.mergeCells(startRow, 1, endRow, 1);  // SKU
      ws.mergeCells(startRow, 2, endRow, 2);  // Product Title
      ws.mergeCells(startRow, 3, endRow, 3);  // Cosmetics.lk Qty
      ws.mergeCells(startRow, 10, endRow, 10); // Stock Available Elsewhere
    }

    // Alignment for merged cells (always applied)
    ws.getCell(startRow, 1).alignment  = { vertical: 'middle', horizontal: 'left' };
    ws.getCell(startRow, 2).alignment  = { vertical: 'middle', horizontal: 'left' };
    ws.getCell(startRow, 3).alignment  = { horizontal: 'center', vertical: 'middle' };
    ws.getCell(startRow, 10).alignment = { horizontal: 'center', vertical: 'middle' };

    const availableCell = ws.getCell(startRow, 10);
    if (row.hasAnyStock) {
      availableCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC8E6C9' } };
      availableCell.font = { bold: true, color: { argb: 'FF1B5E20' } };
    } else {
      availableCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFCDD2' } };
      availableCell.font = { bold: true, color: { argb: 'FFB71C1C' } };
      for (let c = 1; c <= 3; c++) {
        ws.getCell(startRow, c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9C4' } };
      }
    }
  }

  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 10 } };

  const buffer = await workbook.xlsx.writeBuffer();
  return buffer;
}
