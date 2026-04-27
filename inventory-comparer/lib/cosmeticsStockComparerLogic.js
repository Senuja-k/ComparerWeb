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

function parseInventoryFile(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
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
      // Only Active, only out-of-stock (qty ≤ 0)
      if (status.toLowerCase() !== 'active') continue;
      if (qty > 0) continue;
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

export async function generateCosmeticsStockReport(inventoryFile) {
  const { mainMap, otherMap } = parseInventoryFile(inventoryFile.buffer);

  // Build result rows
  const results = [];

  for (const [skuLower, mainEntry] of mainMap) {
    const otherShops = otherMap.get(skuLower) || [];

    // Split by priority and only keep shops that have qty > 0
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
  const sheet = workbook.addWorksheet('Stock Report');

  sheet.columns = [
    { header: 'SKU', key: 'sku', width: 28 },
    { header: 'Product Title', key: 'title', width: 36 },
    { header: 'Cosmetics.lk Qty', key: 'cosmeticsQty', width: 18 },
    { header: 'Priority 1 Shop(s)', key: 'p1shops', width: 46 },
    { header: 'Priority 1 Qty', key: 'p1qty', width: 16 },
    { header: 'Priority 2 Shop(s)', key: 'p2shops', width: 36 },
    { header: 'Priority 2 Qty', key: 'p2qty', width: 16 },
    { header: 'Priority 3 Shop(s)', key: 'p3shops', width: 46 },
    { header: 'Priority 3 Qty', key: 'p3qty', width: 16 },
    { header: 'Stock Available Elsewhere', key: 'available', width: 26 },
  ];

  // Header styling
  const headerRow = sheet.getRow(1);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00695C' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FF004D40' } } };
  });
  headerRow.height = 22;

  function formatShopList(shops) {
    return shops.map((s) => displayShopName(s.shopName)).join('\n');
  }

  function formatQtyList(shops) {
    return shops.map((s) => String(s.qty)).join('\n');
  }

  for (const row of results) {
    const p1ShopText = formatShopList(row.priority1);
    const p1QtyText  = formatQtyList(row.priority1);
    const p2ShopText = formatShopList(row.priority2);
    const p2QtyText  = formatQtyList(row.priority2);
    const p3ShopText = formatShopList(row.priority3);
    const p3QtyText  = formatQtyList(row.priority3);

    const dataRow = sheet.addRow({
      sku: row.sku,
      title: row.productTitle,
      cosmeticsQty: row.cosmeticsQty,
      p1shops: p1ShopText,
      p1qty: p1QtyText,
      p2shops: p2ShopText,
      p2qty: p2QtyText,
      p3shops: p3ShopText,
      p3qty: p3QtyText,
      available: row.hasAnyStock ? 'YES' : 'NO',
    });

    // Wrap text in multi-value cells
    ['p1shops', 'p1qty', 'p2shops', 'p2qty', 'p3shops', 'p3qty'].forEach((key) => {
      dataRow.getCell(key).alignment = { wrapText: true, vertical: 'top' };
    });

    dataRow.getCell('cosmeticsQty').alignment = { horizontal: 'center', vertical: 'middle' };
    dataRow.getCell('available').alignment = { horizontal: 'center', vertical: 'middle' };

    if (row.hasAnyStock) {
      // Green highlight — stock found elsewhere
      dataRow.getCell('available').fill = {
        type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC8E6C9' },
      };
      dataRow.getCell('available').font = { bold: true, color: { argb: 'FF1B5E20' } };
    } else {
      // Red highlight — no stock anywhere
      dataRow.getCell('available').fill = {
        type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFCDD2' },
      };
      dataRow.getCell('available').font = { bold: true, color: { argb: 'FFB71C1C' } };
      dataRow.eachCell((cell, colNumber) => {
        if (colNumber <= 3) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9C4' } };
        }
      });
    }
  }

  sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 10 } };

  const buffer = await workbook.xlsx.writeBuffer();
  return buffer;
}
