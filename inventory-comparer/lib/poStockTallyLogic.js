import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';

// ── Company alias mapping ──────────────────────────────────────────────────────

const COMPANY_ALIAS_TO_CODE = new Map([
  ['AJS', 'OUT200'],
  ['KRIBATHGODA', 'OUT200'],
  ['KIRI', 'OUT200'],
  ['MNK', 'OUT100'],
  ['COOLPLANET', 'OUT100'],
  ['CP', 'OUT100'],
  ['LMJ', 'OUT400'],
  ['PEPILIYANA', 'OUT400'],
  ['PEP', 'OUT400'],
  ['LWK', 'OUT300'],
  ['OGF', 'OUT300'],
  ['DRO', 'OUT700'],
  ['MAH', 'OUT700'],
  ['MAHARAGAMA', 'OUT700'],
  ['CHAMI', 'OUT500'],
  ['SPK', 'OUT800'],
  ['COSMETICS', 'OUT600'],
  ['COS', 'OUT600'],
  ['OUT010', 'OUT010'],
  ['OUT100', 'OUT100'],
  ['OUT200', 'OUT200'],
  ['OUT300', 'OUT300'],
  ['OUT400', 'OUT400'],
  ['OUT500', 'OUT500'],
  ['OUT600', 'OUT600'],
  ['OUT700', 'OUT700'],
  ['OUT800', 'OUT800'],
]);

const VALID_COMPANY_CODES = new Set(COMPANY_ALIAS_TO_CODE.values());

// ── Helpers ────────────────────────────────────────────────────────────────────

function safeTrim(s) {
  return s == null ? '' : String(s).trim();
}

function isEmpty(s) {
  return s == null || String(s).trim().length === 0;
}

function getCellValue(row, colIndex) {
  if (colIndex == null || row == null) return '';
  const cell = row[colIndex];
  if (cell == null) return '';
  if (cell.w != null) return String(cell.w).trim();
  if (cell.v != null) return String(cell.v).trim();
  return '';
}

function deriveCompanyCodeFromFileName(fileName) {
  const upper = fileName.toUpperCase();
  for (const [alias, code] of COMPANY_ALIAS_TO_CODE) {
    if (upper.includes(alias.toUpperCase())) return code;
  }
  return '';
}

function deriveCompanyCodeFromSupplierCell(supplier) {
  if (supplier == null) return '';
  const upper = supplier.toUpperCase();
  for (const [alias, code] of COMPANY_ALIAS_TO_CODE) {
    if (upper.includes(alias.toUpperCase())) return code;
  }
  return supplier.toUpperCase();
}

function parseDate(s) {
  if (isEmpty(s)) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function isWithinOneWeek(poDateStr, stockDateStr) {
  const d1 = parseDate(poDateStr);
  const d2 = parseDate(stockDateStr);
  if (!d1 || !d2) return false;
  const diffMs = Math.abs(d1.getTime() - d2.getTime());
  const diffDays = diffMs / (1000 * 60 * 60 * 24);
  return diffDays <= 7;
}

// ── File readers ───────────────────────────────────────────────────────────────

function readSheetRows(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true, dateNF: 'yyyy-mm-dd' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  if (!sheet) return [];
  const range = XLSX.utils.decode_range(sheet['!ref'] ?? 'A1');
  const rows = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      row[c] = sheet[addr];
    }
    rows.push(row);
  }
  return rows;
}

function readPurchaseOrderFile(name, buffer) {
  const records = [];
  const shopName = name.replace(/\.(xlsx|xls)$/i, '');

  try {
    const rows = readSheetRows(buffer);
    if (rows.length === 0) return records;

    const headerRow = rows[0];
    const columnMap = {};
    for (let c = 0; c < headerRow.length; c++) {
      const header = getCellValue(headerRow, c);
      switch (header) {
        case 'Purchase Order': columnMap['PurchaseOrder'] = c; break;
        case 'Supplier': columnMap['Supplier'] = c; break;
        case 'Product': columnMap['Product'] = c; break;
        case 'SKU': columnMap['SKU'] = c; break;
        case 'Barcode': columnMap['Barcode'] = c; break;
        case 'PO Date': columnMap['PODate'] = c; break;
        case 'Quantity': columnMap['Quantity'] = c; break;
      }
    }

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row) continue;

      const po = getCellValue(row, columnMap['PurchaseOrder']);
      const supplierCell = getCellValue(row, columnMap['Supplier']);
      const product = getCellValue(row, columnMap['Product']);
      const sku = getCellValue(row, columnMap['SKU']);
      const barcode = getCellValue(row, columnMap['Barcode']);
      const date = getCellValue(row, columnMap['PODate']);
      const qtyStr = getCellValue(row, columnMap['Quantity']);

      const supplierCode = deriveCompanyCodeFromSupplierCell(supplierCell);
      if (!VALID_COMPANY_CODES.has(supplierCode) && supplierCode !== 'OUT010') continue;

      if (!isEmpty(po) && !isEmpty(supplierCode)) {
        const qty = parseInt(qtyStr.trim(), 10);
        if (!isNaN(qty) && qty > 0) {
          records.push({
            purchaseOrderNo: po.trim(),
            supplier: supplierCode,
            product: safeTrim(product),
            sku: safeTrim(sku),
            barcode: safeTrim(barcode),
            date: safeTrim(date),
            quantity: qty,
            shop: safeTrim(shopName),
          });
        }
      }
    }
  } catch (e) {
    console.error(`Error reading PO file '${name}':`, e);
  }
  return records;
}

function readStockAdjustmentFile(name, buffer) {
  const records = [];
  const companyName = name.replace(/\.(xlsx|xls)$/i, '').trim();
  const companyCode = deriveCompanyCodeFromFileName(companyName);

  try {
    const rows = readSheetRows(buffer);
    if (rows.length === 0) return records;

    const headerRow = rows[0];
    const columnMap = {};
    for (let c = 0; c < headerRow.length; c++) {
      const header = getCellValue(headerRow, c);
      switch (header) {
        case 'SKU': columnMap['SKU'] = c; break;
        case 'Barcode': columnMap['Barcode'] = c; break;
        case 'Date': columnMap['Date'] = c; break;
        case 'Reason': columnMap['Reason'] = c; break;
        case 'Adjustment': columnMap['Adjustment'] = c; break;
        case 'No.': columnMap['SAID'] = c; break;
      }
    }

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row) continue;

      const sku = getCellValue(row, columnMap['SKU']);
      const barcode = getCellValue(row, columnMap['Barcode']);
      const date = getCellValue(row, columnMap['Date']);
      const reason = getCellValue(row, columnMap['Reason']);
      const adjStr = getCellValue(row, columnMap['Adjustment']);
      const saId = getCellValue(row, columnMap['SAID']);

      if (!isEmpty(sku)) {
        const adjustment = parseInt(adjStr.trim(), 10);
        if (!isNaN(adjustment)) {
          records.push({
            sku: safeTrim(sku),
            barcode: safeTrim(barcode),
            date: safeTrim(date),
            reason: safeTrim(reason),
            adjustment,
            company: safeTrim(companyName),
            saId: safeTrim(saId),
            companyCode: safeTrim(companyCode),
            sourceFile: name,
          });
        }
      }
    }
  } catch (e) {
    console.error(`Error reading Stock file '${name}':`, e);
  }
  return records;
}

// ── ID conflict detection ──────────────────────────────────────────────────────

function detectIDConflicts(records) {
  const dateSkuGroups = new Map();
  const dateBarcodeGroups = new Map();

  for (const r of records) {
    if (!isEmpty(r.date)) {
      const skuKey = `${r.date}|${r.sku ?? ''}`;
      if (!dateSkuGroups.has(skuKey)) dateSkuGroups.set(skuKey, []);
      dateSkuGroups.get(skuKey).push(r);

      if (!isEmpty(r.barcode) && r.barcode.toLowerCase() !== 'no barcode') {
        const barcodeKey = `${r.date}|${r.barcode}`;
        if (!dateBarcodeGroups.has(barcodeKey)) dateBarcodeGroups.set(barcodeKey, []);
        dateBarcodeGroups.get(barcodeKey).push(r);
      }
    }
  }

  for (const group of dateSkuGroups.values()) {
    const barcodes = new Set(
      group
        .map(rec => rec.barcode)
        .filter(b => !isEmpty(b) && b.toLowerCase() !== 'no barcode')
    );
    if (barcodes.size > 1) {
      const conflict = `Same SKU different barcodes: ${[...barcodes].join(', ')}`;
      for (const rec of group) rec.idConflict = conflict;
    }
  }

  for (const group of dateBarcodeGroups.values()) {
    const skus = new Set(
      group.map(rec => rec.sku).filter(s => !isEmpty(s))
    );
    if (skus.size > 1) {
      const conflict = `Same barcode different SKUs: ${[...skus].join(', ')}`;
      for (const rec of group) {
        if (isEmpty(rec.idConflict)) rec.idConflict = conflict;
        else rec.idConflict += '; ' + conflict;
      }
    }
  }
}

// ── Pass 1 – Exact matching ───────────────────────────────────────────────────

function isExactMatch(po, stock) {
  const dateMatch = po.date === stock.date;
  const skuMatch = po.sku === stock.sku;
  const barcodeMatch =
    (isEmpty(po.barcode) && isEmpty(stock.barcode)) || po.barcode === stock.barcode;
  const qtyMatch = po.quantity === Math.abs(stock.quantity);
  return dateMatch && skuMatch && barcodeMatch && qtyMatch;
}

function performTallyMatching(records) {
  const poList = records.filter(r => !isEmpty(r.poNo) && r.quantity > 0);
  const stockList = records.filter(r => !isEmpty(r.company) && r.quantity < 0);
  const matchedStocks = new Set();

  for (const po of poList) {
    for (const stock of stockList) {
      if (matchedStocks.has(stock)) continue;
      if (isEmpty(po.supplier) || isEmpty(stock.companyCode)) continue;

      if (
        (po.supplier === 'OUT010' && stock.companyCode === 'OUT600') ||
        po.supplier.toUpperCase() === stock.companyCode.toUpperCase()
      ) {
        if (isExactMatch(po, stock)) {
          if (isEmpty(po.company)) {
            po.company = stock.company;
            po.companyMatched = true;
          }
          if (isEmpty(po.saId)) po.saId = stock.saId;
          po.remarks = 'Tally';
          stock.poNo = po.poNo;
          if (isEmpty(stock.supplier)) stock.supplier = po.supplier;
          if (isEmpty(stock.shop)) {
            stock.shop = po.shop;
            stock.shopMatched = true;
          }
          stock.remarks = 'Tally';
          matchedStocks.add(stock);
          break;
        }
      }
    }
  }
}

// ── Pass 2 – SKU + within-one-week ────────────────────────────────────────────

function performSecondPassMatching(records) {
  const unmatchedPO = records.filter(
    r => !isEmpty(r.poNo) && (isEmpty(r.remarks) || !r.remarks.startsWith('Tally'))
  );
  const unmatchedStock = records.filter(
    r => !isEmpty(r.company) && (isEmpty(r.remarks) || !r.remarks.startsWith('Tally'))
  );
  const matchedStocks = new Set();

  for (const po of unmatchedPO) {
    for (const stock of unmatchedStock) {
      if (matchedStocks.has(stock)) continue;
      if (isEmpty(po.supplier) || isEmpty(stock.companyCode)) continue;

      if (
        (po.supplier === 'OUT010' && stock.companyCode === 'OUT600') ||
        po.supplier.toUpperCase() === stock.companyCode.toUpperCase()
      ) {
        if (
          !isEmpty(po.sku) &&
          po.sku === stock.sku &&
          Math.abs(stock.quantity) === po.quantity &&
          isWithinOneWeek(po.date, stock.date)
        ) {
          if (isEmpty(po.company)) {
            po.company = stock.company;
            po.companyMatched = true;
          }
          if (isEmpty(po.saId)) po.saId = stock.saId;
          po.remarks = 'Tally (2nd pass)';
          stock.poNo = po.poNo;
          if (isEmpty(stock.shop)) {
            stock.shop = po.shop;
            stock.shopMatched = true;
          }
          stock.remarks = 'Tally (2nd pass)';
          matchedStocks.add(stock);
          break;
        }
      }
    }
  }
}

// ── Pass 3 – Stock self-cancel (same file, same SKU, same date, opposite qty) ─

function performThirdPassStockMatching(records) {
  const unmatchedStock = records.filter(
    r => !isEmpty(r.company) && (isEmpty(r.remarks) || !r.remarks.startsWith('Tally'))
  );
  const matchedStocks = new Set();

  for (let i = 0; i < unmatchedStock.length; i++) {
    const s1 = unmatchedStock[i];
    if (matchedStocks.has(s1)) continue;

    for (let j = i + 1; j < unmatchedStock.length; j++) {
      const s2 = unmatchedStock[j];
      if (matchedStocks.has(s2)) continue;

      if (s1.sourceFile !== s2.sourceFile) continue;

      if (
        !isEmpty(s1.sku) &&
        s1.sku === s2.sku &&
        s1.date === s2.date &&
        s1.quantity === -s2.quantity
      ) {
        s1.remarks = 'Tally (3rd pass)';
        s2.remarks = 'Tally (3rd pass)';
        matchedStocks.add(s1);
        matchedStocks.add(s2);
        break;
      }
    }
  }
}

// ── Update unmatched remarks ───────────────────────────────────────────────────

function updateRemarksForUnmatched(records) {
  for (const r of records) {
    if (isEmpty(r.remarks) || r.remarks.toLowerCase() === 'pending') {
      if (!isEmpty(r.poNo)) {
        r.remarks = 'Mismatch: no matching stock adjustment';
      } else if (!isEmpty(r.company)) {
        r.remarks = 'Mismatch: no matching purchase order';
      } else {
        r.remarks = 'Unmatched';
      }
    }
  }
}

// ── Tally record generation ────────────────────────────────────────────────────

function generateTallyRecords(poRecords, stockRecords) {
  const tally = [];

  for (const p of poRecords) {
    tally.push({
      poNo: p.purchaseOrderNo,
      company: '',
      companyCode: '',
      supplier: p.supplier,
      shop: p.shop,
      product: p.product,
      sku: p.sku,
      barcode: p.barcode,
      date: p.date,
      quantity: p.quantity,
      reason: '',
      idConflict: '',
      remarks: 'Pending',
      saId: '',
      sourceFile: '',
      companyMatched: false,
      shopMatched: false,
    });
  }

  for (const s of stockRecords) {
    tally.push({
      poNo: '',
      company: s.company,
      companyCode: s.companyCode,
      supplier: '',
      shop: '',
      product: '',
      sku: s.sku,
      barcode: s.barcode,
      date: s.date,
      quantity: s.adjustment,
      reason: s.reason,
      idConflict: '',
      remarks: 'Pending',
      saId: s.saId,
      sourceFile: s.sourceFile,
      companyMatched: false,
      shopMatched: false,
    });
  }

  detectIDConflicts(tally);
  performTallyMatching(tally);
  performSecondPassMatching(tally);
  performThirdPassStockMatching(tally);
  updateRemarksForUnmatched(tally);

  return tally;
}

// ── Excel output (ExcelJS) ─────────────────────────────────────────────────────

async function writeTallyReport(records) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Tally Report');

  const headers = [
    'PO No', 'Company', 'Supplier', 'Shop', 'In/Out',
    'Product', 'SKU', 'Barcode', 'Date', 'Quantity',
    'Reason', 'ID Conflict', 'Remarks', 'SA ID',
  ];
  sheet.addRow(headers);

  const greenFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF90EE90' }, // light green
  };

  const orangeFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFC878' }, // rgb(255,200,120)
  };

  for (const r of records) {
    const inOut = r.quantity > 0 ? 'IN' : r.quantity < 0 ? 'OUT' : '';

    const row = sheet.addRow([
      r.poNo,
      r.company,
      r.supplier,
      r.shop,
      inOut,
      r.product,
      r.sku,
      r.barcode,
      r.date,
      r.quantity,
      r.reason,
      r.idConflict,
      r.remarks,
      r.saId,
    ]);

    // Company cell (col 2) — green if matched
    if (r.companyMatched) {
      row.getCell(2).fill = greenFill;
    }

    // Shop cell (col 4) — orange if matched
    if (r.shopMatched) {
      row.getCell(4).fill = orangeFill;
    }
  }

  const arrayBuffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}

// ── Main exported function ─────────────────────────────────────────────────────

export async function generateReport(
  poFiles,
  saFiles,
  excludeSAIds = [],
) {
  const excludeSet = new Set(excludeSAIds);

  const allPORecords = [];
  for (const file of poFiles) {
    allPORecords.push(...readPurchaseOrderFile(file.name, file.buffer));
  }

  const allStockRecords = [];
  for (const file of saFiles) {
    const recs = readStockAdjustmentFile(file.name, file.buffer).filter(
      r => isEmpty(r.saId) || !excludeSet.has(r.saId)
    );
    allStockRecords.push(...recs);
  }

  const tallyRecords = generateTallyRecords(allPORecords, allStockRecords);
  return writeTallyReport(tallyRecords);
}
