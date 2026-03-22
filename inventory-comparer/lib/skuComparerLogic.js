import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { detectOgfRemarkFromSku, processLocationFiles } from './fileProcessor';

// ── Helpers ────────────────────────────────────────────────────────────────────

function isOgfName(name) {
  if (!name) return false;
  const n = name.toLowerCase();
  return n.startsWith('temp_sku_ogf') || n.includes('ogf');
}

function isPlaceholderValue(value) {
  if (!value || value.trim() === '') return true;
  const lower = value.trim().toLowerCase();
  return (
    lower === 'no barcode' ||
    lower === 'n/a' ||
    lower === 'na' ||
    lower === 'none' ||
    lower === 'null' ||
    lower === 'no barcode available' ||
    lower === 'missing barcode' ||
    (lower.startsWith('no ') && lower.includes('barcode'))
  );
}

function hasValidProductTitle(title) {
  if (!title) return false;
  const trimmed = title.trim();
  return (
    trimmed.length >= 2 &&
    trimmed.toLowerCase() !== 'null' &&
    trimmed.toLowerCase() !== 'n/a' &&
    trimmed.toLowerCase() !== 'na' &&
    trimmed.toLowerCase() !== 'none' &&
    trimmed.toLowerCase() !== 'default title'
  );
}

function getFormattedCellValue(sheet, row, col) {
  const addr = XLSX.utils.encode_cell({ r: row, c: col });
  const cell = sheet[addr];
  if (!cell) return '';
  if (cell.t === 'n') {
    const val = cell.v;
    return val === Math.floor(val) ? String(Math.trunc(val)) : String(val);
  }
  return cell.w !== undefined ? String(cell.w).trim() : String(cell.v ?? '').trim();
}

function stripExtension(name) {
  return name.replace(/\.xlsx$/i, '').replace(/\.xls$/i, '');
}

// ── Data Structures ────────────────────────────────────────────────────────────

function createItemSourceData(
  rawSku,
  rawBarcode,
  rawProductName,
  isDuplicate,
  hasShortBarcode,
  ogfRemark,
  isTempOgfFile
) {
  const rSku = rawSku?.trim() ?? '';
  const rBarcode = rawBarcode?.trim() ?? '';
  const rProductName = rawProductName?.trim() ?? '';
  const shortBc = hasShortBarcode || (rBarcode.length > 0 && rBarcode.length < 3);

  let remark;
  if (isTempOgfFile && (!ogfRemark || ogfRemark.trim() === '')) {
    remark = detectOgfRemarkFromSku(rawSku ?? '');
  } else {
    remark = ogfRemark?.trim() ?? '';
  }

  return {
    rawSku: rSku,
    rawBarcode: rBarcode,
    rawProductName: rProductName,
    cleanSku: rSku, // cleanSkuForComparison just trims
    isDuplicateInSource: isDuplicate,
    hasShortBarcode: shortBc,
    ogfRemark: remark,
    isSkuDuplicate: false,
    isBarcodeDuplicate: false,
  };
}

function createItem(sku, barcode) {
  return {
    primarySku: sku?.trim() ?? '',
    primaryBarcode: barcode?.trim() ?? '',
    consolidatedProductName: '',
    primarySkuSource: '',
    isCosmeticsGroupItem: false,
    isOgfGroupItem: false,
    allSkus: new Set(),
    allBarcodes: new Set(),
    sourceData: new Map(),
    finalRemarks: [],
    simpleStatus: '',
    conflictStatus: '',
  };
}

function itemKey(item) {
  return `${item.primarySku.toLowerCase()}||${item.primaryBarcode.toLowerCase()}`;
}

function addSourceData(item, sourceName, data) {
  if (!item.sourceData.has(sourceName)) {
    item.sourceData.set(sourceName, data);
  }
  if (data.rawSku && data.rawSku.trim() !== '') {
    const sourceSku = data.rawSku.trim();
    if (sourceSku.toLowerCase() === item.primarySku.toLowerCase()) {
      if (item.primarySkuSource === '') {
        item.primarySkuSource = sourceName;
      }
    }
  }
  if (data.cleanSku) item.allSkus.add(data.cleanSku);
  if (data.rawBarcode) item.allBarcodes.add(data.rawBarcode);
}

function markAsCosmeticsGroupItem(item, locationName) {
  const lower = locationName.toLowerCase();
  if (lower.includes('cosmetics') || lower.includes('cos')) {
    item.isCosmeticsGroupItem = true;
  }
}

function markAsOgfGroupItem(item, isOgfFile) {
  if (isOgfFile) item.isOgfGroupItem = true;
}

function getDataForLocation(item, locationName) {
  return item.sourceData.get(locationName);
}

function isPresentIn(item, locationName) {
  return item.sourceData.has(locationName);
}

function getAllSourceBarcodes(item) {
  const all = new Set();
  if (item.primaryBarcode && !isPlaceholderValue(item.primaryBarcode)) {
    all.add(item.primaryBarcode.trim().toLowerCase());
  }
  for (const sd of item.sourceData.values()) {
    if (sd.rawBarcode && !isPlaceholderValue(sd.rawBarcode)) {
      all.add(sd.rawBarcode.trim().toLowerCase());
    }
  }
  return all;
}

// ── Read Items from a single sheet ─────────────────────────────────────────────

function readItems(buffer, fileName, isTempOgfFile, skipInternalValidation) {
  const uniqueItems = new Map();
  const duplicateSourceData = [];

  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return { uniqueItems, duplicateSourceData };

  const range = XLSX.utils.decode_range(sheet['!ref'] ?? 'A1');
  const lastRow = range.e.r;
  const lastCol = range.e.c;

  // Find header columns
  let skuCol = -1, barcodeCol = -1, nameCol = -1, remarkCol = -1;
  for (let c = 0; c <= lastCol; c++) {
    const val = getFormattedCellValue(sheet, 0, c).toLowerCase();
    if (val.includes('sku')) skuCol = c;
    if (val.includes('barcode')) barcodeCol = c;
    if (val === 'product' || val.includes('title')) nameCol = c;
    if (val.includes('remark')) remarkCol = c;
  }

  if (skuCol === -1 && barcodeCol === -1) {
    throw new Error(`Could not find SKU or Barcode columns in file: ${fileName}`);
  }

  // First pass: collect all SKUs and barcodes with row numbers for duplicate detection
  const skuRowMap = new Map();
  const barcodeRowMap = new Map();

  for (let i = 1; i <= lastRow; i++) {
    const rawSku = skuCol >= 0 ? getFormattedCellValue(sheet, i, skuCol) : '';
    const rawBarcode = barcodeCol >= 0 ? getFormattedCellValue(sheet, i, barcodeCol) : '';

    if (rawSku && !isPlaceholderValue(rawSku)) {
      const lower = rawSku.toLowerCase();
      if (!skuRowMap.has(lower)) skuRowMap.set(lower, []);
      skuRowMap.get(lower).push(i + 1);
    }
    if (rawBarcode && !isPlaceholderValue(rawBarcode)) {
      const lower = rawBarcode.toLowerCase();
      if (!barcodeRowMap.has(lower)) barcodeRowMap.set(lower, []);
      barcodeRowMap.get(lower).push(i + 1);
    }
  }

  // Identify duplicates (values appearing more than once)
  const duplicateSkus = new Set();
  for (const [key, rows] of skuRowMap) {
    if (rows.length > 1) duplicateSkus.add(key);
  }
  const duplicateBarcodes = new Set();
  for (const [key, rows] of barcodeRowMap) {
    if (rows.length > 1) duplicateBarcodes.add(key);
  }

  // Second pass: process all rows
  for (let i = 1; i <= lastRow; i++) {
    const originalRawSku = skuCol >= 0 ? getFormattedCellValue(sheet, i, skuCol) : '';
    const rawSku = originalRawSku;
    const rawBarcode = barcodeCol >= 0 ? getFormattedCellValue(sheet, i, barcodeCol) : '';
    const rawProductName = nameCol >= 0 ? getFormattedCellValue(sheet, i, nameCol) : '';
    let ogfRemark = remarkCol >= 0 ? getFormattedCellValue(sheet, i, remarkCol) : '';

    if (isTempOgfFile && (!ogfRemark || ogfRemark.trim() === '')) {
      ogfRemark = detectOgfRemarkFromSku(originalRawSku);
    }

    if (rawSku || rawBarcode) {
      const lowerSku = rawSku.toLowerCase();
      const lowerBarcode = rawBarcode.toLowerCase();
      let isDuplicate = false;
      let isShortBarcode = false;
      let isSkuDup = false;
      let isBarcodeDup = false;

      if (!skipInternalValidation) {
        isShortBarcode = rawBarcode.length > 0 && rawBarcode.trim().length < 3 && !isPlaceholderValue(rawBarcode);

        if (rawSku && !isPlaceholderValue(rawSku) && duplicateSkus.has(lowerSku)) {
          isDuplicate = true;
          isSkuDup = true;
        }
        if (rawBarcode && !isPlaceholderValue(rawBarcode) && duplicateBarcodes.has(lowerBarcode)) {
          isDuplicate = true;
          isBarcodeDup = true;
        }
      }

      const tempSourceData = createItemSourceData(rawSku, rawBarcode, rawProductName, isDuplicate, isShortBarcode, ogfRemark, isTempOgfFile);
      tempSourceData.isSkuDuplicate = isSkuDup;
      tempSourceData.isBarcodeDuplicate = isBarcodeDup;

      if (isDuplicate) {
        const dupData = createItemSourceData(rawSku, rawBarcode, rawProductName, true, isShortBarcode, ogfRemark, isTempOgfFile);
        duplicateSourceData.push(dupData);
      }

      const newItem = createItem(tempSourceData.cleanSku, tempSourceData.rawBarcode);
      addSourceData(newItem, 'TEMP_KEY', tempSourceData);
      const key = itemKey(newItem);
      if (!uniqueItems.has(key)) {
        uniqueItems.set(key, newItem);
      }
    }
  }

  return { uniqueItems, duplicateSourceData };
}

// ── Detection passes ───────────────────────────────────────────────────────────

function detectSkuBarcodeMismatches(allItems) {
  const barcodeToItems = new Map();

  for (const item of allItems) {
    const itemBarcodes = getAllSourceBarcodes(item);
    for (const barcode of itemBarcodes) {
      if (!barcodeToItems.has(barcode)) barcodeToItems.set(barcode, []);
      const list = barcodeToItems.get(barcode);
      if (!list.includes(item)) list.push(item);
    }
  }

  for (const [barcode, duplicateItems] of barcodeToItems) {
    if (duplicateItems.length <= 1) continue;

    const uniqueSkus = new Set(
      duplicateItems.map((it) => it.primarySku.trim().toLowerCase()).filter((s) => s !== '')
    );

    if (uniqueSkus.size > 1) {
      for (const item of duplicateItems) {
        if (!item.conflictStatus.includes('DUPLICATE_BARCODE_ACROSS_SKUS')) {
          item.conflictStatus = item.conflictStatus
            ? item.conflictStatus + ' + DUPLICATE_BARCODE_ACROSS_SKUS'
            : 'DUPLICATE_BARCODE_ACROSS_SKUS';
        }
        const otherSkus = duplicateItems
          .filter((o) => o.primarySku !== item.primarySku)
          .map((o) => o.primarySku);

        const alreadyHasRemark = item.finalRemarks.some(
          (r) => r.includes(`Barcode ${barcode} shared with other SKU`)
        );
        if (!alreadyHasRemark) {
          item.finalRemarks.push(`🚫 CRITICAL: Barcode ${barcode} shared with other SKU(s): ${otherSkus.join(', ')}`);
        }
      }
    }
  }
}

function detectCrossItemBarcodeDuplicates(allItems) {
  const barcodeToItems = new Map();

  for (const item of allItems) {
    const itemBarcodes = getAllSourceBarcodes(item);
    for (const barcode of itemBarcodes) {
      if (!barcodeToItems.has(barcode)) barcodeToItems.set(barcode, []);
      const list = barcodeToItems.get(barcode);
      if (!list.includes(item)) list.push(item);
    }
  }

  for (const [barcode, duplicateItems] of barcodeToItems) {
    if (duplicateItems.length <= 1) continue;

    for (const item of duplicateItems) {
      if (!item.conflictStatus.includes('DUPLICATE_BARCODE_ACROSS_ITEMS')) {
        item.conflictStatus = item.conflictStatus
          ? item.conflictStatus + ' + DUPLICATE_BARCODE_ACROSS_ITEMS'
          : 'DUPLICATE_BARCODE_ACROSS_ITEMS';
      }
      const otherSkus = duplicateItems
        .filter((o) => o.primarySku !== item.primarySku)
        .map((o) => o.primarySku)
        .join(', ');
      item.finalRemarks.push(`🚫 Barcode ${barcode} shared with other SKUs: ${otherSkus}`);
    }
  }
}

function detectInternalInconsistencies(item) {
  const skusInThisItem = new Set();
  const barcodesInThisItem = new Set();
  const withinFileDuplicateReasons = [];

  for (const [sourceName, data] of item.sourceData) {
    if (data.rawSku.trim()) skusInThisItem.add(data.rawSku.trim());
    if (data.rawBarcode.trim()) barcodesInThisItem.add(data.rawBarcode.trim());

    if (data.isDuplicateInSource) {
      let reason = `Duplicate in '${sourceName}'`;
      if (data.isSkuDuplicate && data.isBarcodeDuplicate) {
        reason += ` - SKU '${data.rawSku}' and Barcode '${data.rawBarcode}' appear multiple times in this file`;
      } else if (data.isSkuDuplicate) {
        reason += ` - SKU '${data.rawSku}' appears multiple times in this file`;
      } else if (data.isBarcodeDuplicate) {
        reason += ` - Barcode '${data.rawBarcode}' appears multiple times in this file`;
      }
      withinFileDuplicateReasons.push(reason);
    }
  }

  if (skusInThisItem.size > 1) {
    if (!item.conflictStatus.includes('INCONSISTENT_SKU')) {
      item.conflictStatus = item.conflictStatus ? item.conflictStatus + ' + INCONSISTENT_SKU' : 'INCONSISTENT_SKU';
    }
    item.finalRemarks.push(`Different SKUs for same item across files: ${[...skusInThisItem].join(' vs ')}`);
  }

  if (barcodesInThisItem.size > 1) {
    if (!item.conflictStatus.includes('INCONSISTENT_BARCODE')) {
      item.conflictStatus = item.conflictStatus ? item.conflictStatus + ' + INCONSISTENT_BARCODE' : 'INCONSISTENT_BARCODE';
    }
    item.finalRemarks.push(`Different barcodes for same item across files: ${[...barcodesInThisItem].join(' vs ')}`);
  }

  if (withinFileDuplicateReasons.length > 0) {
    if (!item.conflictStatus.includes('FILE_DUPLICATE')) {
      item.conflictStatus = item.conflictStatus ? item.conflictStatus + ' + FILE_DUPLICATE' : 'FILE_DUPLICATE';
    }
    item.finalRemarks.push(...withinFileDuplicateReasons);
  }
}

function detectCrossFileDifferences(item) {
  const allSkus = new Set();
  const allBarcodes = new Set();
  const skuToSource = new Map();
  const barcodeToSource = new Map();

  for (const [source, data] of item.sourceData) {
    if (data.rawSku.trim()) {
      allSkus.add(data.rawSku.trim());
      skuToSource.set(data.rawSku.trim(), source);
    }
    if (data.rawBarcode.trim()) {
      allBarcodes.add(data.rawBarcode.trim());
      barcodeToSource.set(data.rawBarcode.trim(), source);
    }
  }

  if (allSkus.size > 1) {
    if (!item.conflictStatus.includes('INCONSISTENT_SKU')) {
      item.conflictStatus = item.conflictStatus ? item.conflictStatus + ' + INCONSISTENT_SKU' : 'INCONSISTENT_SKU';
    }
    const details = [...allSkus].map((s) => `${s}(${skuToSource.get(s)})`);
    item.finalRemarks.push(`Different SKUs across files: ${details.join(' vs ')}`);
  }

  if (allBarcodes.size > 1) {
    if (!item.conflictStatus.includes('INCONSISTENT_BARCODE')) {
      item.conflictStatus = item.conflictStatus ? item.conflictStatus + ' + INCONSISTENT_BARCODE' : 'INCONSISTENT_BARCODE';
    }
    const details = [...allBarcodes].map((b) => `${b}(${barcodeToSource.get(b)})`);
    item.finalRemarks.push(`Different barcodes across files: ${details.join(' vs ')}`);
  }

  if (item.primarySku && allSkus.size > 1 && !allSkus.has(item.primarySku)) {
    item.finalRemarks.push(`Primary SKU '${item.primarySku}' doesn't match other files`);
  }
  if (item.primaryBarcode && allBarcodes.size > 1 && !allBarcodes.has(item.primaryBarcode)) {
    item.finalRemarks.push(`Primary barcode '${item.primaryBarcode}' doesn't match other files`);
  }
}

function detectShortBarcodes(item) {
  const shortBarcodeSources = [];
  let hasShortBarcode = false;

  for (const [source, data] of item.sourceData) {
    if (data.rawBarcode.trim() && data.rawBarcode.trim().length < 3) {
      shortBarcodeSources.push(`${source}('${data.rawBarcode.trim()}')`);
      hasShortBarcode = true;
    }
  }

  if (item.primaryBarcode && item.primaryBarcode.length < 3) {
    const alreadyDetected = shortBarcodeSources.some((s) => s.includes(`'${item.primaryBarcode}'`));
    if (!alreadyDetected) {
      shortBarcodeSources.push(`Primary('${item.primaryBarcode}')`);
      hasShortBarcode = true;
    }
  }

  if (hasShortBarcode) {
    if (!item.conflictStatus.includes('SHORT_BARCODE')) {
      item.conflictStatus = item.conflictStatus ? item.conflictStatus + ' + SHORT_BARCODE' : 'SHORT_BARCODE';
    }
    item.finalRemarks.push(`Short barcodes (<3 chars) in: ${shortBarcodeSources.join(', ')}`);
  }
}

// ── "In ANY UNLISTED?" logic ───────────────────────────────────────────────────

function isItemInAnyRelevantUnlisted(
  item,
  locationNames,
  unlistedNames,
  cosmeticLocationNames,
  useOgfRules
) {
  if (useOgfRules) {
    let presentLocations;
    if (item.isOgfGroupItem) {
      presentLocations = locationNames.filter((n) => isOgfName(n) && isPresentIn(item, n));
    } else {
      presentLocations = locationNames.filter((n) => !isOgfName(n) && isPresentIn(item, n));
    }
    if (item.isOgfGroupItem && presentLocations.length > 0) {
      return unlistedNames.filter(isOgfName).some((n) => isPresentIn(item, n));
    } else if (!item.isOgfGroupItem && presentLocations.length > 0) {
      return unlistedNames.filter((n) => !isOgfName(n)).some((n) => isPresentIn(item, n));
    }
    return false;
  } else {
    // Cosmetics rules
    const isCosName = (n) =>
      cosmeticLocationNames.has(n) || n.toUpperCase().includes('COSMETIC') || n.toUpperCase().includes('COS');

    let presentLocations;
    if (item.isCosmeticsGroupItem) {
      presentLocations = locationNames.filter((n) => isCosName(n) && isPresentIn(item, n));
    } else {
      presentLocations = locationNames.filter((n) => !isCosName(n) && isPresentIn(item, n));
    }

    if (item.isCosmeticsGroupItem && presentLocations.length > 0) {
      return unlistedNames.filter((n) => n.toUpperCase().includes('WEB')).some((n) => isPresentIn(item, n));
    } else if (!item.isCosmeticsGroupItem && presentLocations.length > 0) {
      return unlistedNames.filter((n) => !n.toUpperCase().includes('WEB')).some((n) => isPresentIn(item, n));
    }
    return false;
  }
}

// ── Final Remarks Generation ───────────────────────────────────────────────────

function generateFinalRemarksWithFilteredUnlisted(
  item,
  locationNames,
  unlistedNames,
  cosmeticLocationNames,
  useOgfRules
) {
  // Preserve existing duplicate barcode status
  const hadDuplicateBarcodeBefore = item.conflictStatus.includes('DUPLICATE_BARCODE_ACROSS_SKUS');
  const hadDuplicateBarcodeAcrossItemsBefore = item.conflictStatus.includes('DUPLICATE_BARCODE_ACROSS_ITEMS');
  const existingDuplicateRemarks = item.finalRemarks.filter(
    (r) => r.includes('CRITICAL: Barcode') && r.includes('shared with other SKU')
  );

  // Clear but preserve duplicate barcode info
  item.finalRemarks = [];
  item.conflictStatus = '';

  if (hadDuplicateBarcodeBefore) {
    item.conflictStatus = 'DUPLICATE_BARCODE_ACROSS_SKUS';
    item.finalRemarks.push(...existingDuplicateRemarks);
  }
  if (hadDuplicateBarcodeAcrossItemsBefore) {
    item.conflictStatus = item.conflictStatus
      ? item.conflictStatus + ' + DUPLICATE_BARCODE_ACROSS_ITEMS'
      : 'DUPLICATE_BARCODE_ACROSS_ITEMS';
  }

  // Data quality checks
  detectInternalInconsistencies(item);
  detectShortBarcodes(item);
  detectCrossFileDifferences(item);
  const hasDataIssues = item.conflictStatus !== '';

  // Determine presence
  const presentLocations = new Set(locationNames.filter((loc) => isPresentIn(item, loc)));
  const presentUnlisted = new Set(unlistedNames.filter((unl) => isPresentIn(item, unl)));
  const missingLocations = new Set(locationNames.filter((loc) => !presentLocations.has(loc)));

  let isBad = false;
  const badReasons = [];
  const ruleViolations = [];

  if (useOgfRules) {
    // ── OGF RULES ──
    const ogfLocationFile = locationNames.find(isOgfName) ?? null;
    const ogfUnlistedFile = unlistedNames.find(isOgfName) ?? null;
    const nonOgfUnlisted = new Set(unlistedNames.filter((u) => ogfUnlistedFile == null || u !== ogfUnlistedFile));

    const inOgfLoc = ogfLocationFile != null && presentLocations.has(ogfLocationFile);
    const inOgfUnl = ogfUnlistedFile != null && presentUnlisted.has(ogfUnlistedFile);
    const inAnyNonOgfUnlisted = [...nonOgfUnlisted].some((u) => presentUnlisted.has(u));

    if (inOgfLoc && inOgfUnl) {
      ruleViolations.push(`OGF item should not appear in both ${ogfLocationFile} and ${ogfUnlistedFile}`);
    }
    if (inAnyNonOgfUnlisted && presentLocations.size > 0) {
      ruleViolations.push('Non-OGF unlisted item should not appear in any location files');
    }

    if (presentLocations.size > 0 && missingLocations.size > 0) {
      const unjustifiedMissing = [];
      for (const missingLoc of missingLocations) {
        if (isOgfName(missingLoc)) {
          if (!inOgfUnl) unjustifiedMissing.push(missingLoc);
        } else {
          if (!inAnyNonOgfUnlisted) unjustifiedMissing.push(missingLoc);
        }
      }
      if (unjustifiedMissing.length > 0) {
        ruleViolations.push(`Item missing from locations: ${unjustifiedMissing.join(', ')}`);
      }
    }

    if (item.isOgfGroupItem) {
      if (!inOgfLoc && !inOgfUnl) {
        ruleViolations.push('OGF item missing from both OGF location and OGF unlisted');
      }
    } else {
      if (presentLocations.size === 0 && !inAnyNonOgfUnlisted) {
        ruleViolations.push('Non-OGF item missing from all locations and non-OGF unlisted files');
      }
    }
  } else {
    // ── COSMETICS RULES ──
    const isCosName = (n) =>
      cosmeticLocationNames.has(n) || n.toUpperCase().includes('COSMETIC') || n.toUpperCase().includes('COS');

    const cosmeticsLocationFile = locationNames.find(isCosName) ?? null;
    const webUnlistedFile = unlistedNames.find((n) => n.toUpperCase().includes('WEB')) ?? null;
    const nonWebUnlisted = new Set(unlistedNames.filter((u) => webUnlistedFile == null || u !== webUnlistedFile));

    const inCosLoc = cosmeticsLocationFile != null && presentLocations.has(cosmeticsLocationFile);
    const inWebUnl = webUnlistedFile != null && presentUnlisted.has(webUnlistedFile);
    const inAnyNonWebUnlisted = [...nonWebUnlisted].some((u) => presentUnlisted.has(u));

    if (inCosLoc && inWebUnl) {
      ruleViolations.push(`Cosmetics item should not appear in both ${cosmeticsLocationFile} and ${webUnlistedFile}`);
    }
    if (inAnyNonWebUnlisted) {
      const nonCosmeticsPresentLocations = [...presentLocations].filter((loc) => !isCosName(loc));
      if (nonCosmeticsPresentLocations.length > 0) {
        ruleViolations.push(
          `Non-WEB unlisted item should not appear in non-cosmetics locations: ${nonCosmeticsPresentLocations.join(', ')}`
        );
      }
    }

    if (presentLocations.size > 0 && missingLocations.size > 0) {
      const unjustifiedMissing = [];
      for (const missingLoc of missingLocations) {
        if (isCosName(missingLoc)) {
          if (!inWebUnl) unjustifiedMissing.push(missingLoc);
        } else {
          if (!inAnyNonWebUnlisted) unjustifiedMissing.push(missingLoc);
        }
      }
      if (unjustifiedMissing.length > 0) {
        ruleViolations.push(`Item missing from locations: ${unjustifiedMissing.join(', ')}`);
      }
    }

    if (item.isCosmeticsGroupItem) {
      if (!inCosLoc && !inWebUnl) {
        ruleViolations.push('Cosmetics item missing from both cosmetics location and WEB unlisted');
      }
    } else {
      if (presentLocations.size === 0 && presentUnlisted.size === 0) {
        ruleViolations.push('Non-cosmetics item missing from all locations and all unlisted files');
      }
    }
  }

  // Mark as bad if rule violations
  if (ruleViolations.length > 0) {
    isBad = true;
    badReasons.push(...ruleViolations);
  }

  // Determine final status
  if (presentLocations.size === 0 && presentUnlisted.size === 0) {
    item.simpleStatus = 'No Data Found - BAD';
    item.finalRemarks.push('Item not found in any location or unlisted files');
  } else if (isBad || hasDataIssues) {
    const parts = [];
    if (ruleViolations.length > 0 && hasDataIssues) {
      parts.push('Rule Violation + DATA ISSUES');
    } else if (ruleViolations.length > 0) {
      parts.push('Rule Violation');
    } else if (hasDataIssues) {
      parts.push('DATA ISSUES');
    }
    if (item.conflictStatus.includes('DUPLICATE_BARCODE_ACROSS_SKUS')) {
      parts.push('CRITICAL: Duplicate Barcode');
    }
    parts.push('- BAD');
    item.simpleStatus = parts.join(' + ');

    badReasons.forEach((r) => item.finalRemarks.push(`🚫 RULE VIOLATION: ${r}`));
    if (hasDataIssues) {
      item.finalRemarks.push(`⚠️ DATA ISSUE: ${item.conflictStatus}`);
    }
  } else {
    item.simpleStatus = 'GOOD';
    if (presentLocations.size > 0 && presentUnlisted.size === 0) {
      item.finalRemarks.push('✅ Item correctly placed in all locations and not in any unlisted files');
    } else if (presentLocations.size === 0 && presentUnlisted.size > 0) {
      item.finalRemarks.push(`✅ Item correctly only in unlisted files: ${[...presentUnlisted].join(', ')}`);
    } else {
      item.finalRemarks.push('✅ Item follows all location/unlisted pairing rules');
    }
  }

  // Presence info
  if (presentLocations.size > 0) {
    item.finalRemarks.push(`📋 Present in locations: ${[...presentLocations].join(', ')}`);
  }
  if (presentUnlisted.size > 0) {
    item.finalRemarks.push(`📋 Present in unlisted: ${[...presentUnlisted].join(', ')}`);
  }
  if (missingLocations.size > 0 && presentLocations.size > 0) {
    item.finalRemarks.push(`📋 Missing from locations: ${[...missingLocations].join(', ')}`);
  }
}

// ── Main Export ────────────────────────────────────────────────────────────────

export async function generateSkuReport(
  locationFiles,
  unlistedFiles,
  ogfRulesChecked
) {
  if (!locationFiles || locationFiles.length === 0) {
    throw new Error('Must provide at least one location file for comparison.');
  }

  // Pre-process location files (OGF cleanup)
  const processedLocationFiles = processLocationFiles(locationFiles, ogfRulesChecked);

  const locationNames = processedLocationFiles.map((f) => stripExtension(f.name));
  const unlistedNames = (unlistedFiles ?? []).map((f) => stripExtension(f.name));

  const consolidatedItemsBySku = new Map();
  const consolidatedItemsWithNoSku = [];

  const cosmeticLocationNames = new Set(
    locationNames.filter((n) => n.toLowerCase().includes('cosmetics') || n.toLowerCase().includes('cos'))
  );

  // ── Read location files and merge ──
  for (let fi = 0; fi < processedLocationFiles.length; fi++) {
    const file = processedLocationFiles[fi];
    const fileName = stripExtension(file.name);
    const isTempOgfFile = isOgfName(fileName);

    const { uniqueItems, duplicateSourceData } = readItems(file.buffer, file.name, isTempOgfFile, false);

    const duplicatesKeySet = new Set(duplicateSourceData.map((d) => `${d.cleanSku.toLowerCase()}||${d.rawBarcode.toLowerCase()}`));

    for (const newItem of uniqueItems.values()) {
      const currentData = getDataForLocation(newItem, 'TEMP_KEY');
      if (!currentData) continue;
      const currentSku = currentData.cleanSku;
      const isDuplicateInSource = duplicatesKeySet.has(itemKey(newItem));
      const isShortBarcode = currentData.hasShortBarcode;

      const finalData = createItemSourceData(
        currentData.rawSku, currentData.rawBarcode, currentData.rawProductName,
        isDuplicateInSource, isShortBarcode, currentData.ogfRemark, isTempOgfFile
      );

      let existingItem;
      if (currentSku) {
        existingItem = consolidatedItemsBySku.get(currentSku.toLowerCase());
      }

      if (!existingItem) {
        let itemToUse = null;
        if (currentSku) {
          itemToUse = createItem(currentSku, currentData.rawBarcode);
          addSourceData(itemToUse, fileName, finalData);
          consolidatedItemsBySku.set(currentSku.toLowerCase(), itemToUse);
        } else if (currentData.rawBarcode) {
          let merged = false;
          for (const noSkuItem of consolidatedItemsWithNoSku) {
            if (noSkuItem.primaryBarcode.toLowerCase() === currentData.rawBarcode.toLowerCase()) {
              addSourceData(noSkuItem, fileName, finalData);
              merged = true;
              break;
            }
          }
          if (!merged) {
            itemToUse = createItem('', currentData.rawBarcode);
            addSourceData(itemToUse, fileName, finalData);
            consolidatedItemsWithNoSku.push(itemToUse);
          }
        } else {
          continue;
        }
        if (itemToUse) {
          markAsOgfGroupItem(itemToUse, isTempOgfFile);
          markAsCosmeticsGroupItem(itemToUse, fileName);
        }
      } else {
        addSourceData(existingItem, fileName, finalData);
        markAsOgfGroupItem(existingItem, isTempOgfFile);
        markAsCosmeticsGroupItem(existingItem, fileName);
      }
    }
  }

  // ── Read unlisted files ──
  if (unlistedFiles && unlistedFiles.length > 0) {
    for (const file of unlistedFiles) {
      const fileName = stripExtension(file.name);
      const isTempOgfFile = fileName.toUpperCase().replace(/[^A-Z0-9]/g, '').includes('OGF');

      const { uniqueItems } = readItems(file.buffer, file.name, isTempOgfFile, true);

      for (const newItem of uniqueItems.values()) {
        const currentData = getDataForLocation(newItem, 'TEMP_KEY');
        if (!currentData) continue;
        const currentSku = currentData.cleanSku;

        const finalData = createItemSourceData(
          currentData.rawSku, currentData.rawBarcode, currentData.rawProductName,
          false, false, currentData.ogfRemark, isTempOgfFile
        );

        let existingItem;
        if (currentSku) {
          existingItem = consolidatedItemsBySku.get(currentSku.toLowerCase());
        }

        if (!existingItem) {
          if (currentSku) {
            const itemToUse = createItem(currentSku, currentData.rawBarcode);
            itemToUse.primarySkuSource = fileName;
            addSourceData(itemToUse, fileName, finalData);
            consolidatedItemsBySku.set(currentSku.toLowerCase(), itemToUse);
          } else if (currentData.rawBarcode) {
            let merged = false;
            for (const noSkuItem of consolidatedItemsWithNoSku) {
              if (noSkuItem.primaryBarcode.toLowerCase() === currentData.rawBarcode.toLowerCase()) {
                addSourceData(noSkuItem, fileName, finalData);
                merged = true;
                break;
              }
            }
            if (!merged) {
              const itemToUse = createItem('', currentData.rawBarcode);
              addSourceData(itemToUse, fileName, finalData);
              consolidatedItemsWithNoSku.push(itemToUse);
            }
          }
        } else {
          addSourceData(existingItem, fileName, finalData);
          if (existingItem.primarySkuSource === '' && currentSku) {
            existingItem.primarySkuSource = fileName;
          }
        }
      }
    }
  }

  // ── Write comparison report ──
  return writeComparisonReport(
    consolidatedItemsBySku,
    consolidatedItemsWithNoSku,
    locationNames,
    unlistedNames,
    cosmeticLocationNames,
    ogfRulesChecked
  );
}

// ── Write Report with ExcelJS ──────────────────────────────────────────────────

async function writeComparisonReport(
  consolidatedItemsBySku,
  consolidatedItemsWithNoSku,
  locationNames,
  unlistedNames,
  cosmeticLocationNames,
  useOgfRules
) {
  const allConsolidatedItems = [...consolidatedItemsBySku.values(), ...consolidatedItemsWithNoSku];

  // Detection passes
  detectCrossItemBarcodeDuplicates(allConsolidatedItems);
  detectSkuBarcodeMismatches(allConsolidatedItems);

  // ── Product title resolution (5 strategies) ──
  for (const item of allConsolidatedItems) {
    item.consolidatedProductName = '';

    // STRATEGY 1: Product title from primarySkuSource
    if (item.primarySkuSource) {
      const sourceData = getDataForLocation(item, item.primarySkuSource);
      if (sourceData && hasValidProductTitle(sourceData.rawProductName)) {
        item.consolidatedProductName = sourceData.rawProductName.trim();
      }
    }

    // STRATEGY 2: Discover source with matching SKU
    if (!item.consolidatedProductName) {
      for (const [sourceName, data] of item.sourceData) {
        if (data.rawSku && data.rawSku.trim().toLowerCase() === item.primarySku.toLowerCase()) {
          if (hasValidProductTitle(data.rawProductName)) {
            item.consolidatedProductName = data.rawProductName.trim();
            break;
          }
        }
      }
    }

    // STRATEGY 3: Check all location files
    if (!item.consolidatedProductName) {
      for (const location of locationNames) {
        const locationData = getDataForLocation(item, location);
        if (locationData && hasValidProductTitle(locationData.rawProductName)) {
          item.consolidatedProductName = locationData.rawProductName.trim();
          break;
        }
      }
    }

    // STRATEGY 4: Check all unlisted files for best title
    if (!item.consolidatedProductName) {
      let bestTitle = '';
      for (const unlisted of unlistedNames) {
        for (const [key, data] of item.sourceData) {
          if (key.toLowerCase() === unlisted.toLowerCase() || key.toLowerCase().includes(unlisted.toLowerCase())) {
            if (hasValidProductTitle(data.rawProductName)) {
              const current = data.rawProductName.trim();
              if (current.length > bestTitle.length) {
                bestTitle = current;
              }
            }
          }
        }
      }
      if (bestTitle) {
        item.consolidatedProductName = bestTitle;
      }
    }

    // STRATEGY 5: Default Title fallback
    if (!item.consolidatedProductName) {
      item.consolidatedProductName = 'Default Title';
    }
  }

  // Sort items
  allConsolidatedItems.sort((a, b) => {
    const skuCmp = a.primarySku.toLowerCase().localeCompare(b.primarySku.toLowerCase());
    if (skuCmp !== 0) return skuCmp;
    return a.primaryBarcode.toLowerCase().localeCompare(b.primaryBarcode.toLowerCase());
  });

  // ── Build display names ──
  const locationDisplayNames = new Map();
  for (const name of locationNames) {
    locationDisplayNames.set(name, isOgfName(name) ? 'OGF Location' : name);
  }

  const unlistedDisplayNames = new Map();
  for (const name of unlistedNames) {
    unlistedDisplayNames.set(name, isOgfName(name) ? 'OGF Unlisted' : name);
  }

  // ── Build workbook ──
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Inventory Comparison Report');

  // Header style
  const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC0C0C0' } };
  const headerFont = { bold: true };
  const thinBorder = {
    top: { style: 'thin' }, bottom: { style: 'thin' },
    left: { style: 'thin' }, right: { style: 'thin' },
  };

  // Build header columns
  const columns = [
    'Primary SKU (Consolidated)',
    'Primary Barcode (Consolidated)',
    'Product Name',
  ];

  for (const unlistedName of unlistedNames) {
    const display = unlistedDisplayNames.get(unlistedName) ?? unlistedName;
    columns.push(`SKU (${display})`);
    columns.push(`Barcode (${display})`);
  }

  for (const location of locationNames) {
    const display = locationDisplayNames.get(location) ?? location;
    columns.push(`SKU (${display})`);
    columns.push(`Barcode (${display})`);
    columns.push(`OGF Remark (${display})`);
  }

  columns.push('In ALL Locations?', 'In ANY UNLISTED?', 'Simple Status', 'ID / Data Problem', 'CONSOLIDATED REMARKS');

  // Write header row
  const headerRow = sheet.addRow(columns);
  headerRow.eachCell((cell) => {
    cell.fill = headerFill;
    cell.font = headerFont;
    cell.border = thinBorder;
    cell.alignment = { wrapText: true };
  });

  // ── Data rows ──
  for (const item of allConsolidatedItems) {
    generateFinalRemarksWithFilteredUnlisted(item, locationNames, unlistedNames, cosmeticLocationNames, useOgfRules);

    const rowData = [
      item.primarySku,
      item.primaryBarcode,
      item.consolidatedProductName,
    ];

    // Unlisted columns
    for (const unlistedName of unlistedNames) {
      const usd = getDataForLocation(item, unlistedName);
      rowData.push(usd?.rawSku ?? '');
      rowData.push(usd?.rawBarcode ?? '');
    }

    // Location columns
    let presentCount = 0;
    for (const location of locationNames) {
      const sd = getDataForLocation(item, location);
      if (sd) {
        rowData.push(sd.rawSku);
        rowData.push(sd.rawBarcode);
        rowData.push(sd.ogfRemark);
        presentCount++;
      } else {
        rowData.push('', '', '');
      }
    }

    const presentInAll = presentCount === locationNames.length && locationNames.length > 0;
    rowData.push(presentInAll ? 'YES' : 'NO');

    const statusInAnyRelevantUnlisted = isItemInAnyRelevantUnlisted(item, locationNames, unlistedNames, cosmeticLocationNames, useOgfRules);
    rowData.push(statusInAnyRelevantUnlisted ? 'YES' : 'NO');

    rowData.push(item.simpleStatus);
    rowData.push(item.conflictStatus);
    rowData.push(item.finalRemarks.join(' | '));

    sheet.addRow(rowData);
  }

  // Auto-size columns (approximate)
  sheet.columns.forEach((col) => {
    let maxLen = 10;
    col.eachCell?.({ includeEmpty: false }, (cell) => {
      const len = String(cell.value ?? '').length;
      if (len > maxLen) maxLen = len;
    });
    col.width = Math.min(maxLen + 2, 60);
  });

  // Write to buffer
  const arrayBuffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}
