import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { processLocationFiles, cleanupSkuForPriceComparison } from './fileProcessor';

// ── Header names ───────────────────────────────────────────────────────────────

const SKU_HEADER = 'SKU';
const NAME_HEADER = 'Product Name';
const PRICE_HEADER = 'Price';
const COMPARED_PRICE_HEADER = 'Compare at price';
const AVAILABLE_HEADER = 'Available';

// ── Data structures ────────────────────────────────────────────────────────────

function createReferenceItem(
  sku,
  productName,
  referencePrice,
  referenceCompareAtPrice,
  locationPrices
) {
  return {
    sku,
    productName,
    referencePrice,
    referenceCompareAtPrice,
    discrepancies: [],
    basicRemarks: [],
    locationPrices,
    locationCompareAtPrices: new Map(),
    compareAtPriceDifferences: new Map(),
    locationPricesUsed: new Map(),
    locationStock: new Map(),
    status: '',
    ogfPercentageDiff: 0,
    ogfPercentageRemark: '',
    statusReason: '',
    differenceExplanation: '',
    totalStock: 0,
    ogfDiscrepancies: [],
    nonOgfDiscrepancies: [],
    locationDiscountPercentages: new Map(),
    discountDiscrepancies: [],
  };
}

// ── Cell value helpers ─────────────────────────────────────────────────────────

function getCellValue(sheet, row, col) {
  const addr = XLSX.utils.encode_cell({ r: row, c: col });
  const cell = sheet[addr];
  if (!cell) return '';
  if (cell.t === 'n') return String(cell.v);
  if (cell.t === 'b') return String(cell.v);
  if (cell.t === 's') return String(cell.v ?? '').trim();
  // formula
  if (cell.v !== undefined) return String(cell.v);
  return '';
}

function parseFormattedNumber(value) {
  if (!value || value.trim() === '') return null;
  const cleaned = value.replace(/[,\sRs]/g, '').replace(/[^\d.\-]/g, '').trim();
  if (cleaned === '') return null;
  const num = parseFloat(cleaned);
  return isNaN(num) ? null : num;
}

function getEnhancedNumericCellValue(sheet, row, col) {
  const addr = XLSX.utils.encode_cell({ r: row, c: col });
  const cell = sheet[addr];
  if (!cell) return null;
  if (cell.t === 'n') return cell.v;
  if (cell.t === 's') return parseFormattedNumber(cell.v);
  if (cell.v !== undefined) {
    const n = Number(cell.v);
    if (!isNaN(n)) return n;
    return parseFormattedNumber(String(cell.v));
  }
  return null;
}

function parseFormattedInteger(value) {
  if (!value || value.trim() === '') return null;
  const cleaned = value.replace(/[^\d\-]/g, '').trim();
  if (cleaned === '') return null;
  const num = parseInt(cleaned, 10);
  return isNaN(num) ? null : num;
}

function getIntegerCellValue(sheet, row, col) {
  const addr = XLSX.utils.encode_cell({ r: row, c: col });
  const cell = sheet[addr];
  if (!cell) return null;
  if (cell.t === 'n') return Math.floor(cell.v);
  if (cell.t === 's') return parseFormattedInteger(cell.v);
  if (cell.v !== undefined) {
    const n = Number(cell.v);
    if (!isNaN(n)) return Math.floor(n);
    return parseFormattedInteger(String(cell.v));
  }
  return null;
}

// ── Header detection ───────────────────────────────────────────────────────────

function findHeaderIndices(sheet) {
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
  let skuCol = -1, nameCol = -1, priceCol = -1, comparedPriceCol = -1, availableCol = -1;

  for (let c = range.s.c; c <= range.e.c; c++) {
    const header = getCellValue(sheet, 0, c).trim();
    if (!header) continue;
    const h = header.toLowerCase();

    if (h === SKU_HEADER.toLowerCase()) {
      skuCol = c;
    } else if (
      h === NAME_HEADER.toLowerCase() ||
      h === 'product' ||
      h === 'product name' ||
      h === 'product title'
    ) {
      nameCol = c;
    } else if (h === PRICE_HEADER.toLowerCase()) {
      priceCol = c;
    } else if (
      h === COMPARED_PRICE_HEADER.toLowerCase() ||
      h === 'compare at price' ||
      h === 'compare price'
    ) {
      comparedPriceCol = c;
    } else if (
      h === AVAILABLE_HEADER.toLowerCase() ||
      h === 'stock' ||
      h === 'quantity' ||
      h === 'inventory quantity'
    ) {
      availableCol = c;
    }
  }

  if (skuCol === -1 || nameCol === -1 || priceCol === -1) {
    console.error('CRITICAL ERROR: Failed to find required headers (SKU, Product Name, Price).');
    return null;
  }

  return { skuCol, nameCol, priceCol, comparedPriceCol, availableCol };
}

// ── Read reference data ────────────────────────────────────────────────────────

function readReferenceData(
  file,
  prices,
  referenceCompareAtPrices,
  reportItems,
  locationFileNames
) {
  const workbook = XLSX.read(file.buffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return null;

  const indices = findHeaderIndices(sheet);
  if (!indices) return null;

  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');

  for (let i = 1; i <= range.e.r; i++) {
    const sku = getCellValue(sheet, i, indices.skuCol).toUpperCase();
    const name = getCellValue(sheet, i, indices.nameCol);
    const price = getEnhancedNumericCellValue(sheet, i, indices.priceCol);

    let compareAtPrice = null;
    if (indices.comparedPriceCol >= 0) {
      compareAtPrice = getEnhancedNumericCellValue(sheet, i, indices.comparedPriceCol);
    }
    if (compareAtPrice == null || compareAtPrice === 0) {
      compareAtPrice = 0.0;
    }

    if (sku && price != null) {
      prices.set(sku, price);
      referenceCompareAtPrices.set(sku, compareAtPrice);

      const locPrices = new Map();
      locationFileNames.forEach((ln) => locPrices.set(ln, null));

      const item = createReferenceItem(sku, name, price, compareAtPrice, locPrices);

      // Read stock from reference file if available
      if (indices.availableCol >= 0) {
        const stock = getIntegerCellValue(sheet, i, indices.availableCol);
        if (stock != null) {
          item.locationStock.set(file.name, stock);
        }
      }

      reportItems.set(sku, item);
    }
  }

  return indices;
}

// ── OGF comparison ─────────────────────────────────────────────────────────────

function performOgfComparison(
  originalFileName,
  file,
  refPrices,
  reportItems
) {
  const workbook = XLSX.read(file.buffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return;

  const locIndices = findHeaderIndices(sheet);
  if (!locIndices) {
    console.error(`Skipping ${originalFileName}: Could not find required headers.`);
    return;
  }

  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');

  for (let i = 1; i <= range.e.r; i++) {
    const sku = getCellValue(sheet, i, locIndices.skuCol).toUpperCase();
    const locationPrice = getEnhancedNumericCellValue(sheet, i, locIndices.priceCol);

    let compareAtPrice = null;
    if (locIndices.comparedPriceCol >= 0) {
      compareAtPrice = getEnhancedNumericCellValue(sheet, i, locIndices.comparedPriceCol);
    }

    let availableStock = null;
    if (locIndices.availableCol >= 0) {
      availableStock = getIntegerCellValue(sheet, i, locIndices.availableCol);
    }

    if (sku && reportItems.has(sku)) {
      const referencePrice = refPrices.get(sku);
      const item = reportItems.get(sku);

      // Store both price values
      item.locationPrices.set(originalFileName, locationPrice);
      if (compareAtPrice != null) {
        item.locationCompareAtPrices.set(originalFileName, compareAtPrice);
      } else {
        item.locationCompareAtPrices.set(originalFileName, 0.0);
      }

      // Calculate and store discount percentage
      if (compareAtPrice != null && locationPrice != null && compareAtPrice > 0) {
        const discount = ((compareAtPrice - locationPrice) / compareAtPrice) * 100;
        if (Math.abs(discount) > 0.5) {
          item.locationDiscountPercentages.set(originalFileName, discount);
        }
      }

      // Store available stock
      if (availableStock != null) {
        item.locationStock.set(originalFileName, availableStock);
      }

      // For OGF, always use Price column for comparison
      const priceToUse = locationPrice;
      if (priceToUse != null) {
        item.locationPricesUsed.set(originalFileName, priceToUse);
      }

      // Calculate and store Compare at price difference
      if (compareAtPrice != null && locationPrice != null) {
        const compareAtDifference = compareAtPrice - locationPrice;
        item.compareAtPriceDifferences.set(originalFileName, compareAtDifference);
      }

      // For OGF files, check if price is LESS THAN 22% of reference price
      if (priceToUse != null && referencePrice != null && referencePrice > 0) {
        const percentageDiff = ((priceToUse - referencePrice) / referencePrice) * 100;
        if (percentageDiff < 22.0) {
          const priceType = compareAtPrice != null ? 'Compare at price' : 'Price';
          const discrepancy = `${originalFileName}: Below 22% range (${percentageDiff.toFixed(2)}%) (${priceType})`;
          item.discrepancies.push(discrepancy);
        }
      }
    }
  }
}

// ── Regular (non-OGF) comparison ───────────────────────────────────────────────

function performRegularComparison(
  originalFileName,
  file,
  refPrices,
  reportItems
) {
  const workbook = XLSX.read(file.buffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return;

  const locIndices = findHeaderIndices(sheet);
  if (!locIndices) {
    console.error(`Skipping ${originalFileName}: Could not find required headers.`);
    return;
  }

  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');

  for (let i = 1; i <= range.e.r; i++) {
    const sku = getCellValue(sheet, i, locIndices.skuCol).toUpperCase();
    const locationPrice = getEnhancedNumericCellValue(sheet, i, locIndices.priceCol);

    let compareAtPrice = null;
    if (locIndices.comparedPriceCol >= 0) {
      compareAtPrice = getEnhancedNumericCellValue(sheet, i, locIndices.comparedPriceCol);
    }

    let availableStock = null;
    if (locIndices.availableCol >= 0) {
      availableStock = getIntegerCellValue(sheet, i, locIndices.availableCol);
    }

    if (sku && reportItems.has(sku)) {
      const referencePrice = refPrices.get(sku);
      const item = reportItems.get(sku);

      // Store both price values
      item.locationPrices.set(originalFileName, locationPrice);
      if (compareAtPrice != null) {
        item.locationCompareAtPrices.set(originalFileName, compareAtPrice);
      } else {
        item.locationCompareAtPrices.set(originalFileName, 0.0);
      }

      // Calculate and store discount percentage
      if (compareAtPrice != null && locationPrice != null && compareAtPrice > 0) {
        const discount = ((compareAtPrice - locationPrice) / compareAtPrice) * 100;
        if (Math.abs(discount) > 0.5) {
          item.locationDiscountPercentages.set(originalFileName, discount);
        }
      }

      // Store available stock
      if (availableStock != null) {
        item.locationStock.set(originalFileName, availableStock);
      }

      // For non-OGF files, always use Price column for comparison
      const priceToUse = locationPrice;
      if (priceToUse != null) {
        item.locationPricesUsed.set(originalFileName, priceToUse);
      }

      // Calculate and store Compare at price difference
      if (compareAtPrice != null && locationPrice != null) {
        const compareAtDifference = compareAtPrice - locationPrice;
        item.compareAtPriceDifferences.set(originalFileName, compareAtDifference);
      }

      // Compare the selected price with reference price
      if (priceToUse != null && referencePrice != null && Math.abs(priceToUse - referencePrice) > 0.01) {
        const difference = priceToUse - referencePrice;
        const sign = difference > 0 ? '+' : '-';
        const priceType = compareAtPrice != null ? 'Compare at price' : 'Price';
        const discrepancy = `${originalFileName}: ${sign}Rs.${Math.abs(difference).toFixed(2)} (${priceType})`;
        item.discrepancies.push(discrepancy);
      }
    }
  }
}

// ── Compare file (routes to OGF or regular) ────────────────────────────────────

function compareFile(
  referenceFile,
  locationFile,
  referencePrices,
  reportItems
) {
  const originalFileName = locationFile.name;
  const isOgfFile = originalFileName.toLowerCase().includes('ogf');

  if (isOgfFile) {
    // For OGF files, clean up SKUs first
    const fileToCompare = cleanupSkuForPriceComparison(locationFile);
    performOgfComparison(originalFileName, fileToCompare, referencePrices, reportItems);
  } else {
    performRegularComparison(originalFileName, locationFile, referencePrices, reportItems);
  }
}

// ── Calculate total stock ──────────────────────────────────────────────────────

function calculateTotalStockForItems(reportItems) {
  for (const item of reportItems.values()) {
    let totalStock = 0;
    for (const stock of item.locationStock.values()) {
      if (stock != null) {
        totalStock += stock;
      }
    }
    item.totalStock = totalStock;
  }
}

// ── Discount consistency check ─────────────────────────────────────────────────

function checkDiscountConsistency(item, locationFileNames) {
  item.discountDiscrepancies.length = 0;

  const hasDiscount = new Map();
  const discountValues = new Map();

  for (const locName of locationFileNames) {
    const discount = item.locationDiscountPercentages.get(locName);
    const hasDisc = discount != null && Math.abs(discount) > 0.5;
    hasDiscount.set(locName, hasDisc);
    if (hasDisc) {
      discountValues.set(locName, discount);
    }
  }

  const withDiscountCount = [...hasDiscount.values()].filter((b) => b).length;
  const withoutDiscountCount = hasDiscount.size - withDiscountCount;

  // Check consistency
  if (withDiscountCount > 0 && withoutDiscountCount > 0) {
    const discCompanies = [];
    const noDiscCompanies = [];

    for (const [name, hasDisc] of hasDiscount) {
      const cleanName = name.replace('.xlsx', '').replace('.xls', '');
      if (hasDisc) {
        discCompanies.push(cleanName);
      } else {
        noDiscCompanies.push(cleanName);
      }
    }

    const issue = `Discount inconsistency: ${discCompanies.length} have discounts, ${noDiscCompanies.length} don't`;
    item.discountDiscrepancies.push(issue);

    if (discCompanies.length > 0) {
      item.discountDiscrepancies.push(` Discounted in: ${discCompanies.join(', ')}`);
    }
  }

  // Check if discount percentages vary significantly
  if (discountValues.size > 1) {
    let total = 0;
    for (const v of discountValues.values()) total += v;
    const average = total / discountValues.size;

    for (const [name, val] of discountValues) {
      if (Math.abs(val - average) > 10.0) {
        const cleanName = name.replace('.xlsx', '').replace('.xls', '');
        item.discountDiscrepancies.push(`${cleanName}: ${val.toFixed(1)}% discount differs from others`);
      }
    }
  }
}

// ── Analyze compare-at-price margins ───────────────────────────────────────────

function analyzeCompareAtPriceMargins(
  item,
  consistencyIssues,
  compareAtDiffFiles
) {
  const compareAtPrices = item.locationCompareAtPrices;
  const regularPrices = item.locationPrices;

  const margins = new Map();
  for (const [fileName, compareAtPrice] of compareAtPrices) {
    const regularPrice = regularPrices.get(fileName);
    if (regularPrice != null && compareAtPrice != null) {
      margins.set(fileName, compareAtPrice - regularPrice);
    }
  }

  if (margins.size > 1) {
    // Find the mode (most frequent margin)
    const frequency = new Map();
    for (const margin of margins.values()) {
      frequency.set(margin, (frequency.get(margin) || 0) + 1);
    }

    let modeMargin = 0;
    let maxCount = 0;
    for (const [margin, count] of frequency) {
      if (count > maxCount) {
        maxCount = count;
        modeMargin = margin;
      }
    }

    if (frequency.size > 1) {
      for (const [fileName, margin] of margins) {
        const difference = margin - modeMargin;
        if (Math.abs(difference) > 1.0) {
          const direction = difference > 0 ? 'higher' : 'lower';
          const issue = `${fileName}: Rs.${Math.abs(difference).toFixed(2)} ${direction} margin`;
          consistencyIssues.push(issue);
          compareAtDiffFiles.push(fileName);
        }
      }

      if (consistencyIssues.length > 0) {
        const discrepancy = 'Varying Compare at price margins: ' + consistencyIssues.join(' | ');
        item.discrepancies.push(discrepancy);
      }
    }
  }
}

// ── Generate simple explanation ────────────────────────────────────────────────

function generateSimpleExplanation(
  item,
  ogfDifferences,
  nonOgfDifferences,
  compareAtDiffFiles,
  nonOgfDifferenceFiles,
  ogfPercentageDiff,
  status
) {
  if (status === 'Good') {
    return 'All prices match correctly';
  }

  const explanations = [];

  if (ogfDifferences.length > 0) {
    if (ogfPercentageDiff < 22.0) {
      explanations.push(`OGF price below 22% threshold (${ogfPercentageDiff.toFixed(2)}%)`);
    }
  }

  if (nonOgfDifferenceFiles.length > 0) {
    explanations.push('Price difference in: ' + nonOgfDifferenceFiles.join(', '));
  }

  if (compareAtDiffFiles.length > 0) {
    explanations.push('Compare at price difference in: ' + compareAtDiffFiles.join(', '));
  }

  if (explanations.length === 0) {
    return 'Check detailed discrepancies';
  }

  return explanations.join('; ');
}

// ── Calculate status for all items ─────────────────────────────────────────────

function calculateStatusForItems(
  reportItems,
  locationFileNames
) {
  for (const item of reportItems.values()) {
    const ogfDifferences = [];
    const nonOgfDifferences = [];
    const differenceFiles = [];
    const compareAtDiffFiles = [];

    // Analyze Compare at price consistency across files
    const compareAtPrices = item.locationCompareAtPrices;
    const hasCompareAtPrices = compareAtPrices.size > 0;
    const compareAtConsistencyIssues = [];

    if (hasCompareAtPrices && compareAtPrices.size > 1) {
      analyzeCompareAtPriceMargins(item, compareAtConsistencyIssues, compareAtDiffFiles);
    }

    // Track OGF price and percentage difference
    let ogfPrice = null;
    let ogfPercentageDiff = 0.0;
    let ogfFileName = '';

    // Check ALL location files for price differences
    for (const fileName of locationFileNames) {
      const priceUsed = item.locationPricesUsed.get(fileName);
      const isOgfFile = fileName.toLowerCase().includes('ogf');

      if (priceUsed != null) {
        if (isOgfFile) {
          const referenceValue =
            item.referenceCompareAtPrice > 0 ? item.referenceCompareAtPrice : item.referencePrice;
          const percentageDiff = ((priceUsed - referenceValue) / referenceValue) * 100;

          item.ogfPercentageDiff = percentageDiff;

          let ogfPercentageRemark;
          if (percentageDiff < 22.0) {
            ogfPercentageRemark = `Below 22% threshold (${percentageDiff.toFixed(2)}%)`;
          } else if (percentageDiff >= 22.0 && percentageDiff <= 25.0) {
            ogfPercentageRemark = `Within 22-25% range (${percentageDiff.toFixed(2)}%)`;
          } else {
            ogfPercentageRemark = `Above 25% (${percentageDiff.toFixed(2)}%)`;
          }
          item.ogfPercentageRemark = ogfPercentageRemark;

          if (percentageDiff < 22.0) {
            differenceFiles.push(fileName);
            ogfFileName = fileName;
            ogfPrice = priceUsed;
            ogfPercentageDiff = percentageDiff;

            const discrepancyKey = fileName + ':';
            const hasDiscrepancy = item.discrepancies.some((d) => d.startsWith(discrepancyKey));

            if (!hasDiscrepancy) {
              const compareAtPrice = item.locationCompareAtPrices.get(fileName);
              const priceType = compareAtPrice != null ? 'Compare at price' : 'Price';
              const discrepancy = `${fileName}: Below 22% range (${priceType})`;
              ogfDifferences.push(discrepancy);
              item.discrepancies.push(discrepancy);
            }
          }
        } else {
          // For non-OGF files, check exact match
          if (Math.abs(priceUsed - item.referencePrice) > 0.01) {
            differenceFiles.push(fileName);

            const discrepancyKey = fileName + ':';
            const hasDiscrepancy = item.discrepancies.some((d) => d.startsWith(discrepancyKey));

            if (!hasDiscrepancy) {
              const difference = priceUsed - item.referencePrice;
              const sign = difference > 0 ? '+' : '-';
              const compareAtPrice = item.locationCompareAtPrices.get(fileName);
              const priceType = compareAtPrice != null ? 'Compare at price' : 'Price';
              const discrepancy = `${fileName}: ${sign}Rs.${Math.abs(difference).toFixed(2)} (${priceType})`;
              nonOgfDifferences.push(discrepancy);
              item.discrepancies.push(discrepancy);
            }
          }
        }

        // Track OGF specifically for percentage calculation
        if (isOgfFile && ogfPrice == null) {
          ogfFileName = fileName;
          ogfPrice = priceUsed;
          if (ogfPrice != null) {
            const referenceValue =
              item.referenceCompareAtPrice > 0 ? item.referenceCompareAtPrice : item.referencePrice;
            if (referenceValue > 0) {
              ogfPercentageDiff = ((ogfPrice - referenceValue) / referenceValue) * 100;
            }
          }
        }
      }
    }

    // Also check existing discrepancies to ensure we capture everything
    for (const discrepancy of item.discrepancies) {
      const fileName = discrepancy.split(':')[0].trim();
      if (fileName.toLowerCase().includes('ogf')) {
        if (!ogfDifferences.includes(discrepancy)) {
          ogfDifferences.push(discrepancy);
        }
      } else {
        if (!nonOgfDifferences.includes(discrepancy)) {
          nonOgfDifferences.push(discrepancy);
        }
      }
    }

    // Check discount consistency
    checkDiscountConsistency(item, locationFileNames);

    let status;
    const statusReasons = [];

    const hasNonOgfDifferences = nonOgfDifferences.length > 0;
    const hasCompareAtInconsistency = compareAtConsistencyIssues.length > 0;
    const hasAnyDifferences =
      ogfDifferences.length > 0 || hasNonOgfDifferences || hasCompareAtInconsistency;
    const hasDiscountIssues = item.discountDiscrepancies.length > 0;

    if (!hasAnyDifferences && !hasDiscountIssues) {
      status = 'Good';
      statusReasons.push('No price differences found');
    } else if (hasCompareAtInconsistency) {
      status = 'Bad';
      statusReasons.push('Inconsistent Compare at price margins across files');
      statusReasons.push(...compareAtConsistencyIssues);
    } else if (ogfDifferences.length > 0) {
      if (ogfPercentageDiff >= 22.0) {
        if (hasNonOgfDifferences) {
          status = 'Bad';
          statusReasons.push('OGF within acceptable range');
          statusReasons.push('Other files have differences');
          statusReasons.push(...nonOgfDifferences);
        } else {
          status = 'Good';
          statusReasons.push('OGF within acceptable range');
        }
      } else {
        status = 'Bad';
        statusReasons.push('OGF price below acceptable range (less than 22%)');
        if (hasNonOgfDifferences) {
          statusReasons.push('Other files also have differences');
          statusReasons.push(...nonOgfDifferences);
        }
      }
    } else {
      status = 'Bad';
      statusReasons.push('Non-OGF files have differences:');
      statusReasons.push(...nonOgfDifferences);
    }

    item.status = status;
    item.statusReason = statusReasons.join(' | ');

    // Store separated discrepancies
    item.ogfDiscrepancies.length = 0;
    item.ogfDiscrepancies.push(...ogfDifferences);

    item.nonOgfDiscrepancies.length = 0;
    item.nonOgfDiscrepancies.push(...nonOgfDifferences);

    // Track non-OGF difference files for explanation
    const nonOgfDifferenceFiles = [];
    for (const fileName of locationFileNames) {
      if (fileName.toLowerCase().includes('ogf')) continue;
      const priceUsed = item.locationPricesUsed.get(fileName);
      if (priceUsed != null && Math.abs(priceUsed - item.referencePrice) > 0.01) {
        nonOgfDifferenceFiles.push(fileName);
      }
    }

    // Generate simple difference explanation
    const simpleExplanation = generateSimpleExplanation(
      item,
      ogfDifferences,
      nonOgfDifferences,
      compareAtDiffFiles,
      nonOgfDifferenceFiles,
      ogfPercentageDiff,
      status
    );
    item.differenceExplanation = simpleExplanation;

    // Sort discrepancies: OGF first
    item.discrepancies.sort((d1, d2) => {
      const d1IsOgf = d1.toLowerCase().includes('ogf');
      const d2IsOgf = d2.toLowerCase().includes('ogf');
      if (d1IsOgf && !d2IsOgf) return -1;
      if (!d1IsOgf && d2IsOgf) return 1;
      return d1.localeCompare(d2);
    });
  }
}

// ── Write comparison report (ExcelJS with styling) ─────────────────────────────

async function writeComparisonReport(
  reportItems,
  locationFileNames
) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Price Comparison Report');

  // ── Build header row ──
  const headers = [
    'SKU',
    'Product Name',
    'Stock Available',
    'Reference Price',
    'Reference Compare at Price',
  ];

  const sellingPriceColMap = new Map();
  const originalPriceColMap = new Map();
  const discountColMap = new Map();

  let colIndex = headers.length + 1; // ExcelJS is 1-based

  for (const locName of locationFileNames) {
    const displayName = locName.replace('.xlsx', '').replace('.xls', '');

    headers.push(`${displayName} - Price`);
    sellingPriceColMap.set(locName, colIndex++);

    headers.push(`${displayName} - Compare at price`);
    originalPriceColMap.set(locName, colIndex++);

    headers.push(`${displayName} - Disc %`);
    discountColMap.set(locName, colIndex++);
  }

  headers.push('Status');
  const statusColIndex = colIndex++;

  headers.push('OGF Percentage');
  const ogfPercentageColIndex = colIndex++;

  headers.push('Difference Explanation');
  const explanationColIndex = colIndex++;

  headers.push('OGF Differences');
  const ogfDifferencesColIndex = colIndex++;

  headers.push('Non-OGF Differences');
  const nonOgfDifferencesColIndex = colIndex++;

  headers.push('Discount Issues');
  const discountIssuesColIndex = colIndex++;

  // Add header row
  const headerRow = sheet.addRow(headers);

  // Style header row
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF4472C4' },
    };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };
  });

  // ── Data rows ──
  for (const item of reportItems.values()) {
    const rowValues = [
      item.sku,
      item.productName,
      item.totalStock,
      item.referencePrice,
      item.referenceCompareAtPrice,
    ];

    // Fill location columns in header order
    for (const locName of locationFileNames) {
      // Selling price
      const sellingPrice = item.locationPrices.get(locName);
      rowValues.push(sellingPrice != null ? sellingPrice : 'N/A');

      // Compare at price
      const originalPrice = item.locationCompareAtPrices.get(locName);
      rowValues.push(originalPrice != null ? originalPrice : 0);

      // Discount %
      const discount = item.locationDiscountPercentages.get(locName);
      if (discount != null && Math.abs(discount) > 0.5) {
        rowValues.push(`${discount.toFixed(1)}%`);
      } else {
        rowValues.push('');
      }
    }

    // Status columns
    rowValues.push(item.status);
    rowValues.push(`${item.ogfPercentageDiff.toFixed(2)}%`);
    rowValues.push(item.differenceExplanation);

    // OGF Differences
    rowValues.push(
      item.ogfDiscrepancies.length === 0
        ? 'No OGF differences'
        : item.ogfDiscrepancies.join('\n')
    );

    // Non-OGF Differences
    rowValues.push(
      item.nonOgfDiscrepancies.length === 0
        ? 'No differences'
        : item.nonOgfDiscrepancies.join('\n')
    );

    // Discount Issues
    rowValues.push(
      item.discountDiscrepancies.length === 0
        ? 'No discount issues'
        : item.discountDiscrepancies.join('\n')
    );

    const dataRow = sheet.addRow(rowValues);

    // Style status cell
    const statusCell = dataRow.getCell(statusColIndex);
    if (item.status === 'Good') {
      statusCell.font = { bold: true, color: { argb: 'FF006100' } };
      statusCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFC6EFCE' },
      };
    } else if (item.status === 'Bad') {
      statusCell.font = { bold: true, color: { argb: 'FF9C0006' } };
      statusCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFC7CE' },
      };
    }

    // Style price cells that differ
    for (const locName of locationFileNames) {
      const sellingCol = sellingPriceColMap.get(locName);
      const priceUsed = item.locationPricesUsed.get(locName);
      const isOgf = locName.toLowerCase().includes('ogf');

      if (priceUsed != null) {
        if (isOgf) {
          const referenceValue =
            item.referenceCompareAtPrice > 0 ? item.referenceCompareAtPrice : item.referencePrice;
          const pctDiff = ((priceUsed - referenceValue) / referenceValue) * 100;
          if (pctDiff < 22.0) {
            // Red highlight for below threshold
            dataRow.getCell(sellingCol).fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFFFC7CE' },
            };
          }
        } else {
          if (Math.abs(priceUsed - item.referencePrice) > 0.01) {
            // Yellow highlight for non-OGF mismatch
            dataRow.getCell(sellingCol).fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFFFFF00' },
            };
          }
        }
      }
    }

    // Apply thin borders to all data cells
    dataRow.eachCell((cell) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
      cell.alignment = { vertical: 'middle', wrapText: true };
    });
  }

  // Auto-size columns (approximate)
  sheet.columns.forEach((col) => {
    let maxLength = 12;
    col.eachCell?.({ includeEmpty: false }, (cell) => {
      const cellLength = cell.value ? String(cell.value).length : 0;
      if (cellLength > maxLength) maxLength = cellLength;
    });
    col.width = Math.min(maxLength + 2, 40);
  });

  const arrayBuffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}

// ── Main exported function ─────────────────────────────────────────────────────

export async function generatePriceReport(
  referenceFile,
  locationFiles
) {
  const referencePrices = new Map();
  const referenceCompareAtPrices = new Map();
  const reportItems = new Map();
  const locationFileNames = locationFiles.map((f) => f.name);

  // Process files through FileProcessor (OGF cleanup)
  const allFiles = [referenceFile, ...locationFiles];
  const processedFiles = processLocationFiles(allFiles, true);

  // Extract processed reference and location files
  const processedRefFile = processedFiles[0];
  const processedLocFiles = processedFiles.slice(1);

  // 1. Read Reference File and Initialize Report Map
  const refIndices = readReferenceData(
    processedRefFile,
    referencePrices,
    referenceCompareAtPrices,
    reportItems,
    locationFileNames
  );

  if (reportItems.size === 0 || refIndices == null) {
    console.error('ERROR: Could not find required columns or read any data from the Reference File.');
    return writeComparisonReport(reportItems, locationFileNames);
  }

  // 2. Process and Compare Location Files
  for (const locFile of processedLocFiles) {
    compareFile(processedRefFile, locFile, referencePrices, reportItems);
  }

  // 3. Calculate Status for Each Item
  calculateStatusForItems(reportItems, locationFileNames);

  // 4. Calculate total stock for each item
  calculateTotalStockForItems(reportItems);

  // 5. Write Report
  return writeComparisonReport(reportItems, locationFileNames);
}
