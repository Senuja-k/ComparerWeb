import * as XLSX from 'xlsx';

export const OGF_PREFIX = "OGF-";
export const OGF_FILENAME_PATTERN = "ogf";

/**
 * Detects OGF-related remarks from a raw SKU string.
 */
export function detectOgfRemarkFromSku(originalRawSku) {
  if (!originalRawSku || originalRawSku.trim() === "") return "";

  const sku = originalRawSku.trim();
  const upperSku = sku.toUpperCase();

  const hasOgfPrefix =
    upperSku.startsWith("OGF-") ||
    upperSku.startsWith("OGF_") ||
    upperSku.startsWith("TEMP_SKU_OGF") ||
    upperSku.startsWith("TEMP-OGF") ||
    upperSku.includes("-OGF") ||
    upperSku.includes("_OGF");

  const hasOgfAnywhere = upperSku.includes("OGF");

  if (hasOgfPrefix) {
    return `OGF prefix found: '${sku}'`;
  } else if (hasOgfAnywhere) {
    return `OGF detected in SKU: '${sku}'`;
  } else {
    return `WARNING: No OGF prefix in SKU: '${sku}'`;
  }
}

/**
 * Separates OGF files from non-OGF files and cleans SKU columns in OGF files.
 */
export function processLocationFiles(files, ogfActive) {
  if (!ogfActive) {
    return files;
  }

  const ogfFiles = files.filter(f => f.name.toLowerCase().includes(OGF_FILENAME_PATTERN));
  const nonOgfFiles = files.filter(f => !f.name.toLowerCase().includes(OGF_FILENAME_PATTERN));

  if (ogfFiles.length === 0) {
    return files;
  }

  const finalLocationFiles = [];

  const referenceFile = files[0];
  finalLocationFiles.push(referenceFile);

  const nonOgfFiltered = nonOgfFiles.filter(f => f !== referenceFile);

  for (const ogfFile of ogfFiles) {
    if (ogfFile !== referenceFile) {
      try {
        const cleaned = cleanupSkuForPriceComparison(ogfFile);
        finalLocationFiles.push(cleaned);
      } catch (e) {
        console.error(`Error processing OGF file for Price Comparer: ${ogfFile.name} - ${e}`);
        finalLocationFiles.push(ogfFile);
      }
    }
  }

  finalLocationFiles.push(...nonOgfFiltered);

  const seen = new Set();
  return finalLocationFiles.filter(f => {
    if (seen.has(f)) return false;
    seen.add(f);
    return true;
  });
}

/**
 * Cleans SKU column in an OGF file by removing the "OGF" prefix.
 */
export function cleanupSkuForPriceComparison(file) {
  const workbook = XLSX.read(file.buffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  if (!sheet) {
    console.error(`⚠️ No sheet found in ${file.name}`);
    return file;
  }

  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');

  let skuColIndex = -1;
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cellAddr = XLSX.utils.encode_cell({ r: 0, c });
    const cell = sheet[cellAddr];
    if (cell && String(cell.v).trim().toUpperCase() === "SKU") {
      skuColIndex = c;
      break;
    }
  }

  if (skuColIndex === -1) {
    console.error(`⚠️ SKU column not found in ${file.name}`);
    return file;
  }

  let remarkColIndex = -1;
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cellAddr = XLSX.utils.encode_cell({ r: 0, c });
    const cell = sheet[cellAddr];
    if (cell && String(cell.v).trim().toUpperCase() === "REMARK") {
      remarkColIndex = c;
      break;
    }
  }
  if (remarkColIndex === -1) {
    remarkColIndex = range.e.c + 1;
    const remarkHeaderAddr = XLSX.utils.encode_cell({ r: 0, c: remarkColIndex });
    sheet[remarkHeaderAddr] = { t: 's', v: 'Remark' };
    range.e.c = remarkColIndex;
    sheet['!ref'] = XLSX.utils.encode_range(range);
  }

  for (let r = 1; r <= range.e.r; r++) {
    const skuAddr = XLSX.utils.encode_cell({ r, c: skuColIndex });
    const skuCell = sheet[skuAddr];
    const sku = skuCell ? String(skuCell.v ?? '').trim() : '';

    let skuRemark = '';
    if (sku !== '') {
      const hasOgf = sku.toUpperCase().startsWith(OGF_PREFIX);
      skuRemark = hasOgf
        ? "OGF- prefix found."
        : "WARNING: OGF- prefix missing from SKU.";
    }

    const remarkAddr = XLSX.utils.encode_cell({ r, c: remarkColIndex });
    sheet[remarkAddr] = { t: 's', v: skuRemark };

    if (sku !== '' && sku.toUpperCase().startsWith(OGF_PREFIX)) {
      const cleanedSku = sku.substring(OGF_PREFIX.length);
      sheet[skuAddr] = { t: 's', v: cleanedSku };
    }
  }

  const newBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  return { name: file.name, buffer: Buffer.from(newBuffer) };
}
