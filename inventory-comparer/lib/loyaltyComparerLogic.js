import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';

// ── Phone Number Normalization (Sri Lankan +94 format) ──────────────────

function normalizePhoneNumber(phoneNumber) {
  if (phoneNumber == null) return '';

  // Step 1: Remove any leading single quote
  let cleaned = phoneNumber.trim();
  if (cleaned.startsWith("'")) {
    cleaned = cleaned.substring(1);
  }

  // Step 2: Remove all non-digit characters
  const digitsOnly = cleaned.replace(/[^\d]/g, '');

  // Step 3: If empty after cleaning, return empty
  if (digitsOnly.length === 0) return '';

  // Step 4: Handle different starting patterns
  let processedNumber;

  if (digitsOnly.startsWith('+94')) {
    // won't happen since we stripped non-digits, but kept for parity
    processedNumber = digitsOnly.substring(3);
  } else if (digitsOnly.startsWith('94') && digitsOnly.length === 11) {
    // Already starts with 94 and has 11 digits, use as is
    return formatPhoneNumber(digitsOnly);
  } else if (digitsOnly.startsWith('0')) {
    processedNumber = digitsOnly.substring(1);
  } else if (digitsOnly.startsWith('1') || digitsOnly.startsWith('7')) {
    processedNumber = digitsOnly;
  } else {
    processedNumber = digitsOnly;
  }

  // Step 5: Add 94 prefix
  let withCountryCode = '94' + processedNumber;

  // Step 6: Only take numbers with exactly 11 digits total
  if (withCountryCode.length !== 11) {
    if (digitsOnly.length === 10) {
      withCountryCode = '94' + digitsOnly;
      if (withCountryCode.length === 12) {
        withCountryCode = '94' + digitsOnly.substring(1);
      }
    }
    if (withCountryCode.length !== 11) {
      return '';
    }
  }

  // Step 7: Format as 94-XXX-XXX-XXX
  return formatPhoneNumber(withCountryCode);
}

function formatPhoneNumber(phoneNumber) {
  if (phoneNumber == null || phoneNumber.length !== 11) {
    return phoneNumber ?? '';
  }
  return (
    phoneNumber.substring(0, 2) + '-' +
    phoneNumber.substring(2, 5) + '-' +
    phoneNumber.substring(5, 8) + '-' +
    phoneNumber.substring(8)
  );
}

// ── Cell Value Extraction ───────────────────────────────────────────────

function getCellValue(cell) {
  if (cell == null) return '';

  if (cell.t === 's') {
    return String(cell.v ?? '').trim();
  }
  if (cell.t === 'n') {
    const num = cell.v;
    if (num === Math.floor(num)) {
      return String(Math.floor(num));
    }
    return String(num);
  }
  if (cell.t === 'b') {
    return String(cell.v);
  }
  if (cell.t === 'e') {
    return '';
  }
  // date or other
  if (cell.w) return cell.w;
  return cell.v != null ? String(cell.v) : '';
}

function getCellValueFromRow(sheet, row, col) {
  const addr = XLSX.utils.encode_cell({ r: row, c: col });
  return getCellValue(sheet[addr]);
}

// ── Column Detection ────────────────────────────────────────────────────

function findPhoneColumn(sheet) {
  const range = XLSX.utils.decode_range(sheet['!ref'] ?? 'A1');
  const headerRow = range.s.r;

  // Exact matches
  for (let c = range.s.c; c <= range.e.c; c++) {
    const val = getCellValueFromRow(sheet, headerRow, c).toLowerCase().trim();
    if (
      val === 'tp_number' || val === 'phone' || val === 'mobile' ||
      val === 'contact' || val === 'ph' || val === 'tel' ||
      val === 'telephone' || val === 'phone number' ||
      val === 'mobile number' || val === 'contact number'
    ) {
      return c;
    }
  }

  // Partial matches
  for (let c = range.s.c; c <= range.e.c; c++) {
    const val = getCellValueFromRow(sheet, headerRow, c).toLowerCase().trim();
    if (
      val.includes('phone') || val.includes('mobile') ||
      val.includes('contact') || val.includes('number') ||
      val.includes('tel') || val.includes('ph')
    ) {
      return c;
    }
  }

  return -1;
}

function findColumnIndices(sheet) {
  const range = XLSX.utils.decode_range(sheet['!ref'] ?? 'A1');
  const headerRow = range.s.r;
  let phone = -1;
  let tags = -1;

  for (let c = range.s.c; c <= range.e.c; c++) {
    const val = getCellValueFromRow(sheet, headerRow, c).toLowerCase().trim();

    if (
      phone === -1 &&
      (val === 'tp_number' || val === 'phone' || val === 'mobile' ||
        val === 'contact' || val === 'ph' || val === 'tel' ||
        val === 'telephone' || val === 'phone number' ||
        val === 'mobile number' || val === 'contact number')
    ) {
      phone = c;
    }

    if (
      tags === -1 &&
      (val === 'tag' || val === 'tags' || val === 'category' ||
        val === 'type' || val === 'loyalty' || val === 'status')
    ) {
      tags = c;
    }
  }

  return { phone, tags };
}

function findAnyPhoneColumn(sheet) {
  const range = XLSX.utils.decode_range(sheet['!ref'] ?? 'A1');
  const headerRow = range.s.r;

  for (let c = range.s.c; c <= range.e.c; c++) {
    const val = getCellValueFromRow(sheet, headerRow, c).toLowerCase().trim();
    if (
      val.includes('no') || val.includes('num') ||
      val.includes('phone') || val.includes('tp_number') ||
      val.includes('mobile') || val.includes('contact') ||
      val.length <= 3
    ) {
      return c;
    }
  }
  return -1;
}

// ── Loyalty Tag Detection ───────────────────────────────────────────────

function determineLoyaltyTypeFromTags(tags) {
  if (tags == null) return 'Not Loyalty';
  if (tags.includes('Loyalty Customer G2')) return 'Loyalty Customer G2';
  if (tags.includes('Loyalty Customer')) return 'Loyalty Customer';
  return 'Not Loyalty';
}

function findLoyaltyInOtherColumns(sheet, row) {
  const range = XLSX.utils.decode_range(sheet['!ref'] ?? 'A1');
  for (let c = range.s.c; c <= range.e.c; c++) {
    const val = getCellValueFromRow(sheet, row, c);
    if (val.includes('Loyalty Customer G2')) return 'Loyalty Customer G2';
    if (val.includes('Loyalty Customer')) return 'Loyalty Customer';
  }
  return 'Not Loyalty';
}

// ── File Reading ────────────────────────────────────────────────────────

function readReferenceFilePhoneNumbers(file) {
  const workbook = XLSX.read(file.buffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error('No sheets found in reference file: ' + file.name);

  const sheet = workbook.Sheets[sheetName];
  const range = XLSX.utils.decode_range(sheet['!ref'] ?? 'A1');

  const phoneCol = findPhoneColumn(sheet);
  if (phoneCol === -1) {
    throw new Error('Phone number column not found in reference file: ' + file.name);
  }

  const phoneNumbers = [];
  // Skip header row (row 0), start from row 1
  for (let r = range.s.r + 1; r <= range.e.r; r++) {
    const original = getCellValueFromRow(sheet, r, phoneCol);
    const normalized = normalizePhoneNumber(original);
    if (normalized.trim().length > 0) {
      phoneNumbers.push(normalized);
    }
  }

  if (phoneNumbers.length === 0) {
    throw new Error('No valid phone numbers found in reference file. Please check the file format.');
  }

  return phoneNumbers;
}

function readLocationFile(file) {
  const workbook = XLSX.read(file.buffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error('No sheets found in file: ' + file.name);

  const sheet = workbook.Sheets[sheetName];
  const range = XLSX.utils.decode_range(sheet['!ref'] ?? 'A1');

  const indices = findColumnIndices(sheet);
  let phoneCol = indices.phone;
  const tagsCol = indices.tags;

  if (phoneCol === -1) {
    phoneCol = findAnyPhoneColumn(sheet);
    if (phoneCol === -1) {
      throw new Error(
        'Phone number column not found in file: ' + file.name +
        '. Please ensure your Excel file has a column for phone numbers.'
      );
    }
  }

  const records = [];

  for (let r = range.s.r + 1; r <= range.e.r; r++) {
    const originalPhone = getCellValueFromRow(sheet, r, phoneCol);
    const normalized = normalizePhoneNumber(originalPhone);
    if (!normalized || normalized.trim().length === 0) continue;

    let loyaltyType;
    let tags = '';

    if (tagsCol !== -1) {
      tags = getCellValueFromRow(sheet, r, tagsCol);
      loyaltyType = determineLoyaltyTypeFromTags(tags);
    } else {
      loyaltyType = findLoyaltyInOtherColumns(sheet, r);
    }

    records.push({ phoneNumber: normalized, loyaltyType, tags });
  }

  return records;
}

// ── Consistency Checking ────────────────────────────────────────────────

function shouldIgnore(locationLoyaltyStatuses) {
  return !locationLoyaltyStatuses.some(
    s => s === 'Loyalty Customer' || s === 'Loyalty Customer G2'
  );
}

function checkConsistency(locationLoyaltyStatuses) {
  if (locationLoyaltyStatuses.length === 0) return true;

  // If any location doesn't have the phone number, it's inconsistent
  if (locationLoyaltyStatuses.some(s => s === 'NOT FOUND' || s === 'FILE ERROR')) {
    return false;
  }

  // All statuses must match
  const first = locationLoyaltyStatuses[0];
  return locationLoyaltyStatuses.every(s => s === first);
}

function getDifferenceDetails(
  locationLoyaltyStatuses,
  locationNames
) {
  let hasMissingPhone = false;
  let hasDifferentStatus = false;
  const missingLocations = [];
  const differentStatusLocations = [];

  let firstValidStatus = null;
  for (const status of locationLoyaltyStatuses) {
    if (status !== 'NOT FOUND' && status !== 'FILE ERROR') {
      firstValidStatus = status;
      break;
    }
  }

  for (let i = 0; i < locationLoyaltyStatuses.length; i++) {
    const status = locationLoyaltyStatuses[i];
    const locName = locationNames[i];

    if (status === 'NOT FOUND' || status === 'FILE ERROR') {
      hasMissingPhone = true;
      missingLocations.push(locName + ': ' + status);
    } else if (firstValidStatus != null && status !== firstValidStatus) {
      hasDifferentStatus = true;
      differentStatusLocations.push(locName + ' has ' + status);
    }
  }

  const parts = [];

  if (hasMissingPhone) {
    parts.push('Missing in: ' + missingLocations.join('; ') + '; ');
  }

  if (hasDifferentStatus) {
    let s = 'Status differences: ';
    if (firstValidStatus != null) {
      s += 'Expected: ' + firstValidStatus + '; ';
    }
    s += differentStatusLocations.join('; ') + '; ';
    parts.push(s);
  }

  if (!hasMissingPhone && !hasDifferentStatus) {
    if (firstValidStatus != null) {
      return 'All locations have same status: ' + firstValidStatus;
    }
    return 'No issues found';
  }

  return parts.join(' | ');
}

// ── Helpers ─────────────────────────────────────────────────────────────

function getShortFileName(fileName) {
  if (fileName == null) return 'Unknown';
  if (fileName.length > 20) return fileName.substring(0, 17) + '...';
  return fileName;
}

// ── Excel Styles ────────────────────────────────────────────────────────

const THIN_BORDER = { style: 'thin' };
const ALL_BORDERS = {
  top: THIN_BORDER, bottom: THIN_BORDER,
  left: THIN_BORDER, right: THIN_BORDER,
};

function applyHeaderStyle(cell) {
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9BC2E6' } }; // light blue
  cell.border = ALL_BORDERS;
  cell.font = { bold: true };
}

function applyDataStyle(cell) {
  cell.border = ALL_BORDERS;
}

function applyLoyaltyStyle(cell) {
  cell.border = ALL_BORDERS;
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF92D050' } }; // light green
}

function applyLoyaltyG2Style(cell) {
  cell.border = ALL_BORDERS;
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9BC2E6' } }; // light blue
}

function applyNonLoyaltyStyle(cell) {
  cell.border = ALL_BORDERS;
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC000' } }; // light orange
}

function applyNotFoundStyle(cell) {
  cell.border = ALL_BORDERS;
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC0C0C0' } }; // grey 25%
}

function applyGoodStyle(cell) {
  cell.border = ALL_BORDERS;
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00B050' } }; // bright green
  cell.font = { bold: true };
}

function applyBadStyle(cell) {
  cell.border = ALL_BORDERS;
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } }; // red
  cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
}

function applyIgnoreStyle(cell) {
  cell.border = ALL_BORDERS;
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; // light yellow
  cell.font = { bold: true };
}

// ── Report Generation ───────────────────────────────────────────────────

export async function generateLoyaltyReport(
  referenceFile,
  locationFiles
) {
  // 1. Read reference phone numbers
  const referencePhoneNumbers = readReferenceFilePhoneNumbers(referenceFile);

  // 2. Read all location files
  const locationRecordsMap = new Map();
  for (const locFile of locationFiles) {
    const records = readLocationFile(locFile);
    locationRecordsMap.set(locFile.name, records);
  }

  // 3. Build the output workbook with ExcelJS
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Loyalty Comparison');

  const locationNames = locationFiles.map(f => getShortFileName(f.name));

  // Header row
  const headers = [
    'Phone Number',
    ...locationNames.map(n => n + ' Loyalty Status'),
    'Consistency',
    'Difference Details',
  ];

  const headerRow = ws.addRow(headers);
  headerRow.eachCell(cell => applyHeaderStyle(cell));

  // Data rows
  for (const phoneNumber of referencePhoneNumbers) {
    const locationLoyaltyStatuses = [];
    const rowValues = [phoneNumber];

    // Collect loyalty status per location file
    for (const locFile of locationFiles) {
      const records = locationRecordsMap.get(locFile.name);
      if (!records) {
        rowValues.push('FILE ERROR');
        locationLoyaltyStatuses.push('FILE ERROR');
        continue;
      }

      const record = records.find(r => r.phoneNumber === phoneNumber);
      if (record) {
        rowValues.push(record.loyaltyType);
        locationLoyaltyStatuses.push(record.loyaltyType);
      } else {
        rowValues.push('PHONE NOT FOUND');
        locationLoyaltyStatuses.push('NOT FOUND');
      }
    }

    // Consistency check
    const ignore = shouldIgnore(locationLoyaltyStatuses);
    let consistencyVal;
    let differenceVal;

    if (ignore) {
      consistencyVal = 'Ignore';
      differenceVal = 'No loyalty in any location file';
    } else {
      const isConsistent = checkConsistency(locationLoyaltyStatuses);
      consistencyVal = isConsistent ? 'Good' : 'Bad';
      differenceVal = getDifferenceDetails(locationLoyaltyStatuses, locationNames);
    }

    rowValues.push(consistencyVal, differenceVal);

    const row = ws.addRow(rowValues);

    // ── Apply cell styles ──
    // Phone number cell
    applyDataStyle(row.getCell(1));

    // Location status cells
    for (let i = 0; i < locationFiles.length; i++) {
      const cell = row.getCell(i + 2); // 1-indexed, skip phone col
      const status = locationLoyaltyStatuses[i];

      if (status === 'NOT FOUND' || status === 'FILE ERROR') {
        applyNotFoundStyle(cell);
      } else if (status === 'Loyalty Customer') {
        applyLoyaltyStyle(cell);
      } else if (status === 'Loyalty Customer G2') {
        applyLoyaltyG2Style(cell);
      } else {
        applyNonLoyaltyStyle(cell);
      }
    }

    // Consistency cell
    const consistencyCell = row.getCell(locationFiles.length + 2);
    const differenceCell = row.getCell(locationFiles.length + 3);

    if (ignore) {
      applyIgnoreStyle(consistencyCell);
      applyIgnoreStyle(differenceCell);
    } else {
      if (consistencyVal === 'Good') {
        applyGoodStyle(consistencyCell);
      } else {
        applyBadStyle(consistencyCell);
      }
      applyDataStyle(differenceCell);
    }
  }

  // Auto-fit column widths (approximate)
  ws.columns.forEach(col => {
    let maxLen = 10;
    col.eachCell?.({ includeEmpty: false }, cell => {
      const len = String(cell.value ?? '').length;
      if (len > maxLen) maxLen = len;
    });
    col.width = Math.min(maxLen + 2, 50);
  });

  // 4. Write to buffer
  const arrayBuffer = await wb.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}
