package com.example.InventoryComparer.logic;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.util.IOUtils;

import java.io.*;
import java.util.*;
import java.util.stream.Collectors;

public class LoyaltyComparerLogic {

    static {
        // Set memory limit for reading large files - THIS MUST BE AT CLASS LEVEL
        IOUtils.setByteArrayMaxOverride(150_000_000);
    }

    public static void generateReport(File referenceFile, List<File> locationFiles, File outputFile) throws Exception {
        System.out.println("=== Starting Loyalty Comparison Report ===");
        System.out.println("Reference file: " + referenceFile.getAbsolutePath());
        System.out.println("Location files: " + locationFiles.size());
        for (File loc : locationFiles) {
            System.out.println("  - " + loc.getAbsolutePath());
        }
        System.out.println("Output file: " + outputFile.getAbsolutePath());

        // Validate input files
        if (!referenceFile.exists()) {
            throw new FileNotFoundException("Reference file not found: " + referenceFile.getAbsolutePath());
        }

        for (File locationFile : locationFiles) {
            if (!locationFile.exists()) {
                throw new FileNotFoundException("Location file not found: " + locationFile.getAbsolutePath());
            }
        }

        // Read reference file data - ONLY PHONE NUMBERS
        System.out.println("Reading reference file (phone numbers only)...");
        List<String> referencePhoneNumbers = readReferenceFilePhoneNumbers(referenceFile);
        System.out.println("Found " + referencePhoneNumbers.size() + " phone numbers in reference file");

        if (referencePhoneNumbers.isEmpty()) {
            throw new RuntimeException("No valid phone numbers found in reference file. Please check the file format.");
        }

        // Read all location files data - PHONE NUMBERS AND LOYALTY TAGS
        Map<String, List<CustomerRecord>> locationRecordsMap = new HashMap<>();
        for (File locationFile : locationFiles) {
            System.out.println("Reading location file: " + locationFile.getName());
            List<CustomerRecord> locationRecords = readLocationFile(locationFile);
            locationRecordsMap.put(locationFile.getName(), locationRecords);
            System.out.println("Found " + locationRecords.size() + " records in " + locationFile.getName());
        }

        // Generate comparison report
        System.out.println("Generating Excel report...");
        generateLoyaltyReport(referencePhoneNumbers, locationRecordsMap, outputFile, locationFiles);
        System.out.println("=== Report generation completed successfully! ===");
    }

    private static List<String> readReferenceFilePhoneNumbers(File file) throws Exception {
        List<String> phoneNumbers = new ArrayList<>();

        // Set memory limit again for safety
        IOUtils.setByteArrayMaxOverride(150_000_000);

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = file.getName().toLowerCase().endsWith(".xlsx") ?
                     new XSSFWorkbook(fis) : new HSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            if (sheet == null) {
                throw new RuntimeException("No sheets found in file: " + file.getName());
            }

            Iterator<Row> rowIterator = sheet.iterator();

            // Find phone number column
            int phoneNumberCol = findPhoneColumn(sheet);
            if (phoneNumberCol == -1) {
                throw new RuntimeException("Phone number column not found in reference file: " + file.getName());
            }

            // Skip header row
            if (rowIterator.hasNext()) {
                rowIterator.next();
            }

            // Process data rows - ONLY EXTRACT PHONE NUMBERS
            int rowCount = 0;
            int totalRows = 0;

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                totalRows++;
                // Skip empty rows
                if (row == null) continue;

                Cell phoneCell = row.getCell(phoneNumberCol);
                if (phoneCell != null) {
                    String originalPhone = getCellValue(phoneCell);
                    String normalizedPhone = normalizePhoneNumber(originalPhone);

                    // Debug: show first few rows
                    if (totalRows <= 5) {
                        System.out.println("DEBUG Reference Row " + totalRows + ": '" + originalPhone + "' -> '" + normalizedPhone + "'");
                    }

                    if (normalizedPhone != null && !normalizedPhone.trim().isEmpty()) {
                        phoneNumbers.add(normalizedPhone);
                        rowCount++;
                    }
                }
            }
            System.out.println("Processed " + rowCount + " valid phone numbers from reference file out of " + totalRows + " total rows");
        } catch (Exception e) {
            System.err.println("Error reading reference file: " + file.getName());
            e.printStackTrace();
            throw e;
        }

        return phoneNumbers;
    }

    private static List<CustomerRecord> readLocationFile(File file) throws Exception {
        List<CustomerRecord> records = new ArrayList<>();

        // Set memory limit again for safety
        IOUtils.setByteArrayMaxOverride(150_000_000);

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = file.getName().toLowerCase().endsWith(".xlsx") ?
                     new XSSFWorkbook(fis) : new HSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            if (sheet == null) {
                throw new RuntimeException("No sheets found in file: " + file.getName());
            }

            Iterator<Row> rowIterator = sheet.iterator();

            // Find column indices
            Map<String, Integer> columnIndices = findColumnIndices(sheet);
            System.out.println("Columns found in " + file.getName() + ": " + columnIndices);

            int phoneNumberCol = columnIndices.getOrDefault("phone", -1);
            int tagsCol = columnIndices.getOrDefault("tags", -1);

            if (phoneNumberCol == -1) {
                // Try to find any column that might contain phone numbers
                phoneNumberCol = findAnyPhoneColumn(sheet);
                if (phoneNumberCol == -1) {
                    throw new RuntimeException("Phone number column not found in file: " + file.getName() +
                            ". Please ensure your Excel file has a column for phone numbers.");
                }
            }

            // Skip header row
            if (rowIterator.hasNext()) {
                rowIterator.next();
            }

            // Process data rows
            int rowCount = 0;
            int loyaltyCount = 0;
            int totalRows = 0;

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                totalRows++;
                // Skip empty rows
                if (row == null) continue;

                CustomerRecord record = createRecordFromRow(row, phoneNumberCol, tagsCol);
                if (record != null && record.phoneNumber != null && !record.phoneNumber.trim().isEmpty()) {
                    records.add(record);
                    rowCount++;
                    if (!"Not Loyalty".equals(record.loyaltyType)) {
                        loyaltyCount++;
                    }
                }
            }
            System.out.println("Processed " + rowCount + " valid rows from " + file.getName() + " out of " + totalRows + " total rows" +
                    " (" + loyaltyCount + " loyalty customers)");
        } catch (Exception e) {
            System.err.println("Error reading file: " + file.getName());
            e.printStackTrace();
            throw e;
        }

        return records;
    }

    private static int findPhoneColumn(Sheet sheet) {
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return -1;

        // Look for phone number columns
        for (Cell cell : headerRow) {
            if (cell == null) continue;

            String cellValue = getCellValue(cell).toLowerCase().trim();

            // Exact matches for phone columns
            if (cellValue.equals("tp_number") || cellValue.equals("phone") ||
                    cellValue.equals("mobile") || cellValue.equals("contact") ||
                    cellValue.equals("ph") || cellValue.equals("tel") ||
                    cellValue.equals("telephone") || cellValue.equals("phone number") ||
                    cellValue.equals("mobile number") || cellValue.equals("contact number")) {
                return cell.getColumnIndex();
            }
        }

        // Fallback: look for partial matches
        for (Cell cell : headerRow) {
            if (cell == null) continue;

            String cellValue = getCellValue(cell).toLowerCase().trim();
            if (cellValue.contains("phone") || cellValue.contains("mobile") ||
                    cellValue.contains("contact") || cellValue.contains("number") ||
                    cellValue.contains("tel") || cellValue.contains("ph")) {
                return cell.getColumnIndex();
            }
        }

        return -1;
    }

    private static Map<String, Integer> findColumnIndices(Sheet sheet) {
        Map<String, Integer> columnIndices = new HashMap<>();
        Row headerRow = sheet.getRow(0);

        if (headerRow == null) {
            throw new RuntimeException("No header row found in the Excel file");
        }

        // Look for phone number columns
        for (Cell cell : headerRow) {
            if (cell == null) continue;

            String cellValue = getCellValue(cell).toLowerCase().trim();

            // Exact matches for phone columns
            if (cellValue.equals("tp_number") || cellValue.equals("phone") ||
                    cellValue.equals("mobile") || cellValue.equals("contact") ||
                    cellValue.equals("ph") || cellValue.equals("tel") ||
                    cellValue.equals("telephone") || cellValue.equals("phone number") ||
                    cellValue.equals("mobile number") || cellValue.equals("contact number")) {
                if (!columnIndices.containsKey("phone")) {
                    columnIndices.put("phone", cell.getColumnIndex());
                    System.out.println("  -> Identified as PHONE column");
                }
            }

            // Look for tags columns
            if (cellValue.equals("tag") || cellValue.equals("tags") ||
                    cellValue.equals("category") || cellValue.equals("type") ||
                    cellValue.equals("loyalty") || cellValue.equals("status")) {
                if (!columnIndices.containsKey("tags")) {
                    columnIndices.put("tags", cell.getColumnIndex());
                    System.out.println("  -> Identified as TAGS column");
                }
            }
        }

        return columnIndices;
    }

    private static int findAnyPhoneColumn(Sheet sheet) {
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return -1;

        for (Cell cell : headerRow) {
            if (cell == null) continue;

            String cellValue = getCellValue(cell).toLowerCase().trim();
            if (cellValue.contains("no") || cellValue.contains("num") ||
                    cellValue.contains("phone") || cellValue.contains("tp_number") ||
                    cellValue.contains("mobile") || cellValue.contains("contact") ||
                    cellValue.length() <= 3) {
                System.out.println("Trying column '" + cellValue + "' as potential phone column");
                return cell.getColumnIndex();
            }
        }

        return -1;
    }

    private static CustomerRecord createRecordFromRow(Row row, int phoneNumberCol, int tagsCol) {
        if (row == null) {
            return null;
        }

        Cell phoneCell = row.getCell(phoneNumberCol);
        if (phoneCell == null) {
            return null;
        }

        String originalPhone = getCellValue(phoneCell);

        CustomerRecord record = new CustomerRecord();

        // Get phone number and normalize
        record.phoneNumber = normalizePhoneNumber(originalPhone);

        // Skip if phone number is empty after normalization
        if (record.phoneNumber == null || record.phoneNumber.trim().isEmpty()) {
            return null;
        }

        // Get tags if available
        if (tagsCol != -1) {
            Cell tagsCell = row.getCell(tagsCol);
            if (tagsCell != null) {
                record.tags = getCellValue(tagsCell);
                record.loyaltyType = determineLoyaltyTypeFromTags(record.tags);
            } else {
                record.loyaltyType = "Not Loyalty";
            }
        } else {
            // If no tags column, try to find loyalty in other columns
            record.loyaltyType = findLoyaltyInOtherColumns(row);
        }

        return record;
    }

    private static String determineLoyaltyTypeFromTags(String tags) {
        if (tags == null) return "Not Loyalty";

        if (tags.contains("Loyalty Customer G2")) {
            return "Loyalty Customer G2";
        } else if (tags.contains("Loyalty Customer")) {
            return "Loyalty Customer";
        } else {
            return "Not Loyalty";
        }
    }

    private static String findLoyaltyInOtherColumns(Row row) {
        if (row == null) return "Not Loyalty";

        for (int i = 0; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell != null) {
                String cellValue = getCellValue(cell);
                if (cellValue.contains("Loyalty Customer G2")) {
                    return "Loyalty Customer G2";
                } else if (cellValue.contains("Loyalty Customer")) {
                    return "Loyalty Customer";
                }
            }
        }
        return "Not Loyalty";
    }

    private static String normalizePhoneNumber(String phoneNumber) {
        if (phoneNumber == null) return "";

        // Step 1: Remove any leading single quote (')
        String cleaned = phoneNumber.trim();
        if (cleaned.startsWith("'")) {
            cleaned = cleaned.substring(1);
        }

        // Step 2: Remove all non-digit characters including spaces, dashes, parentheses, etc.
        String digitsOnly = cleaned.replaceAll("[^\\d]", "");

        // Step 3: If empty after cleaning, return empty
        if (digitsOnly.isEmpty()) return "";

        // Step 4: Handle different starting patterns
        String processedNumber;

        if (digitsOnly.startsWith("+94")) {
            // Remove +94 prefix and keep the rest
            processedNumber = digitsOnly.substring(3);
        } else if (digitsOnly.startsWith("94") && digitsOnly.length() == 11) {
            // Already starts with 94 and has 11 digits, use as is
            return formatPhoneNumber(digitsOnly);
        } else if (digitsOnly.startsWith("0")) {
            // Remove leading 0 and keep the rest
            processedNumber = digitsOnly.substring(1);
        } else if (digitsOnly.startsWith("1") || digitsOnly.startsWith("7")) {
            // Keep as is (we'll add 94 later)
            processedNumber = digitsOnly;
        } else {
            // For any other starting pattern, use as is
            processedNumber = digitsOnly;
        }

        // Step 5: Add 94 prefix to all numbers
        String withCountryCode = "94" + processedNumber;

        // Step 6: Only take numbers with exactly 11 digits total (including the 94)
        if (withCountryCode.length() != 11) {
            // Try alternative: maybe it's already 10 digits without country code
            if (digitsOnly.length() == 10) {
                withCountryCode = "94" + digitsOnly;
                if (withCountryCode.length() == 12) {
                    // Too long, try removing first digit
                    withCountryCode = "94" + digitsOnly.substring(1);
                }
            }

            if (withCountryCode.length() != 11) {
                return "";
            }
        }

        // Step 7: Format as 94-XXX-XXX-XXX
        return formatPhoneNumber(withCountryCode);
    }

    private static String formatPhoneNumber(String phoneNumber) {
        if (phoneNumber == null || phoneNumber.length() != 11) {
            return phoneNumber;
        }
        return phoneNumber.substring(0, 2) + "-" +
                phoneNumber.substring(2, 5) + "-" +
                phoneNumber.substring(5, 8) + "-" +
                phoneNumber.substring(8);
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    // For phone numbers stored as numbers, we need to handle them carefully
                    double num = cell.getNumericCellValue();
                    // If the number is too large for integer, use the double value
                    if (num > Long.MAX_VALUE) {
                        return String.valueOf(num).replace(".0", "");
                    } else if (num == (long) num) {
                        return String.valueOf((long) num);
                    } else {
                        return String.valueOf(num);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (Exception e) {
                    try {
                        double num = cell.getNumericCellValue();
                        if (num == (long) num) {
                            return String.valueOf((long) num);
                        } else {
                            return String.valueOf(num);
                        }
                    } catch (Exception ex) {
                        return cell.getCellFormula();
                    }
                }
            default:
                return "";
        }
    }

    private static void generateLoyaltyReport(List<String> referencePhoneNumbers,
                                              Map<String, List<CustomerRecord>> locationRecordsMap,
                                              File outputFile,
                                              List<File> locationFiles) throws Exception {

        int totalRecords = referencePhoneNumbers.size();
        int totalLocations = locationFiles.size();
        int estimatedCells = totalRecords * (1 + totalLocations + 2) * 3; // Phone + locations + 2 new columns

        System.out.println("Data size estimation:");
        System.out.println("  - Total records: " + totalRecords);
        System.out.println("  - Location files: " + totalLocations);
        System.out.println("  - Estimated cells: " + estimatedCells);

        if (estimatedCells > 300000 || totalRecords > 10000) {
            System.out.println("Using streaming workbook for large dataset (" + totalRecords + " records)");
            generateWithStreamingWorkbook(referencePhoneNumbers, locationRecordsMap, outputFile, locationFiles);
        } else {
            System.out.println("Using standard workbook");
            generateWithStandardWorkbook(referencePhoneNumbers, locationRecordsMap, outputFile, locationFiles);
        }
    }

    private static void generateWithStreamingWorkbook(List<String> referencePhoneNumbers,
                                                      Map<String, List<CustomerRecord>> locationRecordsMap,
                                                      File outputFile,
                                                      List<File> locationFiles) throws Exception {
        try (SXSSFWorkbook workbook = new SXSSFWorkbook(100)) {
            workbook.setCompressTempFiles(true);

            Sheet comparisonSheet = workbook.createSheet("Loyalty Comparison");
            if (comparisonSheet instanceof SXSSFSheet) {
                ((SXSSFSheet) comparisonSheet).trackAllColumnsForAutoSizing();
            }
            createComparisonSheet(workbook, comparisonSheet, referencePhoneNumbers, locationRecordsMap, locationFiles);

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                workbook.write(fos);
            }

            workbook.dispose();
        }
    }

    private static void generateWithStandardWorkbook(List<String> referencePhoneNumbers,
                                                     Map<String, List<CustomerRecord>> locationRecordsMap,
                                                     File outputFile,
                                                     List<File> locationFiles) throws Exception {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet comparisonSheet = workbook.createSheet("Loyalty Comparison");
            createComparisonSheet(workbook, comparisonSheet, referencePhoneNumbers, locationRecordsMap, locationFiles);

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                workbook.write(fos);
            }
        }
    }

    private static void createComparisonSheet(Workbook workbook, Sheet sheet,
                                              List<String> referencePhoneNumbers,
                                              Map<String, List<CustomerRecord>> locationRecordsMap,
                                              List<File> locationFiles) {

        CellStyle headerStyle = createHeaderStyle(workbook);
        CellStyle dataStyle = createDataStyle(workbook);
        CellStyle loyaltyStyle = createLoyaltyStyle(workbook);
        CellStyle loyaltyG2Style = createLoyaltyG2Style(workbook);
        CellStyle nonLoyaltyStyle = createNonLoyaltyStyle(workbook);
        CellStyle notFoundStyle = createNotFoundStyle(workbook);
        CellStyle goodStyle = createGoodStyle(workbook);
        CellStyle badStyle = createBadStyle(workbook);
        CellStyle ignoreStyle = createIgnoreStyle(workbook);

        // Create header row - Phone Number, then location columns, then Consistency, then Difference Details
        Row headerRow = sheet.createRow(0);
        String[] headers = new String[locationFiles.size() + 3]; // Phone + locations + 2 new columns
        headers[0] = "Phone Number";

        for (int i = 0; i < locationFiles.size(); i++) {
            headers[i + 1] = getShortFileName(locationFiles.get(i).getName()) + " Loyalty Status";
        }

        headers[locationFiles.size() + 1] = "Consistency";
        headers[locationFiles.size() + 2] = "Difference Details";

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }

        int rowNum = 1;
        int notFoundCount = 0;
        int ignoreCount = 0;

        for (String phoneNumber : referencePhoneNumbers) {
            Row row = sheet.createRow(rowNum++);

            // Phone Number
            Cell phoneCell = row.createCell(0);
            phoneCell.setCellValue(phoneNumber);
            phoneCell.setCellStyle(dataStyle);

            // Store loyalty statuses for consistency check (ONLY FROM LOCATION FILES)
            List<String> locationLoyaltyStatuses = new ArrayList<>();

            // Location file loyalty status
            for (int i = 0; i < locationFiles.size(); i++) {
                String locationFileName = locationFiles.get(i).getName();
                List<CustomerRecord> locationRecords = locationRecordsMap.get(locationFileName);

                if (locationRecords == null) {
                    Cell locationCell = row.createCell(i + 1);
                    locationCell.setCellValue("FILE ERROR");
                    locationCell.setCellStyle(notFoundStyle);
                    locationLoyaltyStatuses.add("FILE ERROR");
                    continue;
                }

                Map<String, CustomerRecord> locationMap = locationRecords.stream()
                        .collect(Collectors.toMap(r -> r.phoneNumber, r -> r, (r1, r2) -> r1));

                CustomerRecord locationRecord = locationMap.get(phoneNumber);
                Cell locationCell = row.createCell(i + 1);

                if (locationRecord != null) {
                    String locationStatus = locationRecord.loyaltyType;
                    locationCell.setCellValue(locationStatus);
                    locationLoyaltyStatuses.add(locationStatus);

                    CellStyle locStyle = dataStyle;
                    if ("Loyalty Customer".equals(locationRecord.loyaltyType)) {
                        locStyle = loyaltyStyle;
                    } else if ("Loyalty Customer G2".equals(locationRecord.loyaltyType)) {
                        locStyle = loyaltyG2Style;
                    } else {
                        locStyle = nonLoyaltyStyle;
                    }
                    locationCell.setCellStyle(locStyle);
                } else {
                    locationCell.setCellValue("PHONE NOT FOUND");
                    locationCell.setCellStyle(notFoundStyle);
                    locationLoyaltyStatuses.add("NOT FOUND");
                    notFoundCount++;
                }
            }

            // Consistency Check - Two new columns at the end
            int consistencyCol = locationFiles.size() + 1;
            int differenceCol = locationFiles.size() + 2;

            Cell consistencyCell = row.createCell(consistencyCol);
            Cell differenceCell = row.createCell(differenceCol);

            // Check if none of the files have Loyalty Customer or Loyalty Customer G2
            boolean shouldIgnore = shouldIgnore(locationLoyaltyStatuses);

            if (shouldIgnore) {
                consistencyCell.setCellValue("Ignore");
                differenceCell.setCellValue("No loyalty in any location file");
                consistencyCell.setCellStyle(ignoreStyle);
                differenceCell.setCellStyle(ignoreStyle);
                ignoreCount++;
            } else {
                // Check if all loyalty statuses are the same (ONLY ACROSS LOCATION FILES)
                boolean isConsistent = checkConsistency(locationLoyaltyStatuses);
                String consistencyResult = isConsistent ? "Good" : "Bad";
                String differenceDetails = getDifferenceDetails(locationLoyaltyStatuses, locationFiles);

                consistencyCell.setCellValue(consistencyResult);
                differenceCell.setCellValue(differenceDetails);

                // Apply styles to new columns
                if (isConsistent) {
                    consistencyCell.setCellStyle(goodStyle);
                } else {
                    consistencyCell.setCellStyle(badStyle);
                }
                differenceCell.setCellStyle(dataStyle);
            }
        }

        System.out.println("Total phone not found occurrences: " + notFoundCount);
        System.out.println("Total ignored (no loyalty): " + ignoreCount);

        try {
            for (int i = 0; i < headers.length; i++) {
                sheet.autoSizeColumn(i);
            }
        } catch (IllegalStateException e) {
            System.out.println("Warning: Could not auto-size columns in streaming mode.");
        }
    }

    private static boolean shouldIgnore(List<String> locationLoyaltyStatuses) {
        // Check if none of the location files have "Loyalty Customer" or "Loyalty Customer G2"
        boolean hasLoyaltyInAnyFile = false;

        for (String status : locationLoyaltyStatuses) {
            if ("Loyalty Customer".equals(status) || "Loyalty Customer G2".equals(status)) {
                hasLoyaltyInAnyFile = true;
                break;
            }
        }

        // If no file has loyalty, then we should ignore this record
        return !hasLoyaltyInAnyFile;
    }

    private static boolean checkConsistency(List<String> locationLoyaltyStatuses) {
        if (locationLoyaltyStatuses.isEmpty()) {
            return true;
        }

        // Check if ALL locations have the phone number (no "NOT FOUND" or "FILE ERROR")
        boolean allLocationsHavePhone = true;
        for (String status : locationLoyaltyStatuses) {
            if ("NOT FOUND".equals(status) || "FILE ERROR".equals(status)) {
                allLocationsHavePhone = false;
                break;
            }
        }

        // If any location doesn't have the phone number, it's inconsistent
        if (!allLocationsHavePhone) {
            return false;
        }

        // Now check if all loyalty statuses are the same
        String firstStatus = locationLoyaltyStatuses.get(0);
        for (String status : locationLoyaltyStatuses) {
            if (!firstStatus.equals(status)) {
                return false;
            }
        }

        return true;
    }

    private static String getDifferenceDetails(List<String> locationLoyaltyStatuses, List<File> locationFiles) {
        StringBuilder details = new StringBuilder();

        boolean hasMissingPhone = false;
        boolean hasDifferentStatus = false;
        List<String> missingLocations = new ArrayList<>();
        List<String> differentStatusLocations = new ArrayList<>();

        String firstValidStatus = null;

        // Find first valid status for comparison
        for (String status : locationLoyaltyStatuses) {
            if (!"NOT FOUND".equals(status) && !"FILE ERROR".equals(status)) {
                firstValidStatus = status;
                break;
            }
        }

        // Check for missing phones and different statuses
        for (int i = 0; i < locationLoyaltyStatuses.size(); i++) {
            String locationStatus = locationLoyaltyStatuses.get(i);
            String locationName = getShortFileName(locationFiles.get(i).getName());

            if ("NOT FOUND".equals(locationStatus) || "FILE ERROR".equals(locationStatus)) {
                hasMissingPhone = true;
                missingLocations.add(locationName + ": " + locationStatus);
            } else if (firstValidStatus != null && !firstValidStatus.equals(locationStatus)) {
                hasDifferentStatus = true;
                differentStatusLocations.add(locationName + " has " + locationStatus);
            }
        }

        if (hasMissingPhone) {
            details.append("Missing in: ");
            for (String missing : missingLocations) {
                details.append(missing).append("; ");
            }
        }

        if (hasDifferentStatus) {
            if (hasMissingPhone) {
                details.append(" | ");
            }
            details.append("Status differences: ");
            if (firstValidStatus != null) {
                details.append("Expected: ").append(firstValidStatus).append("; ");
            }
            for (String diff : differentStatusLocations) {
                details.append(diff).append("; ");
            }
        }

        // If no issues found
        if (!hasMissingPhone && !hasDifferentStatus) {
            if (firstValidStatus != null) {
                details.append("All locations have same status: ").append(firstValidStatus);
            } else {
                details.append("No issues found");
            }
        }

        return details.toString();
    }

    private static CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    private static CellStyle createDataStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private static CellStyle createLoyaltyStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(createDataStyle(workbook));
        style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private static CellStyle createLoyaltyG2Style(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(createDataStyle(workbook));
        style.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private static CellStyle createNonLoyaltyStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(createDataStyle(workbook));
        style.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private static CellStyle createNotFoundStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(createDataStyle(workbook));
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private static CellStyle createGoodStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(createDataStyle(workbook));
        style.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    private static CellStyle createBadStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(createDataStyle(workbook));
        style.setFillForegroundColor(IndexedColors.RED.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);
        return style;
    }

    private static CellStyle createIgnoreStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(createDataStyle(workbook));
        style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    private static String getShortFileName(String fileName) {
        if (fileName == null) return "Unknown";
        if (fileName.length() > 20) {
            return fileName.substring(0, 17) + "...";
        }
        return fileName;
    }

    // Helper class to store customer record data for location files
    private static class CustomerRecord {
        String phoneNumber;
        String loyaltyType;
        String tags;

        @Override
        public String toString() {
            return "CustomerRecord{phone='" + phoneNumber + "', loyaltyType='" + loyaltyType + "'}";
        }
    }
}