package com.example.InventoryComparer.logic;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class PriceComparerLogic {

    //HEADER NAMES
    private static final String SKU_HEADER = "SKU";
    private static final String NAME_HEADER = "Product Name";
    private static final String PRICE_HEADER = "Price";
    private static final String COMPARED_PRICE_HEADER = "Compare at price";
    private static final String AVAILABLE_HEADER = "Available"; // NEW: Stock header

    //Data Structures
    private record ColumnIndices(int skuCol, int nameCol, int priceCol, int comparedPriceCol, int availableCol) {} // UPDATED: Added availableCol

    private static class ReferenceItem {
        private final String sku;
        private final String productName;
        private final double referencePrice;
        private final List<String> discrepancies;
        private final List<String> basicRemarks;
        private final Map<String, Double> locationPrices;
        private final Map<String, Double> locationCompareAtPrices;
        private final Map<String, Double> compareAtPriceDifferences;
        private final Map<String, Double> locationPricesUsed; // Track which price was actually used
        private final Map<String, Integer> locationStock; // NEW: Track stock per location
        private String status;
        private String statusReason;
        private String differenceExplanation; // NEW: Simple explanation field
        private int totalStock; // NEW: Total stock across all locations
        private final List<String> ogfDiscrepancies; // NEW: Separate OGF discrepancies
        private final List<String> nonOgfDiscrepancies; // NEW: Separate non-OGF discrepancies

        public ReferenceItem(String sku, String productName, double referencePrice,
                             List<String> discrepancies, Map<String, Double> locationPrices, String status) {
            this.sku = sku;
            this.productName = productName;
            this.referencePrice = referencePrice;
            this.discrepancies = discrepancies;
            this.basicRemarks = new ArrayList<>();
            this.locationPrices = locationPrices;
            this.locationCompareAtPrices = new HashMap<>();
            this.compareAtPriceDifferences = new HashMap<>();
            this.locationPricesUsed = new HashMap<>(); // Track price used for comparison
            this.locationStock = new HashMap<>(); // NEW: Initialize stock map
            this.status = status;
            this.statusReason = "";
            this.differenceExplanation = ""; // Initialize
            this.totalStock = 0; // NEW: Initialize total stock
            this.ogfDiscrepancies = new ArrayList<>(); // NEW: Initialize OGF discrepancies
            this.nonOgfDiscrepancies = new ArrayList<>(); // NEW: Initialize non-OGF discrepancies
        }

        public String sku() { return sku; }
        public String productName() { return productName; }
        public double referencePrice() { return referencePrice; }
        public List<String> discrepancies() { return discrepancies; }
        public List<String> basicRemarks() { return basicRemarks; }
        public Map<String, Double> locationPrices() { return locationPrices; }
        public Map<String, Double> locationCompareAtPrices() { return locationCompareAtPrices; }
        public Map<String, Double> compareAtPriceDifferences() { return compareAtPriceDifferences; }
        public Map<String, Double> locationPricesUsed() { return locationPricesUsed; }
        public Map<String, Integer> locationStock() { return locationStock; } // NEW: Getter for location stock
        public String status() { return status; }
        public void setStatus(String status) { this.status = status; }
        public String statusReason() { return statusReason; }
        public void setStatusReason(String statusReason) { this.statusReason = statusReason; }
        public String differenceExplanation() { return differenceExplanation; }
        public void setDifferenceExplanation(String differenceExplanation) { this.differenceExplanation = differenceExplanation; }
        public int totalStock() { return totalStock; } // NEW: Getter for total stock
        public void setTotalStock(int totalStock) { this.totalStock = totalStock; } // NEW: Setter for total stock
        public List<String> ogfDiscrepancies() { return ogfDiscrepancies; } // NEW: Getter for OGF discrepancies
        public List<String> nonOgfDiscrepancies() { return nonOgfDiscrepancies; } // NEW: Getter for non-OGF discrepancies
    }

    // UPDATED: Public Entry Point with original file names map
    public static void generateReport(File referenceFile, List<File> locationFiles, File outputFile,
                                      Map<File, String> originalFileNames) throws IOException {
        Map<String, Double> referencePrices = new HashMap<>();
        Map<String, ReferenceItem> reportItems = new LinkedHashMap<>();
        List<String> locationFileNames = new ArrayList<>();

        // UPDATED: Use the provided original file names map
        for (File file : locationFiles) {
            String originalName = getOriginalFileName(file, originalFileNames);
            locationFileNames.add(originalName);
        }

        //1. Read Reference File and Initialize Report Map
        ColumnIndices refIndices = readReferenceData(referenceFile, referencePrices, reportItems, locationFileNames, originalFileNames);

        if (reportItems.isEmpty() || refIndices == null) {
            System.err.println("ERROR: Could not find required columns or read any data from the Reference File. Check headers.");
            writeComparisonReport(outputFile, reportItems, locationFileNames);
            return;
        }

        //2. Process and Compare Location Files
        for (File locationFile : locationFiles) {
            compareFile(referenceFile, locationFile, referencePrices, reportItems, refIndices, originalFileNames);
        }

        //3. Calculate Status for Each Item
        calculateStatusForItems(reportItems, locationFileNames);

        //4. NEW: Calculate total stock for each item
        calculateTotalStockForItems(reportItems);

        //5. Write Report
        writeComparisonReport(outputFile, reportItems, locationFileNames);

        System.out.println("Price Comparison Complete. Report saved to: " + outputFile.getAbsolutePath());
    }

    // UPDATED: Helper method to get original file name using the provided map
    private static String getOriginalFileName(File file, Map<File, String> originalFileNames) {
        // First check if we have the original name in our map
        if (originalFileNames.containsKey(file)) {
            return originalFileNames.get(file);
        }

        String name = file.getName();
        // Remove temp file prefixes if present
        if (name.startsWith("temp_") || name.startsWith("temp_price_ogf_") || name.startsWith("upload_")) {
            // Try to extract original name from temp file
            return name.replaceFirst("^temp_", "")
                    .replaceFirst("^temp_price_ogf_", "")
                    .replaceFirst("^upload_\\d+_", "");
        }
        return name;
    }

    //CORE LOGIC METHODS

    private static ColumnIndices findHeaderIndices(Sheet sheet) {
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return null;

        int skuCol = -1, nameCol = -1, priceCol = -1, comparedPriceCol = -1, availableCol = -1;

        for (Cell cell : headerRow) {
            if (cell == null) continue;
            String header = getCellValue(cell).trim();
            int index = cell.getColumnIndex();

            if (header.equalsIgnoreCase(SKU_HEADER)) {
                skuCol = index;
            } else if (header.equalsIgnoreCase(NAME_HEADER) || header.equalsIgnoreCase("Product")) {
                nameCol = index;
            } else if (header.equalsIgnoreCase(PRICE_HEADER)) {
                priceCol = index;
            } else if (header.equalsIgnoreCase(COMPARED_PRICE_HEADER) ||
                    header.equalsIgnoreCase("Compare At Price") ||
                    header.equalsIgnoreCase("Compare Price")) {
                comparedPriceCol = index;
            } else if (header.equalsIgnoreCase(AVAILABLE_HEADER) ||
                    header.equalsIgnoreCase("Stock") ||
                    header.equalsIgnoreCase("Quantity")) { // NEW: Look for stock/quantity headers
                availableCol = index;
            }
        }

        if (skuCol == -1 || nameCol == -1 || priceCol == -1) {
            System.err.println("CRITICAL ERROR: Failed to find required headers (SKU, Product Name, Price).");
            return null;
        }

        return new ColumnIndices(skuCol, nameCol, priceCol, comparedPriceCol, availableCol);
    }

    // UPDATED: Added originalFileNames parameter
    private static ColumnIndices readReferenceData(File file, Map<String, Double> prices,
                                                   Map<String, ReferenceItem> reportItems,
                                                   List<String> locationFileNames,
                                                   Map<File, String> originalFileNames) throws IOException {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            ColumnIndices indices = findHeaderIndices(sheet);
            if (indices == null) return null;

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String sku = getCellValue(row.getCell(indices.skuCol())).toUpperCase();
                String name = getCellValue(row.getCell(indices.nameCol()));

                // Use Compare at price if available, otherwise use regular Price
                Double price = null;

                // First try to get Compare at price if column exists and has value
                if (indices.comparedPriceCol() >= 0) {
                    price = getEnhancedNumericCellValue(row.getCell(indices.comparedPriceCol()));
                }

                // If no Compare at price, use regular Price column
                if (price == null) {
                    price = getEnhancedNumericCellValue(row.getCell(indices.priceCol()));
                }

                if (!sku.isEmpty() && price != null) {
                    prices.put(sku, price);
                    Map<String, Double> locPrices = new HashMap<>();
                    locationFileNames.forEach(locName -> locPrices.put(locName, null));

                    ReferenceItem item = new ReferenceItem(sku, name, price, new ArrayList<>(), locPrices, "");

                    // NEW: Read stock from reference file if available
                    if (indices.availableCol() >= 0) {
                        Integer stock = getIntegerCellValue(row.getCell(indices.availableCol()));
                        if (stock != null) {
                            // UPDATED: Use the original file name from the map
                            item.locationStock().put(getOriginalFileName(file, originalFileNames), stock);
                        }
                    }

                    reportItems.put(sku, item);
                }
            }
            return indices;

        } catch (Exception e) {
            System.err.println("Error reading reference file: " + e.getMessage());
            return null;
        }
    }

    // UPDATED: Added originalFileNames parameter
    private static void compareFile(File refFile, File locationFile, Map<String, Double> referencePrices,
                                    Map<String, ReferenceItem> reportItems, ColumnIndices refIndices,
                                    Map<File, String> originalFileNames) throws IOException {
        String originalFileName = getOriginalFileName(locationFile, originalFileNames);
        boolean isOgfFile = originalFileName.toLowerCase().contains("ogf");

        if (isOgfFile) {
            // For OGF files, clean up SKUs first using FileProccessor
            File fileToCompare = FileProccessor.cleanupSkuForPriceComparison(locationFile);
            // Store the original name for the temp file in the map
            originalFileNames.put(fileToCompare, originalFileName);

            try {
                // For OGF files, apply special 15-20% rule
                performOgfComparison(originalFileName, fileToCompare, referencePrices, reportItems, originalFileNames);
            } finally {
                // Clean up the temporary file
                FileProccessor.cleanUpTempFiles(Collections.singletonList(fileToCompare));
                originalFileNames.remove(fileToCompare);
            }
        } else {
            // For non-OGF files, use Compare at price if available, otherwise use Price
            performRegularComparison(originalFileName, locationFile, referencePrices, reportItems, originalFileNames);
        }
    }

    // UPDATED: Added originalFileNames parameter (though not used in this method, for consistency)
    private static void performOgfComparison(String originalFileName, File file, Map<String, Double> refPrices,
                                             Map<String, ReferenceItem> reportItems,
                                             Map<File, String> originalFileNames) throws IOException {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            ColumnIndices locIndices = findHeaderIndices(sheet);
            if (locIndices == null) {
                System.err.println("Skipping " + originalFileName + ": Could not find required headers.");
                return;
            }

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String sku = getCellValue(row.getCell(locIndices.skuCol())).toUpperCase();

                // Get both Price and Compare at price
                Double locationPrice = getEnhancedNumericCellValue(row.getCell(locIndices.priceCol()));
                Double compareAtPrice = null;

                // Check if Compare at price column exists and has value
                if (locIndices.comparedPriceCol() >= 0) {
                    compareAtPrice = getEnhancedNumericCellValue(row.getCell(locIndices.comparedPriceCol()));
                }

                // NEW: Get available stock
                Integer availableStock = null;
                if (locIndices.availableCol() >= 0) {
                    availableStock = getIntegerCellValue(row.getCell(locIndices.availableCol()));
                }

                if (!sku.isEmpty() && reportItems.containsKey(sku)) {
                    Double referencePrice = refPrices.get(sku);
                    ReferenceItem item = reportItems.get(sku);

                    // Store both price values using original file name
                    item.locationPrices().put(originalFileName, locationPrice);
                    if (compareAtPrice != null) {
                        item.locationCompareAtPrices().put(originalFileName, compareAtPrice);
                    }

                    // NEW: Store available stock
                    if (availableStock != null) {
                        item.locationStock().put(originalFileName, availableStock);
                    }

                    // For OGF, use Compare at price if available, otherwise use Price
                    Double priceToUse = (compareAtPrice != null) ? compareAtPrice : locationPrice;

                    // Store which price we actually used for comparison
                    item.locationPricesUsed().put(originalFileName, priceToUse);

                    // Calculate and store Compare at price difference (if both prices exist)
                    if (compareAtPrice != null && locationPrice != null) {
                        double compareAtDifference = compareAtPrice - locationPrice;
                        item.compareAtPriceDifferences().put(originalFileName, compareAtDifference);
                    }

                    // For OGF files, check if price is LESS THAN 15% of reference price
                    if (priceToUse != null && referencePrice != null && referencePrice > 0) {
                        double percentageDiff = ((priceToUse - referencePrice) / referencePrice) * 100;

                        // Check if LESS THAN 15%
                        if (percentageDiff < 15.0) {
                            // UPDATED: Remove exact percentage from remark
                            String priceType = (compareAtPrice != null) ? "Compare at price" : "Price";
                            String discrepancy = String.format("%s: Below 15%% range (%s)",
                                    originalFileName, priceType);
                            item.discrepancies().add(discrepancy);
                        }
                    }
                }
            }
        }
    }

    // UPDATED: Added originalFileNames parameter (though not used in this method, for consistency)
    private static void performRegularComparison(String originalFileName, File file, Map<String, Double> refPrices,
                                                 Map<String, ReferenceItem> reportItems,
                                                 Map<File, String> originalFileNames) throws IOException {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            ColumnIndices locIndices = findHeaderIndices(sheet);
            if (locIndices == null) {
                System.err.println("Skipping " + originalFileName + ": Could not find required headers.");
                return;
            }

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String sku = getCellValue(row.getCell(locIndices.skuCol())).toUpperCase();

                // Get both Price and Compare at price
                Double locationPrice = getEnhancedNumericCellValue(row.getCell(locIndices.priceCol()));
                Double compareAtPrice = null;

                // Check if Compare at price column exists and has value
                if (locIndices.comparedPriceCol() >= 0) {
                    compareAtPrice = getEnhancedNumericCellValue(row.getCell(locIndices.comparedPriceCol()));
                }

                // NEW: Get available stock
                Integer availableStock = null;
                if (locIndices.availableCol() >= 0) {
                    availableStock = getIntegerCellValue(row.getCell(locIndices.availableCol()));
                }

                if (!sku.isEmpty() && reportItems.containsKey(sku)) {
                    Double referencePrice = refPrices.get(sku);
                    ReferenceItem item = reportItems.get(sku);

                    // Store both price values using original file name
                    item.locationPrices().put(originalFileName, locationPrice);
                    if (compareAtPrice != null) {
                        item.locationCompareAtPrices().put(originalFileName, compareAtPrice);
                    }

                    // NEW: Store available stock
                    if (availableStock != null) {
                        item.locationStock().put(originalFileName, availableStock);
                    }

                    // For non-OGF files, use Compare at price if available, otherwise use Price
                    Double priceToUse = (compareAtPrice != null) ? compareAtPrice : locationPrice;

                    // Store which price we actually used for comparison
                    item.locationPricesUsed().put(originalFileName, priceToUse);

                    // Calculate and store Compare at price difference (if both prices exist)
                    if (compareAtPrice != null && locationPrice != null) {
                        double compareAtDifference = compareAtPrice - locationPrice;
                        item.compareAtPriceDifferences().put(originalFileName, compareAtDifference);
                    }

                    // Compare the selected price with reference price
                    if (priceToUse != null && referencePrice != null &&
                            Math.abs(priceToUse - referencePrice) > 0.01) {
                        double difference = priceToUse - referencePrice;
                        String sign = difference > 0 ? "+" : "-";
                        String priceType = (compareAtPrice != null) ? "Compare at price" : "Price";
                        String discrepancy = String.format("%s: %sRs.%.2f (%s)",
                                originalFileName, sign, Math.abs(difference), priceType);
                        item.discrepancies().add(discrepancy);
                    }
                }
            }
        }
    }

    // NEW: Method to calculate total stock for each item
    private static void calculateTotalStockForItems(Map<String, ReferenceItem> reportItems) {
        for (ReferenceItem item : reportItems.values()) {
            int totalStock = 0;

            // Sum up stock from all locations
            for (Integer stock : item.locationStock().values()) {
                if (stock != null) {
                    totalStock += stock;
                }
            }

            item.setTotalStock(totalStock);
        }
    }

    private static void calculateStatusForItems(Map<String, ReferenceItem> reportItems, List<String> locationFileNames) {
        for (ReferenceItem item : reportItems.values()) {
            List<String> ogfDifferences = new ArrayList<>();
            List<String> nonOgfDifferences = new ArrayList<>();
            List<String> differenceFiles = new ArrayList<>();
            List<String> compareAtDiffFiles = new ArrayList<>();

            // Analyze Compare at price consistency across files
            Map<String, Double> compareAtPrices = item.locationCompareAtPrices();
            boolean hasCompareAtPrices = !compareAtPrices.isEmpty();
            List<String> compareAtConsistencyIssues = new ArrayList<>();

            if (hasCompareAtPrices && compareAtPrices.size() > 1) {
                // Check for differences in Compare at price margins across files
                analyzeCompareAtPriceMargins(item, compareAtConsistencyIssues, compareAtDiffFiles);
            }

            // Track OGF price and percentage difference
            Double ogfPrice = null;
            double ogfPercentageDiff = 0.0;
            String ogfFileName = "";

            // Check ALL location files for price differences using the actual prices used for comparison
            for (String fileName : locationFileNames) {
                Double priceUsed = item.locationPricesUsed().get(fileName);
                boolean isOgfFile = fileName.toLowerCase().contains("ogf");

                if (priceUsed != null) {
                    if (isOgfFile) {
                        // For OGF files, calculate percentage difference
                        double percentageDiff = ((priceUsed - item.referencePrice()) / item.referencePrice()) * 100;

                        // Check if LESS THAN 15%
                        if (percentageDiff < 15.0) {
                            differenceFiles.add(fileName);
                            ogfFileName = fileName;
                            ogfPrice = priceUsed;
                            ogfPercentageDiff = percentageDiff;

                            // Add to discrepancy list if not already there
                            String discrepancyKey = fileName + ":";
                            boolean hasDiscrepancy = item.discrepancies().stream()
                                    .anyMatch(d -> d.startsWith(discrepancyKey));

                            if (!hasDiscrepancy) {
                                // UPDATED: Remove exact percentage from remark
                                Double compareAtPrice = item.locationCompareAtPrices().get(fileName);
                                String priceType = (compareAtPrice != null) ? "Compare at price" : "Price";
                                String discrepancy = String.format("%s: Below 15%% range (%s)",
                                        fileName, priceType);
                                ogfDifferences.add(discrepancy);
                                item.discrepancies().add(discrepancy);
                            }
                        }
                    } else {
                        // For non-OGF files, check exact match
                        if (Math.abs(priceUsed - item.referencePrice()) > 0.01) {
                            differenceFiles.add(fileName);

                            // Add to discrepancy list if not already there
                            String discrepancyKey = fileName + ":";
                            boolean hasDiscrepancy = item.discrepancies().stream()
                                    .anyMatch(d -> d.startsWith(discrepancyKey));

                            if (!hasDiscrepancy) {
                                double difference = priceUsed - item.referencePrice();
                                String sign = difference > 0 ? "+" : "-";
                                Double compareAtPrice = item.locationCompareAtPrices().get(fileName);
                                String priceType = (compareAtPrice != null) ? "Compare at price" : "Price";
                                String discrepancy = String.format("%s: %sRs.%.2f (%s)",
                                        fileName, sign, Math.abs(difference), priceType);
                                nonOgfDifferences.add(discrepancy);
                                item.discrepancies().add(discrepancy);
                            }
                        }
                    }

                    // Track OGF specifically for percentage calculation
                    if (isOgfFile && ogfPrice == null) {
                        ogfFileName = fileName;
                        ogfPrice = priceUsed;
                        if (ogfPrice != null && item.referencePrice() > 0) {
                            ogfPercentageDiff = ((ogfPrice - item.referencePrice()) / item.referencePrice()) * 100;
                        }
                    }
                }
            }

            // Also check existing discrepancies to ensure we capture everything
            for (String discrepancy : item.discrepancies()) {
                String fileName = discrepancy.split(":")[0].trim();
                if (fileName.toLowerCase().contains("ogf")) {
                    if (!ogfDifferences.contains(discrepancy)) {
                        ogfDifferences.add(discrepancy);
                    }
                } else {
                    if (!nonOgfDifferences.contains(discrepancy)) {
                        nonOgfDifferences.add(discrepancy);
                    }
                }
            }

            String status;
            List<String> statusReasons = new ArrayList<>();

            // Check if other files have differences (excluding OGF)
            boolean hasNonOgfDifferences = !nonOgfDifferences.isEmpty();
            boolean hasCompareAtInconsistency = !compareAtConsistencyIssues.isEmpty();
            boolean hasAnyDifferences = !ogfDifferences.isEmpty() || hasNonOgfDifferences || hasCompareAtInconsistency;

            if (!hasAnyDifferences) {
                status = "Good";
                statusReasons.add("No price differences found");
            } else if (hasCompareAtInconsistency) {
                status = "Bad";
                statusReasons.add("Inconsistent Compare at price margins across files");
                statusReasons.addAll(compareAtConsistencyIssues);
            } else if (!ogfDifferences.isEmpty()) {
                // OGF has differences - check if within 15-20% range
                if (ogfPercentageDiff >= 15.0 && ogfPercentageDiff <= 20.0) {
                    // OGF difference is acceptable (15-20%), but check other files
                    if (hasNonOgfDifferences) {
                        status = "Bad";
                        // UPDATED: Remove exact percentage from status reason
                        statusReasons.add("OGF within acceptable range (15-20%)");
                        statusReasons.add("Other files have differences");
                        statusReasons.addAll(nonOgfDifferences);
                    } else {
                        status = "Good";
                        // UPDATED: Remove exact percentage from status reason
                        statusReasons.add("OGF within acceptable range (15-20%)");
                    }
                } else {
                    // OGF difference is outside acceptable range (now only below 15%)
                    status = "Bad";
                    // UPDATED: Remove exact percentage from status reason
                    statusReasons.add("OGF price below acceptable range (less than 15%)");
                    if (hasNonOgfDifferences) {
                        statusReasons.add("Other files also have differences");
                        statusReasons.addAll(nonOgfDifferences);
                    }
                }
            } else {
                // Only non-OGF differences
                status = "Bad";
                statusReasons.add("Non-OGF files have differences:");
                statusReasons.addAll(nonOgfDifferences);
            }

            item.setStatus(status);
            item.setStatusReason(String.join(" | ", statusReasons));

            // NEW: Store separated discrepancies in their respective lists
            item.ogfDiscrepancies().clear();
            item.ogfDiscrepancies().addAll(ogfDifferences);

            item.nonOgfDiscrepancies().clear();
            item.nonOgfDiscrepancies().addAll(nonOgfDifferences);

            // Create consolidated basic remarks (we'll keep this logic but won't use it in the report)
            List<String> consolidatedRemarks = new ArrayList<>();

            // Check for price differences in non-OGF files
            List<String> nonOgfDifferenceFiles = new ArrayList<>();
            for (String fileName : locationFileNames) {
                if (fileName.toLowerCase().contains("ogf")) continue;

                Double priceUsed = item.locationPricesUsed().get(fileName);

                if (priceUsed != null && Math.abs(priceUsed - item.referencePrice()) > 0.01) {
                    nonOgfDifferenceFiles.add(fileName);
                }
            }

            if (!nonOgfDifferenceFiles.isEmpty()) {
                consolidatedRemarks.add("Price difference in: " + String.join(", ", nonOgfDifferenceFiles));
            }

            if (!compareAtDiffFiles.isEmpty()) {
                consolidatedRemarks.add("Compare at price difference in: " + String.join(", ", compareAtDiffFiles));
            }

            // NEW: Generate simple difference explanation
            String simpleExplanation = generateSimpleExplanation(item, ogfDifferences, nonOgfDifferences,
                    compareAtDiffFiles, nonOgfDifferenceFiles,
                    ogfPercentageDiff, status);
            item.setDifferenceExplanation(simpleExplanation);

            item.discrepancies().sort((d1, d2) -> {
                boolean d1IsOgf = d1.toLowerCase().contains("ogf");
                boolean d2IsOgf = d2.toLowerCase().contains("ogf");
                if (d1IsOgf && !d2IsOgf) return -1;
                if (!d1IsOgf && d2IsOgf) return 1;
                return d1.compareTo(d2);
            });
        }
    }

    // NEW: Generate simple explanation for differences
    private static String generateSimpleExplanation(ReferenceItem item,
                                                    List<String> ogfDifferences,
                                                    List<String> nonOgfDifferences,
                                                    List<String> compareAtDiffFiles,
                                                    List<String> nonOgfDifferenceFiles,
                                                    double ogfPercentageDiff,
                                                    String status) {
        if (status.equals("Good")) {
            return "All prices match correctly";
        }

        List<String> explanations = new ArrayList<>();

        // Check OGF differences
        if (!ogfDifferences.isEmpty()) {
            if (ogfPercentageDiff < 15.0) {
                // UPDATED: Remove exact percentage from explanation
                explanations.add("OGF price below 15% range");
            }
        }

        // Check non-OGF differences
        if (!nonOgfDifferenceFiles.isEmpty()) {
            explanations.add("Price difference in: " + String.join(", ", nonOgfDifferenceFiles));
        }

        // Check compare at price margin differences
        if (!compareAtDiffFiles.isEmpty()) {
            explanations.add("Compare at price difference in: " + String.join(", ", compareAtDiffFiles));
        }

        if (explanations.isEmpty()) {
            return "Check detailed discrepancies";
        }

        return String.join("; ", explanations);
    }

    private static void analyzeCompareAtPriceMargins(ReferenceItem item, List<String> consistencyIssues, List<String> compareAtDiffFiles) {
        Map<String, Double> compareAtPrices = item.locationCompareAtPrices();
        Map<String, Double> regularPrices = item.locationPrices();

        // Calculate margins (difference between compare at price and regular price) for each file
        Map<String, Double> margins = new HashMap<>();
        for (Map.Entry<String, Double> entry : compareAtPrices.entrySet()) {
            String fileName = entry.getKey();
            Double compareAtPrice = entry.getValue();
            Double regularPrice = regularPrices.get(fileName);

            if (regularPrice != null && compareAtPrice != null) {
                double margin = compareAtPrice - regularPrice;
                margins.put(fileName, margin);
            }
        }

        // Check if margins are consistent across files
        if (margins.size() > 1) {
            // Find the most common margin (mode)
            Map<Double, Integer> frequency = new HashMap<>();
            for (Double margin : margins.values()) {
                frequency.put(margin, frequency.getOrDefault(margin, 0) + 1);
            }

            // Find the mode (most frequent margin)
            Double modeMargin = null;
            int maxCount = 0;
            for (Map.Entry<Double, Integer> entry : frequency.entrySet()) {
                if (entry.getValue() > maxCount) {
                    maxCount = entry.getValue();
                    modeMargin = entry.getKey();
                }
            }

            // If all margins are the same, no issue
            if (frequency.size() > 1) {
                // Check each file against the mode margin
                for (Map.Entry<String, Double> entry : margins.entrySet()) {
                    String fileName = entry.getKey();
                    double margin = entry.getValue();
                    double difference = margin - modeMargin;

                    // If this file's margin is significantly different from the mode
                    if (Math.abs(difference) > 1.0) {
                        // Include file name in the consistency issue
                        String direction = difference > 0 ? "higher" : "lower";
                        String issue = String.format("%s: Rs.%.2f %s margin", fileName, Math.abs(difference), direction);
                        consistencyIssues.add(issue);
                        compareAtDiffFiles.add(fileName);
                    }
                }

                if (!consistencyIssues.isEmpty()) {
                    // File names are now included in the consistency issues
                    String discrepancy = "Varying Compare at price margins: " + String.join(" | ", consistencyIssues);
                    item.discrepancies().add(discrepancy);
                }
            }
        }
    }

    private static void writeComparisonReport(File outputFile, Map<String, ReferenceItem> reportItems,
                                              List<String> locationFileNames) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Price Comparison Report");

            // Create header row
            Row headerRow = sheet.createRow(0);
            int colIndex = 0;
            headerRow.createCell(colIndex++).setCellValue("SKU");
            headerRow.createCell(colIndex++).setCellValue("Product Name");
            headerRow.createCell(colIndex++).setCellValue("Stock Available"); // NEW: Stock column
            headerRow.createCell(colIndex++).setCellValue("Reference Price");

            Map<String, Integer> locationColumnMap = new HashMap<>();
            for (String locName : locationFileNames) {
                headerRow.createCell(colIndex).setCellValue(locName.replace(".xlsx", "").replace(".xls", ""));
                locationColumnMap.put(locName, colIndex);
                colIndex++;
            }

            // UPDATED: Added separate columns for OGF and Non-OGF differences
            headerRow.createCell(colIndex++).setCellValue("Status");
            int statusColIndex = colIndex - 1;

            headerRow.createCell(colIndex++).setCellValue("Difference Explanation");
            int explanationColIndex = colIndex - 1;

            // NEW: Separate columns for OGF and Non-OGF differences
            headerRow.createCell(colIndex++).setCellValue("OGF Differences");
            int ogfDifferencesColIndex = colIndex - 1;

            headerRow.createCell(colIndex++).setCellValue("Non-OGF Differences");
            int nonOgfDifferencesColIndex = colIndex - 1;

            // Data rows
            int rowNum = 1;
            for (ReferenceItem item : reportItems.values()) {
                Row row = sheet.createRow(rowNum++);

                row.createCell(0).setCellValue(item.sku());
                row.createCell(1).setCellValue(item.productName());
                row.createCell(2).setCellValue(item.totalStock()); // NEW: Total stock
                row.createCell(3).setCellValue(item.referencePrice());

                for (String locName : locationFileNames) {
                    int currentLocCol = locationColumnMap.get(locName);
                    // Display the price that was actually used for comparison
                    Double priceUsed = item.locationPricesUsed().get(locName);
                    if (priceUsed != null) {
                        row.createCell(currentLocCol).setCellValue(priceUsed);
                    } else {
                        row.createCell(currentLocCol).setCellValue("N/A");
                    }
                }

                row.createCell(statusColIndex).setCellValue(item.status());
                row.createCell(explanationColIndex).setCellValue(item.differenceExplanation());

                // UPDATED: Separate OGF and Non-OGF differences
                String ogfDifferences = item.ogfDiscrepancies().isEmpty()
                        ? "No OGF differences"
                        : String.join("\n", item.ogfDiscrepancies());
                row.createCell(ogfDifferencesColIndex).setCellValue(ogfDifferences);

                String nonOgfDifferences = item.nonOgfDiscrepancies().isEmpty()
                        ? "No differences"
                        : String.join("\n", item.nonOgfDiscrepancies());
                row.createCell(nonOgfDifferencesColIndex).setCellValue(nonOgfDifferences);
            }

            // Auto-size columns
            for (int i = 0; i < colIndex; i++) {
                sheet.autoSizeColumn(i);
            }

            try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
                workbook.write(fileOut);
            }
        }
    }

    // --- Enhanced Cell Value Methods ---
    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        try {
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue().trim();
                case NUMERIC:
                    return String.valueOf(cell.getNumericCellValue());
                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());
                case FORMULA:
                    try {
                        return String.valueOf(cell.getNumericCellValue());
                    } catch (Exception e) {
                        return cell.getStringCellValue().trim();
                    }
                default:
                    return "";
            }
        } catch (Exception e) {
            return "";
        }
    }

    private static Double getNumericCellValue(Cell cell) {
        if (cell == null) return null;
        try {
            switch (cell.getCellType()) {
                case NUMERIC:
                    return cell.getNumericCellValue();
                case FORMULA:
                    return cell.getNumericCellValue();
                default:
                    return null;
            }
        } catch (Exception e) {
            return null;
        }
    }

    // NEW: Method to get integer cell value for stock
    private static Integer getIntegerCellValue(Cell cell) {
        if (cell == null) return null;
        try {
            switch (cell.getCellType()) {
                case NUMERIC:
                    return (int) cell.getNumericCellValue();
                case FORMULA:
                    try {
                        return (int) cell.getNumericCellValue();
                    } catch (Exception e) {
                        // If formula returns string, try to parse it
                        String stringValue = cell.getStringCellValue();
                        return parseFormattedInteger(stringValue);
                    }
                case STRING:
                    String stringValue = cell.getStringCellValue().trim();
                    return parseFormattedInteger(stringValue);
                default:
                    return null;
            }
        } catch (Exception e) {
            return null;
        }
    }

    // NEW: Helper method to parse formatted integers for stock
    private static Integer parseFormattedInteger(String value) {
        if (value == null || value.isEmpty()) return null;

        try {
            // Remove commas and any non-digit characters (except minus sign)
            String cleanedValue = value.replaceAll("[^\\d-]", "").trim();
            if (cleanedValue.isEmpty()) return null;

            return Integer.parseInt(cleanedValue);
        } catch (NumberFormatException e) {
            return null;
        }
    }

    // Enhanced method that handles both numeric cells and string cells with formatted numbers
    private static Double getEnhancedNumericCellValue(Cell cell) {
        if (cell == null) return null;

        try {
            switch (cell.getCellType()) {
                case NUMERIC:
                    return cell.getNumericCellValue();
                case FORMULA:
                    try {
                        return cell.getNumericCellValue();
                    } catch (Exception e) {
                        // If formula returns string, try to parse it
                        String stringValue = cell.getStringCellValue();
                        return parseFormattedNumber(stringValue);
                    }
                case STRING:
                    String stringValue = cell.getStringCellValue().trim();
                    return parseFormattedNumber(stringValue);
                default:
                    return null;
            }
        } catch (Exception e) {
            return null;
        }
    }

    // Helper method to parse formatted numbers like "8,000.00"
    private static Double parseFormattedNumber(String value) {
        if (value == null || value.isEmpty()) return null;

        try {
            // Remove commas and any currency symbols, then parse
            String cleanedValue = value.replaceAll("[,\\sRs\\p{Sc}]", "").trim();
            if (cleanedValue.isEmpty()) return null;

            return Double.parseDouble(cleanedValue);
        } catch (NumberFormatException e) {
            return null;
        }
    }
}