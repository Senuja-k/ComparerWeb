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
        private final double referenceCompareAtPrice;
        private final List<String> discrepancies;
        private final List<String> basicRemarks;
        private final Map<String, Double> locationPrices;
        private final Map<String, Double> locationCompareAtPrices;
        private final Map<String, Double> compareAtPriceDifferences;
        private final Map<String, Double> locationPricesUsed; // Track which price was actually used
        private final Map<String, Integer> locationStock; // NEW: Track stock per location
        private String status;
        private double ogfPercentageDiff;
        private String ogfPercentageRemark;
        private String statusReason;
        private String differenceExplanation; // NEW: Simple explanation field
        private int totalStock; // NEW: Total stock across all locations
        private final List<String> ogfDiscrepancies; // NEW: Separate OGF discrepancies
        private final List<String> nonOgfDiscrepancies; // NEW: Separate non-OGF discrepancies
        private final Map<String, Double> locationDiscountPercentages; // NEW FIELD 1
        private final List<String> discountDiscrepancies;

        public ReferenceItem(String sku, String productName, double referencePrice, double referenceCompareAtPrice,
                             List<String> discrepancies, Map<String, Double> locationPrices, String status) {
            this.sku = sku;
            this.productName = productName;
            this.referencePrice = referencePrice;
            this.referenceCompareAtPrice = referenceCompareAtPrice;
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
            this.locationDiscountPercentages = new HashMap<>();
            this.discountDiscrepancies = new ArrayList<>();
        }

        public String sku() { return sku; }
        public String productName() { return productName; }
        public double referencePrice() { return referencePrice; }
        public double referenceCompareAtPrice() { return referenceCompareAtPrice; }
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
        public double ogfPercentageDiff() { return ogfPercentageDiff; }
        public void setOgfPercentageDiff(double ogfPercentageDiff) { this.ogfPercentageDiff = ogfPercentageDiff; }
        public String ogfPercentageRemark() { return ogfPercentageRemark; }
        public void setOgfPercentageRemark(String ogfPercentageRemark) { this.ogfPercentageRemark = ogfPercentageRemark; }
        public Map<String, Double> locationDiscountPercentages() { return locationDiscountPercentages; }
        public List<String> discountDiscrepancies() { return discountDiscrepancies; }
    }

    // UPDATED: Public Entry Point with original file names map
    public static void generateReport(File referenceFile, List<File> locationFiles, File outputFile,
                                      Map<File, String> originalFileNames) throws IOException {
        Map<String, Double> referencePrices = new HashMap<>();
        Map<String, Double> referenceCompareAtPrices = new HashMap<>();
        Map<String, ReferenceItem> reportItems = new LinkedHashMap<>();
        List<String> locationFileNames = new ArrayList<>();

        // UPDATED: Use the provided original file names map
        for (File file : locationFiles) {
            String originalName = getOriginalFileName(file, originalFileNames);
            locationFileNames.add(originalName);
        }

        //1. Read Reference File and Initialize Report Map
        ColumnIndices refIndices = readReferenceData(referenceFile, referencePrices, referenceCompareAtPrices, reportItems, locationFileNames, originalFileNames);

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
            } else if (header.equalsIgnoreCase(NAME_HEADER) ||
                    header.equalsIgnoreCase("Product") ||
                    header.equalsIgnoreCase("Product Name") ||
                    header.equalsIgnoreCase("Product Title")) {
                nameCol = index;
            } else if (header.equalsIgnoreCase(PRICE_HEADER)) {
                priceCol = index;
            } else if (header.equalsIgnoreCase(COMPARED_PRICE_HEADER) ||
                    header.equalsIgnoreCase("Compare At Price") ||
                    header.equalsIgnoreCase("Compare Price")) {
                comparedPriceCol = index;
            } else if (header.equalsIgnoreCase(AVAILABLE_HEADER) ||
                    header.equalsIgnoreCase("Stock") ||
                    header.equalsIgnoreCase("Quantity") ||
                    header.equalsIgnoreCase("Inventory Quantity")) { // NEW: Look for stock/quantity headers
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
                                                   Map<String, Double> referenceCompareAtPrices,
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

                // Get both Price and Compare at price
                Double price = getEnhancedNumericCellValue(row.getCell(indices.priceCol()));
                Double compareAtPrice = null;

                // Get Compare at price if column exists
                if (indices.comparedPriceCol() >= 0) {
                    compareAtPrice = getEnhancedNumericCellValue(row.getCell(indices.comparedPriceCol()));
                }

                // If compare at price is null or 0, it means no promotion
                if (compareAtPrice == null || compareAtPrice == 0) {
                    compareAtPrice = 0.0; // If no promotion, compare at price = regular price
                }


                if (!sku.isEmpty() && price != null) {
                    prices.put(sku, price);
                    referenceCompareAtPrices.put(sku, compareAtPrice);
                    Map<String, Double> locPrices = new HashMap<>();
                    locationFileNames.forEach(locName -> locPrices.put(locName, null));

                    ReferenceItem item = new ReferenceItem(sku, name, price, compareAtPrice, new ArrayList<>(), locPrices, "");

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
            // For OGF files, clean up SKUs first using FileProcessor
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
                    } else {
                        // If compare at price is null, set it to 0 (meaning no promotion)
                        item.locationCompareAtPrices().put(originalFileName, 0.0); // CHANGE #4: Show 0 instead of null
                    }

                    // NEW: Calculate and store discount percentage
                    if (compareAtPrice != null && locationPrice != null && compareAtPrice > 0) {
                        double discount = ((compareAtPrice - locationPrice) / compareAtPrice) * 100;

                        // Only store if there's a meaningful discount (> 0.5%)
                        if (Math.abs(discount) > 0.5) {
                            item.locationDiscountPercentages().put(originalFileName, discount);
                        }
                    }

                    // NEW: Store available stock
                    if (availableStock != null) {
                        item.locationStock().put(originalFileName, availableStock);
                    }

                    //For OGF, always use Price column for comparison (not Compare at price)
                    Double priceToUse = locationPrice;

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
                        if (percentageDiff < 22.0) {
                            String priceType = (compareAtPrice != null) ? "Compare at price" : "Price";
                            // Also added percentage value to the discrepancy message
                            String discrepancy = String.format("%s: Below 22%% range (%.2f%%) (%s)",
                                    originalFileName, percentageDiff, priceType);
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
                    } else {
                        // If compare at price is null, set it to 0 (meaning no promotion)
                        item.locationCompareAtPrices().put(originalFileName, 0.0); // CHANGE #4: Show 0 instead of null
                    }

                    // NEW: Calculate and store discount percentage
                    if (compareAtPrice != null && locationPrice != null && compareAtPrice > 0) {
                        double discount = ((compareAtPrice - locationPrice) / compareAtPrice) * 100;

                        // Only store if there's a meaningful discount (> 0.5%)
                        if (Math.abs(discount) > 0.5) {
                            item.locationDiscountPercentages().put(originalFileName, discount);
                        }
                    }

                    // NEW: Store available stock
                    if (availableStock != null) {
                        item.locationStock().put(originalFileName, availableStock);
                    }

                    // For non-OGF files, always use Price column for comparison (not Compare at price)
                    Double priceToUse = locationPrice;

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

    // NEW: Method to check discount consistency across companies
    private static void checkDiscountConsistency(ReferenceItem item, List<String> locationFileNames) {
        item.discountDiscrepancies().clear();

        // Track discount status for each company
        Map<String, Boolean> hasDiscount = new HashMap<>();
        Map<String, Double> discountValues = new HashMap<>();

        for (String locName : locationFileNames) {
            Double discount = item.locationDiscountPercentages().get(locName);
            boolean hasDisc = (discount != null && Math.abs(discount) > 0.5);
            hasDiscount.put(locName, hasDisc);
            if (hasDisc) {
                discountValues.put(locName, discount);
            }
        }

        // Count companies with/without discounts
        long withDiscountCount = hasDiscount.values().stream().filter(b -> b).count();
        long withoutDiscountCount = hasDiscount.size() - withDiscountCount;

        // Check consistency
        if (withDiscountCount > 0 && withoutDiscountCount > 0) {
            // Inconsistent - some have discounts, some don't
            List<String> discCompanies = new ArrayList<>();
            List<String> noDiscCompanies = new ArrayList<>();

            for (Map.Entry<String, Boolean> entry : hasDiscount.entrySet()) {
                String cleanName = entry.getKey().replace(".xlsx", "").replace(".xls", "");
                if (entry.getValue()) {
                    discCompanies.add(cleanName);
                } else {
                    noDiscCompanies.add(cleanName);
                }
            }

            String issue = String.format("Discount inconsistency: %s have discounts, %s don't",
                    discCompanies.size(), noDiscCompanies.size());
            item.discountDiscrepancies().add(issue);

            // Add details if needed
            if (!discCompanies.isEmpty()) {
                item.discountDiscrepancies().add(" Discounted in: " + String.join(", ", discCompanies));
            }
        }

        // Also check if discount percentages vary significantly (optional)
        if (discountValues.size() > 1) {
            // Calculate average discount
            double total = discountValues.values().stream().mapToDouble(Double::doubleValue).sum();
            double average = total / discountValues.size();

            // Check each discount against average
            for (Map.Entry<String, Double> entry : discountValues.entrySet()) {
                if (Math.abs(entry.getValue() - average) > 10.0) { // More than 10% difference
                    String cleanName = entry.getKey().replace(".xlsx", "").replace(".xls", "");
                    item.discountDiscrepancies().add(
                            String.format("%s: %.1f%% discount differs from others", cleanName, entry.getValue())
                    );
                }
            }
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
                        double referenceValue = item.referenceCompareAtPrice() > 0 ?
                                item.referenceCompareAtPrice() : item.referencePrice();
                        double percentageDiff = ((priceUsed - referenceValue) / referenceValue) * 100;

                        // NEW: Store OGF percentage for display
                        item.setOgfPercentageDiff(percentageDiff);

                        // NEW: Generate remark based on percentage
                        String ogfPercentageRemark;
                        if (percentageDiff < 22.0) {
                            ogfPercentageRemark = String.format("Below 22%% threshold (%.2f%%)", percentageDiff);
                        } else if (percentageDiff >= 22.0 && percentageDiff <= 25.0) {
                            ogfPercentageRemark = String.format("Within 22-25%% range (%.2f%%)", percentageDiff);
                        } else {
                            ogfPercentageRemark = String.format("Above 25%% (%.2f%%)", percentageDiff);
                        }
                        item.setOgfPercentageRemark(ogfPercentageRemark);

                        // Check if LESS THAN 22%
                        if (percentageDiff < 22.0) {
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
                                String discrepancy = String.format("%s: Below 22%% range (%s)",
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
                        if (ogfPrice != null) {
                            double referenceValue = item.referenceCompareAtPrice() > 0?
                                    item.referenceCompareAtPrice() : item.referencePrice();
                            if (referenceValue > 0) {
                                ogfPercentageDiff = ((ogfPrice - referenceValue) / referenceValue) * 100;
                            }
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

            // NEW: Check discount consistency (add this after analyzing price differences)
            checkDiscountConsistency(item, locationFileNames);

            String status;
            List<String> statusReasons = new ArrayList<>();

            // Check if other files have differences (excluding OGF)
            boolean hasNonOgfDifferences = !nonOgfDifferences.isEmpty();
            boolean hasCompareAtInconsistency = !compareAtConsistencyIssues.isEmpty();
            boolean hasAnyDifferences = !ogfDifferences.isEmpty() || hasNonOgfDifferences || hasCompareAtInconsistency;

            // NEW: Also check for discount discrepancies
            boolean hasDiscountIssues = !item.discountDiscrepancies().isEmpty();

            if (!hasAnyDifferences && !hasDiscountIssues) {
                status = "Good";
                statusReasons.add("No price differences found");
            } else if (hasCompareAtInconsistency) {
                status = "Bad";
                statusReasons.add("Inconsistent Compare at price margins across files");
                statusReasons.addAll(compareAtConsistencyIssues);
            } else if (!ogfDifferences.isEmpty()) {
                // OGF has differences - check if within 15-20% range
                if (ogfPercentageDiff >= 22.0) {
                    // OGF difference is acceptable (15-20%), but check other files
                    if (hasNonOgfDifferences) {
                        status = "Bad";
                        // UPDATED: Remove exact percentage from status reason
                        statusReasons.add("OGF within acceptable range");
                        statusReasons.add("Other files have differences");
                        statusReasons.addAll(nonOgfDifferences);
                    } else {
                        status = "Good";
                        // UPDATED: Remove exact percentage from status reason
                        statusReasons.add("OGF within acceptable range");
                    }
                } else {
                    // OGF difference is outside acceptable range (now only below 15%)
                    status = "Bad";
                    // UPDATED: Remove exact percentage from status reason
                    statusReasons.add("OGF price below acceptable range (less than 22%)");
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
            if (ogfPercentageDiff < 22.0) {
                // UPDATED: Remove exact percentage from explanation
                explanations.add(String.format("OGF price below 22%% threshold (%.2f%%)", ogfPercentageDiff));
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

    // Helper method to get discount summary
    private static String getDiscountSummary(ReferenceItem item, List<String> locationFileNames) {
        List<String> issues = new ArrayList<>();

        // Track which companies have discounts
        List<String> withDiscount = new ArrayList<>();
        List<String> withoutDiscount = new ArrayList<>();

        for (String locName : locationFileNames) {
            Double discount = item.locationDiscountPercentages().get(locName);
            String cleanName = locName.replace(".xlsx", "").replace(".xls", "");

            if (discount != null && Math.abs(discount) > 0.5) {
                withDiscount.add(cleanName);
            } else {
                withoutDiscount.add(cleanName);
            }
        }

        // Build summary
        if (!withDiscount.isEmpty() && !withoutDiscount.isEmpty()) {
            return String.format("Discounts not uniform: %s have discounts, %s don't",
                    String.join(", ", withDiscount),
                    String.join(", ", withoutDiscount));
        } else if (!withDiscount.isEmpty()) {
            return "All locations have discounts";
        } else {
            return "No discounts found";
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
            headerRow.createCell(colIndex++).setCellValue("Stock Available");
            headerRow.createCell(colIndex++).setCellValue("Reference Price");
            headerRow.createCell(colIndex++).setCellValue("Reference Compare at Price");

            // Create maps for column indices
            Map<String, Integer> sellingPriceColMap = new HashMap<>();
            Map<String, Integer> originalPriceColMap = new HashMap<>();
            Map<String, Integer> discountColMap = new HashMap<>();

            for (String locName : locationFileNames) {
                String displayName = locName.replace(".xlsx", "").replace(".xls", "");

                // Price column (always shown)
                headerRow.createCell(colIndex).setCellValue(displayName + " - Price"); //
                sellingPriceColMap.put(locName, colIndex);
                colIndex++;

                // Compare At Price column
                headerRow.createCell(colIndex).setCellValue(displayName + " - Compare at price"); //
                originalPriceColMap.put(locName, colIndex);
                colIndex++;

                // Discount % column
                headerRow.createCell(colIndex).setCellValue(displayName + " - Disc %");
                discountColMap.put(locName, colIndex);
                colIndex++;
            }

            // Status and other columns
            headerRow.createCell(colIndex++).setCellValue("Status");
            int statusColIndex = colIndex - 1;

            headerRow.createCell(colIndex++).setCellValue("OGF Percentage");
            int ogfPercentageColIndex = colIndex - 1;

            headerRow.createCell(colIndex++).setCellValue("Difference Explanation");
            int explanationColIndex = colIndex - 1;

            headerRow.createCell(colIndex++).setCellValue("OGF Differences");
            int ogfDifferencesColIndex = colIndex - 1;

            headerRow.createCell(colIndex++).setCellValue("Non-OGF Differences");
            int nonOgfDifferencesColIndex = colIndex - 1;

            // NEW: Discount issues column
            headerRow.createCell(colIndex++).setCellValue("Discount Issues");
            int discountIssuesColIndex = colIndex - 1;

            // Data rows
            int rowNum = 1;
            for (ReferenceItem item : reportItems.values()) {
                Row row = sheet.createRow(rowNum++);

                row.createCell(0).setCellValue(item.sku());
                row.createCell(1).setCellValue(item.productName());
                row.createCell(2).setCellValue(item.totalStock());
                row.createCell(3).setCellValue(item.referencePrice());
                row.createCell(4).setCellValue(item.referenceCompareAtPrice());

                // For each location, show all price info
                for (String locName : locationFileNames) {
                    int sellingCol = sellingPriceColMap.get(locName);
                    int originalCol = originalPriceColMap.get(locName);
                    int discountCol = discountColMap.get(locName);

                    // Get prices from existing fields
                    Double sellingPrice = item.locationPrices().get(locName);
                    Double originalPrice = item.locationCompareAtPrices().get(locName);
                    Double discount = item.locationDiscountPercentages().get(locName);

                    // Fill selling price (always show if available)
                    if (sellingPrice != null) {
                        row.createCell(sellingCol).setCellValue(sellingPrice);
                    } else {
                        row.createCell(sellingCol).setCellValue("N/A");
                    }

                    // Fill compare at price
                    if (originalPrice != null) {
                        row.createCell(originalCol).setCellValue(originalPrice); // CHANGE #7: Show 0 if no promotion
                    } else {
                        row.createCell(originalCol).setCellValue(0.0); // CHANGE #7: Show 0 instead of N/A
                    }

                    // Fill discount percentage
                    if (discount != null && Math.abs(discount) > 0.5) {
                        row.createCell(discountCol).setCellValue(String.format("%.1f%%", discount));
                    } else {
                        row.createCell(discountCol).setCellValue("");
                    }
                }

                // Status and other info
                row.createCell(statusColIndex).setCellValue(item.status());
                row.createCell(ogfPercentageColIndex).setCellValue(String.format("%.2f%%", item.ogfPercentageDiff()));
                row.createCell(explanationColIndex).setCellValue(item.differenceExplanation());

                // Discrepancies
                String ogfDifferences = item.ogfDiscrepancies().isEmpty()
                        ? "No OGF differences"
                        : String.join("\n", item.ogfDiscrepancies());
                row.createCell(ogfDifferencesColIndex).setCellValue(ogfDifferences);

                String nonOgfDifferences = item.nonOgfDiscrepancies().isEmpty()
                        ? "No differences"
                        : String.join("\n", item.nonOgfDiscrepancies());
                row.createCell(nonOgfDifferencesColIndex).setCellValue(nonOgfDifferences);

                // Show discount issues
                String discountIssues = item.discountDiscrepancies().isEmpty()
                        ? "No discount issues"
                        : String.join("\n", item.discountDiscrepancies());
                row.createCell(discountIssuesColIndex).setCellValue(discountIssues);
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