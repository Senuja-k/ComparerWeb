// File: SKUComparerLogic.java
package com.example.InventoryComparer.logic;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;
import java.util.stream.Collectors;

public class SKUComparerLogic {

    // Set this flag based on whether the "OGF Rules" box is checked.
    private static boolean useOgfRules = false;

    // DataFormatter to handle all cell types consistently
    private static final DataFormatter dataFormatter = new DataFormatter();

    // helper: consistent OGF detection (case-insensitive)
    private static boolean isOgfName(String name) {
        if (name == null) return false;
        String n = name.toLowerCase();
        // Accept both the temp prefix *and* any name containing "ogf"
        return n.startsWith("temp_sku_ogf") || n.contains("ogf");
    }

    // --- Data Structures ---

    static class ItemSourceData {
        final String rawSku, rawBarcode, rawProductName;
        final String cleanSku;

        private static String cleanSkuForComparison(String sku, boolean isTempOgfFile) {
            return sku == null ? "" : sku.trim();
        }

        boolean isDuplicateInSource;
        boolean hasShortBarcode;
        String ogfRemark;

        // NEW: Store which specific values are duplicates
        boolean isSkuDuplicate;
        boolean isBarcodeDuplicate;

        ItemSourceData(String rawSku, String rawBarcode, String rawProductName, boolean isDuplicate, boolean hasShortBarcode, String ogfRemark, boolean isTempOgfFile) {
            this.rawSku = rawSku == null ? "" : rawSku.trim();
            this.rawBarcode = rawBarcode == null ? "" : rawBarcode.trim();
            this.rawProductName = rawProductName == null ? "" : rawProductName.trim();
            this.isDuplicateInSource = isDuplicate;

            // FIXED: Ensure short barcode flag is set correctly for <3 characters
            this.hasShortBarcode = hasShortBarcode || (!this.rawBarcode.isEmpty() && this.rawBarcode.length() < 3);

            // ENHANCED: Auto-detect OGF remark if not provided and this is an OGF file
            if (isTempOgfFile && (ogfRemark == null || ogfRemark.trim().isEmpty())) {
                this.ogfRemark = SKUComparerLogic.detectOgfRemarkFromSku(rawSku); // Use the raw SKU before trimming
            } else {
                this.ogfRemark = ogfRemark == null ? "" : ogfRemark.trim();
            }

            this.cleanSku = cleanSkuForComparison(this.rawSku, isTempOgfFile);

            // Initialize duplicate flags
            this.isSkuDuplicate = false;
            this.isBarcodeDuplicate = false;
        }
    }
    // Main data model for an item across all files
    static class Item {
        final String primarySku, primaryBarcode;
        String consolidatedProductName = "";

        // NEW: Track which source provided the primary SKU
        String primarySkuSource = "";

        // REINTRODUCED: Flag to track if the item is part of the Cosmetics Group
        boolean isCosmeticsGroupItem = false;
        // RETAINED: Flag to track if the item originated from an OGF Location file
        boolean isOgfGroupItem = false;

        Set<String> allSkus = new HashSet<>();
        Set<String> allBarcodes = new HashSet<>();
        Map<String, ItemSourceData> sourceData = new HashMap<>();

        List<String> finalRemarks = new ArrayList<>();
        String simpleStatus = "";
        String conflictStatus = "";

        Item(String sku, String barcode) {
            this.primarySku = sku == null ? "" : sku.trim();
            this.primaryBarcode = barcode == null ? "" : barcode.trim();
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (!(o instanceof Item)) return false;
            Item i = (Item) o;
            return primarySku.equalsIgnoreCase(i.primarySku) && primaryBarcode.equalsIgnoreCase(i.primaryBarcode);
        }

        @Override
        public int hashCode() {
            return Objects.hash(primarySku.toLowerCase(), primaryBarcode.toLowerCase());
        }

        public void addSourceData(String sourceName, ItemSourceData data) {
            if (!sourceData.containsKey(sourceName)) {
                sourceData.put(sourceName, data);
            }

            // FIXED: Enhanced SKU source tracking - set primarySkuSource if this source matches our primary SKU
            if (data.rawSku != null && !data.rawSku.trim().isEmpty()) {
                String sourceSku = data.rawSku.trim();
                String currentPrimarySku = this.primarySku.trim();

                // If this source's SKU matches our primary SKU (case-insensitive)
                if (sourceSku.equalsIgnoreCase(currentPrimarySku)) {
                    // If we don't have a source yet, set it to this one
                    if (primarySkuSource.isEmpty()) {
                        primarySkuSource = sourceName;
                        System.out.println("DEBUG: Set primarySkuSource to '" + sourceName + "' for SKU: " + currentPrimarySku);
                    }
                }
            }

            if (!data.cleanSku.isEmpty()) this.allSkus.add(data.cleanSku);
            if (!data.rawBarcode.isEmpty()) this.allBarcodes.add(data.rawBarcode);
        }

        // ADDED: Method to set Cosmetics group status
        public void markAsCosmeticsGroupItem(String locationName) {
            if (locationName.toLowerCase().contains("cosmetics") || locationName.toLowerCase().contains("cos")) {
                this.isCosmeticsGroupItem = true;
            }
        }

        // RETAINED: Method to set OGF group status
        public void markAsOgfGroupItem(boolean isOgfFile) {
            if (isOgfFile) {
                this.isOgfGroupItem = true;
            }
        }

        public ItemSourceData getDataForLocation(String locationName) { return sourceData.get(locationName); }
        public boolean isPresentIn(String locationName) { return sourceData.containsKey(locationName); }
    }

    private static Map<Set<Item>, List<ItemSourceData>> readItems(File file, boolean isTempOgfFile, boolean skipInternalValidation) throws IOException {
        Set<Item> uniqueItems = new HashSet<>();
        List<ItemSourceData> duplicateSourceData = new ArrayList<>();
        Set<String> uniqueSkus = new HashSet<>();
        Set<String> uniqueBarcodes = new HashSet<>();

        if (!file.exists()) {
            return Map.of(uniqueItems, duplicateSourceData);
        }

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) return Map.of(uniqueItems, duplicateSourceData);

            int skuCol = -1, barcodeCol = -1, nameCol = -1, remarkCol = -1;
            for (Cell cell : headerRow) {
                if (cell != null) {
                    // Use DataFormatter to get cell value as string regardless of cell type
                    String value = dataFormatter.formatCellValue(cell).trim().toLowerCase();
                    if (value.contains("sku")) skuCol = cell.getColumnIndex();
                    if (value.contains("barcode")) barcodeCol = cell.getColumnIndex();
                    if (value.equalsIgnoreCase("product") || value.contains("title")) nameCol = cell.getColumnIndex();
                    if (value.contains("remark")) remarkCol = cell.getColumnIndex();
                }
            }

            if (skuCol == -1 && barcodeCol == -1)
                throw new IllegalArgumentException("Could not find SKU or Barcode columns in file: " + file.getName());

            // Enhanced duplicate tracking
            Map<String, List<Integer>> skuRowMap = new HashMap<>();
            Map<String, List<Integer>> barcodeRowMap = new HashMap<>();

            // First pass: collect ALL SKUs and barcodes with their row numbers
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                // Use DataFormatter to get cell values as strings
                String rawSku = skuCol >= 0 ? getFormattedCellValue(row.getCell(skuCol)) : "";
                String rawBarcode = barcodeCol >= 0 ? getFormattedCellValue(row.getCell(barcodeCol)) : "";

                // NEW: Skip placeholder barcode values
                if (!rawSku.isEmpty() && !isPlaceholderValue(rawSku)) {
                    String lowerSku = rawSku.toLowerCase();
                    skuRowMap.computeIfAbsent(lowerSku, k -> new ArrayList<>()).add(i + 1);
                }
                if (!rawBarcode.isEmpty() && !isPlaceholderValue(rawBarcode)) {
                    String lowerBarcode = rawBarcode.toLowerCase();
                    barcodeRowMap.computeIfAbsent(lowerBarcode, k -> new ArrayList<>()).add(i + 1);
                }
            }

            // Identify ALL duplicates (any value that appears more than once)
            Set<String> duplicateSkus = skuRowMap.entrySet().stream()
                    .filter(entry -> entry.getValue().size() > 1)
                    .map(Map.Entry::getKey)
                    .collect(Collectors.toSet());

            Set<String> duplicateBarcodes = barcodeRowMap.entrySet().stream()
                    .filter(entry -> entry.getValue().size() > 1)
                    .map(Map.Entry::getKey)
                    .collect(Collectors.toSet());

            // Second pass: process ALL rows and mark ALL duplicates
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                // CAPTURE ORIGINAL VALUES BEFORE ANY PROCESSING - using DataFormatter
                String originalRawSku = skuCol >= 0 ? getFormattedCellValue(row.getCell(skuCol)) : "";
                String rawSku = originalRawSku; // Keep original for OGF detection
                String rawBarcode = barcodeCol >= 0 ? getFormattedCellValue(row.getCell(barcodeCol)) : "";
                String rawProductName = nameCol >= 0 ? getFormattedCellValue(row.getCell(nameCol)) : "";
                String ogfRemark = remarkCol >= 0 ? getFormattedCellValue(row.getCell(remarkCol)) : "";

                // NEW: Detect OGF remark from ORIGINAL SKU before any cleaning
                if (isTempOgfFile && (ogfRemark == null || ogfRemark.trim().isEmpty())) {
                    ogfRemark = detectOgfRemarkFromSku(originalRawSku);
                    System.out.println("DEBUG: OGF Remark auto-detected - File: " + file.getName() +
                            ", Original SKU: '" + originalRawSku + "', Remark: '" + ogfRemark + "'");
                }

                if (!rawSku.isEmpty() || !rawBarcode.isEmpty()) {
                    String lowerSku = rawSku.toLowerCase();
                    String lowerBarcode = rawBarcode.toLowerCase();
                    boolean isDuplicate = false;
                    boolean isShortBarcode = false;
                    boolean isSkuDuplicate = false;
                    boolean isBarcodeDuplicate = false;

                    if (!skipInternalValidation) {
                        isShortBarcode = !rawBarcode.isEmpty() && rawBarcode.trim().length() < 3 && !isPlaceholderValue(rawBarcode);

                        // ENHANCED: Use the complete duplicate detection to mark ALL duplicates
                        // NEW: Skip placeholder values for duplicate detection
                        if (!rawSku.isEmpty() && !isPlaceholderValue(rawSku) && duplicateSkus.contains(lowerSku)) {
                            isDuplicate = true;
                            isSkuDuplicate = true;
                        }

                        if (!rawBarcode.isEmpty() && !isPlaceholderValue(rawBarcode) && duplicateBarcodes.contains(lowerBarcode)) {
                            isDuplicate = true;
                            isBarcodeDuplicate = true;
                        }

                        // Maintain original sets (but skip placeholders)
                        if (!rawSku.isEmpty() && !isPlaceholderValue(rawSku)) {
                            uniqueSkus.add(lowerSku);
                        }
                        if (!rawBarcode.isEmpty() && !isPlaceholderValue(rawBarcode)) {
                            uniqueBarcodes.add(lowerBarcode);
                        }
                    }

                    ItemSourceData tempSourceData = new ItemSourceData(
                            rawSku, rawBarcode, rawProductName,
                            isDuplicate, isShortBarcode, ogfRemark, isTempOgfFile
                    );

                    tempSourceData.isSkuDuplicate = isSkuDuplicate;
                    tempSourceData.isBarcodeDuplicate = isBarcodeDuplicate;

                    if (isDuplicate) {
                        ItemSourceData duplicateData = new ItemSourceData(
                                rawSku, rawBarcode, rawProductName,
                                true, isShortBarcode, ogfRemark, isTempOgfFile
                        );
                        duplicateSourceData.add(duplicateData);
                    }

                    Item newItem = new Item(tempSourceData.cleanSku, tempSourceData.rawBarcode);
                    newItem.addSourceData("TEMP_KEY", tempSourceData);
                    uniqueItems.add(newItem);
                }
            }

            // DEBUG: Print ALL duplicate information
            System.out.println("=== DUPLICATE DEBUG INFO for " + file.getName() + " ===");
            System.out.println("Total rows in sheet: " + sheet.getLastRowNum());
            System.out.println("Duplicate SKUs found: " + duplicateSkus.size());
            for (String sku : duplicateSkus) {
                List<Integer> rows = skuRowMap.get(sku);
                System.out.println("SKU '" + sku + "' appears " + rows.size() + " times in rows: " + rows);
            }
            System.out.println("Duplicate Barcodes found: " + duplicateBarcodes.size());
            for (String barcode : duplicateBarcodes) {
                List<Integer> rows = barcodeRowMap.get(barcode);
                System.out.println("Barcode '" + barcode + "' appears " + rows.size() + " times in rows: " + rows);
            }
            System.out.println("Total items marked as duplicates: " + duplicateSourceData.size());
            System.out.println("=====================================");

        }
        return Map.of(uniqueItems, duplicateSourceData);
    }

    // NEW: Helper method using DataFormatter to get cell value as string
    private static String getFormattedCellValue(Cell cell) {
        if (cell == null) return "";
        return dataFormatter.formatCellValue(cell).trim();
    }

    // NEW: Helper method to identify placeholder values that shouldn't be treated as duplicates
    private static boolean isPlaceholderValue(String value) {
        if (value == null || value.trim().isEmpty()) {
            return true;
        }

        String lowerValue = value.trim().toLowerCase();

        // Common placeholder values that shouldn't be treated as duplicates
        return lowerValue.equals("no barcode") ||
                lowerValue.equals("n/a") ||
                lowerValue.equals("na") ||
                lowerValue.equals("none") ||
                lowerValue.equals("null") ||
                lowerValue.equals("no barcode available") ||
                lowerValue.equals("missing barcode") ||
                lowerValue.startsWith("no ") && lowerValue.contains("barcode");
    }

    // CHANGED: make this public static so FileProccessor can call it
    public static String detectOgfRemarkFromSku(String originalRawSku) {
        if (originalRawSku == null || originalRawSku.trim().isEmpty()) return "";

        String sku = originalRawSku.trim();
        String upperSku = sku.toUpperCase();

        // Check for various OGF prefix patterns
        boolean hasOgfPrefix = upperSku.startsWith("OGF-") ||
                upperSku.startsWith("OGF_") ||
                upperSku.startsWith("TEMP_SKU_OGF") ||
                upperSku.startsWith("TEMP-OGF") ||
                upperSku.contains("-OGF") ||
                upperSku.contains("_OGF");

        // Check for OGF anywhere in SKU (case insensitive)
        boolean hasOgfAnywhere = upperSku.contains("OGF");

        if (hasOgfPrefix) {
            return "OGF prefix found: '" + sku + "'";
        } else if (hasOgfAnywhere) {
            return "OGF detected in SKU: '" + sku + "'";
        } else {
            return "WARNING: No OGF prefix in SKU: '" + sku + "'";
        }
    }

    @SuppressWarnings("unchecked")
    // MODIFIED: Added parameter to simulate the checkbox state
    public static void generateReport(List<File> locationFiles, List<File> unlistedFiles, File output, boolean ogfRulesChecked) throws IOException {
        if (locationFiles == null || locationFiles.isEmpty()) {
            throw new IllegalArgumentException("Must provide at least one location file for comparison.");
        }

        // SET THE GLOBAL FLAG
        useOgfRules = ogfRulesChecked;

        List<String> locationNames = locationFiles.stream().map(f -> f.getName().replace(".xlsx", "").replace(".xls", "")).collect(Collectors.toList());
        List<String> unlistedNames = (unlistedFiles != null) ?
                unlistedFiles.stream().map(f -> f.getName().replace(".xlsx", "").replace(".xls", "")).collect(Collectors.toList()) :
                new ArrayList<>();

        Map<String, Item> consolidatedItemsBySku = new HashMap<>();
        List<Item> consolidatedItemsWithNoSku = new ArrayList<>();

        // Needed for Cosmetics Rule Logic
        Set<String> cosmeticLocationNames = locationNames.stream()
                .filter(name -> name.toLowerCase().contains("cosmetics") || name.toLowerCase().contains("cos"))
                .collect(Collectors.toSet());

        // Read location files and merge
        for (File file : locationFiles) {
            String fileName = file.getName().replace(".xlsx", "").replace(".xls", "");
            boolean isTempOgfFile = isOgfName(fileName);

            Map<Set<Item>, List<ItemSourceData>> data = readItems(file, isTempOgfFile, false);
            Set<Item> uniqueItems = (Set<Item>) data.keySet().iterator().next();
            List<ItemSourceData> duplicateSourceData = (List<ItemSourceData>) data.values().iterator().next();

            Set<Item> duplicatesSet = new HashSet<>();
            for(ItemSourceData d : duplicateSourceData) { duplicatesSet.add(new Item(d.cleanSku, d.rawBarcode)); }

            for (Item newItem : uniqueItems) {
                ItemSourceData currentData = newItem.getDataForLocation("TEMP_KEY");
                String currentSku = currentData.cleanSku;

                boolean isDuplicateInSource = duplicatesSet.contains(newItem);
                boolean isShortBarcode = currentData.hasShortBarcode;

                ItemSourceData finalData = new ItemSourceData(currentData.rawSku, currentData.rawBarcode, currentData.rawProductName,
                        isDuplicateInSource, isShortBarcode, currentData.ogfRemark, isTempOgfFile);

                Item existingItem = null;
                if (!currentSku.isEmpty()) {
                    existingItem = consolidatedItemsBySku.get(currentSku.toLowerCase());
                }

                if (existingItem == null) {
                    Item itemToUse;
                    if (!currentSku.isEmpty()) {
                        itemToUse = new Item(currentSku, currentData.rawBarcode);
                        itemToUse.addSourceData(fileName, finalData);
                        consolidatedItemsBySku.put(currentSku.toLowerCase(), itemToUse);
                    } else if (!currentData.rawBarcode.isEmpty()) {
                        boolean merged = false;
                        for (Item noSkuItem : consolidatedItemsWithNoSku) {
                            if (noSkuItem.primaryBarcode.equalsIgnoreCase(currentData.rawBarcode)) {
                                noSkuItem.addSourceData(fileName, finalData);
                                merged = true;
                                break;
                            }
                        }
                        if (!merged) {
                            itemToUse = new Item("", currentData.rawBarcode);
                            itemToUse.addSourceData(fileName, finalData);
                            consolidatedItemsWithNoSku.add(itemToUse);
                        } else {
                            itemToUse = null;
                        }
                    } else {
                        continue;
                    }
                    if (itemToUse != null) {
                        itemToUse.markAsOgfGroupItem(isTempOgfFile);
                        // ADDED: Mark item as Cosmetics if it came from a Cosmetics location file
                        itemToUse.markAsCosmeticsGroupItem(fileName);
                    }

                } else {
                    existingItem.addSourceData(fileName, finalData);
                    existingItem.markAsOgfGroupItem(isTempOgfFile);
                    // ADDED: Mark existing item as Cosmetics if new file is Cosmetics
                    existingItem.markAsCosmeticsGroupItem(fileName);
                }
            }
        }

        // FIXED: Process Unlisted/Unavailable files and properly set primarySkuSource for unlisted-only items
        if (unlistedFiles != null) {
            for (File file : unlistedFiles) {
                String fileName = file.getName().replace(".xlsx", "").replace(".xls", "");

                // SIMPLIFIED: Check if unlisted file contains "OGF" (case insensitive)
                boolean isTempOgfFile = fileName.toUpperCase().replaceAll("[^A-Z0-9]", "").contains("OGF");

                Map<Set<Item>, List<ItemSourceData>> data = readItems(file, isTempOgfFile, true);
                Set<Item> uniqueItems = (Set<Item>) data.keySet().iterator().next();

                for (Item newItem : uniqueItems) {
                    ItemSourceData currentData = newItem.getDataForLocation("TEMP_KEY");
                    String currentSku = currentData.cleanSku;

                    ItemSourceData finalData = new ItemSourceData(
                            currentData.rawSku,
                            currentData.rawBarcode,
                            currentData.rawProductName,
                            false,
                            false,
                            currentData.ogfRemark,
                            isTempOgfFile
                    );

                    Item existingItem = null;
                    if (!currentSku.isEmpty()) {
                        existingItem = consolidatedItemsBySku.get(currentSku.toLowerCase());
                    }

                    if (existingItem == null) {
                        // Item only exists in unlisted file, create new Item to track it
                        Item itemToUse;
                        if (!currentSku.isEmpty()) {
                            itemToUse = new Item(currentSku, currentData.rawBarcode);
                            // CRITICAL FIX: Set primarySkuSource immediately for unlisted-only items BEFORE addSourceData
                            itemToUse.primarySkuSource = fileName;
                            itemToUse.addSourceData(fileName, finalData);
                            consolidatedItemsBySku.put(currentSku.toLowerCase(), itemToUse);
                            System.out.println("DEBUG: Created unlisted-only item with SKU: " + currentSku + " from file: " + fileName);
                        } else if (!currentData.rawBarcode.isEmpty()) {
                            // Item with no SKU but has Barcode - check no-sku list
                            boolean merged = false;
                            for (Item noSkuItem : consolidatedItemsWithNoSku) {
                                if (noSkuItem.primaryBarcode.equalsIgnoreCase(currentData.rawBarcode)) {
                                    noSkuItem.addSourceData(fileName, finalData);
                                    System.out.println("DEBUG: Merged no-SKU unlisted item (barcode " + currentData.rawBarcode + ") into existing item from file: " + fileName);
                                    merged = true;
                                    break;
                                }
                            }
                            if (!merged) {
                                itemToUse = new Item("", currentData.rawBarcode);
                                itemToUse.addSourceData(fileName, finalData);
                                consolidatedItemsWithNoSku.add(itemToUse);
                                System.out.println("DEBUG: Created new no-SKU unlisted item (barcode " + currentData.rawBarcode + ") from file: " + fileName);
                            }
                        }
                    } else {
                        // Item found by SKU, merge source data into existing item
                        existingItem.addSourceData(fileName, finalData);
                        System.out.println("DEBUG: Merging unlisted source '" + fileName + "' into existing SKU " + currentSku);

                        // CRITICAL FIX: If existing item doesn't have a primarySkuSource yet, set it to this unlisted file
                        if (existingItem.primarySkuSource.isEmpty() && !currentSku.isEmpty()) {
                            existingItem.primarySkuSource = fileName;
                            System.out.println("DEBUG: Set primarySkuSource to unlisted file '" + fileName + "' for existing item: " + currentSku);
                        }
                    }
                }
            }
        }

        // Pass consolidated data to the writing method
        writeComparisonReport(
                consolidatedItemsBySku,
                consolidatedItemsWithNoSku,
                locationNames,
                unlistedNames,
                cosmeticLocationNames,
                output
        );
    }


    private static void writeComparisonReport(Map<String, Item> consolidatedItemsBySku,
                                              List<Item> consolidatedItemsWithNoSku,
                                              List<String> locationNames, List<String> unlistedNames, // all unlisted names
                                              Set<String> cosmeticLocationNames, // kept for display name logic
                                              File output) throws IOException {
        // Consolidation and Sorting remains the same
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Inventory Comparison Report");

        // Final Consolidation and Sorting
        List<Item> allConsolidatedItems = new ArrayList<>(consolidatedItemsBySku.values());
        allConsolidatedItems.addAll(consolidatedItemsWithNoSku);

        System.out.println("=== STARTING DUPLICATE DETECTION ===");
        System.out.println("Total items to check: " + allConsolidatedItems.size());

        // NEW: Detect cross-item barcode duplicates BEFORE processing individual items
        detectCrossItemBarcodeDuplicates(allConsolidatedItems);

        // NEW: Detect SKU-barcode mismatches (different SKUs sharing same barcode)
        detectSkuBarcodeMismatches(allConsolidatedItems);

        // UPDATED: Enhanced product title logic to handle unlisted-only items correctly
        for (Item item : allConsolidatedItems) {
            // Always start with empty product name
            item.consolidatedProductName = "";

            System.out.println("=== DEBUG: Processing item with primary SKU: " + item.primarySku + " ===");
            System.out.println("DEBUG: Primary SKU source: '" + item.primarySkuSource + "'");
            System.out.println("DEBUG: All sources: " + item.sourceData.keySet());

            // STRATEGY 1: Try to use product title from the SKU source (including unlisted files)
            if (!item.primarySkuSource.isEmpty()) {
                ItemSourceData sourceData = item.getDataForLocation(item.primarySkuSource);
                if (sourceData != null && hasValidProductTitle(sourceData.rawProductName)) {
                    item.consolidatedProductName = sourceData.rawProductName.trim();
                    System.out.println("DEBUG: SUCCESS - Using product title from SKU source '" + item.primarySkuSource +
                            "' for item: " + item.primarySku + " - Title: '" + item.consolidatedProductName + "'");
                } else {
                    System.out.println("DEBUG: SKU source '" + item.primarySkuSource +
                            "' has no valid product title for item: " + item.primarySku);
                }
            }

            // STRATEGY 2: If no SKU source or no product title from SKU source, discover the source
            if (item.consolidatedProductName.isEmpty()) {
                System.out.println("DEBUG: Attempting to discover SKU source for item: " + item.primarySku);

                // Look through ALL sources to find which one has this SKU and a valid product title
                for (Map.Entry<String, ItemSourceData> entry : item.sourceData.entrySet()) {
                    String sourceName = entry.getKey();
                    ItemSourceData data = entry.getValue();

                    // Check if this source has the same SKU as our primary SKU
                    if (data.rawSku != null && data.rawSku.trim().equalsIgnoreCase(item.primarySku)) {
                        System.out.println("DEBUG: Found matching SKU in source: " + sourceName);

                        if (hasValidProductTitle(data.rawProductName)) {
                            item.consolidatedProductName = data.rawProductName.trim();
                            System.out.println("DEBUG: SUCCESS - Discovered and using product title from source '" + sourceName +
                                    "' for item: " + item.primarySku);
                            break;
                        } else {
                            System.out.println("DEBUG: Source '" + sourceName + "' has no valid product title");
                        }
                    }
                }
            }

            // STRATEGY 3: If still no product title, check ALL location files
            if (item.consolidatedProductName.isEmpty()) {
                System.out.println("DEBUG: Falling back to location files for item: " + item.primarySku);
                for (String location : locationNames) {
                    ItemSourceData locationData = item.getDataForLocation(location);
                    if (locationData != null && hasValidProductTitle(locationData.rawProductName)) {
                        item.consolidatedProductName = locationData.rawProductName.trim();
                        System.out.println("DEBUG: SUCCESS - Fallback to location '" + location +
                                "' for product title: " + item.primarySku);
                        break;
                    }
                }
            }

            // STRATEGY 4: If still no product title, check ALL unlisted files for the best title
            if (item.consolidatedProductName.isEmpty()) {
                System.out.println("DEBUG: Falling back to unlisted files for item: " + item.primarySku);
                String bestUnlistedTitle = "";
                String bestUnlistedSource = "";

                for (String unlisted : unlistedNames) {
                    for (String key : item.sourceData.keySet()) {
                        if (key.equalsIgnoreCase(unlisted) || key.toLowerCase().contains(unlisted.toLowerCase())) {
                            ItemSourceData unlistedData = item.sourceData.get(key);
                            if (unlistedData != null && hasValidProductTitle(unlistedData.rawProductName)) {
                                String currentTitle = unlistedData.rawProductName.trim();
                                if (currentTitle.length() > bestUnlistedTitle.length()) {
                                    bestUnlistedTitle = currentTitle;
                                    bestUnlistedSource = key;
                                }
                                System.out.println("DEBUG: Found valid product title in unlisted '" + key + "': '" + currentTitle + "'");
                            }
                        }
                    }
                }

                if (!bestUnlistedTitle.isEmpty()) {
                    item.consolidatedProductName = bestUnlistedTitle;
                    System.out.println("DEBUG: SUCCESS - Using best product title from unlisted '" + bestUnlistedSource +
                            "' for item: " + item.primarySku);
                }
            }

            // STRATEGY 5: Final backup - use "Default Title"
            if (item.consolidatedProductName.isEmpty()) {
                item.consolidatedProductName = "Default Title";
                System.out.println("DEBUG: FINAL BACKUP - Using 'Default Title' for item: " + item.primarySku);

                // Additional debug for items with no title
                System.out.println("=== NO TITLE FOUND: Item " + item.primarySku + " using 'Default Title' ===");
                for (Map.Entry<String, ItemSourceData> entry : item.sourceData.entrySet()) {
                    String source = entry.getKey();
                    ItemSourceData data = entry.getValue();
                    System.out.println("DEBUG:   Source " + source + " - SKU: '" + data.rawSku + "', Product Title: '" + data.rawProductName + "'");
                }
            }
            System.out.println("=== END Processing item: " + item.primarySku + " ===\n");
        }

        allConsolidatedItems.sort(new Comparator<Item>() {
            @Override
            public int compare(Item i1, Item i2) {
                int skuCompare = i1.primarySku.compareToIgnoreCase(i2.primarySku);
                if (skuCompare != 0) return skuCompare;
                return i1.primaryBarcode.compareToIgnoreCase(i2.primaryBarcode);
            }
        });
        // End Consolidation and Sorting

        // Write Header Row
        Map<String, String> locationDisplayNames = new HashMap<>();
        String originalOgfName = locationNames.stream()
                .filter(name -> isOgfName(name))  // Changed to temp_sku_ogf
                .map(name -> "OGF Location")  // Always use "OGF Location" for OGF files
                .findFirst().orElse(null);

        for (String name : locationNames) {
            if (isOgfName(name)) {  // Changed to temp_sku_ogf
                locationDisplayNames.put(name, "OGF Location");  // Use consistent name
            } else {
                locationDisplayNames.put(name, name);
            }
        }

        Map<String, String> unlistedDisplayNames = new HashMap<>();
        String originalOgfUnlistedName = unlistedNames.stream()
                .filter(name -> isOgfName(name))  // Changed to temp_sku_ogf
                .map(name -> "OGF Unlisted")  // Always use "OGF Unlisted" for OGF files
                .findFirst().orElse(null);

        for (String name : unlistedNames) {
            if (isOgfName(name)) {  // Changed to temp_sku_ogf
                unlistedDisplayNames.put(name, "OGF Unlisted");  // Use consistent name
            } else {
                unlistedDisplayNames.put(name, name);
            }
        }
        CellStyle headerStyle = createHeaderStyle(workbook);
        int rowIdx = 0;
        Row header = sheet.createRow(rowIdx++);
        int colIdx = 0;

        header.createCell(colIdx++).setCellValue("Primary SKU (Consolidated)");
        header.createCell(colIdx++).setCellValue("Primary Barcode (Consolidated)");
        header.createCell(colIdx++).setCellValue("Product Name");

        for (String unlistedName : unlistedNames) {
            String displayName = unlistedDisplayNames.getOrDefault(unlistedName, unlistedName);
            header.createCell(colIdx++).setCellValue("SKU (" + displayName + ")");
            header.createCell(colIdx++).setCellValue("Barcode (" + displayName + ")");
        }

        for (String location : locationNames) {
            String displayName = locationDisplayNames.getOrDefault(location, location);
            header.createCell(colIdx++).setCellValue("SKU (" + displayName + ")");
            header.createCell(colIdx++).setCellValue("Barcode (" + displayName + ")");
            header.createCell(colIdx++).setCellValue("OGF Remark (" + displayName + ")");
        }

        header.createCell(colIdx++).setCellValue("In ALL Locations?");
        header.createCell(colIdx++).setCellValue("In ANY UNLISTED?");
        header.createCell(colIdx++).setCellValue("Simple Status");
        header.createCell(colIdx++).setCellValue("ID / Data Problem");
        header.createCell(colIdx++).setCellValue("CONSOLIDATED REMARKS");

        for (int i = 0; i < colIdx; i++) header.getCell(i).setCellStyle(headerStyle);


        // Start Data Rows Loop
        for (Item item : allConsolidatedItems) {

            // Call the unified remarks generator
            generateFinalRemarksWithFilteredUnlisted(item, locationNames, unlistedNames, cosmeticLocationNames);

            Row row = sheet.createRow(rowIdx++);
            colIdx = 0;

            row.createCell(colIdx++).setCellValue(item.primarySku);
            row.createCell(colIdx++).setCellValue(item.primaryBarcode);
            row.createCell(colIdx++).setCellValue(item.consolidatedProductName);

            // Unlisted Columns DATA (All data is available in item.sourceData)
            for (String unlistedName : unlistedNames) {
                ItemSourceData unlistedSourceData = item.getDataForLocation(unlistedName);
                if (unlistedSourceData != null) {
                    row.createCell(colIdx++).setCellValue(unlistedSourceData.rawSku);
                    row.createCell(colIdx++).setCellValue(unlistedSourceData.rawBarcode);
                } else {
                    row.createCell(colIdx++).setCellValue("");
                    row.createCell(colIdx++).setCellValue("");
                }
            }

            // Location Columns DATA
            int presentCount = 0;
            for (String location : locationNames) {
                ItemSourceData sourceData = item.getDataForLocation(location);

                if (sourceData != null) {
                    row.createCell(colIdx++).setCellValue(sourceData.rawSku);
                    row.createCell(colIdx++).setCellValue(sourceData.rawBarcode);
                    row.createCell(colIdx++).setCellValue(sourceData.ogfRemark);
                    presentCount++;
                } else {
                    row.createCell(colIdx++).setCellValue("");
                    row.createCell(colIdx++).setCellValue("");
                    row.createCell(colIdx++).setCellValue("");
                }
            }

            boolean presentInAll = presentCount == locationNames.size() && locationNames.size() > 0;

            // Report Columns (Status Flags)
            row.createCell(colIdx++).setCellValue(presentInAll ? "YES" : "NO");

            // Check the status for "In ANY UNLISTED?" column
            boolean statusInAnyRelevantUnlisted = isItemInAnyRelevantUnlisted(item, locationNames, unlistedNames, cosmeticLocationNames);
            row.createCell(colIdx++).setCellValue(statusInAnyRelevantUnlisted ? "YES" : "NO");

            row.createCell(colIdx++).setCellValue(item.simpleStatus);
            row.createCell(colIdx++).setCellValue(item.conflictStatus);

            String finalRemark = String.join(" | ", item.finalRemarks);
            row.createCell(colIdx++).setCellValue(finalRemark);
        }
        // End Data Rows Loop
        // Finalize and Write
        for (int i = 0; i < colIdx; i++) sheet.autoSizeColumn(i);

        try (FileOutputStream fos = new FileOutputStream(output)) {
            workbook.write(fos);
        }
        workbook.close();
        finalizeReportGeneration(output);
    }


    private static void detectSkuBarcodeMismatches(List<Item> allItems) {
        System.out.println("=== DETECTING SKU-BARCODE MISMATCHES (DIFFERENT SKUs SHARING SAME BARCODE) ===");

        Map<String, List<Item>> barcodeToItems = new HashMap<>();

        // Group all items by barcode (case-insensitive, ignore empty barcodes)
        for (Item item : allItems) {
            if (!item.primaryBarcode.isEmpty() && !isPlaceholderValue(item.primaryBarcode)) {
                String normalizedBarcode = item.primaryBarcode.trim().toLowerCase();
                barcodeToItems.computeIfAbsent(normalizedBarcode, k -> new ArrayList<>()).add(item);
                System.out.println("DEBUG: Added barcode '" + normalizedBarcode + "' for SKU: " + item.primarySku);
            }
        }

        // Flag barcodes used by multiple items with DIFFERENT SKUs
        for (Map.Entry<String, List<Item>> entry : barcodeToItems.entrySet()) {
            if (entry.getValue().size() > 1) {
                String barcode = entry.getKey();
                List<Item> duplicateItems = entry.getValue();

                // Check if these are actually different SKUs (not just the same item from multiple files)
                Set<String> uniqueSkus = duplicateItems.stream()
                        .map(item -> item.primarySku.trim().toLowerCase())
                        .filter(sku -> !sku.isEmpty())
                        .collect(Collectors.toSet());

                // Only flag if there are genuinely different SKUs sharing the same barcode
                if (uniqueSkus.size() > 1) {
                    System.out.println("🚨 CRITICAL: SKU-BARCODE MISMATCH FOUND: Barcode '" + barcode +
                            "' is shared by " + uniqueSkus.size() + " different SKUs:");

                    // Collect all SKUs that share this barcode
                    List<String> duplicateSkus = new ArrayList<>();
                    for (Item item : duplicateItems) {
                        duplicateSkus.add(item.primarySku);
                        System.out.println("   - SKU: " + item.primarySku + " | Product: " + item.consolidatedProductName);
                    }

                    // Flag ALL items that share this barcode with different SKUs
                    for (Item item : duplicateItems) {
                        System.out.println("DEBUG: Flagging item " + item.primarySku + " with duplicate barcode");
                        if (!item.conflictStatus.contains("DUPLICATE_BARCODE_ACROSS_SKUS")) {
                            item.conflictStatus = item.conflictStatus.isEmpty() ?
                                    "DUPLICATE_BARCODE_ACROSS_SKUS" : item.conflictStatus + " + DUPLICATE_BARCODE_ACROSS_SKUS";
                        }

                        // Create detailed remark showing all conflicting SKUs
                        // FIXED: Use traditional loop instead of stream to avoid compilation error
                        List<String> otherSkuList = new ArrayList<>();
                        for (Item other : duplicateItems) {
                            if (!other.primarySku.equals(item.primarySku)) {
                                otherSkuList.add(other.primarySku);
                            }
                        }
                        String otherSkus = String.join(", ", otherSkuList);

                        // Only add this remark once to avoid duplication
                        boolean alreadyHasRemark = item.finalRemarks.stream()
                                .anyMatch(remark -> remark.contains("Barcode " + barcode + " shared with other SKU"));

                        if (!alreadyHasRemark) {
                            item.finalRemarks.add("🚫 CRITICAL: Barcode " + barcode + " shared with other SKU(s): " + otherSkus);
                            System.out.println("DEBUG: Added duplicate barcode remark to item: " + item.primarySku);
                        }
                    }
                } else {
                    System.out.println("DEBUG: Barcode '" + barcode + "' has " + duplicateItems.size() +
                            " items but only " + uniqueSkus.size() + " unique SKUs - not flagging as mismatch");
                }
            }
        }

        System.out.println("=== SKU-BARCODE MISMATCH DETECTION COMPLETE ===");
    }


    private static void detectCrossItemBarcodeDuplicates(List<Item> allItems) {
        System.out.println("=== DETECTING CROSS-ITEM BARCODE DUPLICATES ===");

        Map<String, List<Item>> barcodeToItems = new HashMap<>();

        // Group all items by barcode (case-insensitive, ignore empty barcodes)
        for (Item item : allItems) {
            if (!item.primaryBarcode.isEmpty() && !isPlaceholderValue(item.primaryBarcode)) {
                String normalizedBarcode = item.primaryBarcode.trim().toLowerCase();
                barcodeToItems.computeIfAbsent(normalizedBarcode, k -> new ArrayList<>()).add(item);
            }
        }

        // Flag barcodes used by multiple items
        for (Map.Entry<String, List<Item>> entry : barcodeToItems.entrySet()) {
            if (entry.getValue().size() > 1) {
                String barcode = entry.getKey();
                List<Item> duplicateItems = entry.getValue();

                System.out.println("🚨 CROSS-ITEM BARCODE DUPLICATE FOUND: Barcode '" + barcode +
                        "' is shared by " + duplicateItems.size() + " different items:");

                // Collect all SKUs that share this barcode
                List<String> duplicateSkus = new ArrayList<>();
                for (Item item : duplicateItems) {
                    duplicateSkus.add(item.primarySku);
                    System.out.println("   - SKU: " + item.primarySku + " | Product: " + item.consolidatedProductName);
                }

                // Flag ALL items that share this barcode
                for (Item item : duplicateItems) {
                    if (!item.conflictStatus.contains("DUPLICATE_BARCODE_ACROSS_ITEMS")) {
                        item.conflictStatus = item.conflictStatus.isEmpty() ?
                                "DUPLICATE_BARCODE_ACROSS_ITEMS" : item.conflictStatus + " + DUPLICATE_BARCODE_ACROSS_ITEMS";
                    }

                    // Create detailed remark showing all conflicting SKUs
                    String otherSkus = duplicateItems.stream()
                            .filter(other -> !other.primarySku.equals(item.primarySku))
                            .map(other -> other.primarySku)
                            .collect(Collectors.joining(", "));

                    item.finalRemarks.add("🚫 Barcode " + barcode + " shared with other SKUs: " + otherSkus);
                }
            }
        }

        System.out.println("=== CROSS-ITEM BARCODE DUPLICATE DETECTION COMPLETE ===");
    }


    private static boolean hasValidProductTitle(String productTitle) {
        if (productTitle == null) return false;
        String trimmed = productTitle.trim();
        return !trimmed.isEmpty() &&
                !trimmed.equalsIgnoreCase("null") &&
                !trimmed.equalsIgnoreCase("n/a") &&
                !trimmed.equalsIgnoreCase("na") &&
                !trimmed.equalsIgnoreCase("none") &&
                !trimmed.equalsIgnoreCase("default title") &&
                trimmed.length() >= 2; // Minimum reasonable product title length
    }


    private static boolean isItemInAnyRelevantUnlisted(Item item, List<String> locationNames, List<String> unlistedNames, Set<String> cosmeticLocationNames) {

        if (useOgfRules) {
            // For OGF rules, use the same exclusive logic as in generateFinalRemarksWithFilteredUnlisted
            Set<String> presentLocations;
            if (item.isOgfGroupItem) {
                // OGF item: Only consider OGF locations - FIXED: check for "temp_sku_ogf" prefix
                presentLocations = locationNames.stream()
                        .filter(name -> isOgfName(name))  // Changed to temp_sku_ogf
                        .filter(item::isPresentIn)
                        .collect(Collectors.toSet());
            } else {
                // Non-OGF item: Only consider non-OGF locations
                presentLocations = locationNames.stream()
                        .filter(name -> !isOgfName(name))  // fixed to non-OGF
                        .filter(item::isPresentIn)
                        .collect(Collectors.toSet());
            }

            // Now determine relevant unlisted based on the filtered locations
            if (item.isOgfGroupItem && !presentLocations.isEmpty()) {
                // OGF item in OGF locations: Check OGF unlisted
                return unlistedNames.stream()
                        .filter(SKUComparerLogic::isOgfName)
                        .anyMatch(item::isPresentIn);
            } else if (!item.isOgfGroupItem && !presentLocations.isEmpty()) {
                // Non-OGF item in non-OGF locations: Check non-OGF unlisted
                return unlistedNames.stream()
                        .filter(name -> !isOgfName(name))
                        .anyMatch(item::isPresentIn);
            } else {
                return false;
            }
        } else {
            // Cosmetics Rules - Use the same exclusive logic as in generateFinalRemarksWithFilteredUnlisted
            Set<String> presentLocations;
            if (item.isCosmeticsGroupItem) {
                // Cosmetics item: Only consider cosmetics locations
                presentLocations = locationNames.stream()
                        .filter(name -> cosmeticLocationNames.contains(name) ||
                                name.toUpperCase().contains("COSMETIC") ||
                                name.toUpperCase().contains("COS"))
                        .filter(item::isPresentIn)
                        .collect(Collectors.toSet());
            } else {
                // Non-cosmetics item: Only consider non-cosmetics locations
                presentLocations = locationNames.stream()
                        .filter(name -> !cosmeticLocationNames.contains(name) &&
                                !name.toUpperCase().contains("COSMETIC") &&
                                !name.toUpperCase().contains("COS"))
                        .filter(item::isPresentIn)
                        .collect(Collectors.toSet());
            }

            // Now determine relevant unlisted based on the filtered locations
            if (item.isCosmeticsGroupItem && !presentLocations.isEmpty()) {
                // Cosmetics item in cosmetics locations: Check WEB unlisted
                return unlistedNames.stream()
                        .filter(name -> name.toUpperCase().contains("WEB"))
                        .anyMatch(item::isPresentIn);
            } else if (!item.isCosmeticsGroupItem && !presentLocations.isEmpty()) {
                // Non-cosmetics item in non-cosmetics locations: Check non-WEB unlisted
                return unlistedNames.stream()
                        .filter(name -> !name.toUpperCase().contains("WEB"))
                        .anyMatch(item::isPresentIn);
            } else {
                return false;
            }
        }
    }


    private static void generateFinalRemarksWithFilteredUnlisted(
            Item item, List<String> locationNames, List<String> unlistedNames, Set<String> cosmeticLocationNames) {

        // CRITICAL FIX: Store existing duplicate barcode status BEFORE clearing
        boolean hadDuplicateBarcodeBefore = item.conflictStatus.contains("DUPLICATE_BARCODE_ACROSS_SKUS");
        List<String> existingDuplicateRemarks = new ArrayList<>();
        for (String remark : item.finalRemarks) {
            if (remark.contains("CRITICAL: Barcode") && remark.contains("shared with other SKU")) {
                existingDuplicateRemarks.add(remark);
            }
        }

        // Clear regular remarks but preserve duplicate barcode remarks
        item.finalRemarks.clear();
        item.conflictStatus = "";

        // RESTORE duplicate barcode status and remarks if they existed
        if (hadDuplicateBarcodeBefore) {
            item.conflictStatus = "DUPLICATE_BARCODE_ACROSS_SKUS";
            item.finalRemarks.addAll(existingDuplicateRemarks);
            System.out.println("DEBUG: RESTORED duplicate barcode status for item: " + item.primarySku);
        }

        // --- Data Quality Checks (applied regardless of rule set) ---
        detectInternalInconsistencies(item);
        detectShortBarcodes(item);
        detectCrossFileDifferences(item);
        boolean hasDataIssues = !item.conflictStatus.isEmpty();

        // --- NEW: Check for critical duplicate barcode issues FIRST ---
        boolean hasCriticalDuplicateBarcode = item.conflictStatus.contains("DUPLICATE_BARCODE_ACROSS_SKUS");

        // --- Determine presence ---
        Set<String> presentLocations = locationNames.stream()
                .filter(item::isPresentIn)
                .collect(Collectors.toSet());
        Set<String> presentUnlisted = unlistedNames.stream()
                .filter(item::isPresentIn)
                .collect(Collectors.toSet());

        Set<String> missingLocations = locationNames.stream()
                .filter(loc -> !presentLocations.contains(loc))
                .collect(Collectors.toSet());

        boolean isBad = false;
        List<String> badReasons = new ArrayList<>();

        // --- CRITICAL FIX: If this has duplicate barcode issue, skip normal rule checking ---
        if (hasCriticalDuplicateBarcode) {
            // For critical duplicate barcode issues, mark as BAD regardless of other rules
            isBad = true;
            System.out.println("DEBUG: Item " + item.primarySku + " has CRITICAL duplicate barcode - skipping rule checks");
        } else {
            // Only apply normal rules if no critical duplicate barcode issue
            if (useOgfRules) {
                // --------------------- OGF RULES ---------------------
                String ogfLocationFile = locationNames.stream()
                        .filter(SKUComparerLogic::isOgfName)
                        .findFirst().orElse(null);

                String ogfUnlistedFile = unlistedNames.stream()
                        .filter(SKUComparerLogic::isOgfName)
                        .findFirst()
                        .orElse(null);

                // Get non-OGF unlisted files (all unlisted except OGF unlisted)
                Set<String> nonOgfUnlisted = unlistedNames.stream()
                        .filter(unl -> ogfUnlistedFile == null || !unl.equals(ogfUnlistedFile))
                        .collect(Collectors.toSet());

                boolean inOgfLoc = ogfLocationFile != null && presentLocations.contains(ogfLocationFile);
                boolean inOgfUnl = ogfUnlistedFile != null && presentUnlisted.contains(ogfUnlistedFile);
                boolean inAnyNonOgfUnlisted = nonOgfUnlisted.stream().anyMatch(presentUnlisted::contains);

                // LOGIC 1: OGF unlisted should ONLY be compared to OGF location file
                // If item is in both OGF location AND OGF unlisted → BAD
                if (inOgfLoc && inOgfUnl) {
                    isBad = true;
                    badReasons.add("OGF item should not appear in both " + ogfLocationFile + " and " + ogfUnlistedFile);
                }

                // LOGIC 2: Other unlisted files should be compared to ALL location files including OGF location
                // If item is in any non-OGF unlisted AND in any location (including OGF) → BAD
                if (inAnyNonOgfUnlisted && !presentLocations.isEmpty()) {
                    isBad = true;
                    badReasons.add("Non-OGF unlisted item should not appear in any location files");
                }

                // UPDATED LOGIC 3: Location consistency - but account for unlisted rules
                if (!presentLocations.isEmpty() && !missingLocations.isEmpty()) {
                    // Check if the missing locations are justified by unlisted rules
                    boolean hasUnjustifiedMissingLocations = false;
                    Set<String> unjustifiedMissing = new HashSet<>();

                    for (String missingLoc : missingLocations) {
                        boolean isOgfLocation = isOgfName(missingLoc);

                        if (isOgfLocation) {
                            // Missing from OGF location is OK if item is in OGF unlisted
                            if (!inOgfUnl) {
                                unjustifiedMissing.add(missingLoc);
                                hasUnjustifiedMissingLocations = true;
                            }
                        } else {
                            // Missing from non-OGF location is OK if item is in non-OGF unlisted
                            if (!inAnyNonOgfUnlisted) {
                                unjustifiedMissing.add(missingLoc);
                                hasUnjustifiedMissingLocations = true;
                            }
                        }
                    }

                    if (hasUnjustifiedMissingLocations) {
                        isBad = true;
                        badReasons.add("Item missing from locations: " + String.join(", ", unjustifiedMissing));
                    }
                }

                // LOGIC 4: If item is not in any unlisted file AND not in relevant location file → BAD
                if (item.isOgfGroupItem) {
                    if (!inOgfLoc && !inOgfUnl) {
                        isBad = true;
                        badReasons.add("OGF item missing from both OGF location and OGF unlisted");
                    }
                } else {
                    if (presentLocations.isEmpty() && !inAnyNonOgfUnlisted) {
                        isBad = true;
                        badReasons.add("Non-OGF item missing from all locations and non-OGF unlisted files");
                    }
                }

            } else {
                // --------------------- COSMETICS RULES ---------------------
                String cosmeticsLocationFile = locationNames.stream()
                        .filter(name -> cosmeticLocationNames.contains(name)
                                || name.toUpperCase().contains("COSMETIC")
                                || name.toUpperCase().contains("COS"))
                        .findFirst()
                        .orElse(null);

                String webUnlistedFile = unlistedNames.stream()
                        .filter(name -> name.toUpperCase().contains("WEB"))
                        .findFirst()
                        .orElse(null);

                // Get non-WEB unlisted files (all unlisted except WEB unlisted)
                Set<String> nonWebUnlisted = unlistedNames.stream()
                        .filter(unl -> webUnlistedFile == null || !unl.equals(webUnlistedFile))
                        .collect(Collectors.toSet());

                boolean inCosLoc = cosmeticsLocationFile != null && presentLocations.contains(cosmeticsLocationFile);
                boolean inWebUnl = webUnlistedFile != null && presentUnlisted.contains(webUnlistedFile);
                boolean inAnyNonWebUnlisted = nonWebUnlisted.stream().anyMatch(presentUnlisted::contains);

                // LOGIC 1: Cosmetics file should ONLY be compared to WEB unlisted
                if (inCosLoc && inWebUnl) {
                    isBad = true;
                    badReasons.add("Cosmetics item should not appear in both " + cosmeticsLocationFile + " and " + webUnlistedFile);
                }

                // UPDATED LOGIC 2: Other unlisted files should be compared to NON-COSMETICS locations only
                if (inAnyNonWebUnlisted) {
                    Set<String> nonCosmeticsPresentLocations = presentLocations.stream()
                            .filter(loc -> !cosmeticLocationNames.contains(loc) &&
                                    !loc.toUpperCase().contains("COSMETIC") &&
                                    !loc.toUpperCase().contains("COS"))
                            .collect(Collectors.toSet());

                    if (!nonCosmeticsPresentLocations.isEmpty()) {
                        isBad = true;
                        badReasons.add("Non-WEB unlisted item should not appear in non-cosmetics locations: " + String.join(", ", nonCosmeticsPresentLocations));
                    }
                }

                // UPDATED LOGIC 3: Location consistency - but account for unlisted rules
                if (!presentLocations.isEmpty() && !missingLocations.isEmpty()) {
                    boolean hasUnjustifiedMissingLocations = false;
                    Set<String> unjustifiedMissing = new HashSet<>();

                    for (String missingLoc : missingLocations) {
                        boolean isCosmeticsLocation = cosmeticLocationNames.contains(missingLoc) ||
                                missingLoc.toUpperCase().contains("COSMETIC") ||
                                missingLoc.toUpperCase().contains("COS");

                        if (isCosmeticsLocation) {
                            if (!inWebUnl) {
                                unjustifiedMissing.add(missingLoc);
                                hasUnjustifiedMissingLocations = true;
                            }
                        } else {
                            if (!inAnyNonWebUnlisted) {
                                unjustifiedMissing.add(missingLoc);
                                hasUnjustifiedMissingLocations = true;
                            }
                        }
                    }

                    if (hasUnjustifiedMissingLocations) {
                        isBad = true;
                        badReasons.add("Item missing from locations: " + String.join(", ", unjustifiedMissing));
                    }
                }

                // LOGIC 4: If item is not in any unlisted file AND not in relevant location file → BAD
                if (item.isCosmeticsGroupItem) {
                    if (!inCosLoc && !inWebUnl) {
                        isBad = true;
                        badReasons.add("Cosmetics item missing from both cosmetics location and WEB unlisted");
                    }
                } else {
                    if (presentLocations.isEmpty() && presentUnlisted.isEmpty()) {
                        isBad = true;
                        badReasons.add("Non-cosmetics item missing from all locations and all unlisted files");
                    }
                }
            }
        }

        // --- Determine final status ---
        if (presentLocations.isEmpty() && presentUnlisted.isEmpty()) {
            item.simpleStatus = "No Data Found - BAD";
            item.finalRemarks.add("Item not found in any location or unlisted files");
        } else if (isBad) {
            // CRITICAL FIX: Show duplicate barcode as highest priority issue
            if (hasCriticalDuplicateBarcode) {
                item.simpleStatus = "CRITICAL: Duplicate Barcode - BAD";
                System.out.println("DEBUG: Setting CRITICAL status for item: " + item.primarySku);
            } else if (hasDataIssues) {
                item.simpleStatus = "Rule Violation + DATA ISSUES - BAD";
            } else {
                item.simpleStatus = "Rule Violation - BAD";
            }
            badReasons.forEach(reason -> item.finalRemarks.add("🚫 " + reason));
        } else if (hasDataIssues) {
            item.simpleStatus = "DATA ISSUES - BAD";
            item.finalRemarks.add("Item has data quality issues (short barcode/duplicates/SKU differences)");
        } else {
            item.simpleStatus = "GOOD";
            if (!presentLocations.isEmpty() && presentUnlisted.isEmpty()) {
                item.finalRemarks.add("✅ Item correctly placed in all locations and not in any unlisted files");
            } else if (presentLocations.isEmpty() && !presentUnlisted.isEmpty()) {
                item.finalRemarks.add("✅ Item correctly only in unlisted files: " + String.join(", ", presentUnlisted));
            } else {
                item.finalRemarks.add("✅ Item follows all location/unlisted pairing rules");
            }
        }

        // --- Presence info (always add these for clarity) ---
        if (!presentLocations.isEmpty()) {
            item.finalRemarks.add("Present in locations: " + String.join(", ", presentLocations));
        }
        if (!presentUnlisted.isEmpty()) {
            item.finalRemarks.add("Present in unlisted: " + String.join(", ", presentUnlisted));
        }
        if (!missingLocations.isEmpty() && !presentLocations.isEmpty()) {
            item.finalRemarks.add("Missing from locations: " + String.join(", ", missingLocations));
        }

        System.out.println("DEBUG: Final status for " + item.primarySku + ": " + item.simpleStatus + " | Conflict: " + item.conflictStatus);
    }


    private static void detectInternalInconsistencies(Item item) {
        Set<String> skusInThisItem = new HashSet<>();
        Set<String> barcodesInThisItem = new HashSet<>();

        // Track within-file duplicates
        List<String> withinFileDuplicateReasons = new ArrayList<>();

        // Check for multiple different SKUs/barcodes within this same item across files
        for (Map.Entry<String, ItemSourceData> entry : item.sourceData.entrySet()) {
            String sourceName = entry.getKey();
            ItemSourceData data = entry.getValue();

            if (!data.rawSku.trim().isEmpty()) {
                skusInThisItem.add(data.rawSku.trim());
            }
            if (!data.rawBarcode.trim().isEmpty()) {
                barcodesInThisItem.add(data.rawBarcode.trim());
            }

            // Check for within-file duplicates
            if (data.isDuplicateInSource) {
                StringBuilder reason = new StringBuilder();
                reason.append("Duplicate in '").append(sourceName).append("'");

                if (data.isSkuDuplicate && data.isBarcodeDuplicate) {
                    reason.append(" - SKU '").append(data.rawSku).append("' and Barcode '").append(data.rawBarcode).append("' appear multiple times in this file");
                } else if (data.isSkuDuplicate) {
                    reason.append(" - SKU '").append(data.rawSku).append("' appears multiple times in this file");
                } else if (data.isBarcodeDuplicate) {
                    reason.append(" - Barcode '").append(data.rawBarcode).append("' appears multiple times in this file");
                }

                withinFileDuplicateReasons.add(reason.toString());
            }
        }

        // NEW: Flag INCONSISTENT data (same item has different values across files)
        if (skusInThisItem.size() > 1) {
            if (!item.conflictStatus.contains("INCONSISTENT_SKU")) {
                item.conflictStatus = item.conflictStatus.isEmpty() ? "INCONSISTENT_SKU" : item.conflictStatus + " + INCONSISTENT_SKU";
            }
            item.finalRemarks.add("Different SKUs for same item across files: " + String.join(" vs ", skusInThisItem));
        }

        // NEW: Flag INCONSISTENT data (same item has different values across files)
        if (barcodesInThisItem.size() > 1) {
            if (!item.conflictStatus.contains("INCONSISTENT_BARCODE")) {
                item.conflictStatus = item.conflictStatus.isEmpty() ? "INCONSISTENT_BARCODE" : item.conflictStatus + " + INCONSISTENT_BARCODE";
            }
            item.finalRemarks.add("Different barcodes for same item across files: " + String.join(" vs ", barcodesInThisItem));
        }

        // Flag within-file duplicates
        if (!withinFileDuplicateReasons.isEmpty()) {
            if (!item.conflictStatus.contains("FILE_DUPLICATE")) {
                item.conflictStatus = item.conflictStatus.isEmpty() ? "FILE_DUPLICATE" : item.conflictStatus + " + FILE_DUPLICATE";
            }
            item.finalRemarks.addAll(withinFileDuplicateReasons);
        }
    }


    private static void detectCrossFileDifferences(Item item) {
        Set<String> allSkus = new HashSet<>();
        Set<String> allBarcodes = new HashSet<>();
        Map<String, String> skuToSource = new HashMap<>();
        Map<String, String> barcodeToSource = new HashMap<>();

        // Collect all SKUs and barcodes from ALL files
        for (Map.Entry<String, ItemSourceData> entry : item.sourceData.entrySet()) {
            String source = entry.getKey();
            ItemSourceData data = entry.getValue();

            if (!data.rawSku.trim().isEmpty()) {
                allSkus.add(data.rawSku.trim());
                skuToSource.put(data.rawSku.trim(), source);
            }

            if (!data.rawBarcode.trim().isEmpty()) {
                allBarcodes.add(data.rawBarcode.trim());
                barcodeToSource.put(data.rawBarcode.trim(), source);
            }
        }

        // FIXED: Renamed to INCONSISTENT_SKU for clarity
        if (allSkus.size() > 1) {
            if (!item.conflictStatus.contains("INCONSISTENT_SKU")) {
                item.conflictStatus = item.conflictStatus.isEmpty() ? "INCONSISTENT_SKU" : item.conflictStatus + " + INCONSISTENT_SKU";
            }
            List<String> skuDetails = new ArrayList<>();
            for (String sku : allSkus) {
                skuDetails.add(sku + "(" + skuToSource.get(sku) + ")");
            }
            item.finalRemarks.add("Different SKUs across files: " + String.join(" vs ", skuDetails));
        }

        // FIXED: Renamed to INCONSISTENT_BARCODE for clarity
        if (allBarcodes.size() > 1) {
            if (!item.conflictStatus.contains("INCONSISTENT_BARCODE")) {
                item.conflictStatus = item.conflictStatus.isEmpty() ? "INCONSISTENT_BARCODE" : item.conflictStatus + " + INCONSISTENT_BARCODE";
            }
            List<String> barcodeDetails = new ArrayList<>();
            for (String barcode : allBarcodes) {
                barcodeDetails.add(barcode + "(" + barcodeToSource.get(barcode) + ")");
            }
            item.finalRemarks.add("Different barcodes across files: " + String.join(" vs ", barcodeDetails));
        }

        // Flag if primary SKU/barcode doesn't match other sources
        if (!item.primarySku.isEmpty() && allSkus.size() > 1 && !allSkus.contains(item.primarySku)) {
            item.finalRemarks.add("Primary SKU '" + item.primarySku + "' doesn't match other files");
        }

        if (!item.primaryBarcode.isEmpty() && allBarcodes.size() > 1 && !allBarcodes.contains(item.primaryBarcode)) {
            item.finalRemarks.add("Primary barcode '" + item.primaryBarcode + "' doesn't match other files");
        }
    }


    private static void detectShortBarcodes(Item item) {
        Set<String> shortBarcodeSources = new HashSet<>();
        boolean hasShortBarcode = false;

        System.out.println("DEBUG: Checking short barcodes for item: " + item.primarySku + " | " + item.primaryBarcode);

        for (Map.Entry<String, ItemSourceData> entry : item.sourceData.entrySet()) {
            String source = entry.getKey();
            ItemSourceData data = entry.getValue();

            System.out.println("DEBUG: Source: " + source + " | Barcode: '" + data.rawBarcode + "' | Length: " + data.rawBarcode.length());

            // Check for short barcodes (less than 3 characters)
            if (!data.rawBarcode.trim().isEmpty()) {
                String barcode = data.rawBarcode.trim();
                // Check for less than 3 characters
                if (barcode.length() < 3) {
                    System.out.println("DEBUG: FOUND SHORT BARCODE: '" + barcode + "' in source: " + source);
                    shortBarcodeSources.add(source + "('" + barcode + "')");
                    hasShortBarcode = true;
                }
            }
        }

        // Also check primary barcode if it's not already covered above
        if (!item.primaryBarcode.isEmpty() && item.primaryBarcode.length() < 3) {
            System.out.println("DEBUG: FOUND SHORT PRIMARY BARCODE: '" + item.primaryBarcode + "'");
            // Only add if not already detected in source data
            boolean alreadyDetected = shortBarcodeSources.stream()
                    .anyMatch(source -> source.contains("'" + item.primaryBarcode + "'"));
            if (!alreadyDetected) {
                shortBarcodeSources.add("Primary('" + item.primaryBarcode + "')");
                hasShortBarcode = true;
            }
        }

        if (hasShortBarcode) {
            System.out.println("DEBUG: Setting SHORT_BARCODE conflict for item: " + item.primarySku);
            if (item.conflictStatus.isEmpty()) {
                item.conflictStatus = "SHORT_BARCODE";
            } else if (!item.conflictStatus.contains("SHORT_BARCODE")) {
                item.conflictStatus += " + SHORT_BARCODE";
            }
            item.finalRemarks.add("Short barcodes (<3 chars) in: " + String.join(", ", shortBarcodeSources));
        } else {
            System.out.println("DEBUG: No short barcodes found for item: " + item.primarySku);
        }
    }

    // Keep the original getStringValue method as backup, but use getFormattedCellValue in readItems
    private static String getStringValue(Cell cell) {
        if (cell == null) return "";
        try {
            switch (cell.getCellType()) {
                case STRING: return cell.getStringCellValue().trim();
                case NUMERIC:
                    double val = cell.getNumericCellValue();
                    return (val == Math.floor(val)) ? String.valueOf((long) val) : String.valueOf(val);
                case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
                case FORMULA:
                    CellType cachedType = cell.getCachedFormulaResultType();
                    if (cachedType == CellType.STRING) return cell.getStringCellValue().trim();
                    if (cachedType == CellType.NUMERIC) {
                        double valFormula = cell.getNumericCellValue();
                        return (valFormula == Math.floor(valFormula)) ? String.valueOf((long) valFormula) : String.valueOf(valFormula);
                    }
                case BLANK: default: return "";
            }
        } catch (Exception e) { return ""; }
    }

    private static CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setWrapText(true);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private static CellStyle createConflictStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.RED.getIndex());
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.ROSE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setWrapText(true);
        return style;
    }

    private static CellStyle createSuccessStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private static CellStyle createWarningStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private static void writeCell(Row row, int colIdx, String value, CellStyle style) {
        Cell cell = row.createCell(colIdx);
        cell.setCellValue(value);
        if (style != null) {
            cell.setCellStyle(style);
        }
    }

    private static void logReportStatus(String message) {}
    private static String normalizeBarcode(String barcode) {
        if (barcode == null) return "";
        String normalized = barcode.trim().toUpperCase();

        if (normalized.length() < 2 && normalized.length() > 0) {
            logReportStatus("Attempted normalization on short barcode: " + barcode);
            return normalized;
        }

        return normalized;
    }
    private static void finalizeReportGeneration(File output) {
        logReportStatus("Report generation complete: " + output.getAbsolutePath());
    }

} // End of SKUComparerLogic class