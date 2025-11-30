package com.example.InventoryComparer.logic;

import org.apache.poi.ss.usermodel.*;
import java.awt.Color;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.*;
import java.util.stream.Collectors;


public class POStockTallyLogic {

    private static final DataFormatter dataFormatter = new DataFormatter();
    private static final DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");

    private static final Map<String, String> COMPANY_ALIAS_TO_CODE = new LinkedHashMap<>();
    static {
        COMPANY_ALIAS_TO_CODE.put("AJS", "OUT200");
        COMPANY_ALIAS_TO_CODE.put("KRIBATHGODA", "OUT200");
        COMPANY_ALIAS_TO_CODE.put("KIRI", "OUT200");
        COMPANY_ALIAS_TO_CODE.put("MNK", "OUT100");
        COMPANY_ALIAS_TO_CODE.put("COOLPLANET", "OUT100");
        COMPANY_ALIAS_TO_CODE.put("CP", "OUT100");
        COMPANY_ALIAS_TO_CODE.put("LMJ", "OUT400");
        COMPANY_ALIAS_TO_CODE.put("PEPILIYANA", "OUT400");
        COMPANY_ALIAS_TO_CODE.put("PEP", "OUT400");
        COMPANY_ALIAS_TO_CODE.put("LWK", "OUT300");
        COMPANY_ALIAS_TO_CODE.put("OGF", "OUT300");
        COMPANY_ALIAS_TO_CODE.put("DRO", "OUT700");
        COMPANY_ALIAS_TO_CODE.put("MAH", "OUT700");
        COMPANY_ALIAS_TO_CODE.put("MAHARAGAMA", "OUT700");
        COMPANY_ALIAS_TO_CODE.put("CHAMI", "OUT500");
        COMPANY_ALIAS_TO_CODE.put("SPK", "OUT800");
        COMPANY_ALIAS_TO_CODE.put("COSMETICS", "OUT600");
        COMPANY_ALIAS_TO_CODE.put("COS", "OUT600");
        COMPANY_ALIAS_TO_CODE.put("OUT010", "OUT010");
        COMPANY_ALIAS_TO_CODE.put("OUT100", "OUT100");
        COMPANY_ALIAS_TO_CODE.put("OUT200", "OUT200");
        COMPANY_ALIAS_TO_CODE.put("OUT300", "OUT300");
        COMPANY_ALIAS_TO_CODE.put("OUT400", "OUT400");
        COMPANY_ALIAS_TO_CODE.put("OUT500", "OUT500");
        COMPANY_ALIAS_TO_CODE.put("OUT600", "OUT600");
        COMPANY_ALIAS_TO_CODE.put("OUT700", "OUT700");
        COMPANY_ALIAS_TO_CODE.put("OUT800", "OUT800");
    }

    static class PORecord {
        String purchaseOrderNo;
        String supplier;
        String product;
        String sku;
        String barcode;
        String date;
        int quantity;
        String shop;

        PORecord(String purchaseOrderNo, String supplier, String product, String sku,
                 String barcode, String date, int quantity, String shop) {
            this.purchaseOrderNo = purchaseOrderNo;
            this.supplier = supplier;
            this.product = product;
            this.sku = sku;
            this.barcode = barcode;
            this.date = date;
            this.quantity = quantity;
            this.shop = shop;
        }
    }

    static class StockRecord {
        String sku;
        String barcode;
        String date;
        String reason;
        int adjustment;
        String company;
        String saId;
        String companyCode;
        String sourceFile; // keep track of original file

        StockRecord(String sku, String barcode, String date, String reason, int adjustment,
                    String company, String saId, String companyCode, String sourceFile) {
            this.sku = sku;
            this.barcode = barcode;
            this.date = date;
            this.reason = reason;
            this.adjustment = adjustment;
            this.company = company;
            this.saId = saId;
            this.companyCode = companyCode;
            this.sourceFile = sourceFile;
        }
    }

    static class TallyRecord {
        String poNo;
        String company;
        String companyCode;
        String supplier;
        String shop;
        String product;
        String sku;
        String barcode;
        String date;
        int quantity;
        String reason;
        String idConflict;
        String remarks;
        String saId;
        String sourceFile; // track file for third-pass

        boolean companyMatched;
        boolean shopMatched;

        TallyRecord(String poNo, String company, String companyCode, String supplier, String shop, String product,
                    String sku, String barcode, String date, int quantity, String reason, String idConflict,
                    String remarks, String saId, String sourceFile) {
            this.poNo = poNo;
            this.company = company;
            this.companyCode = companyCode;
            this.supplier = supplier;
            this.shop = shop;
            this.product = product;
            this.sku = sku;
            this.barcode = barcode;
            this.date = date;
            this.quantity = quantity;
            this.reason = reason;
            this.idConflict = idConflict;
            this.remarks = remarks;
            this.saId = saId;
            this.sourceFile = sourceFile;
            this.companyMatched = false;
            this.shopMatched = false;
        }
    }

    // Main entry with excludeSAIds
    public static void generateReport(List<File> purchaseOrderFiles, List<File> stockAdjustmentFiles,
                                      File output, List<String> excludeSAIds) throws IOException {
        System.out.println("=== STARTING REPORT GENERATION ===");
        if (excludeSAIds == null) excludeSAIds = new ArrayList<>();
        System.out.println("Exclude SA IDs: " + excludeSAIds);

        final List<String> finalExcludeSAIds = new ArrayList<>(excludeSAIds); // make a final copy

        List<PORecord> allPORecords = new ArrayList<>();
        for (File file : purchaseOrderFiles) {
            List<PORecord> recs = readPurchaseOrderFile(file);
            allPORecords.addAll(recs);
        }

        List<StockRecord> allStockRecords = new ArrayList<>();
        for (File file : stockAdjustmentFiles) {
            List<StockRecord> recs = readStockAdjustmentFile(file).stream()
                    .filter(r -> r.saId == null || !finalExcludeSAIds.contains(r.saId)) // use final copy
                    .collect(Collectors.toList());
            allStockRecords.addAll(recs);
        }

        List<TallyRecord> tallyRecords = generateTallyRecords(allPORecords, allStockRecords);
        writeTallyReport(tallyRecords, output);

        System.out.println("=== REPORT GENERATION COMPLETE ===");
    }

    public static void generateReport(List<File> purchaseOrderFiles, List<File> stockAdjustmentFiles, File output) throws IOException {
        generateReport(purchaseOrderFiles, stockAdjustmentFiles, output, new ArrayList<>());
    }

    private static List<PORecord> readPurchaseOrderFile(File file) throws IOException {
        List<PORecord> records = new ArrayList<>();
        String shopName = file.getName().replaceAll("\\.(xlsx|xls)$", "");

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) return records;

            Map<String, Integer> columnMap = new HashMap<>();
            for (Cell cell : headerRow) {
                String header = dataFormatter.formatCellValue(cell).trim();
                switch (header) {
                    case "Purchase Order": columnMap.put("PurchaseOrder", cell.getColumnIndex()); break;
                    case "Supplier": columnMap.put("Supplier", cell.getColumnIndex()); break;
                    case "Product": columnMap.put("Product", cell.getColumnIndex()); break;
                    case "SKU": columnMap.put("SKU", cell.getColumnIndex()); break;
                    case "Barcode": columnMap.put("Barcode", cell.getColumnIndex()); break;
                    case "PO Date": columnMap.put("PODate", cell.getColumnIndex()); break;
                    case "Quantity": columnMap.put("Quantity", cell.getColumnIndex()); break;
                }
            }

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                String po = getCellValue(row, columnMap.get("PurchaseOrder"));
                String supplierCell = getCellValue(row, columnMap.get("Supplier"));
                String product = getCellValue(row, columnMap.get("Product"));
                String sku = getCellValue(row, columnMap.get("SKU"));
                String barcode = getCellValue(row, columnMap.get("Barcode"));
                String date = getCellValue(row, columnMap.get("PODate"));
                String qtyStr = getCellValue(row, columnMap.get("Quantity"));

                String supplierCode = deriveCompanyCodeFromSupplierCell(supplierCell);
                if (!COMPANY_ALIAS_TO_CODE.containsValue(supplierCode) && !"OUT010".equals(supplierCode)) continue;

                if (po != null && !po.isEmpty() && supplierCode != null && !supplierCode.isEmpty()) {
                    try {
                        int qty = Integer.parseInt(qtyStr.trim());
                        if (qty > 0) {
                            records.add(new PORecord(po.trim(), supplierCode, safeTrim(product),
                                    safeTrim(sku), safeTrim(barcode), safeTrim(date), qty, safeTrim(shopName)));
                        }
                    } catch (NumberFormatException e) {}
                }
            }
        } catch (Exception e) {
            System.out.println("Error reading PO file '" + file.getName() + "': " + e.getMessage());
            e.printStackTrace();
        }
        return records;
    }

    private static List<StockRecord> readStockAdjustmentFile(File file) throws IOException {
        List<StockRecord> records = new ArrayList<>();
        String companyName = file.getName().replaceAll("\\.(xlsx|xls)$", "").trim();
        String companyCode = deriveCompanyCodeFromFileName(companyName);

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) return records;

            Map<String, Integer> columnMap = new HashMap<>();
            for (Cell cell : headerRow) {
                String header = dataFormatter.formatCellValue(cell).trim();
                switch (header) {
                    case "SKU": columnMap.put("SKU", cell.getColumnIndex()); break;
                    case "Barcode": columnMap.put("Barcode", cell.getColumnIndex()); break;
                    case "Date": columnMap.put("Date", cell.getColumnIndex()); break;
                    case "Reason": columnMap.put("Reason", cell.getColumnIndex()); break;
                    case "Adjustment": columnMap.put("Adjustment", cell.getColumnIndex()); break;
                    case "No.": columnMap.put("SAID", cell.getColumnIndex()); break;
                }
            }

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                String sku = getCellValue(row, columnMap.get("SKU"));
                String barcode = getCellValue(row, columnMap.get("Barcode"));
                String date = getCellValue(row, columnMap.get("Date"));
                String reason = getCellValue(row, columnMap.get("Reason"));
                String adjStr = getCellValue(row, columnMap.get("Adjustment"));
                String saId = getCellValue(row, columnMap.get("SAID"));

                if (sku != null && !sku.isEmpty()) {
                    try {
                        int adjustment = Integer.parseInt(adjStr.trim());
                        records.add(new StockRecord(safeTrim(sku), safeTrim(barcode), safeTrim(date),
                                safeTrim(reason), adjustment, safeTrim(companyName), safeTrim(saId), safeTrim(companyCode),
                                file.getName()));
                    } catch (NumberFormatException e) {}
                }
            }
        } catch (Exception e) {
            System.out.println("Error reading Stock file '" + file.getName() + "': " + e.getMessage());
            e.printStackTrace();
        }
        return records;
    }

    private static List<TallyRecord> generateTallyRecords(List<PORecord> poRecords, List<StockRecord> stockRecords) {
        List<TallyRecord> tally = new ArrayList<>();
        for (PORecord p : poRecords) {
            tally.add(new TallyRecord(
                    p.purchaseOrderNo, "", "", p.supplier, p.shop, p.product,
                    p.sku, p.barcode, p.date, p.quantity, "", "", "Pending", "", ""
            ));
        }

        for (StockRecord s : stockRecords) {
            tally.add(new TallyRecord(
                    "", s.company, s.companyCode, "", "", "",
                    s.sku, s.barcode, s.date, s.adjustment, s.reason, "", "Pending", s.saId, s.sourceFile
            ));
        }

        detectIDConflicts(tally);
        performTallyMatching(tally);
        performSecondPassMatching(tally);
        performThirdPassStockMatching(tally);
        updateRemarksForUnmatched(tally);

        return tally;
    }

    //Matching methods
    private static void detectIDConflicts(List<TallyRecord> records) {
        Map<String, List<TallyRecord>> dateSkuGroups = new HashMap<>();
        Map<String, List<TallyRecord>> dateBarcodeGroups = new HashMap<>();

        for (TallyRecord r : records) {
            if (!isEmptyString(r.date)) {
                String skuKey = r.date + "|" + (r.sku == null ? "" : r.sku);
                dateSkuGroups.computeIfAbsent(skuKey, k -> new ArrayList<>()).add(r);

                if (!isEmptyString(r.barcode) && !"No Barcode".equalsIgnoreCase(r.barcode)) {
                    String barcodeKey = r.date + "|" + r.barcode;
                    dateBarcodeGroups.computeIfAbsent(barcodeKey, k -> new ArrayList<>()).add(r);
                }
            }
        }

        for (List<TallyRecord> group : dateSkuGroups.values()) {
            Set<String> barcodes = group.stream()
                    .map(rec -> rec.barcode)
                    .filter(b -> !isEmptyString(b) && !"No Barcode".equalsIgnoreCase(b))
                    .collect(Collectors.toSet());
            if (barcodes.size() > 1) {
                String conflict = "Same SKU different barcodes: " + String.join(", ", barcodes);
                for (TallyRecord rec : group) rec.idConflict = conflict;
            }
        }

        for (List<TallyRecord> group : dateBarcodeGroups.values()) {
            Set<String> skus = group.stream()
                    .map(rec -> rec.sku)
                    .filter(s -> !isEmptyString(s))
                    .collect(Collectors.toSet());
            if (skus.size() > 1) {
                String conflict = "Same barcode different SKUs: " + String.join(", ", skus);
                for (TallyRecord rec : group) {
                    if (isEmptyString(rec.idConflict)) rec.idConflict = conflict;
                    else rec.idConflict += "; " + conflict;
                }
            }
        }
    }

    private static void performTallyMatching(List<TallyRecord> records) {
        List<TallyRecord> poList = records.stream()
                .filter(r -> !isEmptyString(r.poNo) && r.quantity > 0)
                .collect(Collectors.toList());

        List<TallyRecord> stockList = records.stream()
                .filter(r -> !isEmptyString(r.company) && r.quantity < 0)
                .collect(Collectors.toList());

        Set<TallyRecord> matchedStocks = new HashSet<>();

        for (TallyRecord po : poList) {
            for (TallyRecord stock : stockList) {
                if (matchedStocks.contains(stock)) continue;
                if (isEmptyString(po.supplier) || isEmptyString(stock.companyCode)) continue;

                if (("OUT010".equals(po.supplier) && "OUT600".equals(stock.companyCode))
                        || po.supplier.equalsIgnoreCase(stock.companyCode)) {
                    if (isExactMatch(po, stock)) {
                        if (isEmptyString(po.company)) {
                            po.company = stock.company;
                            po.companyMatched = true;
                        }
                        if (isEmptyString(po.saId)) po.saId = stock.saId;
                        po.remarks = "Tally";
                        stock.poNo = po.poNo;
                        if (isEmptyString(stock.supplier)) stock.supplier = po.supplier;
                        if (isEmptyString(stock.shop)) {
                            stock.shop = po.shop;
                            stock.shopMatched = true;
                        }
                        stock.remarks = "Tally";
                        matchedStocks.add(stock);
                        break;
                    }
                }
            }
        }
    }

    private static void performSecondPassMatching(List<TallyRecord> records) {
        List<TallyRecord> unmatchedPO = records.stream()
                .filter(r -> !isEmptyString(r.poNo) && (isEmptyString(r.remarks) || !r.remarks.startsWith("Tally")))
                .collect(Collectors.toList());

        List<TallyRecord> unmatchedStock = records.stream()
                .filter(r -> !isEmptyString(r.company) && (isEmptyString(r.remarks) || !r.remarks.startsWith("Tally")))
                .collect(Collectors.toList());

        Set<TallyRecord> matchedStocks = new HashSet<>();

        for (TallyRecord po : unmatchedPO) {
            for (TallyRecord stock : unmatchedStock) {
                if (matchedStocks.contains(stock)) continue;
                if (isEmptyString(po.supplier) || isEmptyString(stock.companyCode)) continue;

                if (("OUT010".equals(po.supplier) && "OUT600".equals(stock.companyCode))
                        || po.supplier.equalsIgnoreCase(stock.companyCode)) {
                    if (!isEmptyString(po.sku) && po.sku.equals(stock.sku)
                            && Math.abs(stock.quantity) == po.quantity
                            && isWithinOneWeek(po.date, stock.date)) {

                        if (isEmptyString(po.company)) {
                            po.company = stock.company;
                            po.companyMatched = true;
                        }
                        if (isEmptyString(po.saId)) po.saId = stock.saId;
                        po.remarks = "Tally (2nd pass)";
                        stock.poNo = po.poNo;
                        if (isEmptyString(stock.shop)) {
                            stock.shop = po.shop;
                            stock.shopMatched = true;
                        }
                        stock.remarks = "Tally (2nd pass)";
                        matchedStocks.add(stock);
                        break;
                    }
                }
            }
        }
    }

    private static void performThirdPassStockMatching(List<TallyRecord> records) {
        List<TallyRecord> unmatchedStock = records.stream()
                .filter(r -> !isEmptyString(r.company) && (isEmptyString(r.remarks) || !r.remarks.startsWith("Tally")))
                .collect(Collectors.toList());

        Set<TallyRecord> matchedStocks = new HashSet<>();

        for (int i = 0; i < unmatchedStock.size(); i++) {
            TallyRecord s1 = unmatchedStock.get(i);
            if (matchedStocks.contains(s1)) continue;

            for (int j = i + 1; j < unmatchedStock.size(); j++) {
                TallyRecord s2 = unmatchedStock.get(j);
                if (matchedStocks.contains(s2)) continue;

                if (!s1.sourceFile.equals(s2.sourceFile)) continue;

                if (!isEmptyString(s1.sku) && s1.sku.equals(s2.sku)
                        && s1.date.equals(s2.date)
                        && s1.quantity == -s2.quantity) {

                    s1.remarks = "Tally (3rd pass)";
                    s2.remarks = "Tally (3rd pass)";
                    matchedStocks.add(s1);
                    matchedStocks.add(s2);
                    break;
                }
            }
        }
    }

    private static boolean isExactMatch(TallyRecord po, TallyRecord stock) {
        boolean dateMatch = po.date.equals(stock.date);
        boolean skuMatch = po.sku.equals(stock.sku);
        boolean barcodeMatch = (isEmptyString(po.barcode) && isEmptyString(stock.barcode))
                || po.barcode.equals(stock.barcode);
        boolean qtyMatch = po.quantity == Math.abs(stock.quantity);

        return dateMatch && skuMatch && barcodeMatch && qtyMatch;
    }

    private static boolean isWithinOneWeek(String poDateStr, String stockDateStr) {
        try {
            LocalDate poDate = LocalDate.parse(poDateStr, dateFormatter);
            LocalDate stockDate = LocalDate.parse(stockDateStr, dateFormatter);
            long daysBetween = Math.abs(ChronoUnit.DAYS.between(poDate, stockDate));
            return daysBetween <= 7;
        } catch (Exception e) {
            return false;
        }
    }

    private static void updateRemarksForUnmatched(List<TallyRecord> records) {
        for (TallyRecord r : records) {
            if (isEmptyString(r.remarks) || "Pending".equalsIgnoreCase(r.remarks)) {
                if (!isEmptyString(r.poNo)) {
                    r.remarks = "Mismatch: no matching stock adjustment";
                } else if (!isEmptyString(r.company)) {
                    r.remarks = "Mismatch: no matching purchase order";
                } else {
                    r.remarks = "Unmatched"; // fallback
                }
            }
        }
    }



    //Output

    private static void writeTallyReport(List<TallyRecord> records, File output) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Tally Report");

        // Updated header row (Company Code removed)
        String[] headers = {"PO No", "Company", "Supplier", "Shop","In/Out",
                "Product", "SKU", "Barcode", "Date", "Quantity", "Reason", "ID Conflict", "Remarks", "SA ID"};
        Row headerRow = sheet.createRow(0);
        for (int c = 0; c < headers.length; c++) {
            Cell cell = headerRow.createCell(c);
            cell.setCellValue(headers[c]);
        }

        // Define cell styles
        CellStyle greenStyle = workbook.createCellStyle();
        greenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle defaultStyle = workbook.createCellStyle(); // <-- make sure this exists

        // Orange style for Shop column
        XSSFCellStyle orangeStyle = (XSSFCellStyle) workbook.createCellStyle();
        orangeStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        orangeStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(255, 200, 120), null));
        int rowIndex = 1;
        for (TallyRecord r : records) {
            Row row = sheet.createRow(rowIndex++);

            // PO No
            row.createCell(0).setCellValue(r.poNo);

            // Company (color if matched)
            Cell companyCell = row.createCell(1);
            companyCell.setCellValue(r.company);
            if (r.companyMatched) companyCell.setCellStyle(greenStyle);
            else companyCell.setCellStyle(defaultStyle);

            // Supplier
            row.createCell(2).setCellValue(r.supplier);

            // Shop (color if matched)
            Cell shopCell = row.createCell(3);
            shopCell.setCellValue(r.shop);
            if (r.shopMatched) shopCell.setCellStyle(orangeStyle);
            else shopCell.setCellStyle(defaultStyle);

            Cell inOutCell = row.createCell(4);

            if (r.quantity > 0) {
                inOutCell.setCellValue("IN");
            } else if (r.quantity < 0) {
                inOutCell.setCellValue("OUT");
            } else {
                inOutCell.setCellValue("");
            }

            inOutCell.setCellStyle(defaultStyle);
            // Product
            row.createCell(5).setCellValue(r.product);

            // SKU
            row.createCell(6).setCellValue(r.sku);

            // Barcode
            row.createCell(7).setCellValue(r.barcode);

            // Date
            row.createCell(8).setCellValue(r.date);

            // Quantity
            row.createCell(9).setCellValue(r.quantity);

            // Reason
            row.createCell(10).setCellValue(r.reason);

            // ID Conflict
            row.createCell(11).setCellValue(r.idConflict);

            // Remarks
            row.createCell(12).setCellValue(r.remarks);

            // SA ID
            row.createCell(13).setCellValue(r.saId);
        }

        try (FileOutputStream fos = new FileOutputStream(output)) {
            workbook.write(fos);
        }
        workbook.close();
    }



    private static String getCellValue(Row row, Integer colIndex) {
        if (colIndex == null) return "";
        Cell cell = row.getCell(colIndex);
        if (cell == null) return "";
        return dataFormatter.formatCellValue(cell).trim();
    }

    private static String safeTrim(String s) {
        return s == null ? "" : s.trim();
    }

    private static boolean isEmptyString(String s) {
        return s == null || s.trim().isEmpty();
    }

    private static String deriveCompanyCodeFromFileName(String fileName) {
        for (Map.Entry<String, String> e : COMPANY_ALIAS_TO_CODE.entrySet()) {
            if (fileName.toUpperCase().contains(e.getKey().toUpperCase())) return e.getValue();
        }
        return "";
    }

    private static String deriveCompanyCodeFromSupplierCell(String supplier) {
        if (supplier == null) return "";
        for (Map.Entry<String, String> e : COMPANY_ALIAS_TO_CODE.entrySet()) {
            if (supplier.toUpperCase().contains(e.getKey().toUpperCase())) return e.getValue();
        }
        return supplier.toUpperCase();
    }
}
