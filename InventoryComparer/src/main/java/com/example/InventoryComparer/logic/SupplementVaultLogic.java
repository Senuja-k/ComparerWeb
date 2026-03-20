package com.example.InventoryComparer.logic;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;
import java.util.stream.Collectors;

public class SupplementVaultLogic {

    private static final DataFormatter dataFormatter = new DataFormatter();

    // ===== Data Structures =====

    /** One row from a merged order report */
    static class OrderRow {
        String company; // "Origins" or "SupplementVault"
        Map<String, String> data; // column name -> value
        double total;
        String financialStatus;
        String discountCode;

        OrderRow(String company, Map<String, String> data) {
            this.company = company;
            this.data = data;
            this.financialStatus = getVal("Financial Status");
            this.discountCode = normalizeCode(getVal("Discount Code"));
            String totalStr = getVal("Total");
            this.total = parseDouble(totalStr);
        }

        String getVal(String col) {
            if (col == null) return "";
            // direct key
            for (Map.Entry<String, String> e : data.entrySet()) {
                if (e.getKey() != null && e.getKey().equalsIgnoreCase(col)) return e.getValue();
            }

            // relaxed contains match (ignore spaces/underscores/hyphens)
            String normCol = normalizeHeader(col);
            for (Map.Entry<String, String> e : data.entrySet()) {
                String key = e.getKey() == null ? "" : normalizeHeader(e.getKey());
                if (key.contains(normCol) || normCol.contains(key)) return e.getValue();
            }

            return "";
        }
    }

    /** Merchant coupon code mapping */
    static class MerchantCoupon {
        String merchantName;
        String discountCode;
        String type; // "Online" or "Outlet"

        MerchantCoupon(String merchantName, String discountCode, String type) {
            this.merchantName = merchantName == null ? "" : merchantName.trim();
            this.discountCode = discountCode == null ? "" : discountCode.trim();
            this.type = type == null ? "" : type.trim();
        }
    }

    /** A row in the target table */
    static class TargetRow {
        int rowIndex; // row index in the target sheet
        String outlet;
        String merchantName;
        double target;
    }

    /** Sales result per merchant per company */
    static class MerchantSales {
        double originsSale = 0;
        double svSale = 0;

        double totalSale() { return originsSale + svSale; }
    }

    // ===== Main Entry Point =====

    public static void generateReport(
            List<File> orderFiles, File couponFile, File targetFile,
            int daysRemainingOnline, int daysRemainingOutlet,
            int totalDays, int reportDay, File outputFile
    ) throws Exception {

        // 1. Read all order reports, tag with company, merge into ALL
        List<OrderRow> allOrders = new ArrayList<>();
        List<String> orderHeaders = null;

        for (File f : orderFiles) {
            String fileName = f.getName().toLowerCase();
            String company;
            if (fileName.contains("origin")) {
                company = "Origins";
            } else {
                company = "SupplementVault";
            }

            List<Map<String, String>> rows = readExcelOrCsv(f);
            if (!rows.isEmpty() && orderHeaders == null) {
                orderHeaders = new ArrayList<>(rows.get(0).keySet());
                System.out.println("[SV-DEBUG] Order file headers: " + orderHeaders);
            }
            System.out.println("[SV-DEBUG] Order file \"" + f.getName() + "\" -> company=" + company + ", rows=" + rows.size());
            // Log first 5 rows to see what data looks like
            for (int i = 0; i < Math.min(5, rows.size()); i++) {
                Map<String, String> rowData = rows.get(i);
                OrderRow sample = new OrderRow(company, rowData);
                System.out.println("[SV-DEBUG]   Row " + i + ": financialStatus=\"" + sample.financialStatus
                        + "\" discountCode=\"" + sample.discountCode + "\" total=" + sample.total);
            }
            for (Map<String, String> row : rows) {
                allOrders.add(new OrderRow(company, row));
            }
        }

        System.out.println("[SV-DEBUG] Total allOrders: " + allOrders.size());

        // 2. Filter: only keep paid or pending
        List<OrderRow> filtered = allOrders.stream()
                .filter(r -> {
                    String fs = r.financialStatus.trim().toLowerCase();
                    return fs.equals("paid") || fs.equals("pending");
                })
                .collect(Collectors.toList());

        System.out.println("[SV-DEBUG] After filtering (paid/pending): " + filtered.size() + " orders remain");
        // Log first 5 filtered rows
        for (int i = 0; i < Math.min(5, filtered.size()); i++) {
            OrderRow r2 = filtered.get(i);
            System.out.println("[SV-DEBUG]   Filtered[" + i + "]: code=\"" + r2.discountCode + "\" total=" + r2.total + " company=" + r2.company);
        }

        // 3. Read Merchant Coupon Code file
        List<MerchantCoupon> coupons = readCouponFile(couponFile);

        // Build discount code -> merchant mapping (case-insensitive)
        Map<String, MerchantCoupon> codeToMerchant = new HashMap<>();
        for (MerchantCoupon mc : coupons) {
            if (mc.discountCode != null && !mc.discountCode.trim().isEmpty()) {
                String normCode = normalizeCode(mc.discountCode);
                codeToMerchant.put(normCode, mc);
                System.out.println("[SV-DEBUG] Coupon map: normalizedCode=\"" + normCode + "\" -> merchant=\"" + mc.merchantName + "\"");
            }
        }

        // Build merchant name -> type mapping
        Map<String, String> merchantTypeMap = new HashMap<>();
        for (MerchantCoupon mc : coupons) {
            if (!mc.merchantName.isEmpty()) {
                merchantTypeMap.put(mc.merchantName.toLowerCase(), mc.type);
            }
        }

        // 4. Calculate sales per merchant per company
        Map<String, MerchantSales> merchantSalesMap = new LinkedHashMap<>();
        // Track unmatched discount code totals (for DM General/Sandali)
        double unmatchedOrigins = 0;
        double unmatchedSV = 0;

        // Debug: Track unique discount codes seen and their match status
        Map<String, String> codeMatchLog = new LinkedHashMap<>();

        for (OrderRow row : filtered) {
            String rawCode = row.discountCode;  // already normalized in constructor
            String code = normalizeCode(row.discountCode);
            MerchantCoupon mc = codeToMerchant.get(code);

            if (mc != null && !mc.merchantName.isEmpty()) {
                String merchantKey = mc.merchantName;
                MerchantSales sales = merchantSalesMap.computeIfAbsent(merchantKey, k -> new MerchantSales());
                if ("Origins".equals(row.company)) {
                    sales.originsSale += row.total;
                } else {
                    sales.svSale += row.total;
                }
                codeMatchLog.put(rawCode, "MATCHED -> " + merchantKey);
            } else {
                // Unmatched code -> goes to DM General/Sandali
                if ("Origins".equals(row.company)) {
                    unmatchedOrigins += row.total;
                } else {
                    unmatchedSV += row.total;
                }
                if (!code.isEmpty()) {
                    codeMatchLog.put(rawCode, "UNMATCHED (normalized: \"" + code + "\")");
                } else {
                    codeMatchLog.put("(empty)", "UNMATCHED - no discount code");
                }
            }
        }

        System.out.println("[SV-DEBUG] === Discount Code Match Results ===");
        for (Map.Entry<String, String> entry : codeMatchLog.entrySet()) {
            System.out.println("[SV-DEBUG]   Code=\"" + entry.getKey() + "\" => " + entry.getValue());
        }

        // Add unmatched to DM General/Sandali
        if (unmatchedOrigins != 0 || unmatchedSV != 0) {
            MerchantSales dmSales = merchantSalesMap.computeIfAbsent("DM General/Sandali", k -> new MerchantSales());
            dmSales.originsSale += unmatchedOrigins;
            dmSales.svSale += unmatchedSV;
        }

        // 5. Read the target table file
        List<TargetRow> targetRows = readTargetTable(targetFile);

        // Build set of merchant names in the target table (case-insensitive)
        Set<String> targetMerchantNames = new HashSet<>();
        for (TargetRow tr : targetRows) {
            if (!tr.merchantName.isEmpty()) {
                targetMerchantNames.add(tr.merchantName.toLowerCase());
            }
        }

        // Identify merchants with sales but not in target table
        Map<String, MerchantSales> extraMerchants = new LinkedHashMap<>();
        for (Map.Entry<String, MerchantSales> entry : merchantSalesMap.entrySet()) {
            if (!targetMerchantNames.contains(entry.getKey().toLowerCase())) {
                extraMerchants.put(entry.getKey(), entry.getValue());
            }
        }

        // 6. Write the output workbook
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            // --- Sheet 1: ALL (merged order data) ---
            writeAllSheet(wb, allOrders, orderHeaders);

            // --- Sheet 2: Sales (merchant sales breakdown) ---
            writeSalesSheet(wb, merchantSalesMap);

            // --- Sheet 3: Report (filled-in target table) ---
            writeReportSheet(wb, targetFile, targetRows, merchantSalesMap, merchantTypeMap,
                    daysRemainingOnline, daysRemainingOutlet, totalDays, reportDay);

            // --- Sheet 4: Extra Merchants (not in target table) ---
            if (!extraMerchants.isEmpty()) {
                writeExtraMerchantsSheet(wb, extraMerchants);
            }

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                wb.write(fos);
            }
        }
    }

    // ===== Sheet Writers =====

    private static void writeAllSheet(XSSFWorkbook wb, List<OrderRow> allOrders, List<String> headers) {
        Sheet sheet = wb.createSheet("ALL");
        if (headers == null || allOrders.isEmpty()) return;

        // Add "Company" column
        List<String> allHeaders = new ArrayList<>();
        allHeaders.add("Company");
        allHeaders.addAll(headers);

        // Header row
        Row headerRow = sheet.createRow(0);
        CellStyle headerStyle = wb.createCellStyle();
        Font headerFont = wb.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);

        for (int i = 0; i < allHeaders.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(allHeaders.get(i));
            cell.setCellStyle(headerStyle);
        }

        // Data rows
        int rowIdx = 1;
        for (OrderRow or : allOrders) {
            Row row = sheet.createRow(rowIdx++);
            row.createCell(0).setCellValue(or.company);
            for (int i = 0; i < headers.size(); i++) {
                String val = or.data.getOrDefault(headers.get(i), "");
                Cell cell = row.createCell(i + 1);
                try {
                    double d = Double.parseDouble(val);
                    cell.setCellValue(d);
                } catch (NumberFormatException e) {
                    cell.setCellValue(val);
                }
            }
        }
    }

    private static void writeSalesSheet(XSSFWorkbook wb, Map<String, MerchantSales> salesMap) {
        Sheet sheet = wb.createSheet("Sales");

        // Header
        Row headerRow = sheet.createRow(0);
        CellStyle headerStyle = wb.createCellStyle();
        Font headerFont = wb.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);

        String[] cols = {"Merchant Name", "Origins Sale", "SupplementVault Sale", "Total Sale"};
        for (int i = 0; i < cols.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(cols[i]);
            cell.setCellStyle(headerStyle);
        }

        CellStyle numStyle = wb.createCellStyle();
        numStyle.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));

        int rowIdx = 1;
        for (Map.Entry<String, MerchantSales> entry : salesMap.entrySet()) {
            Row row = sheet.createRow(rowIdx++);
            row.createCell(0).setCellValue(entry.getKey());

            Cell c1 = row.createCell(1);
            c1.setCellValue(entry.getValue().originsSale);
            c1.setCellStyle(numStyle);

            Cell c2 = row.createCell(2);
            c2.setCellValue(entry.getValue().svSale);
            c2.setCellStyle(numStyle);

            Cell c3 = row.createCell(3);
            c3.setCellValue(entry.getValue().totalSale());
            c3.setCellStyle(numStyle);
        }
    }

    private static void writeReportSheet(XSSFWorkbook wb, File targetFile,
                                         List<TargetRow> targetRows,
                                         Map<String, MerchantSales> merchantSalesMap,
                                         Map<String, String> merchantTypeMap,
                                         int daysRemainingOnline, int daysRemainingOutlet,
                                         int totalDays, int reportDay) throws Exception {

        // Copy the target file's first sheet as the Report sheet
        try (FileInputStream fis = new FileInputStream(targetFile);
             Workbook srcWb = WorkbookFactory.create(fis)) {

            Sheet srcSheet = srcWb.getSheetAt(0);
            Sheet destSheet = wb.createSheet("Report");

            // Find column indices from the header rows
            int headerRowIdx = findHeaderRow(srcSheet);
            if (headerRowIdx < 0) {
                copySheet(srcSheet, destSheet, wb);
                return;
            }

            // Copy all rows from source to destination (preserving styles)
            copySheet(srcSheet, destSheet, wb);

            // Find ALL column headers by scanning multiple rows (headerRowIdx-1, headerRowIdx, headerRowIdx+1)
            Row hRow = destSheet.getRow(headerRowIdx);
            if (hRow == null) return;
            Row rowAbove = headerRowIdx > 0 ? destSheet.getRow(headerRowIdx - 1) : null;
            Row rowBelow = destSheet.getRow(headerRowIdx + 1);

            int lastCol = hRow.getLastCellNum();
            if (rowAbove != null) lastCol = Math.max(lastCol, rowAbove.getLastCellNum());
            if (rowBelow != null) lastCol = Math.max(lastCol, rowBelow.getLastCellNum());

            System.out.println("[SV-DEBUG] === REPORT COLUMN DETECTION ===");
            System.out.println("[SV-DEBUG] Header row: " + headerRowIdx + ", lastCol: " + lastCol);

            int colMerchant = -1, colTarget = -1, colOriginSale = -1,
                    colSvSale = -1, colTotalSale = -1, colBalance = -1,
                    colPerDayTarget = -1, colForecast = -1;
            List<Integer> achievementPctCols = new ArrayList<>();

            // For each column, collect text from ALL three rows and match against ALL of them
            for (int c = 0; c < lastCol; c++) {
                String above = rowAbove != null ? getCellStr(rowAbove, c).toLowerCase().trim() : "";
                String main = getCellStr(hRow, c).toLowerCase().trim();
                String below = rowBelow != null ? getCellStr(rowBelow, c).toLowerCase().trim() : "";
                // Clean non-printable characters
                above = above.replaceAll("[^a-z0-9% /().,]", "").trim();
                main = main.replaceAll("[^a-z0-9% /().,]", "").trim();
                below = below.replaceAll("[^a-z0-9% /().,]", "").trim();
                String all = above + " " + main + " " + below;

                System.out.println("[SV-DEBUG]   Col " + c + ": above=\"" + above + "\" main=\"" + main + "\" below=\"" + below + "\"");

                // Check each row independently — sub-headers like "ORIGIN Sale" under
                // a merged "Achievement (Invoice Date)" must be found via 'below'
                if (colMerchant < 0 && all.contains("merchant")) colMerchant = c;
                if (colTarget < 0 && (isExactTarget(main) || isExactTarget(above) || isExactTarget(below))) colTarget = c;
                if (colOriginSale < 0 && (below.contains("origin") || (main.contains("origin") && !main.contains("achievement")))) colOriginSale = c;
                if (colSvSale < 0 && (below.contains("sv.lk") || below.contains("sv ") || main.contains("sv.lk") || main.contains("sv "))) colSvSale = c;
                if (colTotalSale < 0 && ((below.contains("total") && below.contains("sale")) || (main.contains("total") && main.contains("sale")))) colTotalSale = c;
                if (colBalance < 0 && all.contains("balance")) colBalance = c;
                if (colPerDayTarget < 0 && all.contains("per day")) colPerDayTarget = c;
                if (colForecast < 0 && all.contains("forecast") && (all.contains("month") || main.equals("forecast") || below.equals("forecast") || above.equals("forecast")) && !all.contains("achievement %")) colForecast = c;

                if (all.contains("achievement") && all.contains("%")) {
                    achievementPctCols.add(c);
                }
            }
            Collections.sort(achievementPctCols);

            int colAchievement = achievementPctCols.size() >= 1 ? achievementPctCols.get(0) : -1;
            int colForecastPct = achievementPctCols.size() >= 2 ? achievementPctCols.get(achievementPctCols.size() - 1) : -1;

            // Fallback for forecast value column
            if (colForecast < 0) {
                for (int c = 0; c < lastCol; c++) {
                    for (int rr = Math.max(0, headerRowIdx - 1); rr <= headerRowIdx + 1; rr++) {
                        Row scanRow = destSheet.getRow(rr);
                        if (scanRow == null) continue;
                        String v = getCellStr(scanRow, c).toLowerCase().trim();
                        if (v.contains("forecast") && (v.contains("month") || v.equals("forecast"))) {
                            colForecast = c;
                            break;
                        }
                    }
                    if (colForecast >= 0) break;
                }
            }

            System.out.println("[SV-DEBUG] Column mapping: Merchant=" + colMerchant + " Target=" + colTarget
                    + " Origin=" + colOriginSale + " SV=" + colSvSale + " Total=" + colTotalSale
                    + " Achievement%=" + colAchievement + " Balance=" + colBalance
                    + " PerDayTarget=" + colPerDayTarget + " Forecast=" + colForecast
                    + " ForecastPct=" + colForecastPct);
            System.out.println("[SV-DEBUG] All achievement% cols: " + achievementPctCols);

            // Fallback styles (only used for cells that don't already have a style)
            CellStyle numStyle = wb.createCellStyle();
            numStyle.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
            CellStyle pctStyle = wb.createCellStyle();
            pctStyle.setDataFormat(wb.createDataFormat().getFormat("0%"));

            // Fill in data for each target row
            for (TargetRow tr : targetRows) {
                Row row = destSheet.getRow(tr.rowIndex);
                if (row == null) continue;

                String merchantKey = tr.merchantName;
                MerchantSales sales = findSalesForMerchant(merchantKey, merchantSalesMap);

                double originSale = sales != null ? sales.originsSale : 0;
                double svSale = sales != null ? sales.svSale : 0;

                // Determine merchant type for days remaining
                String type = getMerchantType(merchantKey, merchantTypeMap);
                int daysRemaining = type.toLowerCase().contains("outlet") ? daysRemainingOutlet : daysRemainingOnline;

                // Excel row number (1-based)
                int excelRow = tr.rowIndex + 1;

                // Write sale values (these are raw data, not calculated)
                if (colOriginSale >= 0) setCellNum(row, colOriginSale, originSale, numStyle);
                if (colSvSale >= 0) setCellNum(row, colSvSale, svSale, numStyle);

                // Column letters for formula references
                String colLetterOrigin = colOriginSale >= 0 ? colLetter(colOriginSale) : "";
                String colLetterSv = colSvSale >= 0 ? colLetter(colSvSale) : "";
                String colLetterTotal = colTotalSale >= 0 ? colLetter(colTotalSale) : "";
                String colLetterTarget = colTarget >= 0 ? colLetter(colTarget) : "";
                String colLetterBalance = colBalance >= 0 ? colLetter(colBalance) : "";
                String colLetterForecast = colForecast >= 0 ? colLetter(colForecast) : "";

                // Total Sale = ORIGIN Sale + SV.LK Sale
                if (colTotalSale >= 0 && colOriginSale >= 0 && colSvSale >= 0) {
                    String formula = colLetterOrigin + excelRow + "+" + colLetterSv + excelRow;
                    setCellFormula(row, colTotalSale, formula, numStyle);
                } else if (colTotalSale >= 0) {
                    setCellNum(row, colTotalSale, originSale + svSale, numStyle);
                }

                // Achievement % = IF(Target=0, 0, Total Sale / Target)
                if (colAchievement >= 0 && colTotalSale >= 0 && colTarget >= 0) {
                    String totalRef = colLetterTotal + excelRow;
                    String targetRef = colLetterTarget + excelRow;
                    String formula = "IF(" + targetRef + "=0,0," + totalRef + "/" + targetRef + ")";
                    setCellFormula(row, colAchievement, formula, pctStyle);
                }

                // Balance = MAX(Target - Total Sale, 0)
                if (colBalance >= 0 && colTarget >= 0 && colTotalSale >= 0) {
                    String totalRef = colLetterTotal + excelRow;
                    String targetRef = colLetterTarget + excelRow;
                    String formula = "MAX(" + targetRef + "-" + totalRef + ",0)";
                    setCellFormula(row, colBalance, formula, numStyle);
                }

                // Per Day Target = IF(Balance=0, 0, Balance / daysRemaining)
                if (colPerDayTarget >= 0 && colBalance >= 0 && daysRemaining > 0) {
                    String balanceRef = colLetterBalance + excelRow;
                    String formula = "IF(" + balanceRef + "=0,0," + balanceRef + "/" + daysRemaining + ")";
                    setCellFormula(row, colPerDayTarget, formula, numStyle);
                }

                // Forecast Month End Achievement = (Total Sale / reportDay) * totalDays
                if (colForecast >= 0 && colTotalSale >= 0 && reportDay > 0) {
                    String totalRef = colLetterTotal + excelRow;
                    String formula = "(" + totalRef + "/" + reportDay + ")*" + totalDays;
                    setCellFormula(row, colForecast, formula, numStyle);
                }

                // Forecast Achievement % = IF(Target=0, 0, Forecast / Target)
                if (colForecastPct >= 0 && colForecast >= 0 && colTarget >= 0) {
                    String forecastRef = colLetterForecast + excelRow;
                    String targetRef = colLetterTarget + excelRow;
                    String formula = "IF(" + targetRef + "=0,0," + forecastRef + "/" + targetRef + ")";
                    setCellFormula(row, colForecastPct, formula, pctStyle);
                }

                System.out.println("[SV-DEBUG] Merchant=\"" + merchantKey + "\" row=" + excelRow
                        + " daysRemaining=" + daysRemaining + " type=" + type);
            }
        }
    }

    private static void writeExtraMerchantsSheet(XSSFWorkbook wb, Map<String, MerchantSales> extraMerchants) {
        Sheet sheet = wb.createSheet("Other Merchants");

        Row headerRow = sheet.createRow(0);
        CellStyle headerStyle = wb.createCellStyle();
        Font headerFont = wb.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);

        String[] cols = {"Merchant Name", "Origins Sale", "SupplementVault Sale", "Total Sale"};
        for (int i = 0; i < cols.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(cols[i]);
            cell.setCellStyle(headerStyle);
        }

        CellStyle numStyle = wb.createCellStyle();
        numStyle.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));

        int rowIdx = 1;
        for (Map.Entry<String, MerchantSales> entry : extraMerchants.entrySet()) {
            Row row = sheet.createRow(rowIdx++);
            row.createCell(0).setCellValue(entry.getKey());
            Cell c1 = row.createCell(1); c1.setCellValue(entry.getValue().originsSale); c1.setCellStyle(numStyle);
            Cell c2 = row.createCell(2); c2.setCellValue(entry.getValue().svSale); c2.setCellStyle(numStyle);
            Cell c3 = row.createCell(3); c3.setCellValue(entry.getValue().totalSale()); c3.setCellStyle(numStyle);
        }
    }

    // ===== File Readers =====

    private static List<Map<String, String>> readExcelOrCsv(File file) throws Exception {
        String name = file.getName().toLowerCase();
        if (name.endsWith(".csv")) {
            return readCsv(file);
        } else {
            return readExcel(file);
        }
    }

    private static List<Map<String, String>> readExcel(File file) throws Exception {
        List<Map<String, String>> rows = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(file);
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet sheet = wb.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) return rows;

            List<String> headers = new ArrayList<>();
            for (int c = 0; c < headerRow.getLastCellNum(); c++) {
                Cell cell = headerRow.getCell(c);
                String hdr = cell != null ? dataFormatter.formatCellValue(cell).trim() : "Column" + c;
                headers.add(hdr);
            }

            System.out.println("[SV-DEBUG] readExcel \"" + file.getName() + "\" headers (" + headers.size() + "): " + headers);
            // Print hex for key columns to spot invisible chars
            for (int i = 0; i < headers.size(); i++) {
                String h = headers.get(i);
                if (h.toLowerCase().contains("discount") || h.toLowerCase().contains("financial") || h.toLowerCase().contains("total")) {
                    StringBuilder hex = new StringBuilder();
                    for (char ch : h.toCharArray()) hex.append(String.format("%04x ", (int) ch));
                    System.out.println("[SV-DEBUG]   Col " + i + " \"" + h + "\" hex=[" + hex.toString().trim() + "]");
                }
            }

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                Map<String, String> map = new LinkedHashMap<>();
                boolean hasData = false;
                for (int c = 0; c < headers.size(); c++) {
                    Cell cell = row.getCell(c);
                    String val = cell != null ? dataFormatter.formatCellValue(cell).trim() : "";
                    if (!val.isEmpty()) hasData = true;
                    map.put(headers.get(c), val);
                }
                if (hasData) rows.add(map);
            }
        }
        return rows;
    }

    private static List<Map<String, String>> readCsv(File file) throws Exception {
        List<Map<String, String>> rows = new ArrayList<>();
        try (BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(file), "UTF-8"))) {
            String headerLine = br.readLine();
            if (headerLine == null) return rows;

            // Handle BOM character
            if (headerLine.startsWith("\uFEFF")) headerLine = headerLine.substring(1);

            List<String> headers = parseCsvLine(headerLine);
            System.out.println("[SV-DEBUG] CSV headers (" + headers.size() + " columns): " + headers);

            String line;
            while ((line = br.readLine()) != null) {
                if (line.trim().isEmpty()) continue;
                List<String> parts = parseCsvLine(line);
                Map<String, String> map = new LinkedHashMap<>();
                boolean hasData = false;
                for (int i = 0; i < headers.size(); i++) {
                    String val = i < parts.size() ? parts.get(i).trim() : "";
                    if (!val.isEmpty()) hasData = true;
                    map.put(headers.get(i), val);
                }
                if (hasData) rows.add(map);
            }
        }
        return rows;
    }

    /** Parse a CSV line respecting quoted fields (handles commas inside quotes) */
    private static List<String> parseCsvLine(String line) {
        List<String> fields = new ArrayList<>();
        StringBuilder current = new StringBuilder();
        boolean inQuotes = false;

        for (int i = 0; i < line.length(); i++) {
            char ch = line.charAt(i);
            if (inQuotes) {
                if (ch == '"') {
                    // Check for escaped quote ""
                    if (i + 1 < line.length() && line.charAt(i + 1) == '"') {
                        current.append('"');
                        i++; // skip next quote
                    } else {
                        inQuotes = false;
                    }
                } else {
                    current.append(ch);
                }
            } else {
                if (ch == '"') {
                    inQuotes = true;
                } else if (ch == ',') {
                    fields.add(current.toString());
                    current.setLength(0);
                } else {
                    current.append(ch);
                }
            }
        }
        fields.add(current.toString());
        return fields;
    }

    private static List<MerchantCoupon> readCouponFile(File file) throws Exception {
        List<MerchantCoupon> coupons = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(file);
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet sheet = wb.getSheetAt(0);

            // Dump first 5 rows of the coupon file for debugging
            System.out.println("[SV-DEBUG] === COUPON FILE RAW DUMP (first 5 rows) ===");
            for (int r = 0; r <= Math.min(sheet.getLastRowNum(), 4); r++) {
                Row row = sheet.getRow(r);
                if (row == null) { System.out.println("[SV-DEBUG]   Row " + r + ": NULL"); continue; }
                StringBuilder sb = new StringBuilder("[SV-DEBUG]   Row " + r + " (cells=" + row.getLastCellNum() + "): ");
                for (int c = 0; c < row.getLastCellNum(); c++) {
                    Cell cell = row.getCell(c);
                    String val = cell != null ? dataFormatter.formatCellValue(cell) : "<null>";
                    sb.append("[").append(c).append("]=\"").append(val).append("\" ");
                }
                System.out.println(sb.toString());
            }

            // Scan rows 0-5 to find the header row (the one containing "coupon" or "code" or "owner")
            int headerRowIdx = -1;
            for (int r = 0; r <= Math.min(sheet.getLastRowNum(), 5); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                for (int c = 0; c < row.getLastCellNum(); c++) {
                    Cell cell = row.getCell(c);
                    if (cell == null) continue;
                    String val = dataFormatter.formatCellValue(cell).trim().toLowerCase();
                    if (val.contains("coupon") || val.contains("owner") || val.contains("mer")) {
                        headerRowIdx = r;
                        break;
                    }
                }
                if (headerRowIdx >= 0) break;
            }

            System.out.println("[SV-DEBUG] Coupon file header row detected at: " + headerRowIdx);
            if (headerRowIdx < 0) {
                System.out.println("[SV-DEBUG] WARNING: Could not find header row in coupon file!");
                return coupons;
            }

            Row headerRow = sheet.getRow(headerRowIdx);

            // Detect columns by scanning headers
            int colMerchant = -1, colCode = -1, colType = -1;
            System.out.println("[SV-DEBUG] Scanning coupon headers from row " + headerRowIdx + ":");
            for (int c = 0; c < headerRow.getLastCellNum(); c++) {
                Cell cell = headerRow.getCell(c);
                String h = cell != null ? dataFormatter.formatCellValue(cell).trim() : "";
                String hLower = h.toLowerCase();
                System.out.println("[SV-DEBUG]   Col " + c + ": \"" + h + "\"");

                // Merchant name column: "Coupon Code Owner"
                if (colMerchant < 0 && hLower.contains("owner")) {
                    colMerchant = c;
                }
                // Coupon Code column: "Coupon Code" (but not "Coupon Code Owner")
                else if (colCode < 0 && hLower.contains("code") && !hLower.contains("owner")) {
                    colCode = c;
                }
                // Type column: "Merchant Type" or "Type"
                else if (colType < 0 && hLower.contains("type")) {
                    colType = c;
                }
            }

            // Fallback: if code column still not found, try any column with "mer" (like MER01)
            if (colCode < 0) {
                // Maybe the header just says "Coupon Code" without "code" — try broader match
                for (int c = 0; c < headerRow.getLastCellNum(); c++) {
                    if (c == colMerchant || c == colType) continue;
                    String hLower = dataFormatter.formatCellValue(headerRow.getCell(c)).trim().toLowerCase();
                    if (hLower.contains("coupon") || hLower.contains("discount") || hLower.contains("code")) {
                        colCode = c;
                        break;
                    }
                }
            }

            System.out.println("[SV-DEBUG] Coupon file columns: Merchant=col" + colMerchant
                    + " Code=col" + colCode + " Type=col" + colType);

            int dataStartRow = headerRowIdx + 1;
            for (int r = dataStartRow; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                String merchantName = colMerchant >= 0 ? dataFormatter.formatCellValue(row.getCell(colMerchant)).trim() : "";
                String code = colCode >= 0 ? dataFormatter.formatCellValue(row.getCell(colCode)).trim() : "";
                String type = colType >= 0 ? dataFormatter.formatCellValue(row.getCell(colType)).trim() : "";

                if (merchantName.isEmpty() && code.isEmpty()) continue;

                coupons.add(new MerchantCoupon(merchantName, code, type));
            }

            System.out.println("[SV-DEBUG] Total coupons loaded: " + coupons.size());
            for (MerchantCoupon mc : coupons) {
                System.out.println("[SV-DEBUG]   Merchant=\"" + mc.merchantName + "\" Code=\"" + mc.discountCode + "\" Type=\"" + mc.type + "\"");
            }
        }
        return coupons;
    }

    private static List<TargetRow> readTargetTable(File file) throws Exception {
        List<TargetRow> result = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(file);
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet sheet = wb.getSheetAt(0);
            int headerRowIdx = findHeaderRow(sheet);
            if (headerRowIdx < 0) return result;

            Row hRow = sheet.getRow(headerRowIdx);
            // Check for sub-header row
            int dataStartRow = headerRowIdx + 1;
            Row nextRow = sheet.getRow(dataStartRow);
            if (nextRow != null) {
                String firstCellVal = getCellStr(nextRow, 0).toLowerCase();
                // If the next row looks like a sub-header (e.g., "ORIGIN Sale", "SV.LK Sale"), skip it
                boolean isSubHeader = false;
                for (int c = 0; c < nextRow.getLastCellNum(); c++) {
                    String v = getCellStr(nextRow, c).toLowerCase();
                    if (v.contains("origin") || v.contains("sv.lk") || v.contains("total sale")) {
                        isSubHeader = true;
                        break;
                    }
                }
                if (isSubHeader) dataStartRow++;
            }

            int colMerchant = -1, colTarget = -1, colOutlet = -1;
            System.out.println("[SV-DEBUG] === TARGET TABLE HEADERS ===");
            // Scan both header row and row above/below for columns
            Row tRowAbove = headerRowIdx > 0 ? sheet.getRow(headerRowIdx - 1) : null;
            Row tRowBelow = sheet.getRow(headerRowIdx + 1);
            int tLastCol = hRow.getLastCellNum();
            if (tRowAbove != null) tLastCol = Math.max(tLastCol, tRowAbove.getLastCellNum());
            if (tRowBelow != null) tLastCol = Math.max(tLastCol, tRowBelow.getLastCellNum());

            for (int c = 0; c < tLastCol; c++) {
                String above = tRowAbove != null ? getCellStr(tRowAbove, c).trim() : "";
                String main = getCellStr(hRow, c).trim();
                String below = tRowBelow != null ? getCellStr(tRowBelow, c).trim() : "";
                // Clean invisible characters
                String val = main.toLowerCase().replaceAll("[^a-z0-9% /().,]", "").trim();
                String allRows = (above + " " + main + " " + below).toLowerCase().replaceAll("[^a-z0-9% /().,]", "").trim();

                Cell cell = hRow.getCell(c);
                String cellType = cell != null ? cell.getCellType().name() : "NULL";
                System.out.println("[SV-DEBUG]   Col " + c + ": above=\"" + above + "\" main=\"" + main + "\" below=\"" + below + "\" type=" + cellType);

                if (colMerchant < 0 && (val.contains("merchant name") || allRows.contains("merchant name"))) colMerchant = c;
                else if (colTarget < 0 && isExactTarget(val)) colTarget = c;
                else if (colOutlet < 0 && (val.equals("outlet") || allRows.contains("outlet"))) colOutlet = c;
            }
            // Fallback: also check above/below rows for Target column
            if (colTarget < 0) {
                for (int c = 0; c < tLastCol; c++) {
                    String above = tRowAbove != null ? getCellStr(tRowAbove, c).toLowerCase().replaceAll("[^a-z0-9 ]", "").trim() : "";
                    String below = tRowBelow != null ? getCellStr(tRowBelow, c).toLowerCase().replaceAll("[^a-z0-9 ]", "").trim() : "";
                    if (isExactTarget(above) || isExactTarget(below)) { colTarget = c; break; }
                }
            }
            System.out.println("[SV-DEBUG] Target cols: Merchant=" + colMerchant + " Target=" + colTarget + " Outlet=" + colOutlet);
            System.out.println("[SV-DEBUG] Data start row: " + dataStartRow);

            if (colMerchant < 0) return result;

            // Read data rows
            for (int r = dataStartRow; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                String merchant = getCellStr(row, colMerchant).trim();
                if (merchant.isEmpty()) continue;

                // Stop if we hit summary rows (Showroom Total, Grand Total, etc.)
                String merchantLower = merchant.toLowerCase();
                if (merchantLower.contains("total") || merchantLower.contains("grand total")) break;

                TargetRow tr = new TargetRow();
                tr.rowIndex = r;
                tr.merchantName = merchant;
                tr.target = colTarget >= 0 ? getCellNumeric(row, colTarget) : 0;
                tr.outlet = colOutlet >= 0 ? getCellStr(row, colOutlet).trim() : "";

                // Debug: log target cell info
                if (colTarget >= 0) {
                    Cell tCell = row.getCell(colTarget);
                    String tType = tCell != null ? tCell.getCellType().name() : "NULL";
                    String tRaw = tCell != null ? dataFormatter.formatCellValue(tCell) : "null";
                    System.out.println("[SV-DEBUG]   Row " + r + ": merchant=\"" + merchant + "\" target=" + tr.target
                            + " (cellType=" + tType + " raw=\"" + tRaw + "\")");
                }

                result.add(tr);
            }
        }
        return result;
    }

    // ===== Helpers =====

    /** Convert 0-based column index to Excel column letter (0=A, 1=B, ... 25=Z, 26=AA, etc.) */
    private static String colLetter(int col) {
        StringBuilder sb = new StringBuilder();
        col++; // 1-based
        while (col > 0) {
            col--;
            sb.insert(0, (char) ('A' + col % 26));
            col /= 26;
        }
        return sb.toString();
    }

    /** Set an Excel formula on a cell, preserving existing style */
    private static void setCellFormula(Row row, int col, String formula, CellStyle fallbackStyle) {
        Cell cell = row.getCell(col);
        if (cell == null) {
            cell = row.createCell(col);
            if (fallbackStyle != null) cell.setCellStyle(fallbackStyle);
        }
        cell.setCellFormula(formula);
    }

    /** Check if a header string matches "Target" but NOT "March Target", "Per Day Target", etc. */
    private static boolean isExactTarget(String headerLower) {
        if (headerLower == null || headerLower.isEmpty()) return false;
        // Must contain "target"
        if (!headerLower.contains("target")) return false;
        // Must not be a compound header
        if (headerLower.contains("march") || headerLower.contains("per day") ||
                headerLower.contains("forecast") || headerLower.contains("achievement") ||
                headerLower.contains("day")) return false;
        return true;
    }

    private static MerchantSales findSalesForMerchant(String merchantKey, Map<String, MerchantSales> salesMap) {
        // Exact match first
        MerchantSales s = salesMap.get(merchantKey);
        if (s != null) return s;

        // Case-insensitive match
        for (Map.Entry<String, MerchantSales> entry : salesMap.entrySet()) {
            if (entry.getKey().equalsIgnoreCase(merchantKey)) return entry.getValue();
        }
        return null;
    }

    private static String getMerchantType(String merchantName, Map<String, String> typeMap) {
        String type = typeMap.get(merchantName.toLowerCase());
        if (type != null) return type;

        // Try partial match
        for (Map.Entry<String, String> e : typeMap.entrySet()) {
            if (e.getKey().contains(merchantName.toLowerCase()) ||
                    merchantName.toLowerCase().contains(e.getKey())) {
                return e.getValue();
            }
        }
        return "Online"; // default
    }

    private static int findHeaderRow(Sheet sheet) {
        for (int r = 0; r <= Math.min(sheet.getLastRowNum(), 10); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            for (int c = 0; c < row.getLastCellNum(); c++) {
                String val = getCellStr(row, c).toLowerCase();
                if (val.contains("merchant name")) return r;
            }
        }
        return -1;
    }

    private static String getCellStr(Row row, int col) {
        if (row == null) return "";
        Cell cell = row.getCell(col);
        if (cell == null) return "";
        return dataFormatter.formatCellValue(cell);
    }

    private static double getCellNumeric(Row row, int col) {
        if (row == null) return 0;
        Cell cell = row.getCell(col);
        if (cell == null) return 0;

        // Try 1: Direct numeric read
        try {
            if (cell.getCellType() == CellType.NUMERIC) return cell.getNumericCellValue();
        } catch (Exception e) {
            System.out.println("[SV-DEBUG] getCellNumeric NUMERIC failed col=" + col + ": " + e.getMessage());
        }

        // Try 2: Formula cached value
        try {
            if (cell.getCellType() == CellType.FORMULA) return cell.getNumericCellValue();
        } catch (Exception e) {
            // formula cells may not have cached numeric
        }

        // Try 3: Parse formatted string
        try {
            String s = dataFormatter.formatCellValue(cell);
            if (s != null && !s.trim().isEmpty()) {
                String cleaned = s.replaceAll("[^\\d.\\-]", "");
                if (!cleaned.isEmpty()) return Double.parseDouble(cleaned);
            }
        } catch (Exception e) {
            System.out.println("[SV-DEBUG] getCellNumeric STRING parse failed col=" + col + ": " + e.getMessage());
        }

        return 0;
    }

    private static void setCellNum(Row row, int col, double value, CellStyle fallbackStyle) {
        Cell cell = row.getCell(col);
        if (cell == null) {
            cell = row.createCell(col);
            if (fallbackStyle != null) cell.setCellStyle(fallbackStyle);
        }
        // Just set the value — preserve existing cell style (colors, borders, fonts)
        cell.setCellValue(value);
    }

    private static double parseDouble(String s) {
        if (s == null || s.isEmpty()) return 0;
        try {
            return Double.parseDouble(s.replaceAll("[^\\d.\\-]", ""));
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    // Normalize discount/coupon codes: just trim and remove Excel artifacts, preserve original case
    private static String normalizeCode(String s) {
        if (s == null) return "";
        String t = s.trim();
        // Remove common Excel artifacts like leading '=' or surrounding '"'
        t = t.replaceAll("^=+", "");
        t = t.replaceAll("^\"|\"$", "");
        return t.trim();
    }

    // Normalize header keys for flexible matching (remove spaces, underscores, hyphens, lowercase)
    private static String normalizeHeader(String s) {
        if (s == null) return "";
        return s.trim().toLowerCase().replaceAll("[\n\r\t ]+", "").replaceAll("[_\\-]", "");
    }

    /** Find a value in a map row by checking if any key contains given keywords */
    private static String findValue(Map<String, String> row, String... keywords) {
        for (Map.Entry<String, String> e : row.entrySet()) {
            String key = e.getKey().toLowerCase();
            boolean match = true;
            for (String kw : keywords) {
                if (!key.contains(kw.toLowerCase())) { match = false; break; }
            }
            if (match) return e.getValue();
        }
        // Fallback: any key containing at least the first keyword
        if (keywords.length > 0) {
            for (Map.Entry<String, String> e : row.entrySet()) {
                if (e.getKey().toLowerCase().contains(keywords[0].toLowerCase())) return e.getValue();
            }
        }
        return "";
    }

    /** Copy a sheet from source to destination workbook, preserving cell styles and colors */
    private static void copySheet(Sheet src, Sheet dest, XSSFWorkbook destWb) {
        // Cache to avoid creating duplicate styles
        Map<Short, CellStyle> styleCache = new HashMap<>();

        for (int r = 0; r <= src.getLastRowNum(); r++) {
            Row srcRow = src.getRow(r);
            if (srcRow == null) continue;
            Row destRow = dest.getRow(r);
            if (destRow == null) destRow = dest.createRow(r);
            destRow.setHeight(srcRow.getHeight());

            for (int c = 0; c < srcRow.getLastCellNum(); c++) {
                Cell srcCell = srcRow.getCell(c);
                if (srcCell == null) continue;
                Cell destCell = destRow.createCell(c);

                // Copy value
                switch (srcCell.getCellType()) {
                    case NUMERIC:
                        destCell.setCellValue(srcCell.getNumericCellValue());
                        break;
                    case STRING:
                        destCell.setCellValue(srcCell.getStringCellValue());
                        break;
                    case BOOLEAN:
                        destCell.setCellValue(srcCell.getBooleanCellValue());
                        break;
                    case FORMULA:
                        try {
                            destCell.setCellFormula(srcCell.getCellFormula());
                        } catch (Exception e) {
                            // If formula can't be copied, copy cached value instead
                            try { destCell.setCellValue(srcCell.getNumericCellValue()); }
                            catch (Exception e2) { destCell.setCellValue(dataFormatter.formatCellValue(srcCell)); }
                        }
                        break;
                    case BLANK:
                        destCell.setBlank();
                        break;
                    default:
                        destCell.setCellValue(dataFormatter.formatCellValue(srcCell));
                        break;
                }

                // Copy cell style (preserves colors, fonts, borders, number formats)
                try {
                    CellStyle srcStyle = srcCell.getCellStyle();
                    short srcIdx = srcStyle.getIndex();
                    CellStyle destStyle = styleCache.get(srcIdx);
                    if (destStyle == null) {
                        destStyle = destWb.createCellStyle();
                        destStyle.cloneStyleFrom(srcStyle);
                        styleCache.put(srcIdx, destStyle);
                    }
                    destCell.setCellStyle(destStyle);
                } catch (Exception e) {
                    // Style copy failed, continue without style
                }
            }
        }

        // Copy merged regions
        for (int i = 0; i < src.getNumMergedRegions(); i++) {
            dest.addMergedRegion(src.getMergedRegion(i));
        }

        // Copy column widths
        for (int c = 0; c < 30; c++) {
            dest.setColumnWidth(c, src.getColumnWidth(c));
        }
    }
}
