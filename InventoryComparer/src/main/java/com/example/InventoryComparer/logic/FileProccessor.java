// File: FileProccessor.java
package com.example.InventoryComparer.logic;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

public class FileProccessor {

    public static final String OGF_PREFIX = "OGF-";
    public static final String OGF_FILENAME_PATTERN = "ogf";
    private static final List<File> tempFiles = new ArrayList<>();

    public static List<File> processLocationFiles(List<File> originalFiles, boolean ogfActive) {
        if (!ogfActive) {
            return originalFiles;
        }

        // Gather OGF files and non-OGF files
        List<File> ogfFiles = originalFiles.stream()
                .filter(f -> f.getName().toLowerCase().contains(OGF_FILENAME_PATTERN))
                .collect(Collectors.toList());

        List<File> nonOgfFiles = originalFiles.stream()
                .filter(f -> !f.getName().toLowerCase().contains(OGF_FILENAME_PATTERN))
                .collect(Collectors.toList());

        if (ogfFiles.isEmpty()) {
            return originalFiles;
        }

        List<File> finalLocationFiles = new ArrayList<>();

        // First file assumed primary reference file: keep original reference file (don't replace with temp)
        File referenceFile = originalFiles.get(0);
        finalLocationFiles.add(referenceFile);
        // remove it from nonOgfFiles if present
        nonOgfFiles.remove(referenceFile);

        // Only create temp cleaned copies for OGF files (and keep their cleaned version in list)
        for (File ogfFile : ogfFiles) {
            try {
                if (!ogfFile.equals(referenceFile)) {
                    File tempFile = cleanupSkuForPriceComparison(ogfFile);
                    if (tempFile == null) tempFile = ogfFile;
                    finalLocationFiles.add(tempFile);
                    tempFiles.add(tempFile);
                }
            } catch (IOException e) {
                System.err.println("Error processing OGF file for Price Comparer: " + ogfFile.getName() + " - " + e.getMessage());
                finalLocationFiles.add(ogfFile);
            }
        }

        // Add non-OGF files as original (no temp creation)
        finalLocationFiles.addAll(nonOgfFiles);

        return finalLocationFiles.stream().distinct().collect(Collectors.toList());
    }

    public static File cleanupSkuForPriceComparison(File originalFile) throws IOException {
        System.out.println("Processing OGF file for Price Comparer SKU cleanup: " + originalFile.getName());

        File tempFile = File.createTempFile("temp_price_ogf_", ".xlsx");

        try (FileInputStream fis = new FileInputStream(originalFile);
             Workbook workbook = new XSSFWorkbook(fis);
             FileOutputStream fos = new FileOutputStream(tempFile)) {

            DataFormatter formatter = new DataFormatter();
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            int skuColIndex = -1;

            if (headerRow != null) {
                for (Cell cell : headerRow) {
                    if (cell != null && formatter.formatCellValue(cell).trim().equalsIgnoreCase("SKU")) {
                        skuColIndex = cell.getColumnIndex();
                        break;
                    }
                }
            }

            if (skuColIndex == -1) {
                System.err.println("⚠️ SKU column not found in " + originalFile.getName());
                return originalFile;
            }

            // --- Update starts here ---
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell skuCell = row.getCell(skuColIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                String sku = formatter.formatCellValue(skuCell);
                String skuRemark = "";

                if (!sku.isEmpty()) {
                    boolean hasOgf = sku.toUpperCase().startsWith(OGF_PREFIX);
                    if (hasOgf) skuRemark = "OGF- prefix found.";
                    else skuRemark = "WARNING: OGF- prefix missing from SKU.";
                }

                // Write remark BEFORE cleaning
                Row headerRowCheck = sheet.getRow(0);
                int remarkColIndex = -1;
                if (headerRowCheck != null) {
                    for (Cell cell : headerRowCheck) {
                        if (formatter.formatCellValue(cell).trim().equalsIgnoreCase("Remark")) {
                            remarkColIndex = cell.getColumnIndex();
                            break;
                        }
                    }
                    if (remarkColIndex == -1) {
                        remarkColIndex = headerRowCheck.getLastCellNum();
                        Cell remarkHeader = headerRowCheck.createCell(remarkColIndex);
                        remarkHeader.setCellValue("Remark");
                    }
                }

                if (remarkColIndex != -1) {
                    Cell remarkCell = row.getCell(remarkColIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    String currentRemark = formatter.formatCellValue(remarkCell);
                    if (!skuRemark.isEmpty()) {
                        if (currentRemark.isEmpty()) remarkCell.setCellValue(skuRemark);
                        else remarkCell.setCellValue(currentRemark + "; " + skuRemark);
                    }
                }

                // Clean SKU AFTER remark is written
                if (!sku.isEmpty() && sku.toUpperCase().contains(OGF_FILENAME_PATTERN.toUpperCase())) {
                    String cleanedSku = sku.replaceAll("(?i)OGF", "").trim();
                    cleanedSku = cleanedSku.replaceAll("^-|-$", "").trim();
                    skuCell.setCellValue(cleanedSku);
                }

                // Re-write every cell in the row as string using the displayed value (preserve commas/decimals)
                for (int c = 0; c < row.getLastCellNum(); c++) {
                    Cell cell = row.getCell(c, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    String displayed = formatter.formatCellValue(cell);
                    cell.setCellType(CellType.STRING);
                    cell.setCellValue(displayed);
                }
            }
            // --- Update ends here ---

            workbook.write(fos);
            tempFiles.add(tempFile);
            return tempFile;

        } catch (Exception e) {
            System.err.println("Error cleaning SKU in OGF file: " + e.getMessage());
            e.printStackTrace();
            return originalFile;
        }
    }

    // --- Shared Utility Methods ---
    private static String getStringValue(Cell cell) {
        if (cell == null) return "";
        try {
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue().trim();
                case NUMERIC:
                    double val = cell.getNumericCellValue();
                    return (val == Math.floor(val)) ? String.valueOf((long) val) : String.valueOf(val);
                case FORMULA:
                    CellType cachedType = cell.getCachedFormulaResultType();
                    if (cachedType == CellType.STRING) return cell.getStringCellValue().trim();
                    if (cachedType == CellType.NUMERIC) {
                        double valFormula = cell.getNumericCellValue();
                        return (valFormula == Math.floor(valFormula)) ? String.valueOf((long) valFormula) : String.valueOf(valFormula);
                    }
                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());
                case BLANK:
                default:
                    return "";
            }
        } catch (Exception e) {
            return "";
        }
    }

    public static void cleanUpTempFiles(List<File> filesToClean) {
        List<File> successfullyCleaned = new ArrayList<>();
        for (File file : filesToClean) {
            if (file.getName().startsWith("temp_") && tempFiles.contains(file)) {
                if (file.delete()) {
                    System.out.println("Cleaned up temporary file: " + file.getName());
                    successfullyCleaned.add(file);
                } else {
                    System.err.println("Could not delete temporary file: " + file.getName());
                    file.deleteOnExit();
                }
            }
        }
        tempFiles.removeAll(successfullyCleaned);
    }
}
