package com.example.InventoryComparer.controller;

import com.example.InventoryComparer.logic.POStockTallyLogic;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;

@CrossOrigin(origins = "http://localhost:3000")
@RestController
@RequestMapping("/api/po-stock")
public class POStockTallyController {

    @PostMapping("/generate")
    public ResponseEntity<byte[]> generateReport(
            @RequestParam("purchaseOrderFiles") List<MultipartFile> purchaseOrderFiles,
            @RequestParam("stockAdjustmentFiles") List<MultipartFile> stockAdjustmentFiles,
            @RequestParam(value = "excludeSAIds", required = false) List<String> excludeSAIds // NEW PARAMETER
    ) {
        System.out.println("=== PO-STOCK CONTROLLER STARTED ===");
        System.out.println("Received " + purchaseOrderFiles.size() + " purchase order files");
        System.out.println("Received " + stockAdjustmentFiles.size() + " stock adjustment files");

        // Debug file names
        for (MultipartFile file : purchaseOrderFiles) {
            System.out.println("PO File: " + file.getOriginalFilename() + " | Size: " + file.getSize() + " | Content Type: " + file.getContentType());
        }
        for (MultipartFile file : stockAdjustmentFiles) {
            System.out.println("Stock File: " + file.getOriginalFilename() + " | Size: " + file.getSize() + " | Content Type: " + file.getContentType());
        }

        if (excludeSAIds != null) {
            System.out.println("Exclude SA IDs: " + excludeSAIds);
        } else {
            excludeSAIds = new ArrayList<>();
            System.out.println("No SA IDs to exclude");
        }

        List<File> poTempFiles = new ArrayList<>();
        List<File> stockTempFiles = new ArrayList<>();
        File outputFile = null;

        try {
            // Convert uploaded MultipartFiles to temporary Files
            System.out.println("Converting multipart files to temporary files...");
            for (MultipartFile mf : purchaseOrderFiles) {
                File tempFile = convertMultipartToFile(mf);
                poTempFiles.add(tempFile);
                System.out.println("Created temp PO file: " + tempFile.getAbsolutePath() + " | Exists: " + tempFile.exists() + " | Size: " + tempFile.length());
            }

            for (MultipartFile mf : stockAdjustmentFiles) {
                File tempFile = convertMultipartToFile(mf);
                stockTempFiles.add(tempFile);
                System.out.println("Created temp Stock file: " + tempFile.getAbsolutePath() + " | Exists: " + tempFile.exists() + " | Size: " + tempFile.length());
            }

            // Temporary file for output
            outputFile = File.createTempFile("PO_Stock_Tally_Report_", ".xlsx");
            System.out.println("Output file: " + outputFile.getAbsolutePath());

            // Call backend logic with excludeSAIds
            System.out.println("Calling POStockTallyLogic.generateReport...");
            POStockTallyLogic.generateReport(poTempFiles, stockTempFiles, outputFile, excludeSAIds);

            // Check if output file was created
            if (!outputFile.exists() || outputFile.length() == 0) {
                System.out.println("ERROR: Output file was not created or is empty!");
                return ResponseEntity.status(500)
                        .body(("Error: Report generation failed - output file is empty").getBytes());
            }

            // Read the generated file into bytes
            byte[] fileContent = Files.readAllBytes(outputFile.toPath());
            System.out.println("Output file size: " + fileContent.length + " bytes");

            // Prepare response headers for download
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            headers.setContentDispositionFormData("attachment", "PO_Stock_Tally_Report.xlsx");

            System.out.println("=== CONTROLLER COMPLETED SUCCESSFULLY ===");
            return ResponseEntity.ok()
                    .headers(headers)
                    .body(fileContent);

        } catch (Exception e) {
            System.out.println("ERROR in controller: " + e.getMessage());
            e.printStackTrace();
            return ResponseEntity.status(500)
                    .body(("Error generating report: " + e.getMessage()).getBytes());
        } finally {
            // Cleanup all temporary files
            System.out.println("Cleaning up temporary files...");
            cleanupTempFiles(poTempFiles);
            cleanupTempFiles(stockTempFiles);
            if (outputFile != null && outputFile.exists()) {
                System.out.println("Deleting output file: " + outputFile.getAbsolutePath());
                outputFile.delete();
            }
        }
    }

    private File convertMultipartToFile(MultipartFile multipart) throws IOException {
        String originalFilename = multipart.getOriginalFilename();
        File convFile = new File(System.getProperty("java.io.tmpdir"), originalFilename);
        System.out.println("Converting multipart to file: " + originalFilename + " -> " + convFile.getAbsolutePath());
        try (FileOutputStream fos = new FileOutputStream(convFile)) {
            fos.write(multipart.getBytes());
        }
        System.out.println("File conversion complete. Size: " + convFile.length() + " bytes");
        return convFile;
    }

    private void cleanupTempFiles(List<File> files) {
        for (File f : files) {
            if (f != null && f.exists()) {
                try {
                    boolean deleted = f.delete();
                    System.out.println("Cleanup: " + f.getName() + " deleted: " + deleted);
                } catch (Exception e) {
                    System.out.println("Cleanup failed for: " + f.getName() + " - " + e.getMessage());
                }
            }
        }
    }
}
