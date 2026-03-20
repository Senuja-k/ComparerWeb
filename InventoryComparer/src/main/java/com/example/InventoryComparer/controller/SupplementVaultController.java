package com.example.InventoryComparer.controller;

import com.example.InventoryComparer.logic.SupplementVaultLogic;
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
@RequestMapping("/api/supplement-vault")
public class SupplementVaultController {

    @PostMapping("/generate")
    public ResponseEntity<byte[]> generateReport(
            @RequestParam("orderFiles") List<MultipartFile> orderFiles,
            @RequestParam("couponFile") MultipartFile couponFile,
            @RequestParam("targetFile") MultipartFile targetFile,
            @RequestParam("daysRemainingOnline") int daysRemainingOnline,
            @RequestParam("daysRemainingOutlet") int daysRemainingOutlet,
            @RequestParam("totalDays") int totalDays,
            @RequestParam("reportDay") int reportDay
    ) {
        List<File> orderTempFiles = new ArrayList<>();
        File couponTempFile = null;
        File targetTempFile = null;
        File outputFile = null;

        try {
            for (MultipartFile mf : orderFiles) {
                orderTempFiles.add(convertMultipartToFile(mf));
            }
            couponTempFile = convertMultipartToFile(couponFile);
            targetTempFile = convertMultipartToFile(targetFile);

            outputFile = File.createTempFile("SupplementVault_Sales_Report_", ".xlsx");

            SupplementVaultLogic.generateReport(
                    orderTempFiles, couponTempFile, targetTempFile,
                    daysRemainingOnline, daysRemainingOutlet,
                    totalDays, reportDay, outputFile
            );

            byte[] fileContent = Files.readAllBytes(outputFile.toPath());

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            headers.setContentDispositionFormData("attachment", "SupplementVault_Sales_Report.xlsx");

            return ResponseEntity.ok().headers(headers).body(fileContent);

        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.status(500)
                    .body(("Error generating report: " + e.getMessage()).getBytes());
        } finally {
            cleanupTempFiles(orderTempFiles);
            if (couponTempFile != null && couponTempFile.exists()) couponTempFile.delete();
            if (targetTempFile != null && targetTempFile.exists()) targetTempFile.delete();
            if (outputFile != null && outputFile.exists()) outputFile.delete();
        }
    }

    private File convertMultipartToFile(MultipartFile multipart) throws IOException {
        File convFile = File.createTempFile("sv_upload_", "_" + multipart.getOriginalFilename());
        try (FileOutputStream fos = new FileOutputStream(convFile)) {
            fos.write(multipart.getBytes());
        }
        return convFile;
    }

    private void cleanupTempFiles(List<File> files) {
        for (File f : files) {
            if (f != null && f.exists()) {
                try { f.delete(); } catch (Exception ignored) {}
            }
        }
    }
}
