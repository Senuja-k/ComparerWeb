package com.example.InventoryComparer.controller;

import com.example.InventoryComparer.logic.FileProccessor;
import com.example.InventoryComparer.logic.SKUComparerLogic;
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
@RequestMapping("/api/comparer")
public class SKUComparerController {

    @PostMapping("/generate")
    public ResponseEntity<byte[]> generateReport(
            @RequestParam("locationFiles") List<MultipartFile> locationFiles,
            @RequestParam(value = "unlistedFiles", required = false) List<MultipartFile> unlistedFiles,
            @RequestParam(value = "ogfRulesChecked", required = false, defaultValue = "false") boolean ogfRulesChecked
    ) {
        List<File> locationTempFiles = new ArrayList<>();
        List<File> processedLocationFiles = new ArrayList<>();
        List<File> unlistedTempFiles = new ArrayList<>();
        File outputFile = null;

        try {
            // ✅ Convert uploaded MultipartFiles to temporary Files (preserving original names)
            for (MultipartFile mf : locationFiles) {
                locationTempFiles.add(convertMultipartToFile(mf));
            }

            // ✅ ADD NULL CHECK HERE - This is the critical fix!
            if (unlistedFiles != null) {
                for (MultipartFile mf : unlistedFiles) {
                    unlistedTempFiles.add(convertMultipartToFile(mf));
                }
            }

            // ✅ Apply OGF logic preprocessing here (controller level, before backend)
            processedLocationFiles = FileProccessor.processLocationFiles(locationTempFiles, ogfRulesChecked);

            // ✅ Temporary file for output
            outputFile = File.createTempFile("Inventory_Comparison_Report_", ".xlsx");

            // ✅ Call backend logic (unchanged signature)
            SKUComparerLogic.generateReport(processedLocationFiles, unlistedTempFiles, outputFile, ogfRulesChecked);

            // ✅ Read the generated file into bytes
            byte[] fileContent = Files.readAllBytes(outputFile.toPath());

            // ✅ Prepare response headers for download
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            headers.setContentDispositionFormData("attachment", "Inventory_Comparison_Report.xlsx");

            return ResponseEntity.ok()
                    .headers(headers)
                    .body(fileContent);

        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.status(500)
                    .body(("Error generating report: " + e.getMessage()).getBytes());
        } finally {
            // ✅ Cleanup all temporary files
            cleanupTempFiles(locationTempFiles);
            cleanupTempFiles(processedLocationFiles);
            cleanupTempFiles(unlistedTempFiles);
            if (outputFile != null && outputFile.exists()) outputFile.delete();
        }
    }

    private File convertMultipartToFile(MultipartFile multipart) throws IOException {
        // Create a temp file with the *original filename*
        File convFile = new File(System.getProperty("java.io.tmpdir"), multipart.getOriginalFilename());
        try (FileOutputStream fos = new FileOutputStream(convFile)) {
            fos.write(multipart.getBytes());
        }
        return convFile;
    }

    private void cleanupTempFiles(List<File> files) {
        for (File f : files) {
            if (f != null && f.exists()) {
                try {
                    f.delete();
                } catch (Exception ignored) {}
            }
        }
    }
}