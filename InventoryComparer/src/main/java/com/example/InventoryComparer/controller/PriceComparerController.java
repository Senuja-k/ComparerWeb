package com.example.InventoryComparer.controller;

import com.example.InventoryComparer.logic.PriceComparerLogic;
import com.example.InventoryComparer.logic.FileProccessor;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/price")
public class PriceComparerController {

    @PostMapping("/generatePrice")
    public ResponseEntity<byte[]> generatePriceReport(
            @RequestParam("referenceFile") MultipartFile referenceFile,
            @RequestParam("locationFiles") MultipartFile[] locationFiles
    ) {
        // Map to track original file names for all files
        Map<File, String> originalFileNames = new HashMap<>();

        try {
            // Convert reference file to File and store original name
            File tempRefFile = convertToFile(referenceFile);
            originalFileNames.put(tempRefFile, referenceFile.getOriginalFilename());

            // Convert location files to List<File> and store original names
            List<File> tempLocationFiles = new ArrayList<>();
            for (MultipartFile mf : locationFiles) {
                File tempFile = convertToFile(mf);
                tempLocationFiles.add(tempFile);
                originalFileNames.put(tempFile, mf.getOriginalFilename());
            }

            // ✅ INTEGRATE FILEPROCESSOR HERE
            List<File> allFiles = new ArrayList<>();
            allFiles.add(tempRefFile);
            allFiles.addAll(tempLocationFiles);

            List<File> processedFiles = FileProccessor.processLocationFiles(allFiles, true);

            // Update original file names for processed files
            // The first file is the reference file, rest are location files
            if (processedFiles.size() == allFiles.size()) {
                for (int i = 0; i < processedFiles.size(); i++) {
                    File originalFile = allFiles.get(i);
                    File processedFile = processedFiles.get(i);
                    String originalName = originalFileNames.get(originalFile);
                    originalFileNames.put(processedFile, originalName);
                }
            }

            // Extract processed reference and location files
            File processedRefFile = processedFiles.get(0);
            List<File> processedLocFiles = processedFiles.subList(1, processedFiles.size());

            // Output report file
            File outputFile = File.createTempFile("Price_Report", ".xlsx");
            outputFile.deleteOnExit();

            // ✅ Call logic with PROCESSED files and original file names
            PriceComparerLogic.generateReport(processedRefFile, processedLocFiles, outputFile, originalFileNames);

            // Read the generated file
            byte[] fileContent = java.nio.file.Files.readAllBytes(outputFile.toPath());

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            headers.setContentDispositionFormData("attachment", "Price_Report.xlsx");

            return ResponseEntity.ok()
                    .headers(headers)
                    .body(fileContent);

        } catch (IOException e) {
            e.printStackTrace();
            return ResponseEntity.status(500).body(null);
        } finally {
            // Clean up temporary files
            cleanupTempFiles(originalFileNames);
        }
    }

    // Helper method to convert MultipartFile to File
    private File convertToFile(MultipartFile file) throws IOException {
        File convFile = File.createTempFile("upload_" + System.currentTimeMillis() + "_", ".tmp");
        convFile.deleteOnExit();
        try (FileOutputStream fos = new FileOutputStream(convFile)) {
            fos.write(file.getBytes());
        }
        return convFile;
    }

    // Helper method to clean up temporary files
    private void cleanupTempFiles(Map<File, String> originalFileNames) {
        for (File file : originalFileNames.keySet()) {
            if (file.exists()) {
                file.delete();
            }
        }
    }
}