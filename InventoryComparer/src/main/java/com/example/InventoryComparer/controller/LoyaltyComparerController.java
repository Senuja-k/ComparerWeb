package com.example.InventoryComparer.controller;

import com.example.InventoryComparer.logic.LoyaltyComparerLogic;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

@CrossOrigin(origins = "http://localhost:3000")
@RestController
@RequestMapping("/loyalty")
public class LoyaltyComparerController {

    @PostMapping("/generateLoyalty")
    public ResponseEntity<InputStreamResource> generateLoyaltyReport(
            @RequestParam("referenceFile") MultipartFile referenceFile,
            @RequestParam("locationsFiles") MultipartFile[] locationFiles) { // <--- FIXED

        // You handle the 'required' check inside the method body.
        if (referenceFile.isEmpty() || locationFiles == null || locationFiles.length == 0) {
            return ResponseEntity.badRequest().build();
        }

        // Convert reference file
        File refFile;
        try {
            refFile = convertMultipartFileToFile(referenceFile);
        } catch (IOException e) {
            return ResponseEntity.internalServerError().build();
        }

        // Convert location files
        List<File> locFiles = new ArrayList<>();
        for (MultipartFile mf : locationFiles) {
            try {
                locFiles.add(convertMultipartFileToFile(mf));
            } catch (IOException e) {
                return ResponseEntity.internalServerError().build();
            }
        }

        // Output file
        File outputFile = new File(System.getProperty("java.io.tmpdir") + "/Loyalty_Comparison_Report.xlsx");

        try {
            LoyaltyComparerLogic.generateReport(refFile, locFiles, outputFile);
        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.internalServerError().build();
        }

        try {
            InputStreamResource resource = new InputStreamResource(new FileInputStream(outputFile));

            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Loyalty_Comparison_Report.xlsx");

            return ResponseEntity.ok()
                    .headers(headers)
                    .contentLength(outputFile.length())
                    .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                    .body(resource);

        } catch (FileNotFoundException e) {
            return ResponseEntity.internalServerError().build();
        }
    }

    private File convertMultipartFileToFile(MultipartFile multipartFile) throws IOException {
        File convFile = new File(System.getProperty("java.io.tmpdir") + "/" + multipartFile.getOriginalFilename());
        try (FileOutputStream fos = new FileOutputStream(convFile)) {
            fos.write(multipartFile.getBytes());
        }
        return convFile;
    }
}
