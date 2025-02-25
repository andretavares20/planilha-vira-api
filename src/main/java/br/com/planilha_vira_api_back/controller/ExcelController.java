package br.com.planilha_vira_api_back.controller;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import br.com.planilha_vira_api_back.service.ExcelService;

@RestController
@RequestMapping("/api/v1/excel")
public class ExcelController {
    private final ExcelService excelService;

    public ExcelController(ExcelService excelService) {
        this.excelService = excelService;
    }

    @PostMapping("/upload")
    public ResponseEntity<Map<String, List<Map<String, Object>>>> uploadExcel(@RequestParam("file") MultipartFile file) {
        try {
            return ResponseEntity.ok(excelService.converterExcelParaJson(file));
        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }
}
