package br.com.planilha_vira_api_back.service;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
public class ExcelService {
    
    private final ExcelParser excelParser;

    public ExcelService(ExcelParser excelParser) {
        this.excelParser = excelParser;
    }

    public Map<String, List<Map<String, Object>>> converterExcelParaJson(MultipartFile file) throws IOException {
        return excelParser.analisarExcel(file);
    }

}
