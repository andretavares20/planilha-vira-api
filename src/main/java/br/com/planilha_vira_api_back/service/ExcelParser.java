package br.com.planilha_vira_api_back.service;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import br.com.planilha_vira_api_back.util.CellUtils;

@Service
public class ExcelParser {
    public Map<String, List<Map<String, Object>>> analisarExcel(MultipartFile file) throws IOException {
        Map<String, List<Map<String, Object>>> data = new HashMap<>();
        
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            for (Sheet sheet : workbook) {
                data.put(sheet.getSheetName(), extrairDadosPlanilha(sheet));
            }
        }
        
        return data;
    }

    private List<Map<String, Object>> extrairDadosPlanilha(Sheet sheet) {
        List<Map<String, Object>> sheetData = new ArrayList<>();
        List<String> headers = obterCabecalhos(sheet.getRow(0));

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            sheetData.add(extrairDadosLinha(row, headers));
        }

        return sheetData;
    }

    private List<String> obterCabecalhos(Row headerRow) {
        List<String> headers = new ArrayList<>();
        for (Cell cell : headerRow) {
            headers.add(cell.getStringCellValue().trim());
        }
        return headers;
    }

    private Map<String, Object> extrairDadosLinha(Row row, List<String> headers) {
        Map<String, Object> rowData = new HashMap<>();
        for (int j = 0; j < headers.size(); j++) {
            rowData.put(headers.get(j), CellUtils.obterValorCelula(row.getCell(j)));
        }
        return rowData;
    }
}
