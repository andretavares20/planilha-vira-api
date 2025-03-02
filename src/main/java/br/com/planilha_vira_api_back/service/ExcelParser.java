package br.com.planilha_vira_api_back.service;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
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

        try (Workbook workbook = carregarWorkbook(file)) {
            for (Sheet sheet : workbook) {
                data.put(sheet.getSheetName(), extrairDadosPlanilha(sheet));
            }
        }

        return data;
    }

    private Workbook carregarWorkbook(MultipartFile file) throws IOException {
        String fileName = file.getOriginalFilename().toLowerCase();

        if (fileName.endsWith(".xlsx")) {
            return new XSSFWorkbook(file.getInputStream());
        } else if (fileName.endsWith(".xls")) {
            return new HSSFWorkbook(file.getInputStream());
        } else {
            throw new IllegalArgumentException("Formato de arquivo não suportado: " + fileName);
        }
    }

    private List<Map<String, Object>> extrairDadosPlanilha(Sheet sheet) {
        List<Map<String, Object>> sheetData = new ArrayList<>();

        int linhaCabecalho = detectarLinhaCabecalho(sheet);
        Row headerRow = sheet.getRow(linhaCabecalho);
        List<String> headers = obterCabecalhos(headerRow);

        for (int i = linhaCabecalho + 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null)
                continue;

            sheetData.add(extrairDadosLinha(row, headers));
        }

        return sheetData;
    }

    private int detectarLinhaCabecalho(Sheet sheet) {
        int linhaCabecalho = -1;
        int maxColunasPreenchidas = 0;

        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (isLinhaVazia(row)) continue;

            int celulasPreenchidas = contarCelulasPreenchidas(row);

            if (celulasPreenchidas > maxColunasPreenchidas) {
                maxColunasPreenchidas = celulasPreenchidas;
                linhaCabecalho = i;
            }
        }

        validarLinhaCabecalho(linhaCabecalho);
        return linhaCabecalho;
    }

    private boolean isLinhaVazia(Row row) {
        return row == null || row.getPhysicalNumberOfCells() == 0;
    }

    private int contarCelulasPreenchidas(Row row) {
        int count = 0;
        for (Cell cell : row) {
            if (isCelulaPreenchida(cell)) {
                count++;
            }
        }
        return count;
    }

    private boolean isCelulaPreenchida(Cell cell) {
        return cell != null && cell.getCellType() != CellType.BLANK;
    }

    private void validarLinhaCabecalho(int linhaCabecalho) {
        if (linhaCabecalho == -1) {
            throw new IllegalArgumentException("Nenhuma linha válida encontrada para cabeçalhos.");
        }
    }

    private List<String> obterCabecalhos(Row headerRow) {
        List<String> headers = new ArrayList<>();
        if (headerRow == null) {
            throw new IllegalArgumentException("A linha de cabeçalho está vazia.");
        }

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
