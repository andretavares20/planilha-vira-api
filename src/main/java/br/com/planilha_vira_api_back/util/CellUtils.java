package br.com.planilha_vira_api_back.util;

import org.apache.poi.ss.usermodel.Cell;

public class CellUtils {
    public static Object obterValorCelula(Cell cell) {
        if (cell == null) return null;
        
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> (cell.getNumericCellValue() % 1 == 0) ? 
                            (int) cell.getNumericCellValue() : 
                            cell.getNumericCellValue();
            case BOOLEAN -> cell.getBooleanCellValue();
            default -> null;
        };
    }
}
