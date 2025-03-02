package br.com.planilha_vira_api_back.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelCellReader {

    private final FormulaEvaluator formulaEvaluator;

    public ExcelCellReader(Workbook workbook) {
        this.formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
    }

    public String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }

        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> getFormulaValue(cell);
            default -> "";
        };
    }

    private String getFormulaValue(Cell cell) {
        CellValue cellValue = formulaEvaluator.evaluate(cell);

        return switch (cellValue.getCellType()) {
            case STRING -> cellValue.getStringValue();
            case NUMERIC -> String.valueOf(cellValue.getNumberValue());
            case BOOLEAN -> String.valueOf(cellValue.getBooleanValue());
            default -> "";
        };
    }
}
