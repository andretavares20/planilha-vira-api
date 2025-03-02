package br.com.planilha_vira_api_back.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

public class CellUtils {

    public static Object obterValorCelula(Cell cell) {
        if (cell == null) {
            return null;
        }

        switch (cell.getCellType()) {
            case STRING:
                return getStringValue(cell);
            case NUMERIC:
                return getNumericValue(cell);
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                return getFormulaValue(cell);
            case BLANK:
                return "";
            default:
                return null;
        }
    }

    private static String getStringValue(Cell cell) {
        return cell.getStringCellValue().trim();
    }

    private static Object getNumericValue(Cell cell) {
        return DateUtil.isCellDateFormatted(cell) ? cell.getDateCellValue() : cell.getNumericCellValue();
    }

    private static Object getFormulaValue(Cell cell) {
        FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
        return evaluateFormula(evaluator, cell);
    }

    private static Object evaluateFormula(FormulaEvaluator evaluator, Cell cell) {
        CellValue cellValue = evaluator.evaluate(cell);

        switch (cellValue.getCellType()) {
            case STRING:
                return cellValue.getStringValue();
            case NUMERIC:
                return cellValue.getNumberValue();
            case BOOLEAN:
                return cellValue.getBooleanValue();
            default:
                return null;
        }
    }
}
