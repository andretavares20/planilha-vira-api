package br.com.planilha_vira_api_back.util;

import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;

public class FormulaEvaluatorProvider {

    public static FormulaEvaluator create(Workbook workbook) {
        return workbook.getCreationHelper().createFormulaEvaluator();
    }
}
