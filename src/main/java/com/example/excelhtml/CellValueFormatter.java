package com.example.excelhtml;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

/**
 * Excel セルの表示値を文字列として取得する。
 * 書式（日付・数値フォーマットなど）は DataFormatter で反映する。
 */
public final class CellValueFormatter {

    private final DataFormatter dataFormatter;
    private final FormulaEvaluator formulaEvaluator;

    public CellValueFormatter(FormulaEvaluator formulaEvaluator) {
        this.dataFormatter = new DataFormatter();
        this.formulaEvaluator = formulaEvaluator;
    }

    public String format(Cell cell) {
        if (cell == null) {
            return "";
        }

        CellType type = cell.getCellType();
        if (type == CellType.FORMULA) {
            // 計算結果の表示値を返す（数式文字列ではない）
            return dataFormatter.formatCellValue(cell, formulaEvaluator);
        }
        return dataFormatter.formatCellValue(cell);
    }
}
