package com.example.excelhtml;

import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * groovy-excel2csv の MyPOIUtil.getCellValue / MyExcelUtils 相当。
 * 取消線付き文字の除去、日付固定書式、数値の DecimalFormat 出力に対応する。
 */
final class TxtCellValueFormatter {

    private static final String DATE_PATTERN = "yyyy/MM/dd";

    private TxtCellValueFormatter() {
    }

    static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        return switch (cell.getCellType()) {
            case NUMERIC -> formatNumeric(cell);
            case STRING -> removeStrikeString(cell);
            case FORMULA -> getStringFormulaValue(cell);
            case BOOLEAN -> Boolean.toString(cell.getBooleanCellValue());
            default -> "";
        };
    }

    private static String formatNumeric(Cell cell) {
        if (DateUtil.isCellDateFormatted(cell)) {
            Date dateValue = cell.getDateCellValue();
            DateFormat dateFormat = new SimpleDateFormat(DATE_PATTERN);
            return dateFormat.format(dateValue);
        }
        DecimalFormat df = new DecimalFormat();
        df.setMinimumFractionDigits(0);
        return df.format(cell.getNumericCellValue());
    }

    private static String getStringFormulaValue(Cell cell) {
        String sheetName = cell.getSheet().getSheetName();
        CellAddress cellAddress = cell.getAddress();
        try {
            return switch (cell.getCachedFormulaResultType()) {
                case NUMERIC -> formatNumeric(cell);
                case STRING -> cell.getStringCellValue();
                case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
                default -> {
                    System.err.println(sheetName + ":" + cellAddress + ":" + cell.getCachedFormulaResultType());
                    yield "";
                }
            };
        } catch (Exception e) {
            Throwable cause = e.getCause();
            String errMsg = cause != null ? cause.toString() : e.getMessage();
            System.err.println(sheetName + ":" + cellAddress + ":" + errMsg);
            return "";
        }
    }

    /** セル文字列から取消線付き部分を除いた文字列を返す。 */
    static String removeStrikeString(Cell cell) {
        CellStyle style = cell.getCellStyle();
        Workbook book = cell.getSheet().getWorkbook();
        Font cellFont = book.getFontAt(style.getFontIndex());
        if (cellFont.getStrikeout()) {
            return "";
        }

        RichTextString richStr = cell.getRichStringCellValue();
        if (richStr instanceof XSSFRichTextString xsrichStr) {
            return removeStrikeXssf(xsrichStr);
        }
        if (richStr instanceof HSSFRichTextString hsrichStr) {
            return removeStrikeHssf(book, hsrichStr);
        }
        return cell.getStringCellValue();
    }

    private static String removeStrikeXssf(XSSFRichTextString xsrichStr) {
        StringBuilder sb = new StringBuilder();
        int cnt = xsrichStr.numFormattingRuns();
        if (cnt == 0) {
            sb.append(xsrichStr.getString());
            return sb.toString();
        }
        for (int i = 0; i < cnt; i++) {
            String text = xsrichStr.getCTRst().getRArray(i).getT();
            XSSFFont xssfFont = xsrichStr.getFontOfFormattingRun(i);
            if (xssfFont == null || !xssfFont.getStrikeout()) {
                sb.append(text);
            }
        }
        return sb.toString();
    }

    static String removeStrikeHssf(Workbook book, HSSFRichTextString hsrichStr) {
        String orgstr = hsrichStr.getString();
        StringBuilder sb = new StringBuilder();
        int cnt = hsrichStr.numFormattingRuns();
        if (cnt == 0) {
            sb.append(orgstr);
            return sb.toString();
        }
        int start = 0;
        boolean addFlag = true;
        for (int i = 0; i < cnt; i++) {
            int end = hsrichStr.getIndexOfFormattingRun(i);
            String tmpStr = orgstr.substring(start, end);
            if (tmpStr.replace("\r", "").replace("\n", "").isEmpty()) {
                addFlag = true;
            }
            if (addFlag) {
                sb.append(tmpStr);
            }
            start = end;
            int fontIdx = hsrichStr.getFontOfFormattingRun(i);
            Font hssfFont = book.getFontAt(fontIdx);
            addFlag = hssfFont == null || !hssfFont.getStrikeout();
        }
        if (addFlag) {
            sb.append(orgstr.substring(start));
        }
        return sb.toString();
    }

    static String nullToEmpty(String value) {
        return value == null ? "" : value;
    }

    static String stripNewlines(String value) {
        return value.replace("\r\n", "").replace("\r", "").replace("\n", "");
    }

    /** 末尾の空白・タブを除去（commons-lang3 StringUtils.stripEnd 相当）。 */
    static String stripEnd(String value) {
        if (value == null || value.isEmpty()) {
            return value == null ? "" : value;
        }
        int end = value.length();
        while (end > 0 && Character.isWhitespace(value.charAt(end - 1))) {
            end--;
        }
        return end == value.length() ? value : value.substring(0, end);
    }

    static String strip(String value) {
        if (value == null) {
            return "";
        }
        return value.strip();
    }

    static boolean isBlank(String value) {
        return value == null || value.isBlank();
    }
}
