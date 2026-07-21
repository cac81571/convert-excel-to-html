package com.example.excelhtml;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Calendar;

/** 動作確認用のサンプル Excel を生成する。 */
public final class CreateSampleExcel {

    public static void main(String[] args) throws Exception {
        Path out = Path.of(args.length > 0 ? args[0] : "sample.xlsx");

        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("サンプル");

            CellStyle headerStyle = wb.createCellStyle();
            Font headerFont = wb.createFont();
            headerFont.setBold(true);
            headerFont.setColor(IndexedColors.WHITE.getIndex());
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            headerStyle.setBorderBottom(BorderStyle.THIN);
            headerStyle.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());

            CellStyle moneyStyle = wb.createCellStyle();
            moneyStyle.setDataFormat(wb.createDataFormat().getFormat("#,##0"));
            moneyStyle.setAlignment(HorizontalAlignment.RIGHT);

            CellStyle dateStyle = wb.createCellStyle();
            dateStyle.setDataFormat(wb.createDataFormat().getFormat("yyyy-mm-dd"));
            dateStyle.setAlignment(HorizontalAlignment.CENTER);

            CellStyle alertStyle = wb.createCellStyle();
            Font alertFont = wb.createFont();
            alertFont.setBold(true);
            alertFont.setColor(IndexedColors.DARK_RED.getIndex());
            alertStyle.setFont(alertFont);
            alertStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
            alertStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            XSSFCellStyle mergeStyle = (XSSFCellStyle) wb.createCellStyle();
            Font mergeFont = wb.createFont();
            mergeFont.setItalic(true);
            mergeStyle.setFont(mergeFont);
            mergeStyle.setAlignment(HorizontalAlignment.CENTER);
            mergeStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            mergeStyle.setFillForegroundColor(new XSSFColor(new byte[] {(byte) 220, (byte) 230, (byte) 242}, null));
            mergeStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            Row header = sheet.createRow(0);
            String[] titles = {"名前", "数量", "単価", "合計", "日付", "備考"};
            for (int i = 0; i < titles.length; i++) {
                Cell cell = header.createCell(i);
                cell.setCellValue(titles[i]);
                cell.setCellStyle(headerStyle);
            }

            Row row1 = sheet.createRow(1);
            row1.createCell(0).setCellValue("りんご");
            row1.createCell(1).setCellValue(3);
            Cell price1 = row1.createCell(2);
            price1.setCellValue(120);
            price1.setCellStyle(moneyStyle);
            Cell total1 = row1.createCell(3);
            total1.setCellFormula("B2*C2");
            total1.setCellStyle(moneyStyle);
            Cell dateCell = row1.createCell(4);
            Calendar cal = Calendar.getInstance();
            cal.set(2026, Calendar.JULY, 21, 0, 0, 0);
            dateCell.setCellValue(cal);
            dateCell.setCellStyle(dateStyle);
            Cell remark = row1.createCell(5);
            remark.setCellValue("A & B <特価>");
            remark.setCellStyle(alertStyle);

            Row row2 = sheet.createRow(2);
            row2.createCell(0).setCellValue("みかん");
            row2.createCell(1).setCellValue(5);
            Cell price2 = row2.createCell(2);
            price2.setCellValue(80);
            price2.setCellStyle(moneyStyle);
            Cell total2 = row2.createCell(3);
            total2.setCellFormula("B3*C3");
            total2.setCellStyle(moneyStyle);
            row2.createCell(5).setCellValue(true);

            sheet.setColumnWidth(0, 12 * 256);
            sheet.setColumnWidth(1, 8 * 256);
            sheet.setColumnWidth(2, 8 * 256);
            sheet.setColumnWidth(3, 10 * 256);
            sheet.setColumnWidth(4, 14 * 256);
            sheet.setColumnWidth(5, 22 * 256);

            Row mergeRow = sheet.createRow(3);
            Cell merged = mergeRow.createCell(0);
            merged.setCellValue("結合セル（A4:C4）");
            merged.setCellStyle(mergeStyle);
            sheet.addMergedRegion(new CellRangeAddress(3, 3, 0, 2));
            Cell vertical = mergeRow.createCell(3);
            vertical.setCellValue("縦結合");
            vertical.setCellStyle(mergeStyle);
            sheet.addMergedRegion(new CellRangeAddress(3, 4, 3, 3));
            Row mergeRow2 = sheet.createRow(4);
            mergeRow2.createCell(0).setCellValue("通常");
            mergeRow2.createCell(1).setCellValue("セル");
            mergeRow2.createCell(2).setCellValue("です");

            Sheet sheet2 = wb.createSheet("メモ");
            CellStyle titleStyle = wb.createCellStyle();
            Font titleFont = wb.createFont();
            titleFont.setBold(true);
            titleFont.setFontHeightInPoints((short) 14);
            titleStyle.setFont(titleFont);
            titleStyle.setAlignment(HorizontalAlignment.CENTER);
            titleStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
            titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            titleStyle.setBorderTop(BorderStyle.MEDIUM);
            titleStyle.setBorderBottom(BorderStyle.MEDIUM);
            titleStyle.setBorderLeft(BorderStyle.MEDIUM);
            titleStyle.setBorderRight(BorderStyle.MEDIUM);
            titleStyle.setTopBorderColor(IndexedColors.DARK_GREEN.getIndex());
            titleStyle.setBottomBorderColor(IndexedColors.DARK_GREEN.getIndex());
            titleStyle.setLeftBorderColor(IndexedColors.DARK_GREEN.getIndex());
            titleStyle.setRightBorderColor(IndexedColors.DARK_GREEN.getIndex());

            CellStyle labelStyle = wb.createCellStyle();
            Font labelFont = wb.createFont();
            labelFont.setBold(true);
            labelStyle.setFont(labelFont);
            labelStyle.setBorderTop(BorderStyle.THIN);
            labelStyle.setBorderBottom(BorderStyle.THIN);
            labelStyle.setBorderLeft(BorderStyle.THIN);
            labelStyle.setBorderRight(BorderStyle.THIN);
            labelStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
            labelStyle.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
            labelStyle.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
            labelStyle.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
            labelStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            labelStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            CellStyle valueStyle = wb.createCellStyle();
            valueStyle.setBorderTop(BorderStyle.THIN);
            valueStyle.setBorderBottom(BorderStyle.THIN);
            valueStyle.setBorderLeft(BorderStyle.THIN);
            valueStyle.setBorderRight(BorderStyle.THIN);
            valueStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
            valueStyle.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
            valueStyle.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
            valueStyle.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());

            // 1行目は空行、A列は空列（HTML 反映確認用）
            sheet2.createRow(0);

            Row noteHeader = sheet2.createRow(1);
            Cell title = noteHeader.createCell(1);
            title.setCellValue("タイトル（結合）");
            title.setCellStyle(titleStyle);
            CellRangeAddress titleRegion = new CellRangeAddress(1, 1, 1, 2);
            sheet2.addMergedRegion(titleRegion);
            // 結合セル全体に罫線を付ける
            RegionUtil.setBorderTop(BorderStyle.MEDIUM, titleRegion, sheet2);
            RegionUtil.setBorderBottom(BorderStyle.MEDIUM, titleRegion, sheet2);
            RegionUtil.setBorderLeft(BorderStyle.MEDIUM, titleRegion, sheet2);
            RegionUtil.setBorderRight(BorderStyle.MEDIUM, titleRegion, sheet2);
            RegionUtil.setTopBorderColor(IndexedColors.DARK_GREEN.getIndex(), titleRegion, sheet2);
            RegionUtil.setBottomBorderColor(IndexedColors.DARK_GREEN.getIndex(), titleRegion, sheet2);
            RegionUtil.setLeftBorderColor(IndexedColors.DARK_GREEN.getIndex(), titleRegion, sheet2);
            RegionUtil.setRightBorderColor(IndexedColors.DARK_GREEN.getIndex(), titleRegion, sheet2);

            Row memoRow1 = sheet2.createRow(2);
            Cell label1 = memoRow1.createCell(1);
            label1.setCellValue("担当");
            label1.setCellStyle(labelStyle);
            Cell value1 = memoRow1.createCell(2);
            value1.setCellValue("山田");
            value1.setCellStyle(valueStyle);

            Row memoRow2 = sheet2.createRow(3);
            Cell label2 = memoRow2.createCell(1);
            label2.setCellValue("状態");
            label2.setCellStyle(labelStyle);
            Cell value2 = memoRow2.createCell(2);
            value2.setCellValue("確認中");
            value2.setCellStyle(valueStyle);

            // フォント名・サイズのサンプル
            CellStyle meiryoStyle = wb.createCellStyle();
            Font meiryoFont = wb.createFont();
            meiryoFont.setFontName("Meiryo");
            meiryoFont.setFontHeightInPoints((short) 16);
            meiryoFont.setBold(true);
            meiryoStyle.setFont(meiryoFont);
            meiryoStyle.setBorderTop(BorderStyle.THIN);
            meiryoStyle.setBorderBottom(BorderStyle.THIN);
            meiryoStyle.setBorderLeft(BorderStyle.THIN);
            meiryoStyle.setBorderRight(BorderStyle.THIN);

            CellStyle smallStyle = wb.createCellStyle();
            Font smallFont = wb.createFont();
            smallFont.setFontName("Arial");
            smallFont.setFontHeightInPoints((short) 9);
            smallFont.setItalic(true);
            smallStyle.setFont(smallFont);
            smallStyle.setBorderTop(BorderStyle.THIN);
            smallStyle.setBorderBottom(BorderStyle.THIN);
            smallStyle.setBorderLeft(BorderStyle.THIN);
            smallStyle.setBorderRight(BorderStyle.THIN);

            Row fontRow = sheet2.createRow(4);
            Cell fontLabel = fontRow.createCell(1);
            fontLabel.setCellValue("Meiryo 16pt");
            fontLabel.setCellStyle(meiryoStyle);
            Cell fontValue = fontRow.createCell(2);
            fontValue.setCellValue("Arial 9pt italic");
            fontValue.setCellStyle(smallStyle);

            // はみ出し（折り返しなし・右隣が空）
            CellStyle overflowStyle = wb.createCellStyle();
            overflowStyle.setWrapText(false);
            overflowStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            Row overflowRow = sheet2.createRow(5);
            Cell overflowCell = overflowRow.createCell(1);
            overflowCell.setCellValue("はみ出しサンプル: 折り返しなしの長い文字列です");
            overflowCell.setCellStyle(overflowStyle);

            // 折り返しあり
            CellStyle wrapStyle = wb.createCellStyle();
            wrapStyle.setWrapText(true);
            wrapStyle.setVerticalAlignment(VerticalAlignment.TOP);
            wrapStyle.setBorderTop(BorderStyle.THIN);
            wrapStyle.setBorderBottom(BorderStyle.THIN);
            wrapStyle.setBorderLeft(BorderStyle.THIN);
            wrapStyle.setBorderRight(BorderStyle.THIN);
            Row wrapRow = sheet2.createRow(6);
            wrapRow.setHeightInPoints(48);
            Cell wrapCell = wrapRow.createCell(1);
            wrapCell.setCellValue("折り返しサンプル:\nセル幅に合わせて\n複数行で表示します");
            wrapCell.setCellStyle(wrapStyle);
            Cell wrapNeighbor = wrapRow.createCell(2);
            wrapNeighbor.setCellValue("右隣あり");
            wrapNeighbor.setCellStyle(valueStyle);

            sheet2.setColumnWidth(0, 8 * 256);   // 空列 A
            sheet2.setColumnWidth(1, 14 * 256);
            sheet2.setColumnWidth(2, 30 * 256);

            try (OutputStream os = Files.newOutputStream(out)) {
                wb.write(os);
            }
        }

        System.out.println("Wrote: " + out.toAbsolutePath());
    }
}
