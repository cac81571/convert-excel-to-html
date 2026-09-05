package com.example.excelhtml;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Shape;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.FileTime;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.StringJoiner;
import java.util.TreeMap;

/**
 * Excel をタブ区切り TXT（groovy-excel2csv 相当）へ変換する。
 * <p>
 * 出力形式（行番号あり）:
 * <pre>
 * [シート名]\tR00001\tセル1\tセル2\t...
 * [シート名]\tA00003\tオートシェイプテキスト
 * </pre>
 */
public final class ExcelToTxtConverter {

    /** デフォルトの最大出力列数（0始まりで maxCol 未満まで）。 */
    public static final int DEFAULT_MAX_COL = 100;

    private final int defaultMaxCol;
    private final Map<String, Integer> maxColBySheet;
    private final boolean withRowNumber;

    public ExcelToTxtConverter() {
        this(true, DEFAULT_MAX_COL, Map.of());
    }

    public ExcelToTxtConverter(boolean withRowNumber) {
        this(withRowNumber, DEFAULT_MAX_COL, Map.of());
    }

    public ExcelToTxtConverter(boolean withRowNumber, int defaultMaxCol, Map<String, Integer> maxColBySheet) {
        this.withRowNumber = withRowNumber;
        this.defaultMaxCol = defaultMaxCol;
        this.maxColBySheet = maxColBySheet == null ? Map.of() : Map.copyOf(maxColBySheet);
    }

    /**
     * Excel を TXT に変換する。
     *
     * @return 出力した TXT のパス
     */
    public Path convert(Path excelPath, Path txtPath) throws IOException {
        ZipSecureFile.setMinInflateRatio(0.001);
        Path absoluteExcel = excelPath.toAbsolutePath().normalize();
        Path absoluteTxt = txtPath.toAbsolutePath().normalize();
        Path parent = absoluteTxt.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }

        try (InputStream in = Files.newInputStream(absoluteExcel);
             Workbook workbook = WorkbookFactory.create(in);
             BufferedWriter out = Files.newBufferedWriter(absoluteTxt, StandardCharsets.UTF_8)) {
            convert(workbook, out);
        }

        FileTime excelTime = Files.getLastModifiedTime(absoluteExcel);
        Files.setLastModifiedTime(absoluteTxt, excelTime);
        return absoluteTxt;
    }

    public void convert(Workbook workbook, Appendable out) throws IOException {
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            String sheetName = TxtCellValueFormatter.strip(sheet.getSheetName());

            if (workbook.getSheetVisibility(i) != SheetVisibility.VISIBLE) {
                System.out.println("[" + sheetName + "] Skip(Not Visible)");
                continue;
            }

            int maxCol = maxColBySheet.getOrDefault(sheetName, defaultMaxCol);
            System.out.println("[" + sheetName + "] Process maxcol=" + maxCol);

            writeSheetRows(sheet, sheetName, maxCol, out);
            writeSheetShapes(workbook, sheet, sheetName, out);
        }
    }

    private void writeSheetRows(Sheet sheet, String sheetName, int maxCol, Appendable out) throws IOException {
        for (Row row : sheet) {
            StringJoiner sjrow = new StringJoiner("\t");
            int first = row.getFirstCellNum();
            int last = row.getLastCellNum();
            if (first < 0) {
                continue;
            }
            for (int idx = 0; idx < first; idx++) {
                sjrow.add("");
            }
            for (int idx = first; idx < last; idx++) {
                if (idx >= maxCol) {
                    break;
                }
                Cell cell = row.getCell(idx);
                String cellValue = "";
                if (cell != null) {
                    cellValue = TxtCellValueFormatter.nullToEmpty(TxtCellValueFormatter.getCellValue(cell));
                    cellValue = TxtCellValueFormatter.stripNewlines(cellValue);
                }
                sjrow.add(cellValue);
            }

            String rowString = TxtCellValueFormatter.stripEnd(sjrow.toString());
            if (TxtCellValueFormatter.isBlank(rowString)) {
                continue;
            }

            StringJoiner sj = new StringJoiner("\t");
            sj.add("[" + sheetName + "]");
            if (withRowNumber) {
                sj.add("R" + leftPad(row.getRowNum() + 1, 5));
            }
            sj.add(rowString);
            out.append(sj.toString()).append("\r\n");
        }
    }

    private void writeSheetShapes(Workbook workbook, Sheet sheet, String sheetName, Appendable out)
            throws IOException {
        Drawing<?> drawing = sheet.getDrawingPatriarch();
        if (drawing == null) {
            return;
        }

        TreeMap<Integer, String> byRow = new TreeMap<>();
        for (Shape shape : iterableShapes(drawing)) {
            ExcelShapeTextExtractor.searchShape(workbook, shape, null, (text, row, col) -> {
                String cleaned = TxtCellValueFormatter.stripNewlines(text);
                StringJoiner sj = new StringJoiner("\t");
                sj.add("[" + sheetName + "]");
                if (withRowNumber) {
                    sj.add("A" + leftPad(row + 1, 5));
                }
                sj.add(cleaned);
                byRow.put(row, sj.toString());
            });
        }

        for (String line : byRow.values()) {
            out.append(line).append("\r\n");
        }
    }

    @SuppressWarnings({"unchecked", "rawtypes"})
    private static Iterable<Shape> iterableShapes(Drawing<?> drawing) {
        if (!(drawing instanceof Iterable)) {
            return java.util.List.of();
        }
        Iterable raw = (Iterable) drawing;
        return () -> {
            Iterator it = raw.iterator();
            return new Iterator<Shape>() {
                @Override
                public boolean hasNext() {
                    return it.hasNext();
                }

                @Override
                public Shape next() {
                    return (Shape) it.next();
                }
            };
        };
    }

    private static String leftPad(int value, int width) {
        String s = Integer.toString(value);
        if (s.length() >= width) {
            return s;
        }
        return "0".repeat(width - s.length()) + s;
    }

    /** Excel ファイル名から対応する TXT ファイル名を作る。 */
    public static String toTxtFileName(String excelFileName) {
        int dot = excelFileName.lastIndexOf('.');
        String base = dot > 0 ? excelFileName.substring(0, dot) : excelFileName;
        return base + ".txt";
    }

    /** HTML パスと同じベース名の TXT パスを返す。 */
    public static Path toTxtPath(Path htmlPath) {
        String name = htmlPath.getFileName().toString();
        String lower = name.toLowerCase();
        String base;
        if (lower.endsWith(".html")) {
            base = name.substring(0, name.length() - 5);
        } else if (lower.endsWith(".htm")) {
            base = name.substring(0, name.length() - 4);
        } else {
            base = name;
        }
        Path parent = htmlPath.getParent();
        Path file = Path.of(base + ".txt");
        return parent == null ? file : parent.resolve(file);
    }

    /** 既定のシート別 maxcol（リクエスト=37, レスポンス=35）。 */
    public static Map<String, Integer> defaultMaxColBySheet() {
        Map<String, Integer> map = new HashMap<>();
        map.put("リクエスト", 37);
        map.put("レスポンス", 35);
        return map;
    }
}
