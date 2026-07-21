package com.example.excelhtml;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.FontScheme;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.model.ThemesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Excel の CellStyle を CSS クラスへ変換・キャッシュする。
 */
final class CellStyleCssRegistry {

    private final Workbook workbook;
    private final Map<Integer, String> cssByStyleIndex = new LinkedHashMap<>();
    private final Map<Integer, String> classByStyleIndex = new LinkedHashMap<>();
    private int nextClassId;

    CellStyleCssRegistry(Workbook workbook) {
        this.workbook = workbook;
    }

    /** スタイルに対応するクラス名。見た目の差分がなければ null。 */
    String className(CellStyle style) {
        if (style == null) {
            return null;
        }
        int index = style.getIndex();
        if (classByStyleIndex.containsKey(index)) {
            return classByStyleIndex.get(index);
        }
        String css = buildCss(style);
        if (css.isEmpty()) {
            classByStyleIndex.put(index, null);
            return null;
        }
        String className = "s" + (nextClassId++);
        classByStyleIndex.put(index, className);
        cssByStyleIndex.put(index, css);
        return className;
    }

    void appendCss(Appendable out) throws IOException {
        out.append("<style>\n");
        // ブックの既定フォントをベースにし、親要素の font 指定に上書きされないようにする
        Font defaultFont = workbook.getNumberOfFonts() > 0 ? workbook.getFontAt(0) : null;
        out.append(".excel-grid td { ");
        appendFont(out, defaultFont, false);
        out.append("}\n");
        for (Map.Entry<Integer, String> entry : cssByStyleIndex.entrySet()) {
            String className = classByStyleIndex.get(entry.getKey());
            if (className == null) {
                continue;
            }
            out.append(".excel-grid td.").append(className).append(" { ").append(entry.getValue()).append(" }\n");
        }
        out.append("</style>\n");
    }

    private String buildCss(CellStyle style) {
        StringBuilder css = new StringBuilder();
        appendFont(css, workbook.getFontAt(style.getFontIndexAsInt()), true);
        appendFill(css, style);
        appendAlignment(css, style);
        appendBorders(css, style);
        if (style.getWrapText()) {
            css.append("white-space: pre-wrap !important; ");
            css.append("overflow-wrap: break-word !important; ");
            css.append("word-break: break-word !important; ");
            css.append("overflow: hidden !important; ");
            css.append("text-overflow: clip !important; ");
        } else {
            // 折り返しなし: 隣の空白セルへはみ出せるようにする
            css.append("white-space: nowrap !important; ");
            css.append("overflow: visible !important; ");
            css.append("text-overflow: clip !important; ");
        }
        int indent = style.getIndention();
        if (indent > 0) {
            css.append("padding-left: ").append(indent * 8).append("px; ");
        }
        return css.toString().trim();
    }

    private void appendFont(Appendable css, Font font, boolean important) throws IOException {
        if (font == null) {
            return;
        }
        String importantSuffix = important ? " !important" : "";
        String name = resolveFontName(font);
        if (name != null && !name.isBlank()) {
            css.append("font-family: \"").append(name.replace("\"", "\\\"")).append("\", \"Yu Gothic UI\", Meiryo, sans-serif")
                    .append(importantSuffix).append("; ");
        }
        short points = font.getFontHeightInPoints();
        if (points > 0) {
            css.append("font-size: ").append(String.valueOf(points)).append("pt")
                    .append(importantSuffix).append("; ");
        }
        if (font.getBold()) {
            css.append("font-weight: 700").append(importantSuffix).append("; ");
        } else if (important) {
            css.append("font-weight: 400").append(importantSuffix).append("; ");
        }
        if (font.getItalic()) {
            css.append("font-style: italic").append(importantSuffix).append("; ");
        } else if (important) {
            css.append("font-style: normal").append(importantSuffix).append("; ");
        }
        boolean underline = font.getUnderline() != Font.U_NONE;
        boolean strike = font.getStrikeout();
        if (underline && strike) {
            css.append("text-decoration: underline line-through").append(importantSuffix).append("; ");
        } else if (underline) {
            css.append("text-decoration: underline").append(importantSuffix).append("; ");
        } else if (strike) {
            css.append("text-decoration: line-through").append(importantSuffix).append("; ");
        } else if (important) {
            css.append("text-decoration: none").append(importantSuffix).append("; ");
        }
        String color = fontColor(font);
        if (color != null) {
            css.append("color: ").append(color).append(importantSuffix).append("; ");
        }
    }

    private void appendFont(StringBuilder css, Font font, boolean important) {
        try {
            appendFont((Appendable) css, font, important);
        } catch (IOException e) {
            throw new IllegalStateException(e);
        }
    }

    /** XSSF のテーマフォント（major/minor）を解決して実フォント名を返す。 */
    private String resolveFontName(Font font) {
        if (font instanceof XSSFFont xssfFont && workbook instanceof XSSFWorkbook xssfWorkbook) {
            ThemesTable theme = xssfWorkbook.getTheme();
            if (theme != null) {
                xssfFont.setThemesTable(theme);
            }
            FontScheme scheme = xssfFont.getScheme();
            String name = xssfFont.getFontName();
            if (name != null && !name.isBlank()
                    && scheme != FontScheme.MAJOR && scheme != FontScheme.MINOR) {
                return name;
            }
            if (name != null && !name.isBlank()) {
                return name;
            }
        }
        return font.getFontName();
    }

    private void appendFill(StringBuilder css, CellStyle style) {
        FillPatternType pattern = style.getFillPattern();
        if (pattern == FillPatternType.NO_FILL || pattern == FillPatternType.SOLID_FOREGROUND) {
            // SOLID は前景色、NO_FILL は何もしない
        }
        if (pattern == FillPatternType.SOLID_FOREGROUND) {
            String color = toCssColor(style.getFillForegroundColorColor());
            if (color != null) {
                css.append("background-color: ").append(color).append("; ");
            }
        } else if (pattern != null && pattern != FillPatternType.NO_FILL) {
            String fg = toCssColor(style.getFillForegroundColorColor());
            String bg = toCssColor(style.getFillBackgroundColorColor());
            if (fg != null) {
                css.append("background-color: ").append(fg).append("; ");
            } else if (bg != null) {
                css.append("background-color: ").append(bg).append("; ");
            }
        }
    }

    private void appendAlignment(StringBuilder css, CellStyle style) {
        HorizontalAlignment align = style.getAlignment();
        if (align != null) {
            switch (align) {
                case CENTER, CENTER_SELECTION -> css.append("text-align: center; ");
                case RIGHT -> css.append("text-align: right; ");
                case LEFT -> css.append("text-align: left; ");
                case JUSTIFY, FILL, DISTRIBUTED -> css.append("text-align: justify; ");
                default -> { /* GENERAL などはそのまま */ }
            }
        }
        VerticalAlignment valign = style.getVerticalAlignment();
        if (valign != null) {
            switch (valign) {
                case TOP -> css.append("vertical-align: top; ");
                case CENTER -> css.append("vertical-align: middle; ");
                case BOTTOM -> css.append("vertical-align: bottom; ");
                default -> { }
            }
        }
    }

    private void appendBorders(StringBuilder css, CellStyle style) {
        appendBorder(css, "border-top", style.getBorderTop(), borderColor(style, "top"));
        appendBorder(css, "border-right", style.getBorderRight(), borderColor(style, "right"));
        appendBorder(css, "border-bottom", style.getBorderBottom(), borderColor(style, "bottom"));
        appendBorder(css, "border-left", style.getBorderLeft(), borderColor(style, "left"));
    }

    private void appendBorder(StringBuilder css, String prop, BorderStyle border, String color) {
        if (border == null || border == BorderStyle.NONE) {
            return;
        }
        String width;
        String styleCss;
        switch (border) {
            case HAIR, DOTTED -> { width = "1px"; styleCss = "dotted"; }
            case DASHED, MEDIUM_DASHED, DASH_DOT, MEDIUM_DASH_DOT,
                 DASH_DOT_DOT, MEDIUM_DASH_DOT_DOT, SLANTED_DASH_DOT -> { width = "1px"; styleCss = "dashed"; }
            case DOUBLE -> { width = "3px"; styleCss = "double"; }
            case MEDIUM -> { width = "2px"; styleCss = "solid"; }
            case THICK -> { width = "3px"; styleCss = "solid"; }
            default -> { width = "1px"; styleCss = "solid"; } // THIN など
        }
        String borderColor = color != null ? color : "#000000";
        css.append(prop).append(": ").append(width).append(' ').append(styleCss).append(' ').append(borderColor).append("; ");
    }

    private String borderColor(CellStyle style, String side) {
        if (style instanceof XSSFCellStyle xssf) {
            XSSFColor color = switch (side) {
                case "top" -> xssf.getTopBorderXSSFColor();
                case "right" -> xssf.getRightBorderXSSFColor();
                case "bottom" -> xssf.getBottomBorderXSSFColor();
                case "left" -> xssf.getLeftBorderXSSFColor();
                default -> null;
            };
            return toCssColor(color);
        }
        short indexed = switch (side) {
            case "top" -> style.getTopBorderColor();
            case "right" -> style.getRightBorderColor();
            case "bottom" -> style.getBottomBorderColor();
            case "left" -> style.getLeftBorderColor();
            default -> 0;
        };
        return indexedColor(indexed);
    }

    private String fontColor(Font font) {
        if (font instanceof XSSFFont xssfFont) {
            return toCssColor(xssfFont.getXSSFColor());
        }
        short indexed = font.getColor();
        if (indexed == Font.COLOR_NORMAL) {
            return null;
        }
        return indexedColor(indexed);
    }

    private String indexedColor(short index) {
        HSSFColor color = HSSFColor.getIndexHash().get((int) index);
        if (color == null) {
            return null;
        }
        short[] t = color.getTriplet();
        return String.format("#%02X%02X%02X", t[0], t[1], t[2]);
    }

    private static String toCssColor(org.apache.poi.ss.usermodel.Color color) {
        if (color == null) {
            return null;
        }
        if (color instanceof XSSFColor xssf) {
            if (xssf.isAuto()) {
                return null;
            }
            byte[] rgb = xssf.getRGBWithTint();
            if (rgb == null) {
                rgb = xssf.getRGB();
            }
            if (rgb == null) {
                return null;
            }
            int offset = rgb.length == 4 ? 1 : 0;
            return String.format("#%02X%02X%02X",
                    rgb[offset] & 0xFF, rgb[offset + 1] & 0xFF, rgb[offset + 2] & 0xFF);
        }
        if (color instanceof HSSFColor hssf) {
            short[] t = hssf.getTriplet();
            if (t == null) {
                return null;
            }
            return String.format("#%02X%02X%02X", t[0], t[1], t[2]);
        }
        return null;
    }
}
