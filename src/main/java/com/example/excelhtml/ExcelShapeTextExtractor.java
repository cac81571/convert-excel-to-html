package com.example.excelhtml;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFShapeGroup;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.ss.usermodel.Shape;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFShapeGroup;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTextParagraph;
import org.apache.poi.xssf.usermodel.XSSFTextRun;

import java.util.StringJoiner;

/**
 * groovy-excel2csv の MyPOIUtil.searchShape 相当。
 * オートシェイプのテキスト（取消線除く）とアンカー位置を取り出す。
 */
final class ExcelShapeTextExtractor {

    @FunctionalInterface
    interface ShapeTextConsumer {
        void accept(String text, int row, int col);
    }

    private ExcelShapeTextExtractor() {
    }

    static void searchShape(Workbook book, Shape shape, Shape parent, ShapeTextConsumer consumer) {
        String text = null;
        int row = -1;
        int col = -1;

        if (shape instanceof XSSFSimpleShape xs) {
            try {
                StringJoiner sj = new StringJoiner("\n");
                for (XSSFTextParagraph p : xs.getTextParagraphs()) {
                    StringBuilder sb = new StringBuilder();
                    for (XSSFTextRun r : p.getTextRuns()) {
                        if (!r.isStrikethrough()) {
                            sb.append(r.getText());
                        }
                    }
                    sj.add(sb.toString());
                }
                text = sj.toString();
            } catch (Exception ignored) {
                // 図形によってはテキスト取得に失敗することがある
            }

            XSSFClientAnchor xca = null;
            if (parent != null && parent.getAnchor() instanceof XSSFClientAnchor parentAnchor) {
                xca = parentAnchor;
            } else if (shape.getAnchor() instanceof XSSFClientAnchor shapeAnchor) {
                xca = shapeAnchor;
            }
            if (xca != null) {
                row = xca.getRow1();
                col = xca.getCol1();
            }
        } else if (shape instanceof HSSFSimpleShape hs) {
            try {
                HSSFRichTextString hsrichStr = hs.getString();
                if (hsrichStr != null) {
                    text = TxtCellValueFormatter.removeStrikeHssf(book, hsrichStr);
                }
            } catch (Exception ignored) {
                // 同上
            }

            HSSFClientAnchor hca = null;
            if (parent != null && parent.getAnchor() instanceof HSSFClientAnchor parentAnchor) {
                hca = parentAnchor;
            } else if (shape.getAnchor() instanceof HSSFClientAnchor shapeAnchor) {
                hca = shapeAnchor;
            }
            if (hca != null) {
                row = hca.getRow1();
                col = hca.getCol1();
            }
        } else if (shape instanceof XSSFShapeGroup xsg) {
            Shape nextParent = parent == null ? shape : parent;
            for (XSSFShape child : xsg) {
                searchShape(book, child, nextParent, consumer);
            }
        } else if (shape instanceof HSSFShapeGroup hsg) {
            Shape nextParent = parent == null ? shape : parent;
            for (HSSFShape child : hsg) {
                searchShape(book, child, nextParent, consumer);
            }
        }

        if (!TxtCellValueFormatter.isBlank(text)) {
            consumer.accept(text, row, col);
        }
    }
}
