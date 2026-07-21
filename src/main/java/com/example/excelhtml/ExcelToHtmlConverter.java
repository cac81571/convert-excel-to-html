package com.example.excelhtml;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

import java.io.IOException;
import java.io.InputStream;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * Excel のセル内容を HTML テーブルへ変換する。
 * 複数シートは上部タブで表示／非表示を切り替える。
 * 列見出し・行見出し・列幅・セル結合・セルスタイルに対応する。
 * シートは {@code #sheet-N} でダイレクトアクセス可能。
 */
public final class ExcelToHtmlConverter {

    public record ConversionResult(List<String> sheetNames, String downloadHref, String downloadFileName) {
        public ConversionResult(List<String> sheetNames) {
            this(sheetNames, null, null);
        }
    }

    public ConversionResult convert(Path excelPath, Path htmlPath) throws IOException {
        Path absoluteExcel = excelPath.toAbsolutePath().normalize();
        Path absoluteHtml = htmlPath.toAbsolutePath().normalize();
        String downloadFileName = Objects.toString(absoluteExcel.getFileName(), "workbook.xlsx");
        String downloadHref = prepareDownloadHref(absoluteExcel, absoluteHtml, downloadFileName);
        try (InputStream in = Files.newInputStream(absoluteExcel);
             Workbook workbook = WorkbookFactory.create(in);
             Writer writer = Files.newBufferedWriter(absoluteHtml, StandardCharsets.UTF_8)) {
            ConversionResult result = convert(workbook, writer, absoluteExcel, downloadHref, downloadFileName);
            return new ConversionResult(result.sheetNames(), downloadHref, downloadFileName);
        }
    }

    public ConversionResult convert(Workbook workbook, Appendable out) throws IOException {
        return convert(workbook, out, null, null, null);
    }

    public ConversionResult convert(Workbook workbook, Appendable out, Path excelSource) throws IOException {
        return convert(workbook, out, excelSource, null, null);
    }

    public ConversionResult convert(Workbook workbook, Appendable out, Path excelSource,
                                    String downloadHref, String downloadFileName) throws IOException {
        CellValueFormatter formatter = new CellValueFormatter(workbook.getCreationHelper().createFormulaEvaluator());
        CellStyleCssRegistry styleRegistry = new CellStyleCssRegistry(workbook);
        int sheetCount = workbook.getNumberOfSheets();
        int activeIndex = Math.max(0, workbook.getActiveSheetIndex());
        List<String> sheetNames = new ArrayList<>(sheetCount);
        Path absoluteExcel = excelSource == null ? null : excelSource.toAbsolutePath().normalize();

        StringBuilder body = new StringBuilder();
        body.append("<div class=\"workbook\">\n");
        if (absoluteExcel != null) {
            appendSourceBar(absoluteExcel, downloadHref, downloadFileName, body);
        }

        body.append("<div class=\"sheet-tabs\" role=\"tablist\">\n");
        for (int s = 0; s < sheetCount; s++) {
            Sheet sheet = workbook.getSheetAt(s);
            String sheetName = sheet.getSheetName();
            sheetNames.add(sheetName);
            boolean selected = s == activeIndex;
            body.append("  <button type=\"button\" class=\"sheet-tab")
                    .append(selected ? " is-active" : "")
                    .append("\" role=\"tab\" id=\"tab-").append(s)
                    .append("\" aria-controls=\"sheet-").append(s)
                    .append("\" aria-selected=\"").append(selected ? "true" : "false")
                    .append("\" data-sheet-index=\"").append(s)
                    .append("\" data-sheet-name=\"").append(escapeHtml(sheetName)).append("\">")
                    .append(escapeHtml(sheetName))
                    .append("</button>\n");
        }
        body.append("</div>\n");

        body.append("<div class=\"sheet-panels\">\n");
        for (int s = 0; s < sheetCount; s++) {
            Sheet sheet = workbook.getSheetAt(s);
            boolean selected = s == activeIndex;
            body.append("  <div class=\"sheet-panel")
                    .append(selected ? " is-active" : "")
                    .append("\" role=\"tabpanel\" id=\"sheet-").append(s)
                    .append("\" aria-labelledby=\"tab-").append(s)
                    .append("\" data-sheet-name=\"").append(escapeHtml(sheet.getSheetName())).append("\"")
                    .append(selected ? "" : " hidden")
                    .append(">\n");
            body.append("    <div class=\"sheet-viewport\">\n");
            appendTable(sheet, formatter, styleRegistry, body);
            body.append("    </div>\n");
            body.append("  </div>\n");
        }
        body.append("</div>\n");
        body.append("</div>\n");

        String pageTitle = absoluteExcel != null
                ? absoluteExcel.getFileName().toString()
                : "Excel to HTML";

        out.append("<!DOCTYPE html>\n");
        out.append("<html lang=\"ja\">\n");
        out.append("<head>\n");
        out.append("<meta charset=\"UTF-8\">\n");
        out.append("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n");
        out.append("<title>").append(escapeHtml(pageTitle)).append("</title>\n");
        appendStyles(out);
        styleRegistry.appendCss(out);
        out.append("</head>\n");
        out.append("<body>\n");
        out.append(body);
        appendScript(out);
        out.append("</body>\n");
        out.append("</html>\n");
        return new ConversionResult(List.copyOf(sheetNames), downloadHref, downloadFileName);
    }

    private static void appendSourceBar(Path excelPath, String downloadHref, String downloadFileName,
                                        Appendable out) throws IOException {
        String displayPath = excelPath.toString();
        String openEditUri = toExcelOpenEditUri(excelPath);
        String fileName = downloadFileName != null
                ? downloadFileName
                : Objects.toString(excelPath.getFileName(), "workbook.xlsx");

        out.append("<div class=\"source-bar\" data-excel-path=\"").append(escapeHtml(displayPath)).append("\">\n");
        out.append("  <div class=\"source-main\">\n");
        out.append("    <span class=\"source-label\">元の Excel</span>\n");
        out.append("    <code class=\"source-path\" title=\"").append(escapeHtml(displayPath)).append("\">")
                .append(escapeHtml(displayPath)).append("</code>\n");
        out.append("  </div>\n");
        out.append("  <div class=\"source-actions\">\n");
        out.append("    <button type=\"button\" class=\"source-btn\" data-copy-excel-path>パスをコピー</button>\n");
        if (downloadHref != null && !downloadHref.isBlank()) {
            out.append("    <a class=\"source-btn source-btn-primary\" href=\"")
                    .append(escapeHtml(downloadHref)).append("\" download=\"")
                    .append(escapeHtml(fileName))
                    .append("\" title=\"元の Excel をダウンロード\">ダウンロード</a>\n");
        }
        out.append("    <a class=\"source-btn\" href=\"").append(escapeHtml(openEditUri))
                .append("\" title=\"元ファイルを Excel で開く\">Excelで開く</a>\n");
        out.append("  </div>\n");
        out.append("</div>\n");
    }

    /** Windows 向けの file:/// URI（ASCII）。 */
    static String toFileUri(Path path) {
        return path.toAbsolutePath().normalize().toUri().toASCIIString();
    }

    /** Excel を起動して指定ファイルを開く（Office URI）。 */
    static String toExcelOpenEditUri(Path path) {
        return "ms-excel:ofe|u|" + toFileUri(path);
    }

    /**
     * HTML から見た Excel の相対パス。作れない場合は HTML 横へコピーしてファイル名を返す。
     */
    static String prepareDownloadHref(Path excelPath, Path htmlPath, String downloadFileName) throws IOException {
        Path absoluteExcel = excelPath.toAbsolutePath().normalize();
        Path absoluteHtml = htmlPath.toAbsolutePath().normalize();
        Path htmlDir = absoluteHtml.getParent();
        if (htmlDir != null) {
            try {
                Path relative = htmlDir.relativize(absoluteExcel);
                if (!relative.toString().isEmpty()) {
                    return relative.toString().replace('\\', '/');
                }
            } catch (IllegalArgumentException ignored) {
                // ドライブルートが違うなど
            }
        }
        Path dest = absoluteHtml.resolveSibling(downloadFileName);
        if (!dest.equals(absoluteExcel)) {
            Files.copy(absoluteExcel, dest, StandardCopyOption.REPLACE_EXISTING);
        }
        return downloadFileName;
    }

    private void appendTable(Sheet sheet, CellValueFormatter formatter, CellStyleCssRegistry styleRegistry,
                             Appendable out) throws IOException {
        MergeIndex merges = MergeIndex.of(sheet);
        int firstRow = sheet.getFirstRowNum();
        int lastRow = sheet.getLastRowNum();
        boolean empty = sheet.getPhysicalNumberOfRows() == 0 || lastRow < 0;
        if (empty) {
            firstRow = 0;
            lastRow = 0;
        }
        int maxCol = empty ? 0 : findMaxColumn(sheet, firstRow, lastRow, merges);
        // 最終行・列の先に空の 1 行・1 列を足して Excel っぽく余白を見せる
        lastRow += 1;
        maxCol += 1;

        int[] colWidthsPx = columnWidthsPx(sheet, maxCol);
        int tableWidthPx = 46; // row header
        for (int px : colWidthsPx) {
            tableWidthPx += px;
        }

        out.append("      <table class=\"excel-grid\" style=\"width:")
                .append(String.valueOf(tableWidthPx)).append("px\">\n");
        appendColGroup(colWidthsPx, out);
        out.append("        <thead>\n");
        out.append("          <tr>");
        out.append("<th class=\"corner\" scope=\"col\"></th>");
        for (int c = 0; c <= maxCol; c++) {
            out.append("<th class=\"col-header\" scope=\"col\" data-c=\"").append(String.valueOf(c)).append("\">")
                    .append(CellReference.convertNumToColString(c))
                    .append("</th>");
        }
        out.append("</tr>\n");
        out.append("        </thead>\n");
        out.append("        <tbody>\n");

        for (int r = firstRow; r <= lastRow; r++) {
            Row row = empty ? null : sheet.getRow(r);
            out.append("          <tr");
            if (row != null && !row.getZeroHeight()) {
                out.append(" style=\"height: ").append(String.valueOf(row.getHeightInPoints())).append("pt\"");
            } else if (row != null && row.getZeroHeight()) {
                out.append(" class=\"row-hidden\"");
            }
            out.append('>');
            out.append("<th class=\"row-header\" scope=\"row\" data-r=\"").append(String.valueOf(r)).append("\">")
                    .append(String.valueOf(r + 1))
                    .append("</th>");
            for (int c = 0; c <= maxCol; c++) {
                CellRangeAddress merge = merges.find(r, c);
                if (merge != null && (r != merge.getFirstRow() || c != merge.getFirstColumn())) {
                    continue;
                }

                int rowSpan = 1;
                int colSpan = 1;
                if (merge != null) {
                    rowSpan = merge.getLastRow() - merge.getFirstRow() + 1;
                    colSpan = merge.getLastColumn() - merge.getFirstColumn() + 1;
                }

                Cell cell = row == null ? null : row.getCell(c);
                String value = formatter.format(cell);
                String styleClass = cell == null ? null : styleRegistry.className(cell.getCellStyle());
                boolean hasContent = value != null && !value.isEmpty();
                boolean hasFill = cell != null && hasSolidFill(cell.getCellStyle());
                boolean wrap = cell != null && cell.getCellStyle().getWrapText();

                out.append("<td");
                StringBuilder classes = new StringBuilder();
                if (styleClass != null) {
                    classes.append(styleClass);
                }
                if (hasContent) {
                    if (!classes.isEmpty()) {
                        classes.append(' ');
                    }
                    classes.append("has-content");
                }
                if (hasFill) {
                    if (!classes.isEmpty()) {
                        classes.append(' ');
                    }
                    classes.append("has-fill");
                }
                if (wrap) {
                    if (!classes.isEmpty()) {
                        classes.append(' ');
                    }
                    classes.append("is-wrap");
                }
                if (!classes.isEmpty()) {
                    out.append(" class=\"").append(classes).append('"');
                }
                out.append(" data-r=\"").append(String.valueOf(r))
                        .append("\" data-c=\"").append(String.valueOf(c))
                        .append("\" data-rs=\"").append(String.valueOf(rowSpan))
                        .append("\" data-cs=\"").append(String.valueOf(colSpan)).append('"');
                if (rowSpan > 1) {
                    out.append(" rowspan=\"").append(String.valueOf(rowSpan)).append('"');
                }
                if (colSpan > 1) {
                    out.append(" colspan=\"").append(String.valueOf(colSpan)).append('"');
                }
                out.append('>').append(escapeHtml(value)).append("</td>");
            }
            out.append("</tr>\n");
        }

        out.append("        </tbody>\n");
        out.append("      </table>\n");
    }

    private static int[] columnWidthsPx(Sheet sheet, int maxCol) {
        int[] widths = new int[maxCol + 1];
        for (int c = 0; c <= maxCol; c++) {
            widths[c] = Math.max(1, Math.round(sheet.getColumnWidthInPixels(c)));
        }
        return widths;
    }

    private static void appendColGroup(int[] colWidthsPx, Appendable out) throws IOException {
        out.append("        <colgroup>\n");
        out.append("          <col class=\"row-header-col\" style=\"width:46px\">\n");
        for (int px : colWidthsPx) {
            out.append("          <col style=\"width:").append(String.valueOf(px)).append("px\">\n");
        }
        out.append("        </colgroup>\n");
    }

    private static void appendStyles(Appendable out) throws IOException {
        out.append("<style>\n");
        out.append("*, *::before, *::after { box-sizing: border-box; }\n");
        out.append("html, body { height: 100%; margin: 0; }\n");
        out.append("body { font-family: \"Segoe UI\", \"Yu Gothic UI\", \"Hiragino Sans\", Meiryo, sans-serif;");
        out.append(" background: #f3f3f3; color: #222; }\n");
        out.append(".workbook { display: flex; flex-direction: column; height: 100vh; min-height: 0; }\n");
        out.append(".source-bar { display: flex; flex-wrap: wrap; align-items: center; gap: 0.5rem 0.75rem;");
        out.append(" padding: 0.4rem 0.65rem; background: #eef3f0; border-bottom: 1px solid #c5d0c9;");
        out.append(" font-size: 12px; flex-shrink: 0; }\n");
        out.append(".source-main { display: flex; align-items: center; gap: 0.5rem; min-width: 0; flex: 1 1 16rem; }\n");
        out.append(".source-label { flex-shrink: 0; font-weight: 600; color: #217346; }\n");
        out.append(".source-path { display: block; min-width: 0; overflow: hidden; text-overflow: ellipsis;");
        out.append(" white-space: nowrap; font-family: Consolas, \"Cascadia Mono\", \"Courier New\", monospace;");
        out.append(" font-size: 11px; color: #333; background: transparent; padding: 0; user-select: text; }\n");
        out.append(".source-actions { display: flex; flex-wrap: wrap; gap: 0.35rem; flex-shrink: 0; }\n");
        out.append(".source-btn { appearance: none; display: inline-flex; align-items: center;");
        out.append(" border: 1px solid #b0b8b3; border-radius: 4px; background: #fff; color: #222;");
        out.append(" padding: 0.2rem 0.55rem; font: inherit; font-size: 11px; text-decoration: none;");
        out.append(" cursor: pointer; line-height: 1.4; }\n");
        out.append(".source-btn:hover, .source-btn:focus-visible { background: #f5f8f6; border-color: #217346;");
        out.append(" outline: none; }\n");
        out.append(".source-btn-primary { background: #217346; border-color: #217346; color: #fff; }\n");
        out.append(".source-btn-primary:hover, .source-btn-primary:focus-visible { background: #1a5c38;");
        out.append(" border-color: #1a5c38; color: #fff; }\n");
        out.append(".sheet-tabs { display: flex; flex-wrap: nowrap; gap: 0; overflow-x: auto;");
        out.append(" background: #f3f3f3; border-bottom: 1px solid #bbb; padding: 0 0.25rem;");
        out.append(" min-height: 28px; align-items: stretch; }\n");
        out.append(".sheet-tab { appearance: none; border: 1px solid #bbb; border-bottom: none;");
        out.append(" background: #e7e7e7; padding: 0.25rem 0.9rem; cursor: pointer;");
        out.append(" border-radius: 4px 4px 0 0; margin: 2px 1px 0; color: #333; font-size: 12px; white-space: nowrap; }\n");
        out.append(".sheet-tab:hover { background: #dedede; }\n");
        out.append(".sheet-tab.is-active { background: #fff; border-color: #bbb; font-weight: 600;");
        out.append(" margin-bottom: -1px; }\n");
        out.append(".sheet-panels { flex: 1; min-height: 0; background: #fff; }\n");
        out.append(".sheet-panel { display: none; height: 100%; }\n");
        out.append(".sheet-panel.is-active { display: block; }\n");
        out.append(".sheet-viewport { height: 100%; overflow: auto; background: #fff; }\n");
        out.append(".excel-grid { border-collapse: separate; border-spacing: 0; table-layout: fixed;");
        out.append(" background: #fff; }\n");
        out.append(".excel-grid th, .excel-grid td { border-right: 1px solid #d0d0d0; border-bottom: 1px solid #d0d0d0;");
        out.append(" padding: 1px 2px; vertical-align: middle; min-height: 22px; line-height: 1.35;");
        out.append(" box-sizing: border-box; }\n");
        /* max-width:0 … nowrap はみ出しで列が押し広げられないようにする（<col> 幅を優先） */
        out.append(".excel-grid td, .excel-grid .col-header { max-width: 0; }\n");
        out.append(".excel-grid td { background-color: transparent; overflow: visible; white-space: nowrap;");
        out.append(" text-overflow: clip; cursor: cell; user-select: none; position: relative; }\n");
        out.append(".excel-grid td.has-content, .excel-grid td.has-fill { background-color: #fff; }\n");
        out.append(".excel-grid td.has-content:not(.is-wrap) { z-index: 1; }\n");
        out.append(".excel-grid td.is-wrap { white-space: pre-wrap; overflow: hidden; overflow-wrap: break-word;");
        out.append(" word-break: break-word; text-overflow: clip; }\n");
        out.append(".excel-grid tr.row-hidden { display: none; }\n");
        out.append(".excel-grid td.is-selected { box-shadow: inset 0 0 0 999px rgba(34, 120, 180, 0.28); z-index: 2; }\n");
        out.append(".excel-grid td.is-active { outline: 2px solid #217346; outline-offset: -2px;");
        out.append(" position: relative; z-index: 3; }\n");
        out.append(".excel-grid .col-header, .excel-grid .row-header, .excel-grid .corner {");
        out.append(" background: #f0f0f0; color: #444; text-align: center; font-weight: 600;");
        out.append(" user-select: none; cursor: pointer; }\n");
        out.append(".excel-grid .col-header:hover, .excel-grid .row-header:hover, .excel-grid .corner:hover {");
        out.append(" background: #e2e2e2; }\n");
        out.append(".excel-grid .col-header.is-header-selected, .excel-grid .row-header.is-header-selected,");
        out.append(" .excel-grid .corner.is-header-selected { background: #c6c6c6; }\n");
        out.append(".excel-grid .col-header { position: sticky; top: 0; z-index: 2;");
        out.append(" border-top: 1px solid #b0b0b0; overflow: visible; }\n");
        out.append(".excel-grid .row-header { position: sticky; left: 0; z-index: 1; max-width: none; width: 46px; }\n");
        out.append(".excel-grid .corner { position: sticky; top: 0; left: 0; z-index: 3;");
        out.append(" border-top: 1px solid #b0b0b0; max-width: none; width: 46px; }\n");
        out.append(".excel-grid .row-header-col { width: 46px; }\n");
        out.append(".excel-grid .col-header.col-resize-hover,");
        out.append(" .excel-grid .row-header.col-resize-hover,");
        out.append(" .excel-grid .corner.col-resize-hover { cursor: col-resize; }\n");
        out.append(".excel-grid .row-header.row-resize-hover,");
        out.append(" .excel-grid .corner.row-resize-hover { cursor: row-resize; }\n");
        out.append("body.is-col-resizing, body.is-col-resizing * { cursor: col-resize !important; user-select: none !important; }\n");
        out.append("body.is-row-resizing, body.is-row-resizing * { cursor: row-resize !important; user-select: none !important; }\n");
        out.append("</style>\n");
    }

    private static void appendScript(Appendable out) throws IOException {
        out.append("<script>\n");
        out.append("(function () {\n");
        out.append("  var tabs = document.querySelectorAll('.sheet-tab');\n");
        out.append("  var panels = document.querySelectorAll('.sheet-panel');\n");
        out.append("  var selected = [];\n");
        out.append("  var activeCell = null;\n");
        out.append("  var anchor = null;\n");
        out.append("\n");
        out.append("  function activateSheet(index, updateHash) {\n");
        out.append("    if (index < 0 || index >= tabs.length) return;\n");
        out.append("    tabs.forEach(function (tab, i) {\n");
        out.append("      var on = i === index;\n");
        out.append("      tab.classList.toggle('is-active', on);\n");
        out.append("      tab.setAttribute('aria-selected', on ? 'true' : 'false');\n");
        out.append("    });\n");
        out.append("    panels.forEach(function (panel, i) {\n");
        out.append("      var on = i === index;\n");
        out.append("      panel.classList.toggle('is-active', on);\n");
        out.append("      if (on) panel.removeAttribute('hidden');\n");
        out.append("      else panel.setAttribute('hidden', '');\n");
        out.append("    });\n");
        out.append("    clearSelection();\n");
        out.append("    activeCell = null;\n");
        out.append("    anchor = null;\n");
        out.append("    if (updateHash !== false) {\n");
        out.append("      var next = '#sheet-' + index;\n");
        out.append("      if (location.hash !== next) history.replaceState(null, '', next);\n");
        out.append("    }\n");
        out.append("  }\n");
        out.append("  function sheetIndexFromHash() {\n");
        out.append("    var hash = (location.hash || '').replace(/^#/, '');\n");
        out.append("    if (!hash) return -1;\n");
        out.append("    var m = /^sheet-(\\d+)$/.exec(hash);\n");
        out.append("    if (m) return Number(m[1]);\n");
        out.append("    var decoded = decodeURIComponent(hash);\n");
        out.append("    for (var i = 0; i < panels.length; i++) {\n");
        out.append("      if (panels[i].getAttribute('data-sheet-name') === decoded) return i;\n");
        out.append("    }\n");
        out.append("    return -1;\n");
        out.append("  }\n");
        out.append("  tabs.forEach(function (tab) {\n");
        out.append("    tab.addEventListener('click', function () {\n");
        out.append("      activateSheet(Number(tab.getAttribute('data-sheet-index')), true);\n");
        out.append("    });\n");
        out.append("  });\n");
        out.append("  window.addEventListener('hashchange', function () {\n");
        out.append("    var idx = sheetIndexFromHash();\n");
        out.append("    if (idx >= 0) activateSheet(idx, false);\n");
        out.append("  });\n");
        out.append("  var initial = sheetIndexFromHash();\n");
        out.append("  if (initial >= 0) activateSheet(initial, false);\n");
        out.append("\n");
        out.append("  var sourceBar = document.querySelector('.source-bar');\n");
        out.append("  if (sourceBar) {\n");
        out.append("    var copyBtn = sourceBar.querySelector('[data-copy-excel-path]');\n");
        out.append("    function copyPath(path) {\n");
        out.append("      if (navigator.clipboard && navigator.clipboard.writeText) {\n");
        out.append("        return navigator.clipboard.writeText(path);\n");
        out.append("      }\n");
        out.append("      return new Promise(function (resolve, reject) {\n");
        out.append("        var ta = document.createElement('textarea');\n");
        out.append("        ta.value = path; ta.setAttribute('readonly', '');\n");
        out.append("        ta.style.position = 'fixed'; ta.style.left = '-9999px';\n");
        out.append("        document.body.appendChild(ta); ta.select();\n");
        out.append("        try {\n");
        out.append("          if (!document.execCommand('copy')) throw new Error('copy failed');\n");
        out.append("          resolve();\n");
        out.append("        } catch (err) { reject(err); }\n");
        out.append("        finally { document.body.removeChild(ta); }\n");
        out.append("      });\n");
        out.append("    }\n");
        out.append("    if (copyBtn) {\n");
        out.append("      copyBtn.addEventListener('click', function () {\n");
        out.append("        var path = sourceBar.getAttribute('data-excel-path') || '';\n");
        out.append("        var label = copyBtn.textContent;\n");
        out.append("        copyPath(path).then(function () {\n");
        out.append("          copyBtn.textContent = 'コピー済';\n");
        out.append("          window.clearTimeout(copyBtn._copyLabelTimer);\n");
        out.append("          copyBtn._copyLabelTimer = window.setTimeout(function () {\n");
        out.append("            copyBtn.textContent = label;\n");
        out.append("          }, 1200);\n");
        out.append("        }).catch(function () {\n");
        out.append("          copyBtn.textContent = '失敗';\n");
        out.append("          window.clearTimeout(copyBtn._copyLabelTimer);\n");
        out.append("          copyBtn._copyLabelTimer = window.setTimeout(function () {\n");
        out.append("            copyBtn.textContent = label;\n");
        out.append("          }, 1200);\n");
        out.append("        });\n");
        out.append("      });\n");
        out.append("    }\n");
        out.append("  }\n");
        out.append("\n");
        out.append("  function clearSelection() {\n");
        out.append("    selected.forEach(function (cell) {\n");
        out.append("      cell.classList.remove('is-selected', 'is-active');\n");
        out.append("    });\n");
        out.append("    selected = [];\n");
        out.append("    document.querySelectorAll('.excel-grid th.is-header-selected').forEach(function (th) {\n");
        out.append("      th.classList.remove('is-header-selected');\n");
        out.append("    });\n");
        out.append("  }\n");
        out.append("  function cellPos(td) {\n");
        out.append("    return {\n");
        out.append("      row: Number(td.getAttribute('data-r')),\n");
        out.append("      col: Number(td.getAttribute('data-c')),\n");
        out.append("      rs: Number(td.getAttribute('data-rs') || 1),\n");
        out.append("      cs: Number(td.getAttribute('data-cs') || 1)\n");
        out.append("    };\n");
        out.append("  }\n");
        out.append("  function findTd(table, row, col) {\n");
        out.append("    var exact = table.querySelector('td[data-r=\"' + row + '\"][data-c=\"' + col + '\"]');\n");
        out.append("    if (exact) return exact;\n");
        out.append("    var cells = table.querySelectorAll('tbody td');\n");
        out.append("    for (var i = 0; i < cells.length; i++) {\n");
        out.append("      var p = cellPos(cells[i]);\n");
        out.append("      if (row >= p.row && row < p.row + p.rs && col >= p.col && col < p.col + p.cs) return cells[i];\n");
        out.append("    }\n");
        out.append("    return null;\n");
        out.append("  }\n");
        out.append("  function activeTable() {\n");
        out.append("    var panel = document.querySelector('.sheet-panel.is-active');\n");
        out.append("    return panel ? panel.querySelector('.excel-grid') : null;\n");
        out.append("  }\n");
        out.append("  function firstDataCell(table) {\n");
        out.append("    return table ? table.querySelector('tbody td') : null;\n");
        out.append("  }\n");
        out.append("  function dataBounds(table) {\n");
        out.append("    var cells = table.querySelectorAll('tbody td');\n");
        out.append("    if (!cells.length) return null;\n");
        out.append("    var minR = Infinity, maxR = -Infinity, minC = Infinity, maxC = -Infinity;\n");
        out.append("    cells.forEach(function (cell) {\n");
        out.append("      var p = cellPos(cell);\n");
        out.append("      minR = Math.min(minR, p.row);\n");
        out.append("      maxR = Math.max(maxR, p.row + p.rs - 1);\n");
        out.append("      minC = Math.min(minC, p.col);\n");
        out.append("      maxC = Math.max(maxC, p.col + p.cs - 1);\n");
        out.append("    });\n");
        out.append("    return { minR: minR, maxR: maxR, minC: minC, maxC: maxC };\n");
        out.append("  }\n");
        out.append("  function selectRange(table, a, b, focus) {\n");
        out.append("    clearSelection();\n");
        out.append("    var r0 = Math.min(a.row, b.row), r1 = Math.max(a.row, b.row);\n");
        out.append("    var c0 = Math.min(a.col, b.col), c1 = Math.max(a.col, b.col);\n");
        out.append("    var seen = [];\n");
        out.append("    for (var r = r0; r <= r1; r++) {\n");
        out.append("      for (var c = c0; c <= c1; c++) {\n");
        out.append("        var cell = findTd(table, r, c);\n");
        out.append("        if (!cell || seen.indexOf(cell) >= 0) continue;\n");
        out.append("        seen.push(cell);\n");
        out.append("        cell.classList.add('is-selected');\n");
        out.append("        selected.push(cell);\n");
        out.append("      }\n");
        out.append("    }\n");
        out.append("    activeCell = focus;\n");
        out.append("    if (activeCell) activeCell.classList.add('is-active');\n");
        out.append("  }\n");
        out.append("  function selectSingle(td) {\n");
        out.append("    var table = td.closest('table');\n");
        out.append("    var pos = cellPos(td);\n");
        out.append("    selectRange(table, pos, { row: pos.row + pos.rs - 1, col: pos.col + pos.cs - 1 }, td);\n");
        out.append("    anchor = { table: table, pos: pos };\n");
        out.append("  }\n");
        out.append("  function selectColumn(table, colIndex, extend) {\n");
        out.append("    var b = dataBounds(table);\n");
        out.append("    if (!b) return;\n");
        out.append("    var c0 = colIndex, c1 = colIndex;\n");
        out.append("    if (extend && anchor && anchor.table === table) {\n");
        out.append("      c0 = Math.min(anchor.pos.col, colIndex);\n");
        out.append("      c1 = Math.max(anchor.pos.col, colIndex);\n");
        out.append("    } else {\n");
        out.append("      anchor = { table: table, pos: { row: b.minR, col: colIndex } };\n");
        out.append("    }\n");
        out.append("    var focus = findTd(table, b.minR, colIndex) || findTd(table, b.minR, c0);\n");
        out.append("    selectRange(table, { row: b.minR, col: c0 }, { row: b.maxR, col: c1 }, focus);\n");
        out.append("    for (var c = c0; c <= c1; c++) {\n");
        out.append("      var header = table.querySelector('th.col-header[data-c=\"' + c + '\"]');\n");
        out.append("      if (header) header.classList.add('is-header-selected');\n");
        out.append("    }\n");
        out.append("  }\n");
        out.append("  function selectRow(table, rowIndex, extend) {\n");
        out.append("    var b = dataBounds(table);\n");
        out.append("    if (!b) return;\n");
        out.append("    var r0 = rowIndex, r1 = rowIndex;\n");
        out.append("    if (extend && anchor && anchor.table === table) {\n");
        out.append("      r0 = Math.min(anchor.pos.row, rowIndex);\n");
        out.append("      r1 = Math.max(anchor.pos.row, rowIndex);\n");
        out.append("    } else {\n");
        out.append("      anchor = { table: table, pos: { row: rowIndex, col: b.minC } };\n");
        out.append("    }\n");
        out.append("    var focus = findTd(table, rowIndex, b.minC) || findTd(table, r0, b.minC);\n");
        out.append("    selectRange(table, { row: r0, col: b.minC }, { row: r1, col: b.maxC }, focus);\n");
        out.append("    for (var r = r0; r <= r1; r++) {\n");
        out.append("      var header = table.querySelector('th.row-header[data-r=\"' + r + '\"]');\n");
        out.append("      if (header) header.classList.add('is-header-selected');\n");
        out.append("    }\n");
        out.append("  }\n");
        out.append("  function selectAll(table) {\n");
        out.append("    var b = dataBounds(table);\n");
        out.append("    if (!b) return;\n");
        out.append("    var focus = findTd(table, b.minR, b.minC);\n");
        out.append("    selectRange(table, { row: b.minR, col: b.minC }, { row: b.maxR, col: b.maxC }, focus);\n");
        out.append("    anchor = { table: table, pos: { row: b.minR, col: b.minC } };\n");
        out.append("    var corner = table.querySelector('.corner');\n");
        out.append("    if (corner) corner.classList.add('is-header-selected');\n");
        out.append("    table.querySelectorAll('th.col-header, th.row-header').forEach(function (th) {\n");
        out.append("      th.classList.add('is-header-selected');\n");
        out.append("    });\n");
        out.append("  }\n");
        out.append("  function ensureActiveCell() {\n");
        out.append("    if (activeCell && activeCell.isConnected) return true;\n");
        out.append("    var td = firstDataCell(activeTable());\n");
        out.append("    if (!td) return false;\n");
        out.append("    selectSingle(td);\n");
        out.append("    return true;\n");
        out.append("  }\n");
        out.append("  function moveActive(dRow, dCol, extend) {\n");
        out.append("    if (!ensureActiveCell()) return;\n");
        out.append("    var table = activeCell.closest('table');\n");
        out.append("    var pos = cellPos(activeCell);\n");
        out.append("    var targetRow = pos.row;\n");
        out.append("    var targetCol = pos.col;\n");
        out.append("    if (dRow > 0) targetRow = pos.row + pos.rs;\n");
        out.append("    else if (dRow < 0) targetRow = pos.row - 1;\n");
        out.append("    if (dCol > 0) targetCol = pos.col + pos.cs;\n");
        out.append("    else if (dCol < 0) targetCol = pos.col - 1;\n");
        out.append("    var next = findTd(table, targetRow, targetCol);\n");
        out.append("    if (!next || next === activeCell) return;\n");
        out.append("    if (extend) {\n");
        out.append("      if (!anchor || anchor.table !== table) anchor = { table: table, pos: pos };\n");
        out.append("      var np = cellPos(next);\n");
        out.append("      selectRange(table, anchor.pos, { row: np.row + np.rs - 1, col: np.col + np.cs - 1 }, next);\n");
        out.append("    } else {\n");
        out.append("      selectSingle(next);\n");
        out.append("    }\n");
        out.append("    next.scrollIntoView({ block: 'nearest', inline: 'nearest' });\n");
        out.append("  }\n");
        out.append("  function selectionText() {\n");
        out.append("    if (!selected.length) return '';\n");
        out.append("    var table = selected[0].closest('table');\n");
        out.append("    var minR = Infinity, maxR = -Infinity, minC = Infinity, maxC = -Infinity;\n");
        out.append("    var selectedSet = selected.slice();\n");
        out.append("    selected.forEach(function (cell) {\n");
        out.append("      var p = cellPos(cell);\n");
        out.append("      minR = Math.min(minR, p.row);\n");
        out.append("      maxR = Math.max(maxR, p.row + p.rs - 1);\n");
        out.append("      minC = Math.min(minC, p.col);\n");
        out.append("      maxC = Math.max(maxC, p.col + p.cs - 1);\n");
        out.append("    });\n");
        out.append("    var lines = [];\n");
        out.append("    for (var r = minR; r <= maxR; r++) {\n");
        out.append("      var line = [];\n");
        out.append("      for (var c = minC; c <= maxC; c++) {\n");
        out.append("        var cell = findTd(table, r, c);\n");
        out.append("        if (cell && selectedSet.indexOf(cell) >= 0) {\n");
        out.append("          var p = cellPos(cell);\n");
        out.append("          line.push(p.row === r && p.col === c ? cell.textContent : '');\n");
        out.append("        } else {\n");
        out.append("          line.push('');\n");
        out.append("        }\n");
        out.append("      }\n");
        out.append("      lines.push(line.join('\\t'));\n");
        out.append("    }\n");
        out.append("    return lines.join('\\n');\n");
        out.append("  }\n");
        out.append("  function copySelection(e) {\n");
        out.append("    if (!selected.length) return false;\n");
        out.append("    var text = selectionText();\n");
        out.append("    if (e && e.clipboardData) {\n");
        out.append("      e.clipboardData.setData('text/plain', text);\n");
        out.append("      e.preventDefault();\n");
        out.append("      return true;\n");
        out.append("    }\n");
        out.append("    if (navigator.clipboard && navigator.clipboard.writeText) {\n");
        out.append("      navigator.clipboard.writeText(text);\n");
        out.append("      return true;\n");
        out.append("    }\n");
        out.append("    var ta = document.createElement('textarea');\n");
        out.append("    ta.value = text;\n");
        out.append("    document.body.appendChild(ta);\n");
        out.append("    ta.select();\n");
        out.append("    document.execCommand('copy');\n");
        out.append("    document.body.removeChild(ta);\n");
        out.append("    return true;\n");
        out.append("  }\n");
        out.append("\n");
        out.append("  var RESIZE_EDGE = 5;\n");
        out.append("  var MIN_COL_PX = 24;\n");
        out.append("  var MIN_ROW_PX = 12;\n");
        out.append("  var resizeState = null;\n");
        out.append("\n");
        out.append("  function nearRightEdge(el, clientX) {\n");
        out.append("    var rect = el.getBoundingClientRect();\n");
        out.append("    return clientX >= rect.right - RESIZE_EDGE && clientX <= rect.right + 1;\n");
        out.append("  }\n");
        out.append("  function nearBottomEdge(el, clientY) {\n");
        out.append("    var rect = el.getBoundingClientRect();\n");
        out.append("    return clientY >= rect.bottom - RESIZE_EDGE && clientY <= rect.bottom + 1;\n");
        out.append("  }\n");
        out.append("  function colElements(table) {\n");
        out.append("    return table.querySelectorAll('colgroup col');\n");
        out.append("  }\n");
        out.append("  function parsePx(value, fallback) {\n");
        out.append("    var n = parseFloat(value);\n");
        out.append("    return isFinite(n) ? n : fallback;\n");
        out.append("  }\n");
        out.append("  function syncTableWidth(table) {\n");
        out.append("    var cols = colElements(table);\n");
        out.append("    var total = 0;\n");
        out.append("    cols.forEach(function (col) {\n");
        out.append("      total += parsePx(col.style.width, col.getBoundingClientRect().width || 0);\n");
        out.append("    });\n");
        out.append("    table.style.width = Math.max(1, Math.round(total)) + 'px';\n");
        out.append("  }\n");
        out.append("  function setColumnWidth(table, colIndex, widthPx) {\n");
        out.append("    var cols = colElements(table);\n");
        out.append("    if (colIndex < 0 || colIndex >= cols.length) return;\n");
        out.append("    var w = Math.max(MIN_COL_PX, Math.round(widthPx));\n");
        out.append("    cols[colIndex].style.width = w + 'px';\n");
        out.append("    if (colIndex === 0) {\n");
        out.append("      table.querySelectorAll('.row-header, .corner').forEach(function (th) {\n");
        out.append("        th.style.width = w + 'px';\n");
        out.append("      });\n");
        out.append("    }\n");
        out.append("    syncTableWidth(table);\n");
        out.append("  }\n");
        out.append("  function setRowHeight(tr, heightPx) {\n");
        out.append("    var h = Math.max(MIN_ROW_PX, Math.round(heightPx));\n");
        out.append("    tr.style.height = h + 'px';\n");
        out.append("  }\n");
        out.append("  function resizeTargetFromEvent(e) {\n");
        out.append("    var header = e.target.closest && e.target.closest('.excel-grid th');\n");
        out.append("    if (!header) return null;\n");
        out.append("    var table = header.closest('table');\n");
        out.append("    if (!table) return null;\n");
        out.append("    if (header.classList.contains('col-header') && nearRightEdge(header, e.clientX)) {\n");
        out.append("      return { type: 'col', table: table, colIndex: Number(header.getAttribute('data-c')) + 1 };\n");
        out.append("    }\n");
        out.append("    if ((header.classList.contains('row-header') || header.classList.contains('corner'))\n");
        out.append("        && nearRightEdge(header, e.clientX)) {\n");
        out.append("      return { type: 'col', table: table, colIndex: 0 };\n");
        out.append("    }\n");
        out.append("    if ((header.classList.contains('row-header') || header.classList.contains('corner'))\n");
        out.append("        && nearBottomEdge(header, e.clientY)) {\n");
        out.append("      if (header.classList.contains('corner')) return null;\n");
        out.append("      return { type: 'row', table: table, tr: header.parentElement };\n");
        out.append("    }\n");
        out.append("    return null;\n");
        out.append("  }\n");
        out.append("  function clearResizeHover() {\n");
        out.append("    document.querySelectorAll('.excel-grid th.col-resize-hover, .excel-grid th.row-resize-hover')\n");
        out.append("      .forEach(function (th) {\n");
        out.append("        th.classList.remove('col-resize-hover', 'row-resize-hover');\n");
        out.append("      });\n");
        out.append("  }\n");
        out.append("  function updateResizeHover(e) {\n");
        out.append("    if (resizeState) return;\n");
        out.append("    clearResizeHover();\n");
        out.append("    var target = resizeTargetFromEvent(e);\n");
        out.append("    if (!target) return;\n");
        out.append("    var header = e.target.closest('.excel-grid th');\n");
        out.append("    if (!header) return;\n");
        out.append("    header.classList.add(target.type === 'col' ? 'col-resize-hover' : 'row-resize-hover');\n");
        out.append("  }\n");
        out.append("  function beginResize(target, e) {\n");
        out.append("    if (target.type === 'col') {\n");
        out.append("      var cols = colElements(target.table);\n");
        out.append("      var col = cols[target.colIndex];\n");
        out.append("      if (!col) return;\n");
        out.append("      resizeState = {\n");
        out.append("        type: 'col',\n");
        out.append("        table: target.table,\n");
        out.append("        colIndex: target.colIndex,\n");
        out.append("        startX: e.clientX,\n");
        out.append("        startY: e.clientY,\n");
        out.append("        startW: parsePx(col.style.width, col.getBoundingClientRect().width),\n");
        out.append("        pending: true\n");
        out.append("      };\n");
        out.append("    } else {\n");
        out.append("      var tr = target.tr;\n");
        out.append("      if (!tr) return;\n");
        out.append("      resizeState = {\n");
        out.append("        type: 'row',\n");
        out.append("        tr: tr,\n");
        out.append("        startX: e.clientX,\n");
        out.append("        startY: e.clientY,\n");
        out.append("        startH: tr.getBoundingClientRect().height,\n");
        out.append("        pending: true\n");
        out.append("      };\n");
        out.append("    }\n");
        out.append("    clearResizeHover();\n");
        out.append("    e.preventDefault();\n");
        out.append("    e.stopPropagation();\n");
        out.append("  }\n");
        out.append("  function moveResize(e) {\n");
        out.append("    if (!resizeState) return;\n");
        out.append("    if (resizeState.pending) {\n");
        out.append("      var dx = Math.abs(e.clientX - resizeState.startX);\n");
        out.append("      var dy = Math.abs(e.clientY - resizeState.startY);\n");
        out.append("      if (dx < 3 && dy < 3) return;\n");
        out.append("      resizeState.pending = false;\n");
        out.append("      document.body.classList.add(resizeState.type === 'col' ? 'is-col-resizing' : 'is-row-resizing');\n");
        out.append("    }\n");
        out.append("    e.preventDefault();\n");
        out.append("    if (resizeState.type === 'col') {\n");
        out.append("      setColumnWidth(resizeState.table, resizeState.colIndex,\n");
        out.append("        resizeState.startW + (e.clientX - resizeState.startX));\n");
        out.append("    } else {\n");
        out.append("      setRowHeight(resizeState.tr, resizeState.startH + (e.clientY - resizeState.startY));\n");
        out.append("    }\n");
        out.append("  }\n");
        out.append("  function endResize() {\n");
        out.append("    if (!resizeState) return;\n");
        out.append("    document.body.classList.remove('is-col-resizing', 'is-row-resizing');\n");
        out.append("    resizeState = null;\n");
        out.append("  }\n");
        out.append("  function contentPadX(el) {\n");
        out.append("    var s = getComputedStyle(el);\n");
        out.append("    return parsePx(s.paddingLeft, 0) + parsePx(s.paddingRight, 0)\n");
        out.append("      + parsePx(s.borderLeftWidth, 0) + parsePx(s.borderRightWidth, 0);\n");
        out.append("  }\n");
        out.append("  function contentPadY(el) {\n");
        out.append("    var s = getComputedStyle(el);\n");
        out.append("    return parsePx(s.paddingTop, 0) + parsePx(s.paddingBottom, 0)\n");
        out.append("      + parsePx(s.borderTopWidth, 0) + parsePx(s.borderBottomWidth, 0);\n");
        out.append("  }\n");
        out.append("  function measureContentWidth(el) {\n");
        out.append("    var text = (el.innerText || el.textContent || '').replace(/\\s+$/g, '');\n");
        out.append("    if (!text) return contentPadX(el);\n");
        out.append("    var s = getComputedStyle(el);\n");
        out.append("    var probe = document.createElement('span');\n");
        out.append("    probe.style.cssText = 'position:absolute;left:-99999px;top:0;visibility:hidden;white-space:pre;'\n");
        out.append("      + 'font:' + s.font + ';letter-spacing:' + s.letterSpacing + ';';\n");
        out.append("    probe.textContent = text;\n");
        out.append("    document.body.appendChild(probe);\n");
        out.append("    var w = probe.offsetWidth + contentPadX(el) + 4;\n");
        out.append("    document.body.removeChild(probe);\n");
        out.append("    return w;\n");
        out.append("  }\n");
        out.append("  function measureContentHeight(el, widthPx) {\n");
        out.append("    var text = el.innerText || el.textContent || '';\n");
        out.append("    if (!text.trim()) return Math.max(MIN_ROW_PX, contentPadY(el) + 14);\n");
        out.append("    var s = getComputedStyle(el);\n");
        out.append("    var probe = document.createElement('div');\n");
        out.append("    var wrap = el.classList.contains('is-wrap') || s.whiteSpace.indexOf('pre') >= 0 || s.whiteSpace === 'normal';\n");
        out.append("    probe.style.cssText = 'position:absolute;left:-99999px;top:0;visibility:hidden;box-sizing:border-box;'\n");
        out.append("      + 'width:' + Math.max(1, Math.round(widthPx)) + 'px;'\n");
        out.append("      + 'font:' + s.font + ';line-height:' + s.lineHeight + ';'\n");
        out.append("      + 'letter-spacing:' + s.letterSpacing + ';'\n");
        out.append("      + 'padding:' + s.paddingTop + ' ' + s.paddingRight + ' ' + s.paddingBottom + ' ' + s.paddingLeft + ';'\n");
        out.append("      + 'white-space:' + (wrap ? (s.whiteSpace === 'nowrap' ? 'pre-wrap' : s.whiteSpace) : 'pre') + ';'\n");
        out.append("      + 'overflow-wrap:' + s.overflowWrap + ';word-break:' + s.wordBreak + ';';\n");
        out.append("    probe.textContent = text;\n");
        out.append("    document.body.appendChild(probe);\n");
        out.append("    var h = probe.offsetHeight + 2;\n");
        out.append("    document.body.removeChild(probe);\n");
        out.append("    return h;\n");
        out.append("  }\n");
        out.append("  function autoFitColumn(table, colIndex) {\n");
        out.append("    var max = MIN_COL_PX;\n");
        out.append("    if (colIndex === 0) {\n");
        out.append("      table.querySelectorAll('th.row-header, th.corner').forEach(function (th) {\n");
        out.append("        max = Math.max(max, measureContentWidth(th));\n");
        out.append("      });\n");
        out.append("    } else {\n");
        out.append("      var dataC = colIndex - 1;\n");
        out.append("      var th = table.querySelector('th.col-header[data-c=\"' + dataC + '\"]');\n");
        out.append("      if (th) max = Math.max(max, measureContentWidth(th));\n");
        out.append("      table.querySelectorAll('td[data-c=\"' + dataC + '\"]').forEach(function (td) {\n");
        out.append("        var cs = Number(td.getAttribute('data-cs') || '1');\n");
        out.append("        if (cs !== 1) return;\n");
        out.append("        max = Math.max(max, measureContentWidth(td));\n");
        out.append("      });\n");
        out.append("    }\n");
        out.append("    setColumnWidth(table, colIndex, max);\n");
        out.append("  }\n");
        out.append("  function autoFitRow(tr) {\n");
        out.append("    if (!tr) return;\n");
        out.append("    var table = tr.closest('table');\n");
        out.append("    var max = MIN_ROW_PX;\n");
        out.append("    var rh = tr.querySelector('th.row-header');\n");
        out.append("    if (rh) max = Math.max(max, measureContentHeight(rh, rh.getBoundingClientRect().width));\n");
        out.append("    tr.querySelectorAll('td').forEach(function (td) {\n");
        out.append("      var rs = Number(td.getAttribute('data-rs') || '1');\n");
        out.append("      if (rs !== 1) return;\n");
        out.append("      max = Math.max(max, measureContentHeight(td, td.getBoundingClientRect().width));\n");
        out.append("    });\n");
        out.append("    setRowHeight(tr, max);\n");
        out.append("  }\n");
        out.append("\n");
        out.append("  var dragging = false;\n");
        out.append("  var dragMode = null;\n");
        out.append("  var dragTable = null;\n");
        out.append("  function cellFromPoint(x, y, table) {\n");
        out.append("    var el = document.elementFromPoint(x, y);\n");
        out.append("    var td = el && el.closest ? el.closest('.excel-grid td') : null;\n");
        out.append("    if (!td || (table && td.closest('table') !== table)) return null;\n");
        out.append("    return td;\n");
        out.append("  }\n");
        out.append("  function colIndexAt(table, x, y) {\n");
        out.append("    var el = document.elementFromPoint(x, y);\n");
        out.append("    if (el && el.closest) {\n");
        out.append("      var th = el.closest('th.col-header');\n");
        out.append("      if (th && th.closest('table') === table) return Number(th.getAttribute('data-c'));\n");
        out.append("      var td = el.closest('td');\n");
        out.append("      if (td && td.closest('table') === table) return Number(td.getAttribute('data-c'));\n");
        out.append("    }\n");
        out.append("    var best = null;\n");
        out.append("    table.querySelectorAll('th.col-header').forEach(function (h) {\n");
        out.append("      var r = h.getBoundingClientRect();\n");
        out.append("      if (x >= r.left && x <= r.right) best = Number(h.getAttribute('data-c'));\n");
        out.append("    });\n");
        out.append("    return best;\n");
        out.append("  }\n");
        out.append("  function rowIndexAt(table, x, y) {\n");
        out.append("    var el = document.elementFromPoint(x, y);\n");
        out.append("    if (el && el.closest) {\n");
        out.append("      var th = el.closest('th.row-header');\n");
        out.append("      if (th && th.closest('table') === table) return Number(th.getAttribute('data-r'));\n");
        out.append("      var td = el.closest('td');\n");
        out.append("      if (td && td.closest('table') === table) return Number(td.getAttribute('data-r'));\n");
        out.append("    }\n");
        out.append("    var best = null;\n");
        out.append("    table.querySelectorAll('th.row-header').forEach(function (h) {\n");
        out.append("      var r = h.getBoundingClientRect();\n");
        out.append("      if (y >= r.top && y <= r.bottom) best = Number(h.getAttribute('data-r'));\n");
        out.append("    });\n");
        out.append("    return best;\n");
        out.append("  }\n");
        out.append("  function extendTo(td) {\n");
        out.append("    if (!td || !anchor) return;\n");
        out.append("    var table = td.closest('table');\n");
        out.append("    if (anchor.table !== table) return;\n");
        out.append("    var pos = cellPos(td);\n");
        out.append("    selectRange(table, anchor.pos, { row: pos.row + pos.rs - 1, col: pos.col + pos.cs - 1 }, td);\n");
        out.append("  }\n");
        out.append("  document.addEventListener('mousedown', function (e) {\n");
        out.append("    if (e.button !== 0) return;\n");
        out.append("    var resizeTarget = resizeTargetFromEvent(e);\n");
        out.append("    if (resizeTarget) {\n");
        out.append("      beginResize(resizeTarget, e);\n");
        out.append("      dragging = false;\n");
        out.append("      dragMode = null;\n");
        out.append("      dragTable = null;\n");
        out.append("      return;\n");
        out.append("    }\n");
        out.append("    var header = e.target.closest('.excel-grid th');\n");
        out.append("    if (header) {\n");
        out.append("      var table = header.closest('table');\n");
        out.append("      if (header.classList.contains('corner')) {\n");
        out.append("        selectAll(table);\n");
        out.append("        dragging = false;\n");
        out.append("        dragMode = null;\n");
        out.append("        dragTable = null;\n");
        out.append("        return;\n");
        out.append("      }\n");
        out.append("      if (header.classList.contains('col-header')) {\n");
        out.append("        e.preventDefault();\n");
        out.append("        selectColumn(table, Number(header.getAttribute('data-c')), e.shiftKey);\n");
        out.append("        dragging = true;\n");
        out.append("        dragMode = 'col';\n");
        out.append("        dragTable = table;\n");
        out.append("        return;\n");
        out.append("      }\n");
        out.append("      if (header.classList.contains('row-header')) {\n");
        out.append("        e.preventDefault();\n");
        out.append("        selectRow(table, Number(header.getAttribute('data-r')), e.shiftKey);\n");
        out.append("        dragging = true;\n");
        out.append("        dragMode = 'row';\n");
        out.append("        dragTable = table;\n");
        out.append("        return;\n");
        out.append("      }\n");
        out.append("      dragging = false;\n");
        out.append("      dragMode = null;\n");
        out.append("      return;\n");
        out.append("    }\n");
        out.append("    var td = e.target.closest('.excel-grid td');\n");
        out.append("    if (!td) {\n");
        out.append("      if (!e.target.closest('.excel-grid')) {\n");
        out.append("        clearSelection();\n");
        out.append("        activeCell = null;\n");
        out.append("        anchor = null;\n");
        out.append("      }\n");
        out.append("      dragging = false;\n");
        out.append("      dragMode = null;\n");
        out.append("      return;\n");
        out.append("    }\n");
        out.append("    e.preventDefault();\n");
        out.append("    var table = td.closest('table');\n");
        out.append("    var pos = cellPos(td);\n");
        out.append("    if (e.shiftKey && anchor && anchor.table === table) {\n");
        out.append("      extendTo(td);\n");
        out.append("      dragging = false;\n");
        out.append("      dragMode = null;\n");
        out.append("      dragTable = null;\n");
        out.append("      return;\n");
        out.append("    }\n");
        out.append("    selectSingle(td);\n");
        out.append("    dragging = true;\n");
        out.append("    dragMode = 'cell';\n");
        out.append("    dragTable = table;\n");
        out.append("  });\n");
        out.append("  document.addEventListener('mousemove', function (e) {\n");
        out.append("    if (resizeState) {\n");
        out.append("      moveResize(e);\n");
        out.append("      return;\n");
        out.append("    }\n");
        out.append("    updateResizeHover(e);\n");
        out.append("    if (!dragging || !dragTable || !dragMode) return;\n");
        out.append("    e.preventDefault();\n");
        out.append("    if (dragMode === 'col') {\n");
        out.append("      var colIndex = colIndexAt(dragTable, e.clientX, e.clientY);\n");
        out.append("      if (colIndex == null || isNaN(colIndex)) return;\n");
        out.append("      selectColumn(dragTable, colIndex, true);\n");
        out.append("      return;\n");
        out.append("    }\n");
        out.append("    if (dragMode === 'row') {\n");
        out.append("      var rowIndex = rowIndexAt(dragTable, e.clientX, e.clientY);\n");
        out.append("      if (rowIndex == null || isNaN(rowIndex)) return;\n");
        out.append("      selectRow(dragTable, rowIndex, true);\n");
        out.append("      return;\n");
        out.append("    }\n");
        out.append("    var td = e.target.closest('.excel-grid td');\n");
        out.append("    if (!td || td.closest('table') !== dragTable) {\n");
        out.append("      td = cellFromPoint(e.clientX, e.clientY, dragTable);\n");
        out.append("    }\n");
        out.append("    if (!td) return;\n");
        out.append("    extendTo(td);\n");
        out.append("  });\n");
        out.append("  document.addEventListener('mouseup', function () {\n");
        out.append("    endResize();\n");
        out.append("    dragging = false;\n");
        out.append("    dragMode = null;\n");
        out.append("    dragTable = null;\n");
        out.append("  });\n");
        out.append("  document.addEventListener('dblclick', function (e) {\n");
        out.append("    var target = resizeTargetFromEvent(e);\n");
        out.append("    if (!target) return;\n");
        out.append("    endResize();\n");
        out.append("    e.preventDefault();\n");
        out.append("    e.stopPropagation();\n");
        out.append("    if (target.type === 'col') autoFitColumn(target.table, target.colIndex);\n");
        out.append("    else autoFitRow(target.tr);\n");
        out.append("  });\n");
        out.append("  document.addEventListener('mouseleave', function (e) {\n");
        out.append("    if (e.target === document.documentElement || e.target === document.body) clearResizeHover();\n");
        out.append("  }, true);\n");
        out.append("\n");
        out.append("  document.addEventListener('keydown', function (e) {\n");
        out.append("    if (e.target && /^(INPUT|TEXTAREA|SELECT|BUTTON)$/.test(e.target.tagName)) return;\n");
        out.append("    var key = e.key;\n");
        out.append("    if (key === 'ArrowUp') { e.preventDefault(); moveActive(-1, 0, e.shiftKey); return; }\n");
        out.append("    if (key === 'ArrowDown') { e.preventDefault(); moveActive(1, 0, e.shiftKey); return; }\n");
        out.append("    if (key === 'ArrowLeft') { e.preventDefault(); moveActive(0, -1, e.shiftKey); return; }\n");
        out.append("    if (key === 'ArrowRight') { e.preventDefault(); moveActive(0, 1, e.shiftKey); return; }\n");
        out.append("    if ((e.ctrlKey || e.metaKey) && (key === 'c' || key === 'C')) {\n");
        out.append("      if (!selected.length) return;\n");
        out.append("      e.preventDefault();\n");
        out.append("      copySelection();\n");
        out.append("    }\n");
        out.append("  });\n");
        out.append("  document.addEventListener('copy', function (e) {\n");
        out.append("    if (!selected.length) return;\n");
        out.append("    copySelection(e);\n");
        out.append("  });\n");
        out.append("})();\n");
        out.append("</script>\n");
    }

    private static boolean hasSolidFill(org.apache.poi.ss.usermodel.CellStyle style) {
        if (style == null) {
            return false;
        }
        FillPatternType pattern = style.getFillPattern();
        return pattern != null && pattern != FillPatternType.NO_FILL;
    }

    private static int findMaxColumn(Sheet sheet, int firstRow, int lastRow, MergeIndex merges) {
        int maxCol = merges.maxColumn();
        for (int r = firstRow; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row != null && row.getLastCellNum() > maxCol) {
                maxCol = row.getLastCellNum() - 1;
            }
        }
        return Math.max(maxCol, 0);
    }

    static String escapeHtml(String text) {
        if (text == null || text.isEmpty()) {
            return "";
        }
        StringBuilder sb = new StringBuilder(text.length());
        for (int i = 0; i < text.length(); i++) {
            char ch = text.charAt(i);
            switch (ch) {
                case '&' -> sb.append("&amp;");
                case '<' -> sb.append("&lt;");
                case '>' -> sb.append("&gt;");
                case '"' -> sb.append("&quot;");
                case '\'' -> sb.append("&#39;");
                default -> sb.append(ch);
            }
        }
        return sb.toString();
    }

    /** シート内の結合セルを高速に参照する。 */
    private static final class MergeIndex {
        private final List<CellRangeAddress> regions;

        private MergeIndex(List<CellRangeAddress> regions) {
            this.regions = regions;
        }

        static MergeIndex of(Sheet sheet) {
            List<CellRangeAddress> list = new ArrayList<>(sheet.getNumMergedRegions());
            for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                list.add(sheet.getMergedRegion(i));
            }
            return new MergeIndex(list);
        }

        CellRangeAddress find(int row, int col) {
            for (CellRangeAddress region : regions) {
                if (region.isInRange(row, col)) {
                    return region;
                }
            }
            return null;
        }

        int maxColumn() {
            int max = -1;
            for (CellRangeAddress region : regions) {
                max = Math.max(max, region.getLastColumn());
            }
            return max;
        }
    }
}
