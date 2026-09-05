package com.example.excelhtml;

import java.io.IOException;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.FileTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;

/**
 * フォルダ内の Excel を一括変換し、HTML / TXT とインデックス HTML を生成する。
 */
public final class FolderExcelToHtmlConverter {

    private static final DateTimeFormatter MODIFIED_FMT =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm").withZone(ZoneId.systemDefault());

    private final ExcelToHtmlConverter converter = new ExcelToHtmlConverter();
    private final ExcelToTxtConverter txtConverter = new ExcelToTxtConverter(
            true, ExcelToTxtConverter.DEFAULT_MAX_COL, ExcelToTxtConverter.defaultMaxColBySheet());

    /**
     * @param inputDir  Excel が入っているフォルダ
     * @param outputDir HTML / TXT の出力先（存在しなければ作成）
     * @return 生成した index.html のパス
     */
    public Path convertAll(Path inputDir, Path outputDir) throws IOException {
        if (!Files.isDirectory(inputDir)) {
            throw new IllegalArgumentException("Not a directory: " + inputDir);
        }
        Files.createDirectories(outputDir);

        List<Path> excelFiles = listExcelFiles(inputDir);
        List<ConvertedFile> converted = new ArrayList<>();
        for (Path excel : excelFiles) {
            String htmlName = toHtmlFileName(excel.getFileName().toString());
            Path htmlPath = outputDir.resolve(htmlName);
            ExcelToHtmlConverter.ConversionResult result = converter.convert(excel, htmlPath);
            Path absoluteExcel = excel.toAbsolutePath().normalize();
            long size = Files.size(htmlPath);
            FileTime modified = Files.getLastModifiedTime(htmlPath);
            converted.add(new ConvertedFile(
                    excel.getFileName().toString(),
                    absoluteExcel.toString(),
                    htmlName,
                    result.downloadHref(),
                    result.downloadFileName(),
                    size,
                    modified,
                    result.sheetNames()));
            System.out.println("Wrote: " + htmlPath.toAbsolutePath());

            Path txtPath = ExcelToTxtConverter.toTxtPath(htmlPath);
            txtConverter.convert(excel, txtPath);
            System.out.println("Wrote: " + txtPath.toAbsolutePath());
        }

        Path indexPath = outputDir.resolve("index.html");
        writeIndex(indexPath, outputDir, converted);
        System.out.println("Wrote: " + indexPath.toAbsolutePath());
        return indexPath;
    }

    static List<Path> listExcelFiles(Path dir) throws IOException {
        List<Path> files = new ArrayList<>();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(dir)) {
            for (Path path : stream) {
                if (!Files.isRegularFile(path)) {
                    continue;
                }
                String name = path.getFileName().toString();
                if (name.startsWith("~$")) {
                    continue;
                }
                String lower = name.toLowerCase(Locale.ROOT);
                if (lower.contains(".readonly.")) {
                    continue;
                }
                if (lower.endsWith(".xlsx") || lower.endsWith(".xls")) {
                    files.add(path);
                }
            }
        }
        files.sort(Comparator.comparing(p -> p.getFileName().toString().toLowerCase(Locale.ROOT)));
        return files;
    }

    static String toHtmlFileName(String excelFileName) {
        int dot = excelFileName.lastIndexOf('.');
        String base = dot > 0 ? excelFileName.substring(0, dot) : excelFileName;
        return base + ".html";
    }

    private static void writeIndex(Path indexPath, Path outputDir, List<ConvertedFile> files) throws IOException {
        String folderLabel = outputDir.getFileName().toString();
        try (Writer out = Files.newBufferedWriter(indexPath, StandardCharsets.UTF_8)) {
            out.append("<!DOCTYPE html>\n");
            out.append("<html lang=\"ja\">\n");
            out.append("<head>\n");
            out.append("<meta charset=\"UTF-8\">\n");
            out.append("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n");
            out.append("<title>").append(escape(folderLabel)).append(" — Excel HTML</title>\n");
            out.append("<link rel=\"preconnect\" href=\"https://fonts.googleapis.com\">\n");
            out.append("<link rel=\"preconnect\" href=\"https://fonts.gstatic.com\" crossorigin>\n");
            out.append("<link href=\"https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;600&family=IBM+Plex+Mono:wght@400;500&display=swap\" rel=\"stylesheet\">\n");
            out.append("<style>\n");
            out.append(":root {\n");
            out.append("  --bg: #edf1f4; --ink: #1a2332; --muted: #5c6b7a;\n");
            out.append("  --line: #c8d2db; --accent: #0f6e56; --accent-soft: #d8f0e8; --row: #ffffffb8;\n");
            out.append("}\n");
            out.append("*, *::before, *::after { box-sizing: border-box; }\n");
            out.append("html, body { height: 100%; margin: 0; }\n");
            out.append("body {\n");
            out.append("  color: var(--ink); font: 13px/1.35 \"IBM Plex Sans\", \"Yu Gothic UI\", Meiryo, sans-serif;\n");
            out.append("  background: linear-gradient(180deg, #e4ece8 0%, var(--bg) 28%);\n");
            out.append("}\n");
            out.append(".wrap {\n");
            out.append("  display: flex; flex-direction: column; height: 100%; min-height: 0;\n");
            out.append("  width: min(1280px, calc(100% - 1.25rem)); margin: 0 auto; padding: 0.55rem 0 0.65rem;\n");
            out.append("}\n");
            out.append(".top {\n");
            out.append("  display: flex; flex-wrap: wrap; align-items: baseline; gap: 0.35rem 0.75rem;\n");
            out.append("  flex-shrink: 0; padding: 0 0.15rem 0.45rem; border-bottom: 1px solid var(--line);\n");
            out.append("}\n");
            out.append(".brand { margin: 0; font-size: 1.05rem; font-weight: 600; letter-spacing: -0.02em; }\n");
            out.append(".folder { color: var(--muted); font-size: 0.82rem; }\n");
            out.append(".count {\n");
            out.append("  margin-left: auto; padding: 0.1rem 0.45rem; border-radius: 4px;\n");
            out.append("  background: var(--accent-soft); color: var(--accent); font-size: 0.72rem; font-weight: 600;\n");
            out.append("}\n");
            out.append(".scroller { flex: 1; min-height: 0; overflow: auto; margin-top: 0.35rem; }\n");
            out.append("table.files {\n");
            out.append("  width: 100%; border-collapse: collapse; table-layout: fixed;\n");
            out.append("}\n");
            out.append("table.files th, table.files td {\n");
            out.append("  padding: 0.28rem 0.4rem; border-bottom: 1px solid var(--line); vertical-align: middle;\n");
            out.append("  text-align: left; font-weight: 400;\n");
            out.append("}\n");
            out.append("table.files thead th {\n");
            out.append("  position: sticky; top: 0; z-index: 1; background: #e8eef1;\n");
            out.append("  color: var(--muted); font-size: 0.68rem; font-weight: 600; text-transform: uppercase;\n");
            out.append("  letter-spacing: 0.04em; border-bottom: 1px solid var(--line);\n");
            out.append("}\n");
            out.append("table.files tbody tr:nth-child(even) { background: var(--row); }\n");
            out.append("table.files tbody tr:hover { background: var(--accent-soft); }\n");
            out.append(".col-file { width: 18%; }\n");
            out.append(".col-sheets { width: 42%; }\n");
            out.append(".col-meta { width: 14%; }\n");
            out.append(".col-src { width: 26%; }\n");
            out.append(".name {\n");
            out.append("  display: block; font-weight: 600; font-size: 0.9rem; color: inherit; text-decoration: none;\n");
            out.append("  overflow: hidden; text-overflow: ellipsis; white-space: nowrap;\n");
            out.append("}\n");
            out.append(".name:hover, .name:focus-visible { color: var(--accent); outline: none; }\n");
            out.append(".excel-name {\n");
            out.append("  display: block; color: var(--muted); font-size: 0.72rem;\n");
            out.append("  overflow: hidden; text-overflow: ellipsis; white-space: nowrap;\n");
            out.append("}\n");
            out.append(".sheets { display: flex; flex-wrap: wrap; gap: 0.2rem; }\n");
            out.append(".sheets a {\n");
            out.append("  display: inline-block; max-width: 9.5rem; padding: 0.08rem 0.35rem;\n");
            out.append("  border: 1px solid var(--line); border-radius: 3px; background: #fff;\n");
            out.append("  color: var(--ink); text-decoration: none; font-size: 0.75rem;\n");
            out.append("  overflow: hidden; text-overflow: ellipsis; white-space: nowrap;\n");
            out.append("}\n");
            out.append(".sheets a:hover, .sheets a:focus-visible {\n");
            out.append("  border-color: var(--accent); background: #fff; color: var(--accent); outline: none;\n");
            out.append("}\n");
            out.append(".meta {\n");
            out.append("  font-family: \"IBM Plex Mono\", ui-monospace, monospace; font-size: 0.7rem;\n");
            out.append("  color: var(--muted); white-space: nowrap;\n");
            out.append("}\n");
            out.append(".meta .size { color: var(--ink); }\n");
            out.append(".source-links { display: flex; flex-wrap: wrap; gap: 0.2rem; }\n");
            out.append(".source-links a, .source-links button {\n");
            out.append("  appearance: none; display: inline-flex; align-items: center; padding: 0.08rem 0.35rem;\n");
            out.append("  border: 1px solid var(--line); border-radius: 3px; background: #fff;\n");
            out.append("  color: var(--muted); text-decoration: none; font-size: 0.7rem; font-weight: 500;\n");
            out.append("  font-family: inherit; cursor: pointer; line-height: 1.3;\n");
            out.append("}\n");
            out.append(".source-links a:hover, .source-links a:focus-visible,\n");
            out.append(".source-links button:hover, .source-links button:focus-visible {\n");
            out.append("  border-color: var(--accent); color: var(--accent); outline: none;\n");
            out.append("}\n");
            out.append(".empty { margin: 1rem 0.15rem; color: var(--muted); font-size: 0.9rem; }\n");
            out.append("@media (max-width: 820px) {\n");
            out.append("  .col-file, .col-sheets, .col-meta, .col-src { width: auto; }\n");
            out.append("  table.files, table.files thead, table.files tbody, table.files th, table.files td, table.files tr {\n");
            out.append("    display: block;\n");
            out.append("  }\n");
            out.append("  table.files thead { display: none; }\n");
            out.append("  table.files tr {\n");
            out.append("    padding: 0.45rem 0.35rem; border-bottom: 1px solid var(--line);\n");
            out.append("  }\n");
            out.append("  table.files td { padding: 0.12rem 0; border: 0; }\n");
            out.append("  .meta { white-space: normal; }\n");
            out.append("}\n");
            out.append("</style>\n");
            out.append("</head>\n");
            out.append("<body>\n");
            out.append("<main class=\"wrap\">\n");
            out.append("  <header class=\"top\">\n");
            out.append("    <h1 class=\"brand\">Excel HTML</h1>\n");
            out.append("    <span class=\"folder\">").append(escape(folderLabel)).append("</span>\n");
            out.append("    <span class=\"count\">").append(String.valueOf(files.size())).append(" files</span>\n");
            out.append("  </header>\n");

            if (files.isEmpty()) {
                out.append("  <p class=\"empty\">変換された HTML はまだありません。</p>\n");
            } else {
                out.append("  <div class=\"scroller\">\n");
                out.append("  <table class=\"files\">\n");
                out.append("    <thead>\n");
                out.append("      <tr>\n");
                out.append("        <th class=\"col-file\">File</th>\n");
                out.append("        <th class=\"col-sheets\">Sheets</th>\n");
                out.append("        <th class=\"col-meta\">Size / Updated</th>\n");
                out.append("        <th class=\"col-src\">Excel</th>\n");
                out.append("      </tr>\n");
                out.append("    </thead>\n");
                out.append("    <tbody>\n");
                for (ConvertedFile file : files) {
                    String baseName = stripHtmlExt(file.htmlName());
                    String openEditUri = ExcelToHtmlConverter.toExcelOpenEditUri(Path.of(file.excelPath()));
                    out.append("      <tr>\n");
                    out.append("        <td class=\"col-file\">\n");
                    out.append("          <a class=\"name\" href=\"").append(escape(file.htmlName())).append("\" title=\"")
                            .append(escape(baseName)).append("\">").append(escape(baseName)).append("</a>\n");
                    out.append("          <span class=\"excel-name\" title=\"").append(escape(file.excelName())).append("\">")
                            .append(escape(file.excelName())).append("</span>\n");
                    out.append("        </td>\n");
                    out.append("        <td class=\"col-sheets\"><div class=\"sheets\">\n");
                    for (int i = 0; i < file.sheetNames().size(); i++) {
                        String sheetName = file.sheetNames().get(i);
                        out.append("          <a href=\"").append(escape(file.htmlName()))
                                .append("#sheet-").append(String.valueOf(i)).append("\" title=\"")
                                .append(escape(sheetName)).append("\">")
                                .append(escape(sheetName)).append("</a>\n");
                    }
                    out.append("        </div></td>\n");
                    out.append("        <td class=\"col-meta\"><div class=\"meta\">\n");
                    out.append("          <span class=\"size\">").append(formatSize(file.size())).append("</span><br>\n");
                    out.append("          ").append(MODIFIED_FMT.format(file.modified().toInstant())).append("\n");
                    out.append("        </div></td>\n");
                    out.append("        <td class=\"col-src\">\n");
                    out.append("          <div class=\"source-links\" data-excel-path=\"")
                            .append(escape(file.excelPath())).append("\">\n");
                    out.append("            <button type=\"button\" data-copy-excel-path title=\"")
                            .append(escape(file.excelPath())).append("\">パスをコピー</button>\n");
                    if (file.downloadHref() != null && !file.downloadHref().isBlank()) {
                        String dlName = file.downloadFileName() != null ? file.downloadFileName() : file.excelName();
                        out.append("            <a href=\"").append(escape(file.downloadHref()))
                                .append("\" download=\"").append(escape(dlName))
                                .append("\" title=\"元の Excel をダウンロード\">ダウンロード</a>\n");
                    }
                    out.append("            <a href=\"").append(escape(openEditUri))
                            .append("\" title=\"元ファイルを Excel で開く\">Excelで開く</a>\n");
                    out.append("          </div>\n");
                    out.append("        </td>\n");
                    out.append("      </tr>\n");
                }
                out.append("    </tbody>\n");
                out.append("  </table>\n");
                out.append("  </div>\n");
                out.append("  <script>\n");
                out.append("  (function () {\n");
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
                out.append("    document.querySelectorAll('[data-copy-excel-path]').forEach(function (btn) {\n");
                out.append("      btn.addEventListener('click', function () {\n");
                out.append("        var wrap = btn.closest('[data-excel-path]');\n");
                out.append("        var path = wrap ? wrap.getAttribute('data-excel-path') : '';\n");
                out.append("        var label = btn.textContent;\n");
                out.append("        copyPath(path).then(function () {\n");
                out.append("          btn.textContent = 'コピー済';\n");
                out.append("          setTimeout(function () { btn.textContent = label; }, 1200);\n");
                out.append("        }).catch(function () {\n");
                out.append("          btn.textContent = '失敗';\n");
                out.append("          setTimeout(function () { btn.textContent = label; }, 1200);\n");
                out.append("        });\n");
                out.append("      });\n");
                out.append("    });\n");
                out.append("  })();\n");
                out.append("  </script>\n");
            }

            out.append("</main>\n");
            out.append("</body>\n");
            out.append("</html>\n");
        }
    }

    static String stripHtmlExt(String htmlName) {
        if (htmlName.toLowerCase(Locale.ROOT).endsWith(".html")) {
            return htmlName.substring(0, htmlName.length() - 5);
        }
        return htmlName;
    }

    static String formatSize(long bytes) {
        if (bytes < 1024) {
            return bytes + " B";
        }
        if (bytes < 1024 * 1024) {
            return String.format(Locale.ROOT, "%.0f KB", bytes / 1024.0);
        }
        return String.format(Locale.ROOT, "%.1f MB", bytes / (1024.0 * 1024.0));
    }

    private static String escape(String text) {
        return ExcelToHtmlConverter.escapeHtml(text);
    }

    private record ConvertedFile(
            String excelName,
            String excelPath,
            String htmlName,
            String downloadHref,
            String downloadFileName,
            long size,
            FileTime modified,
            List<String> sheetNames) {
    }
}
