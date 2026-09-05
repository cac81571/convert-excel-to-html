package com.example.excelhtml;

import java.nio.file.Files;
import java.nio.file.Path;

/**
 * 使い方:
 * <pre>
 *   # 単一ファイル（HTML と TXT を同時出力）
 *   mvn -q exec:java "-Dexec.args=入力.xlsx 出力.html"
 *
 *   # フォルダ一括（出力先省略時は &lt;親&gt;/&lt;フォルダ名&gt;_excelhtml）
 *   mvn -q exec:java "-Dexec.args=入力フォルダ"
 *   mvn -q exec:java "-Dexec.args=入力フォルダ 出力フォルダ"
 * </pre>
 */
public final class ExcelToHtmlApp {

    public static void main(String[] args) throws Exception {
        if (args.length < 1) {
            printUsage();
            System.exit(1);
        }

        Path input = Path.of(args[0]).toAbsolutePath().normalize();

        if (Files.isDirectory(input)) {
            Path outputDir = args.length >= 2
                    ? Path.of(args[1]).toAbsolutePath().normalize()
                    : defaultOutputDir(input);
            Path index = new FolderExcelToHtmlConverter().convertAll(input, outputDir);
            System.out.println("Index: " + index.toAbsolutePath());
            return;
        }

        if (args.length < 2) {
            printUsage();
            System.exit(1);
        }

        Path output = Path.of(args[1]).toAbsolutePath().normalize();
        Path outputParent = output.getParent();
        if (outputParent != null) {
            Files.createDirectories(outputParent);
        }
        new ExcelToHtmlConverter().convert(input, output);
        System.out.println("Wrote: " + output);

        Path txtPath = ExcelToTxtConverter.toTxtPath(output);
        new ExcelToTxtConverter(true, ExcelToTxtConverter.DEFAULT_MAX_COL,
                ExcelToTxtConverter.defaultMaxColBySheet()).convert(input, txtPath);
        System.out.println("Wrote: " + txtPath);
    }

    /** 例: C:/data/samples → C:/data/samples_excelhtml */
    static Path defaultOutputDir(Path inputDir) {
        String name = inputDir.getFileName() != null ? inputDir.getFileName().toString() : "excel";
        Path sibling = inputDir.resolveSibling(name + "_excelhtml");
        if (sibling.equals(inputDir)) {
            return inputDir.resolve(name + "_excelhtml");
        }
        return sibling;
    }

    private static void printUsage() {
        System.err.println("Usage:");
        System.err.println("  ExcelToHtmlApp <input.xlsx|.xls> <output.html>");
        System.err.println("    (also writes <output>.txt beside the HTML)");
        System.err.println("  ExcelToHtmlApp <input-folder> [output-folder]");
        System.err.println("    (default output-folder: <parent>/<folder-name>_excelhtml,");
        System.err.println("     writes *.html, *.txt, and index.html)");
    }
}
