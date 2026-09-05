# convert-excel-to-html

Apache POI を使い、Excel（`.xlsx` / `.xls`）を HTML と TXT（タブ区切り）に変換する Maven プロジェクトです。

セルの表示値に加え、列幅・結合・スタイル（フォント / 色 / 配置 / 枠線）を Excel 風グリッドとして出力します。
あわせて groovy-excel2csv 相当の TXT（シート名・行番号付き TSV）も出力します。

## 必要環境

- Java 17 以上
- Apache Maven 3.x

## ビルド

```bash
mvn -q package
```

実行可能 JAR（依存込み）が `target/convert-excel-to-html.jar` にできます。

```bash
java -jar target/convert-excel-to-html.jar 入力フォルダ
```

## 使い方

### 単一ファイル

```bash
mvn -q exec:java "-Dexec.args=入力.xlsx 出力.html"
```

`出力.html` と同時に、同名の `出力.txt`（タブ区切り）も生成されます。

例:

```bash
mvn -q exec:java "-Dexec.args=sample.xlsx output.html"
# → output.html と output.txt
```

### フォルダ一括変換

フォルダ内の `.xlsx` / `.xls` をすべて変換し、各ファイルの HTML・TXT と `index.html` を生成します。  
インデックスにはシート名リンクがあり、`ファイル.html#sheet-0` のように直接そのシートを開けます。  
各 HTML / インデックスから、元の Excel のフルパスをコピーしたり、ダウンロードしたり、Excel で開けます。  
「ダウンロード」は元ファイルへの相対リンク（`download` 属性）です。  
「Excelで開く」は `ms-excel:ofe` で元ファイルを起動します。  
（Excel の一時ファイル `~$*.xlsx` は除外します）

```bash
# 出力先省略時は <親フォルダ>/<フォルダ名>_excelhtml
mvn -q exec:java "-Dexec.args=入力フォルダ"

# 出力先を指定
mvn -q exec:java "-Dexec.args=入力フォルダ 出力フォルダ"
```

例:

```bash
mvn -q exec:java "-Dexec.args=samples"
# → ../samples_excelhtml/*.html / *.txt と index.html（入力が samples の場合）
```

### ドラッグ＆ドロップ（bat）

プロジェクト直下の `convert-excel-to-html.bat` を使います。

1. bat のショートカットをデスクトップなど好きな場所に作成
2. 変換したいフォルダをショートカットへドラッグ＆ドロップ

出力先未指定時は、変換対象と同じ階層に `<フォルダ名>_excelhtml` を作成して出力します。  
（例: `C:\data\報告書` → `C:\data\報告書_excelhtml`）

```bat
convert-excel-to-html.bat "入力フォルダ"
convert-excel-to-html.bat "入力フォルダ" "出力フォルダ"
```

初回は必要に応じて `mvn package` で `target\convert-excel-to-html.jar` を作成します（bat が無い場合は自動ビルド）。  
起動は `java -jar` です。Java 17 以上が PATH に必要です。

### サンプル Excel の作成

動作確認用の `sample.xlsx` を生成します。

```bash
mvn -q exec:java "-Dexec.mainClass=com.example.excelhtml.CreateSampleExcel" "-Dexec.args=sample.xlsx"
```

## 変換の挙動

### HTML

| セル種別 | 出力 |
|----------|------|
| 文字列 / 数値 / 真偽値 | 表示値 |
| 日付 | Excel の表示書式に従う（例: `yyyy-mm-dd`） |
| 数式 | 計算結果（数式文字列ではない） |
| 空白 | 空の `<td>` |

HTML 特殊文字（`&`, `<`, `>`, `"`）はエスケープされます。

複数シートがある場合は、上部のタブで表示／非表示を切り替えます。初期表示は Excel のアクティブシートです。

各シートは Excel 風グリッドとして出力します。

- 列見出し: `A`, `B`, `C` …
- 行見出し: `1`, `2`, `3` …（Excel の行番号）
- 表示範囲: 使用範囲の最終行・最終列のさらに +1 行・+1 列まで（空の余白）
- 列幅: Excel の列幅をピクセル換算して反映
- セル結合: `rowspan` / `colspan` で反映
- セルスタイル: フォント名・サイズ・太字・斜体・文字色・背景色・配置・折り返し・はみ出し・インデント・枠線
- 行の高さ: Excel の行高を反映
- ビューポート内で縦横スクロール可能
- スクロール時も見出し行・列は sticky で固定
- セル操作:
  - クリック / 矢印キーで移動
  - マウスドラッグ / `Shift+クリック` / `Shift+矢印` で複数セル選択
  - 列見出し（A,B,C…）クリックで列全選択、行見出し（1,2,3…）クリックで行全選択、左上角クリックで全選択
  - `Ctrl+C`（Mac は `Cmd+C`）でコピー（複数セルは TSV）

### TXT（groovy-excel2csv 相当）

HTML と同じベース名で `.txt` を出力します（UTF-8、改行は CRLF、区切りはタブ）。

| 項目 | 内容 |
|------|------|
| 行形式 | `[シート名]` + `R00001`（行番号）+ セル値… |
| 図形 | `[シート名]` + `A00001`（アンカー行）+ テキスト |
| 非表示シート | スキップ |
| 空行 | スキップ |
| セル内改行 | 除去 |
| 取消線 | 該当文字を除去 |
| 日付 | `yyyy/MM/dd` |
| 最大列 | 既定 100（シート名「リクエスト」は 37、「レスポンス」は 35） |

## プロジェクト構成

```
src/main/java/com/example/excelhtml/
├── ExcelToHtmlApp.java              # CLI 入口（単一 / フォルダ）
├── ExcelToHtmlConverter.java        # Workbook → HTML
├── ExcelToTxtConverter.java         # Workbook → TXT（TSV）
├── FolderExcelToHtmlConverter.java  # フォルダ一括 + index.html
├── CellValueFormatter.java          # HTML 用セル表示値
├── TxtCellValueFormatter.java       # TXT 用セル値（取消線除去など）
├── ExcelShapeTextExtractor.java     # オートシェイプテキスト抽出
├── CellStyleCssRegistry.java        # CellStyle → CSS
└── CreateSampleExcel.java           # サンプル xlsx 生成
```

## 依存関係

- [Apache POI](https://poi.apache.org/) `poi-ooxml` 5.5.1

## 今後の予定（未実装）

- 条件付き書式・テーマ色の完全再現
- 図形・画像
