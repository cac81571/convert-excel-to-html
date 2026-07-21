@echo off
setlocal EnableExtensions
chcp 65001 >nul

rem この bat（またはそのショートカット）に変換対象フォルダをドラッグ＆ドロップして実行します。
rem 第2引数で出力先を指定できます。省略時は「<親>/<フォルダ名>_excelhtml」に出力します。

cd /d "%~dp0"

if "%~1"=="" (
    echo.
    echo 使い方:
    echo   1. この bat のショートカットを作る
    echo   2. 変換したいフォルダをショートカットへドラッグ＆ドロップ
    echo.
    echo または:
    echo   %~nx0 "入力フォルダ"
    echo   %~nx0 "入力フォルダ" "出力フォルダ"
    echo.
    pause
    exit /b 1
)

set "INPUT=%~1"
set "OUTPUT=%~2"

if not exist "%INPUT%\" (
    echo エラー: フォルダではありません: %INPUT%
    echo フォルダをドラッグ＆ドロップしてください。
    pause
    exit /b 1
)

where java >nul 2>&1
if errorlevel 1 (
    echo エラー: Java が見つかりません。Java 17 以上をインストールし、PATH を通してください。
    pause
    exit /b 1
)

if not exist "target\classes\com\example\excelhtml\ExcelToHtmlApp.class" (
    where mvn >nul 2>&1
    if errorlevel 1 (
        echo エラー: ビルド成果物がありません。Maven で package するか、mvn を PATH に追加してください。
        pause
        exit /b 1
    )
    echo ビルド中...
    call mvn -q -DskipTests package
    if errorlevel 1 (
        echo エラー: ビルドに失敗しました。
        pause
        exit /b 1
    )
)

if not exist "cp.txt" (
    where mvn >nul 2>&1
    if errorlevel 1 (
        echo エラー: cp.txt がありません。mvn dependency:build-classpath を実行してください。
        pause
        exit /b 1
    )
    echo クラスパスを生成中...
    call mvn -q dependency:build-classpath -Dmdep.outputFile=cp.txt
    if errorlevel 1 (
        echo エラー: クラスパスの生成に失敗しました。
        pause
        exit /b 1
    )
)

set "CP="
for /f "usebackq delims=" %%I in ("cp.txt") do set "CP=%%I"

echo 入力: %INPUT%
if defined OUTPUT (
    echo 出力: %OUTPUT%
    java -cp "target\classes;%CP%" com.example.excelhtml.ExcelToHtmlApp "%INPUT%" "%OUTPUT%"
) else (
    echo 出力: ^(未指定 → フォルダ名_excelhtml^)
    java -cp "target\classes;%CP%" com.example.excelhtml.ExcelToHtmlApp "%INPUT%"
)

set "EXITCODE=%ERRORLEVEL%"
echo.
if not "%EXITCODE%"=="0" (
    echo 変換に失敗しました。^(exit %EXITCODE%^)
) else (
    echo 完了しました。
)
pause
exit /b %EXITCODE%
