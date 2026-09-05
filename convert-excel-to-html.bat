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
set "JAR=target\convert-excel-to-html.jar"

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

if not exist "%JAR%" (
    where mvn >nul 2>&1
    if errorlevel 1 (
        echo エラー: %JAR% がありません。先に mvn package を実行するか、mvn を PATH に追加してください。
        pause
        exit /b 1
    )
    echo ビルド中... ^(mvn package^)
    call mvn -q -DskipTests package
    if errorlevel 1 (
        echo エラー: ビルドに失敗しました。
        pause
        exit /b 1
    )
    if not exist "%JAR%" (
        echo エラー: ビルド後も %JAR% が見つかりません。
        pause
        exit /b 1
    )
)

echo 入力: %INPUT%
if defined OUTPUT (
    echo 出力: %OUTPUT%
    java -jar "%JAR%" "%INPUT%" "%OUTPUT%"
) else (
    echo 出力: ^(未指定 → フォルダ名_excelhtml^)
    java -jar "%JAR%" "%INPUT%"
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
