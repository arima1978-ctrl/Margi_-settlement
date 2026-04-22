@echo off
REM Double-click launcher for monthly settlement generation.
REM 引数なしで実行すると翌月分を対象に、Y キー確認で全サービス生成します。

setlocal
cd /d "%~dp0"

REM UTF-8 で実行 (日本語表示が崩れないように)
chcp 65001 > nul

echo.
echo === 月次清算書 一括生成ツール ===
echo.

set /p MONTH=対象月を入力 (空のまま Enter で翌月、例: 2026-05):

if "%MONTH%"=="" (
    python scripts\run_monthly.py
) else (
    python scripts\run_monthly.py --month %MONTH%
)

echo.
pause
endlocal
