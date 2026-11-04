
@echo off
chcp 65001 >nul
setlocal EnableExtensions DisableDelayedExpansion

REM =================================================
REM PostgreSQL CSV Batch Import
REM 將指定資料夾內所有 CSV 檔案批次匯入 PostgreSQL
REM 張耘馨
REM 修正版：2026.11.03
REM =================================================

REM ====== 設定參數 ======

set "CSV_DIR=C:\Users\Students\Desktop\D&TMOK KLNG_2026_04\filter_result_SOB-Decode\16型檔\CSV_40欄位"

set "DB_HOST=localhost"
set "DB_PORT=5432"
set "DB_USER=postgres"
set "DB_NAME=postgres"
set "TABLE_NAME=public.ship_logtt"
set "DB_PASSWORD=30223044"

REM ====== PostgreSQL 自動登入 ======

set "PGHOST=%DB_HOST%"
set "PGPORT=%DB_PORT%"
set "PGUSER=%DB_USER%"
set "PGPASSWORD=%DB_PASSWORD%"

REM ====== 檢查資料夾 ======

if not exist "%CSV_DIR%\" (
    echo [ERROR] 找不到 CSV 資料夾
    pause
    exit /b 1
)

REM ====== 檢查 psql ======

where psql >nul 2>&1

if errorlevel 1 (
    echo [ERROR] 找不到 psql.exe
    pause
    exit /b 1
)

REM ====== 測試 PostgreSQL 連線 ======

echo.
echo [INFO] 測試 PostgreSQL 自動登入...

psql -X -w -v ON_ERROR_STOP=1 -U "%DB_USER%" -d "%DB_NAME%" -c "SELECT current_database(), current_user;"

if errorlevel 1 (
    echo [ERROR] PostgreSQL 登入失敗
    echo 請確認帳號、密碼及連線設定
    pause
    exit /b 1
)

echo [SUCCESS] PostgreSQL 登入成功

REM ====== 開始匯入 ======

echo.
echo =========================================
echo PostgreSQL CSV 批次匯入
echo =========================================
echo 資料庫：%DB_NAME%
echo 資料表：%TABLE_NAME%
echo CSV資料夾：%CSV_DIR%
echo =========================================
echo.

set /a SUCCESS=0
set /a FAILED=0

for %%F in ("%CSV_DIR%\*.csv") do (
    if exist "%%~fF" (
        call :IMPORT_FILE "%%~fF"
    )
)

REM ====== 匯入結果 ======

echo.
echo =========================================
echo 批次匯入完成
echo =========================================
echo 成功檔案數：%SUCCESS%
echo 失敗檔案數：%FAILED%
echo =========================================

pause
exit /b

REM ====== 單一 CSV 匯入 ======

:IMPORT_FILE

echo.
echo [IMPORT] %~nx1

psql -X -w -v ON_ERROR_STOP=1 -U "%DB_USER%" -d "%DB_NAME%" -c "\copy %TABLE_NAME% FROM '%~1' WITH (FORMAT csv, HEADER true, ENCODING 'UTF8')"

if errorlevel 1 (
    echo [FAILED] %~nx1
    set /a FAILED+=1
) else (
    echo [SUCCESS] %~nx1
    set /a SUCCESS+=1
)

exit /b
