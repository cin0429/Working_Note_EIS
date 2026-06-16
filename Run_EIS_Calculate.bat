@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

cd /d "%~dp0"

set "SCRIPT=monthly_eis_calculator.py"
set "EIS=EIS.xlsx"
set "OUTPUT=EIS_calculated.xlsx"
set "NOTES="
set "COUNT=0"

if not exist "%SCRIPT%" (
    echo [ERROR] 找不到 %SCRIPT%
    echo 請確認 BAT 和 monthly_eis_calculator.py 放在同一個資料夾。
    pause
    exit /b 1
)

if not exist "%EIS%" (
    echo [ERROR] 找不到 %EIS%
    echo 請確認 EIS.xlsx 放在同一個資料夾。
    pause
    exit /b 1
)

for %%F in (WorkingNote_*.xlsx) do (
    if /I not "%%~nxF"=="%OUTPUT%" (
        set "NOTES=!NOTES! "%%F""
        set /a COUNT+=1
    )
)

if "%COUNT%"=="0" (
    echo [ERROR] 找不到任何 WorkingNote_*.xlsx 檔案
    echo.
    echo 請確認檔名格式，例如：
    echo WorkingNote_A.xlsx
    echo WorkingNote_B.xlsx
    echo WorkingNote_Cindy.xlsx
    pause
    exit /b 1
)

echo ========================================
echo Monthly EIS Calculator
echo ========================================
echo 找到 %COUNT% 個 WorkingNote 檔案
echo 開始計算...
echo.

python "%SCRIPT%" --eis "%EIS%" --working-notes %NOTES% --output "%OUTPUT%"

if errorlevel 1 (
    echo.
    echo [ERROR] 執行失敗
    echo 請確認是否已安裝 openpyxl：
    echo pip install openpyxl
    pause
    exit /b 1
)

echo.
echo ========================================
echo 完成！已產生 %OUTPUT%
echo ========================================
pause
