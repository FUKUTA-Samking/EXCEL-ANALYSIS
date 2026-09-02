@echo off
chcp 65001 >nul
setlocal

echo ============================================
echo  Excel 數據分析工具 - Windows 執行檔打包腳本
echo ============================================
echo.

REM 1. 檢查 Python 是否存在
where python >nul 2>nul
if errorlevel 1 (
    echo [錯誤] 找不到 Python，請先安裝 Python 3.9 以上版本
    echo         下載連結: https://www.python.org/downloads/
    echo         安裝時請務必勾選「Add python.exe to PATH」
    pause
    exit /b 1
)

echo [1/4] 建立虛擬環境 (.venv) ...
python -m venv .venv
call .venv\Scripts\activate.bat

echo.
echo [2/4] 安裝相依套件與 PyInstaller ...
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

echo.
echo [3/4] 開始打包 (可能需要 1-3 分鐘) ...
pyinstaller --noconfirm excel_analyzer.spec

echo.
echo [4/4] 打包完成！
echo.
if exist "dist\ExcelAnalyzer.exe" (
    echo 執行檔位置： dist\ExcelAnalyzer.exe
    echo 你可以把這個 .exe 複製到任何 Windows 電腦上直接雙擊執行，
    echo 不需要另外安裝 Python。
) else (
    echo [警告] 沒有找到預期的 exe 檔，請檢查上方是否有錯誤訊息。
)

echo.
pause
