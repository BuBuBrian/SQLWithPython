@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_COMMAND=py"
) else (
    set "PYTHON_COMMAND=python"
)

echo 安裝／確認 openpyxl...
%PYTHON_COMMAND% -m pip install "openpyxl>=3.1,<4"

echo.
echo 啟動 Python UI...
%PYTHON_COMMAND% app.py

if errorlevel 1 (
    echo.
    echo 執行失敗，請查看上方錯誤訊息。
    pause
)
