@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ============================================
echo DemoSQLQueryTool Windows EXE 建置程式
echo ============================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_COMMAND=py"
) else (
    set "PYTHON_COMMAND=python"
)

echo [1/4] 檢查 Python...
%PYTHON_COMMAND% --version
if errorlevel 1 goto :python_error

echo.
echo [2/4] 安裝所需套件...
%PYTHON_COMMAND% -m pip install --upgrade pip
%PYTHON_COMMAND% -m pip install -r requirements.txt
if errorlevel 1 goto :build_error

echo.
echo [3/4] 使用 PyInstaller 建立 EXE...
%PYTHON_COMMAND% -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --name DemoSQLQueryTool ^
    --collect-all openpyxl ^
    --add-data "demo.sql;." ^
    app.py

if errorlevel 1 goto :build_error

echo.
echo [4/4] 將可替換的 demo.sql 複製到 EXE 旁邊...
copy /Y "demo.sql" "dist\demo.sql" >nul

echo.
echo ============================================
echo 建置完成
echo ============================================
echo EXE：
echo %CD%\dist\DemoSQLQueryTool.exe
echo.
echo SQL：
echo %CD%\dist\demo.sql
echo.
echo 可以把 DemoSQLQueryTool.exe 與 demo.sql 一起複製到其他資料夾。
echo 即使只保留 EXE，程式內仍有內嵌的預設 demo.sql。
echo.
start "" "%CD%\dist"
pause
exit /b 0

:python_error
echo.
echo 找不到 Python。
echo 請先安裝 Python 3.12 64-bit，安裝時勾選 Add Python to PATH。
pause
exit /b 1

:build_error
echo.
echo 建置失敗，請查看上方錯誤訊息。
pause
exit /b 1
