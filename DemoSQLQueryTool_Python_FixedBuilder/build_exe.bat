@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "LOG=%CD%\build_log.txt"

> "%LOG%" echo ============================================
>>"%LOG%" echo DemoSQLQueryTool build log
>>"%LOG%" echo Folder: %CD%
>>"%LOG%" echo Date: %DATE% %TIME%
>>"%LOG%" echo ============================================
>>"%LOG%" echo.

echo ============================================
echo DemoSQLQueryTool EXE Builder
echo ============================================
echo.
echo Build folder:
echo %CD%
echo.

rem ------------------------------------------------------------
rem 1. Check required files
rem ------------------------------------------------------------

if not exist "app.py" (
    echo ERROR: app.py was not found.
    >>"%LOG%" echo ERROR: app.py was not found.
    goto :build_error
)

if not exist "demo.sql" (
    echo ERROR: demo.sql was not found.
    >>"%LOG%" echo ERROR: demo.sql was not found.
    goto :build_error
)

echo [1/5] Required files found.
>>"%LOG%" echo [1/5] Required files found.

rem ------------------------------------------------------------
rem 2. Locate Python
rem ------------------------------------------------------------

set "PYTHON_CMD="

py -3.12 -c "import sys; print(sys.executable)" >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=py -3.12"
)

if not defined PYTHON_CMD (
    py -3 -c "import sys; print(sys.executable)" >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON_CMD=py -3"
    )
)

if not defined PYTHON_CMD (
    python -c "import sys; print(sys.executable)" >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON_CMD=python"
    )
)

if not defined PYTHON_CMD (
    echo ERROR: Python was not found.
    >>"%LOG%" echo ERROR: Python was not found.
    goto :python_error
)

echo [2/5] Python found:
%PYTHON_CMD% --version
>>"%LOG%" echo [2/5] Python command: %PYTHON_CMD%
%PYTHON_CMD% --version >>"%LOG%" 2>&1

rem ------------------------------------------------------------
rem 3. Install packages
rem ------------------------------------------------------------

echo.
echo [3/5] Installing required packages...
echo This may take several minutes the first time.

%PYTHON_CMD% -m ensurepip --upgrade >>"%LOG%" 2>&1
%PYTHON_CMD% -m pip install --upgrade pip setuptools wheel >>"%LOG%" 2>&1

%PYTHON_CMD% -m pip install ^
    "openpyxl>=3.1,<4" ^
    "pyinstaller>=6,<7" ^
    >>"%LOG%" 2>&1

if errorlevel 1 (
    echo ERROR: Package installation failed.
    goto :build_error
)

echo Packages installed successfully.
>>"%LOG%" echo Packages installed successfully.

rem ------------------------------------------------------------
rem 4. Clean old build output
rem ------------------------------------------------------------

echo.
echo [4/5] Removing old build files...

if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "DemoSQLQueryTool.spec" del /q "DemoSQLQueryTool.spec"

>>"%LOG%" echo Old build files removed.

rem ------------------------------------------------------------
rem 5. Build EXE
rem ------------------------------------------------------------

echo.
echo [5/5] Building DemoSQLQueryTool.exe...
echo Please wait. The console may appear inactive for a while.

%PYTHON_CMD% -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --name "DemoSQLQueryTool" ^
    --collect-all openpyxl ^
    --add-data "demo.sql;." ^
    "app.py" ^
    >>"%LOG%" 2>&1

if errorlevel 1 (
    echo ERROR: PyInstaller failed.
    goto :build_error
)

if not exist "dist\DemoSQLQueryTool.exe" (
    echo ERROR: PyInstaller finished, but EXE was not found.
    >>"%LOG%" echo ERROR: dist\DemoSQLQueryTool.exe was not found.
    goto :build_error
)

copy /Y "demo.sql" "dist\demo.sql" >>"%LOG%" 2>&1

echo.
echo ============================================
echo BUILD SUCCESSFUL
echo ============================================
echo.
echo EXE created at:
echo %CD%\dist\DemoSQLQueryTool.exe
echo.
echo demo.sql copied to:
echo %CD%\dist\demo.sql
echo.
echo The dist folder will now open.
echo.

>>"%LOG%" echo.
>>"%LOG%" echo BUILD SUCCESSFUL
>>"%LOG%" echo EXE: %CD%\dist\DemoSQLQueryTool.exe

start "" "%CD%\dist"
pause
exit /b 0

:python_error
echo.
echo Python was not found.
echo Install Python 3.12 64-bit first.
echo During installation, enable:
echo Add Python to PATH
echo.
echo Then run this BAT file again.
echo.
start "" notepad "%LOG%"
pause
exit /b 1

:build_error
echo.
echo ============================================
echo BUILD FAILED
echo ============================================
echo.
echo A log file was created:
echo %LOG%
echo.
echo The log will now open in Notepad.
echo Please copy the last red/error-looking lines to ChatGPT.
echo.
start "" notepad "%LOG%"
pause
exit /b 1
