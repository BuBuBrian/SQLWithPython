@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "PYTHON_CMD="

py -3.12 -c "import sys" >nul 2>&1
if not errorlevel 1 set "PYTHON_CMD=py -3.12"

if not defined PYTHON_CMD (
    py -3 -c "import sys" >nul 2>&1
    if not errorlevel 1 set "PYTHON_CMD=py -3"
)

if not defined PYTHON_CMD (
    python -c "import sys" >nul 2>&1
    if not errorlevel 1 set "PYTHON_CMD=python"
)

if not defined PYTHON_CMD (
    echo Python was not found.
    echo Install Python 3.12 64-bit and enable Add Python to PATH.
    pause
    exit /b 1
)

%PYTHON_CMD% -m pip install "openpyxl>=3.1,<4"
if errorlevel 1 (
    echo Failed to install openpyxl.
    pause
    exit /b 1
)

%PYTHON_CMD% app.py

if errorlevel 1 (
    echo.
    echo Python program exited with an error.
    pause
)
