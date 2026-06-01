@echo off
REM Double-click this file to generate the AGM Final Set.
REM Requires Python 3.10+ with python-docx and openpyxl installed.

cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo.
    echo ERROR: Python is not installed or not in PATH.
    echo Install Python from https://www.python.org/downloads/ and tick "Add to PATH" during install.
    echo.
    pause
    exit /b 1
)

python -c "import docx, openpyxl" >nul 2>nul
if errorlevel 1 (
    echo Installing required libraries...
    python -m pip install --quiet python-docx openpyxl
    if errorlevel 1 (
        echo.
        echo ERROR: Failed to install python-docx and openpyxl.
        echo Run this manually:  pip install python-docx openpyxl
        echo.
        pause
        exit /b 1
    )
)

python "generate_final_set.py"
echo.
pause
