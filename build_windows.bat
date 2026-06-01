@echo off
REM Local Windows build. Produces dist\AGM_Final_Set.exe + a release .zip.
REM
REM Prereqs: Python 3.13 on PATH with python-docx, openpyxl, pyinstaller installed:
REM   pip install -r requirements.txt pyinstaller
REM
REM Usage:
REM   build_windows.bat            (build + zip)
REM   build_windows.bat --clean    (also wipe build\ and dist\ first)

setlocal enabledelayedexpansion

if "%1"=="--clean" (
    echo ^>^> Cleaning build\ and dist\
    rmdir /s /q build 2>nul
    rmdir /s /q dist 2>nul
)

echo ^>^> Building demo Master_Input
python sample\build_demo.py
if errorlevel 1 goto :error

echo ^>^> Running PyInstaller
python -m PyInstaller agm_final_set.spec --noconfirm
if errorlevel 1 goto :error

if "%VERSION%"=="" set VERSION=dev
set ZIPNAME=AGM_Final_Set-windows-%VERSION%.zip
set STAGE=dist\AGM_Final_Set-%VERSION%-windows

echo ^>^> Staging release in %STAGE%
rmdir /s /q "%STAGE%" 2>nul
del /q "dist\%ZIPNAME%" 2>nul
mkdir "%STAGE%"
copy dist\AGM_Final_Set.exe "%STAGE%\" >nul
copy README.md "%STAGE%\" >nul
copy LICENSE "%STAGE%\" >nul
if exist USAGE.md copy USAGE.md "%STAGE%\" >nul
if exist PRIVACY.md copy PRIVACY.md "%STAGE%\" >nul
REM Both templates ship side-by-side — user renames whichever they want to use.
copy sample\Master_Input_DEMO.xlsx "%STAGE%\" >nul
copy sample\Master_Input_EMPTY.xlsx "%STAGE%\" >nul

echo ^>^> Zipping
powershell -Command "Compress-Archive -Path '%STAGE%' -DestinationPath 'dist\%ZIPNAME%' -Force"
if errorlevel 1 goto :error

echo ^>^> Done: dist\%ZIPNAME%
exit /b 0

:error
echo BUILD FAILED.
exit /b 1
