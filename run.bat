@echo off
:: ============================================================
:: GC Weekly Stats Dashboard Generator
:: Drop your 3 Excel files in this folder, then double-click
:: this file to regenerate all dashboards and push to GitHub.
:: ============================================================

echo.
echo ========================================
echo   GC Weekly Stats Dashboard
echo ========================================
echo.

:: Change to the folder where this script lives
cd /d "%~dp0"

:: Check that Python is installed
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not in PATH.
    echo         Please install from: https://python.org
    pause
    exit /b 1
)

:: Check that required Excel files are present
if not exist "Employee_Stats.xlsx" (
    echo [ERROR] Employee_Stats.xlsx not found in this folder.
    echo         Please copy your weekly Excel exports here and try again.
    pause
    exit /b 1
)
if not exist "Employee_Return_Stats.xlsx" (
    echo [ERROR] Employee_Return_Stats.xlsx not found in this folder.
    pause
    exit /b 1
)
if not exist "All_Salons.xlsx" (
    echo [ERROR] All_Salons.xlsx not found in this folder.
    pause
    exit /b 1
)

:: Run the data generator
echo [INFO] Generating dashboards at %date% %time%
echo.
python generate_data.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Dashboard generation failed. Check output above.
    pause
    exit /b 1
)

:: Push to GitHub
echo.
echo [INFO] Pushing to GitHub...
git add index.html 3750.html 3800.html 3826.html 4216.html
git commit -m "Update dashboard - %date%"
git push origin main

if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo   Done! Dashboard updated on GitHub.
    echo ========================================
    echo.
    echo   Owner view : https://jmd0515.github.io/GC-Weekly-Stats/
    echo   County Line: https://jmd0515.github.io/GC-Weekly-Stats/3750.html
    echo   Braden River: https://jmd0515.github.io/GC-Weekly-Stats/3800.html
    echo   Kings Crossing: https://jmd0515.github.io/GC-Weekly-Stats/3826.html
    echo   North River Ranch: https://jmd0515.github.io/GC-Weekly-Stats/4216.html
    echo.
    echo Opening owner view for review...
    start "" "index.html"
) else (
    echo.
    echo [WARNING] Push failed. Check your internet connection or GitHub credentials.
    echo           The HTML files were generated locally - you can open index.html to preview.
    start "" "index.html"
    pause
)
