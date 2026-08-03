@echo off
title CRM Counter - Setup
mode con: cols=70 lines=30
cd /d "%~dp0"

echo.
echo  ============================================
echo   CRM Counter - First Time Setup
echo  ============================================
echo.

:: Check Python is installed and on PATH
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python not found on this machine.
    echo.
    echo  Please install Python 3.12 or newer from:
    echo  https://www.python.org/downloads/
    echo.
    echo  IMPORTANT: During install, tick the box that says
    echo  "Add Python to PATH" before clicking Install Now.
    echo.
    echo  After installing Python, run this setup again.
    echo.
    pause
    exit /b 1
)

:: Show which Python version was found
echo  [OK] Python found:
python --version
echo.

:: Check pip is available
pip --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] pip not found. Try reinstalling Python and
    echo  make sure "Add Python to PATH" is ticked.
    echo.
    pause
    exit /b 1
)

echo  [OK] pip found.
echo.
echo  Installing required packages...
echo  This may take a minute on first run.
echo.

:: Install all dependencies
pip install colorama openpyxl reportlab pycaw comtypes keyboard

if errorlevel 1 (
    echo.
    echo  [ERROR] One or more packages failed to install.
    echo  Check your internet connection and try again.
    echo.
    pause
    exit /b 1
)

echo.
echo  ============================================
echo   Setup Complete!
echo  ============================================
echo.
echo  You can now run CRM Counter by double-clicking:
echo  CRM_Counter.bat
echo.
echo  This setup only needs to be run once per machine.
echo.
pause
