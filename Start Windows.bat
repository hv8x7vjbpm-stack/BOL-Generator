@echo off
cd /d "%~dp0"
title BOL Generator - Jackson Pottery
echo ================================
echo    BOL Generator - Jackson Pottery
echo ================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found.
    echo.
    echo Please install Python from https://www.python.org/downloads/
    echo IMPORTANT: Tick "Add Python to PATH" during installation.
    echo.
    pause
    exit
)

echo Python found OK.
echo.

echo Removing old virtual environment...
if exist "venv" rmdir /s /q venv

echo Creating virtual environment...
python -m venv venv
if errorlevel 1 (
    echo ERROR: Failed to create virtual environment.
    pause
    exit
)

echo Activating virtual environment...
call venv\Scripts\activate
if errorlevel 1 (
    echo ERROR: Failed to activate virtual environment.
    pause
    exit
)

echo Installing dependencies...
pip install flask reportlab pypdf openpyxl pdfplumber
if errorlevel 1 (
    echo ERROR: Failed to install dependencies.
    pause
    exit
)

echo.
echo ================================
echo  Starting BOL Generator...
echo ================================
echo.
echo Browser opening at: http://localhost:5001
echo Press CTRL+C to stop the app.
echo.

timeout /t 2 /nobreak >nul
start "" http://localhost:5001
python app.py

echo.
echo App stopped. Press any key to close.
pause
