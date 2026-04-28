@echo off
echo Starting BOL Generator...

python --version >nul 2>&1
if errorlevel 1 (
    echo Python not found. Please install Python from https://www.python.org/downloads/
    echo Make sure to tick "Add Python to PATH" during installation.
    pause
    exit
)

if not exist "venv" (
    echo Setting up for first time...
    python -m venv venv
    call venv\Scripts\activate
    pip install -r requirements.txt
) else (
    call venv\Scripts\activate
)

start http://localhost:5001
python app.py
