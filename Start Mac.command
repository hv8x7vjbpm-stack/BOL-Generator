#!/bin/bash
cd "$(dirname "$0")"
echo "Starting BOL Generator..."

if ! command -v python3 &> /dev/null; then
    echo "Python not found. Install from https://www.python.org/downloads/"
    read -p "Press enter to exit"
    exit
fi

if [ ! -d "venv" ]; then
    echo "Setting up for first time..."
    python3 -m venv venv
    source venv/bin/activate
    pip install -r requirements.txt
else
    source venv/bin/activate
fi

open http://localhost:5001
python app.py
