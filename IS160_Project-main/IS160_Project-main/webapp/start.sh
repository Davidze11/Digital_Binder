#!/bin/bash

echo "========================================"
echo "PlantAI - Plant Disease Detection"
echo "========================================"
echo ""

echo "[1/3] Checking Python installation..."
python3 --version
if [ $? -ne 0 ]; then
    echo "ERROR: Python is not installed or not in PATH"
    exit 1
fi
echo ""

echo "[2/3] Installing dependencies..."
cd "$(dirname "$0")"
pip3 install -r requirements.txt
if [ $? -ne 0 ]; then
    echo "ERROR: Failed to install dependencies"
    exit 1
fi
echo ""

echo "[3/3] Starting PlantAI web application..."
echo ""
echo "========================================"
echo "Access the application at:"
echo "http://localhost:5000"
echo "========================================"
echo ""
echo "Press Ctrl+C to stop the server"
echo ""

python3 app.py

