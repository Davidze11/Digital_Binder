@echo off
echo ========================================
echo PlantAI - Plant Disease Detection
echo ========================================
echo.

echo [1/3] Checking Python installation...
python --version
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    pause
    exit /b 1
)
echo.

echo [2/3] Installing dependencies...
cd "%~dp0"
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies
    pause
    exit /b 1
)
echo.

echo [3/3] Starting PlantAI web application...
echo.
echo ========================================
echo Access the application at:
echo http://localhost:5000
echo ========================================
echo.
echo Press Ctrl+C to stop the server
echo.

python app.py

pause

