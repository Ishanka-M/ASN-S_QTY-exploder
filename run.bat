@echo off
REM ===== ASN S_QTY Exploder - Local Launcher =====
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python not found. Install Python 3.9+ from https://python.org
    pause
    exit /b 1
)

echo Installing requirements...
python -m pip install -r requirements.txt

echo Starting app... (browser will open)
python -m streamlit run streamlit_app.py
pause
