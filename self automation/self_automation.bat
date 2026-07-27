@echo off
cd /d %~dp0

if not exist venv (
    echo Creating virtual environment...
    python -m venv venv
)

call venv\Scripts\activate

pip install -r requirements.txt
playwright install chromium

start /b python app.py
timeout /t 2
start http://127.0.0.1:5000
pause

