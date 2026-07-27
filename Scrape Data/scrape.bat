@echo off
cd /d %~dp0

if not exist venv (
    echo Creating virtual environment...
    python -m venv venv
)

call venv\Scripts\activate

pip install -r requirement.txt
playwright install chromium

python scraper_ui.py

pause
