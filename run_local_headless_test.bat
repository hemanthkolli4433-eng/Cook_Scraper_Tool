@echo off
cd /d "%~dp0"
set HEADLESS=1
set WAIT_SECONDS=120
set MAX_RETRIES=3
set PAUSE_AFTER_SEARCH=5
set COOK_USE_PROFILE=0
python -m pip install -r requirements.txt
python Cook_Search_Tool.py
pause
