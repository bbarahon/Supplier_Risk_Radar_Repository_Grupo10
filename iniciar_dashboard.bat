@echo off
cd /d "%~dp0"
python -m pip install -r requirements.txt --quiet
python dashboard.py
pause
