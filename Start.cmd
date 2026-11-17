@echo off
cd /d "%~dp0"
py -m pip install -r requirements.txt && py -X utf8 main.py
pause
