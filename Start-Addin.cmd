@echo off
cd /d "%~dp0"
where npm >nul 2>nul || (echo Install Node.js 22 LTS first. & pause & exit /b 1)
py -m pip install -r requirements.txt || (pause & exit /b 1)
cd addin
call npm install && call npm start
pause
