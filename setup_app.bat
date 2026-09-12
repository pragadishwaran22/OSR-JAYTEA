@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" python -m venv .venv
".venv\Scripts\python.exe" -m pip install -r requirements.txt
echo.
echo Setup completed. Double-click start_app.bat to open the application.
pause
