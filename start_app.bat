@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo The application environment is missing.
  echo Run setup_app.bat first.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" -m streamlit run streamlit_app.py
