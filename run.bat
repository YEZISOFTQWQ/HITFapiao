@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  powershell -NoProfile -ExecutionPolicy Bypass -File ".\setup.ps1"
  if errorlevel 1 exit /b 1
)
".venv\Scripts\python.exe" ".\app.py"
