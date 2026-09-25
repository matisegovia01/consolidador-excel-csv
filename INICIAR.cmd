@echo off
setlocal
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" gui.py
) else (
  python gui.py
)
if errorlevel 1 (
  echo.
  echo No se pudo iniciar. Consulte la instalacion en README.md.
  pause
)
