@echo off
setlocal
cd /d "%~dp0"
python -m PyInstaller --noconfirm --clean --onedir --windowed --name ConsolidadorExcel --collect-all customtkinter --hidden-import xlrd --exclude-module scipy --exclude-module matplotlib --exclude-module IPython --exclude-module pytest --exclude-module sympy --exclude-module numba --exclude-module tables --exclude-module pyarrow gui.py
if errorlevel 1 (
  echo Error al compilar. Instale requirements-build.txt y vuelva a intentar.
  pause
  exit /b 1
)
echo Ejecutable creado en dist\ConsolidadorExcel\ConsolidadorExcel.exe
pause
