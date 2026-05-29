@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Procesando qPCR...
python procesar_placa.py "%~1" -o "%~dp0..\qPCR_resultados.xlsx"
if errorlevel 1 (
  echo.
  echo Si falla, instale Python y ejecute: pip install openpyxl xlrd
  pause
  exit /b 1
)
echo.
echo Abierto resultados en: %~dp0..\qPCR_resultados.xlsx
pause
