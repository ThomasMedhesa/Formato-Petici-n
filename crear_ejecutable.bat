@echo off
setlocal
cd /d "%~dp0"

set "PLANTILLA="
for %%f in ("EJEMPLO *ELECTRICIDAD.xlsx") do set "PLANTILLA=%%~ff"
if not defined PLANTILLA (
  echo.
  echo No se encontro la plantilla EJEMPLO PETICION ELECTRICIDAD.xlsx
  pause
  exit /b 1
)

python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name FormatearPeticion ^
  --icon "icono.ico" ^
  --add-data "%PLANTILLA%;." ^
  --add-data "icono.ico;." ^
  formatear_peticion.py

if errorlevel 1 (
  echo.
  echo No se pudo crear el ejecutable.
  pause
  exit /b 1
)

echo.
echo Ejecutable creado en: "%~dp0dist\FormatearPeticion.exe"
pause
