@echo off
setlocal
cd /d "%~dp0"

python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name FormatearPeticion ^
  --add-data "EJEMPLO PETICIÓN ELECTRICIDAD.xlsx;." ^
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
