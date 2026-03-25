@echo off
setlocal

REM Build Windows executable with bundled assets and seed data
python -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
  echo PyInstaller no esta instalado en este entorno. Instalando...
  python -m pip install pyinstaller
  if errorlevel 1 (
    echo.
    echo No se pudo instalar PyInstaller.
    exit /b 1
  )
)

python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --onefile ^
  --name CalibracionVC ^
  --icon "img\icono.ico" ^
  --hidden-import ui_shared ^
  --collect-all customtkinter ^
  --collect-all reportlab ^
  --add-data "img;img" ^
  --add-data "data;data" ^
  --add-data "Documentos PDF.py;Documentos PDF.py" ^
  app.py

if errorlevel 1 (
  echo.
  echo Build failed.
  exit /b 1
)

echo.
echo Build completed. Executable path:
echo dist\CalibracionVC.exe
