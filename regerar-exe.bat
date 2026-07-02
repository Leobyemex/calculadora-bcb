@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo === RECOMPILACAO (formato pasta/onedir) ===
echo.

if not exist calculadora_bcb.py (
    echo ERRO: calculadora_bcb.py nao encontrado nesta pasta!
    pause
    exit /b 1
)

echo --- Versao ---
findstr /c:"APP_VERSION  =" calculadora_bcb.py
echo.

echo --- Limpando caches e builds antigos ---
if exist __pycache__ rmdir /s /q __pycache__
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist CalculadoraBCB.spec del CalculadoraBCB.spec

echo --- Dependencias ---
python -m pip install --upgrade pip 2>nul
python -m pip install pyinstaller openpyxl reportlab --upgrade
echo.

echo --- Compilando (formato pasta) - 3 a 5 min ---
python -m PyInstaller --onedir --windowed --name CalculadoraBCB --noconfirm ^
  --collect-all openpyxl --collect-all reportlab ^
  calculadora_bcb.py
if errorlevel 1 (
    echo ERRO na compilacao!
    pause
    exit /b 1
)

if exist build rmdir /s /q build
if exist CalculadoraBCB.spec del CalculadoraBCB.spec

echo.
echo ===== SUCESSO! =====
echo  Pasta do app:  dist\CalculadoraBCB\
echo  (rode o CalculadoraBCB.exe de DENTRO dessa pasta - ele precisa da _internal)
echo.
pause
