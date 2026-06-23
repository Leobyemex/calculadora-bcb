@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo === VERIFICACAO E RECOMPILACAO COMPLETA ===
echo.
echo Pasta atual: %CD%
echo.

REM 1) Confirmar que o arquivo esta na pasta
if not exist calculadora_bcb.py (
    echo ERRO: calculadora_bcb.py nao encontrado nesta pasta!
    pause
    exit /b 1
)

echo --- Data do calculadora_bcb.py ---
dir calculadora_bcb.py | findstr calculadora_bcb.py
echo.

echo --- Versao do arquivo ---
findstr /c:"APP_VERSION" calculadora_bcb.py | findstr /v "config\|getattr"
echo.

echo --- Limpando caches e builds antigos ---
if exist __pycache__ rmdir /s /q __pycache__
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist CalculadoraBCB.spec del CalculadoraBCB.spec
echo Limpeza concluida.
echo.

echo --- Verificando dependencias ---
python -m pip install --upgrade pip 2>nul
python -m pip install pyinstaller openpyxl reportlab --upgrade
echo.

echo --- Compilando exe (vai demorar 3-5 minutos) ---
python -m PyInstaller --onefile --windowed --name CalculadoraBCB --noconfirm ^
  --collect-all openpyxl --collect-all reportlab ^
  calculadora_bcb.py

if errorlevel 1 (
    echo ERRO na compilacao!
    pause
    exit /b 1
)

echo.
echo --- Limpando temporarios ---
if exist build rmdir /s /q build
if exist CalculadoraBCB.spec del CalculadoraBCB.spec

echo.
echo ===== SUCESSO! =====
echo  Novo executavel:  dist\CalculadoraBCB.exe
dir dist\CalculadoraBCB.exe | findstr CalculadoraBCB
echo.
pause
