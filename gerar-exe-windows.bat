@echo off
REM ========================================================
REM  Gerar CalculadoraBCB.exe (Windows) - v2.1 com PDF + XLSX
REM  Requer Python 3 instalado.
REM ========================================================
setlocal

echo.
echo  Calculadora do Cidadao v2.1 - Gerador de Executavel
echo  ====================================================
echo.

where python >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo  ERRO: Python nao encontrado. Instale: https://python.org/downloads/
    pause & exit /b 1
)

echo  [1/5] Instalando PyInstaller...
python -m pip install --upgrade --quiet pyinstaller
if %ERRORLEVEL% NEQ 0 ( echo  ERRO. & pause & exit /b 1 )

echo  [2/5] Instalando openpyxl (XLSX)...
python -m pip install --upgrade --quiet openpyxl

echo  [3/5] Instalando reportlab (PDF)...
python -m pip install --upgrade --quiet reportlab

echo  [4/5] Compilando executavel (pode demorar 3-5 minutos)...
cd /d "%~dp0"
python -m PyInstaller --onefile --windowed --name CalculadoraBCB --noconfirm ^
  --collect-all openpyxl --collect-all reportlab ^
  calculadora_bcb.py
if %ERRORLEVEL% NEQ 0 ( echo  ERRO. & pause & exit /b 1 )

echo  [5/5] Limpando temporarios...
if exist build rmdir /s /q build
if exist CalculadoraBCB.spec del CalculadoraBCB.spec

echo.
echo  ===== SUCESSO! =====
echo  Executavel:  dist\CalculadoraBCB.exe
echo.
pause
