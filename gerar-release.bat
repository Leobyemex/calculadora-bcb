@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo  GERAR RELEASE (formato pasta/onedir) - exe + zip
echo ============================================================
echo.

if not exist calculadora_bcb.py (
    echo ERRO: calculadora_bcb.py nao encontrado nesta pasta!
    pause
    exit /b 1
)

echo --- Versao ---
findstr /c:"APP_VERSION  =" calculadora_bcb.py
echo.

echo --- Limpando builds antigos ---
if exist __pycache__ rmdir /s /q __pycache__
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist CalculadoraBCB.spec del CalculadoraBCB.spec

echo --- Dependencias ---
python -m pip install --upgrade pip >nul 2>&1
python -m pip install pyinstaller openpyxl reportlab --upgrade

echo --- Compilando (formato pasta) - 3 a 5 min ---
python -m PyInstaller --onedir --windowed --name CalculadoraBCB --noconfirm ^
  --collect-all openpyxl --collect-all reportlab ^
  calculadora_bcb.py
if errorlevel 1 (
    echo ERRO na compilacao!
    pause
    exit /b 1
)

echo --- Empacotando .zip para a release ---
REM onedir gera dist\CalculadoraBCB\ (exe + _internal). Zipa a PASTA inteira.
if exist dist\CalculadoraBCB.zip del dist\CalculadoraBCB.zip
powershell -NoProfile -ExecutionPolicy Bypass -Command "Compress-Archive -Path 'dist\CalculadoraBCB' -DestinationPath 'dist\CalculadoraBCB.zip' -Force"

if exist build rmdir /s /q build
if exist CalculadoraBCB.spec del CalculadoraBCB.spec

echo.
echo ============================================================
echo  PRONTO!
echo    Pasta do app:               dist\CalculadoraBCB\  (exe + _internal)
echo    ZIP p/ anexar na release:   dist\CalculadoraBCB.zip
echo.
echo  PROXIMOS PASSOS:
echo   1) No GitHub, crie a release com a tag  v2.9.32
echo   2) Anexe o arquivo  dist\CalculadoraBCB.zip  na release
echo   3) Para novos usuarios, distribua a PASTA dist\CalculadoraBCB
echo      (a pasta inteira, nao apenas o .exe)
echo ============================================================
pause
