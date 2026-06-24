@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo  GERAR RELEASE (.exe + .zip para auto-update)
echo ============================================================
echo.

if not exist calculadora_bcb.py (
    echo ERRO: calculadora_bcb.py nao encontrado nesta pasta!
    pause & exit /b 1
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

echo --- Compilando .exe (3-5 min) ---
python -m PyInstaller --onefile --windowed --name CalculadoraBCB --noconfirm ^
  --collect-all openpyxl --collect-all reportlab ^
  calculadora_bcb.py
if errorlevel 1 ( echo ERRO na compilacao! & pause & exit /b 1 )

echo --- Empacotando .zip para a release ---
REM Estrutura esperada pelo updater: CalculadoraBCB\CalculadoraBCB.exe
if exist dist\_pkg rmdir /s /q dist\_pkg
mkdir dist\_pkg\CalculadoraBCB
copy /Y dist\CalculadoraBCB.exe dist\_pkg\CalculadoraBCB\CalculadoraBCB.exe >nul
if exist dist\CalculadoraBCB.zip del dist\CalculadoraBCB.zip
powershell -NoProfile -ExecutionPolicy Bypass -Command "Compress-Archive -Path 'dist\_pkg\CalculadoraBCB' -DestinationPath 'dist\CalculadoraBCB.zip' -Force"
rmdir /s /q dist\_pkg

REM limpeza de temporarios da compilacao
if exist build rmdir /s /q build
if exist CalculadoraBCB.spec del CalculadoraBCB.spec

echo.
echo ============================================================
echo  PRONTO!
echo    Executavel:                 dist\CalculadoraBCB.exe
echo    ZIP p/ anexar na release:   dist\CalculadoraBCB.zip
echo.
echo  PROXIMO PASSO:
echo   1) No GitHub, crie a release com a tag  v2.9.31
echo   2) Anexe o arquivo  dist\CalculadoraBCB.zip  na release
echo ============================================================
pause
