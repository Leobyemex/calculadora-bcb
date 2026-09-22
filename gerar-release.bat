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
set "VER="
for /f "tokens=3" %%a in ('findstr /c:"APP_VERSION  =" calculadora_bcb.py') do set "VER=%%a"
set VER=%VER:"=%
echo  APP_VERSION = %VER%
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
REM etiqueta de versao gravada na _internal (usada para concluir atualizacoes)
(echo %VER%)> build_version.txt
python -m PyInstaller --onedir --windowed --name CalculadoraBCB --noconfirm ^
  --collect-all openpyxl --collect-all reportlab ^
  --add-data "build_version.txt;." ^
  calculadora_bcb.py
if errorlevel 1 (
    echo ERRO na compilacao!
    pause
    exit /b 1
)

echo --- Empacotando .zip para a release ---
python montar_zip_release.py %VER%
if errorlevel 1 (
    echo ERRO ao montar o zip!
    pause
    exit /b 1
)

if exist build rmdir /s /q build
if exist CalculadoraBCB.spec del CalculadoraBCB.spec
if exist build_version.txt del build_version.txt

echo.
echo ============================================================
echo  PRONTO!
echo    Pasta do app:               dist\CalculadoraBCB\  (exe + _internal)
echo    ZIP p/ anexar na release:   dist\CalculadoraBCB-v%VER%.zip  (ja com o nome certo)
echo.
echo  PROXIMOS PASSOS:
echo   1) No GitHub, crie a release com a tag  v%VER%
echo   2) Anexe o arquivo  dist\CalculadoraBCB-v%VER%.zip  na release
echo   3) Para novos usuarios, distribua a PASTA dist\CalculadoraBCB
echo      (a pasta inteira, nao apenas o .exe)
echo ============================================================
pause
