@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo  Instalando dependencias do app (PDF e Excel)
echo  reportlab = exportar PDF   ^|   openpyxl = exportar Excel
echo ============================================================
echo.

REM Usa exatamente o mesmo Python 3.14 que o executar.bat usa
set "PY=C:\Users\i501925\AppData\Local\Python\pythoncore-3.14-64\python.exe"

if exist "%PY%" (
    "%PY%" -m pip install --upgrade reportlab openpyxl
) else (
    echo Caminho fixo nao encontrado, tentando pelo PATH...
    python -m pip install --upgrade reportlab openpyxl
)

echo.
echo ============================================================
echo  Pronto! Feche o app e abra de novo pelo executar.bat.
echo  O botao "Exportar PDF" devera reaparecer.
echo ============================================================
pause
