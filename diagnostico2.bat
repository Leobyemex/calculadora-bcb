@echo off
chcp 65001 >nul
echo === DIAGNOSTICO 2 - Conteudo do arquivo ===
echo.
echo --- Pasta atual ---
echo %CD%
echo.
echo --- Data e tamanho do arquivo ---
dir calculadora_bcb.py 2>&1
echo.
echo --- Ocorrencias de "Exportar PDF" no arquivo ---
findstr /n /c:"Exportar PDF" calculadora_bcb.py
if errorlevel 1 echo NAO ENCONTRADO - arquivo nao tem suporte a PDF
echo.
echo --- Tem o emoji documento (deveria ser zero) ---
findstr /n /c:"📄" calculadora_bcb.py
if errorlevel 1 echo OK - sem emoji
echo.
echo --- HAS_PDF e botoes na classe ---
findstr /n "if HAS_PDF" calculadora_bcb.py
echo.
echo === FIM ===
pause
