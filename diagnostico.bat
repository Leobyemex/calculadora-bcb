@echo off
chcp 65001 >nul
echo === DIAGNOSTICO Calculadora BCB ===
echo.
echo --- Diretorio atual ---
echo %CD%
echo.
echo --- Python ativo ---
python -c "import sys; print(sys.executable); print('Versao:', sys.version)"
echo.
echo --- openpyxl ---
python -c "import openpyxl; print('OK', openpyxl.__version__)" 2>&1
echo.
echo --- reportlab ---
python -c "import reportlab; print('OK', reportlab.Version)" 2>&1
echo.
echo --- Imports do reportlab usados no script ---
python -c "from reportlab.lib import colors; from reportlab.lib.pagesizes import A4, landscape, portrait; from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle; from reportlab.lib.units import cm, mm; from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image; from reportlab.platypus.flowables import KeepTogether; from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT; from reportlab.pdfgen import canvas; print('TUDO OK')" 2>&1
echo.
echo --- Arquivos suspeitos (reportlab*) na pasta ---
dir reportlab* 2>nul
if errorlevel 1 echo (nenhum)
echo.
echo --- Estado das flags HAS_XLSX e HAS_PDF dentro do script ---
python -c "import calculadora_bcb as m; print('APP_VERSION:', m.APP_VERSION); print('HAS_XLSX:', m.HAS_XLSX); print('HAS_PDF:', m.HAS_PDF)" 2>&1
echo.
echo === FIM ===
echo.
pause
