#!/usr/bin/env bash
# Gera executável standalone (Linux ou Mac) - v2.1
set -e
cd "$(dirname "${BASH_SOURCE[0]}")"

echo ""
echo "  Calculadora do Cidadão v2.1 - Gerador de Executável"
echo "  ====================================================="
echo ""

PY=""
if command -v python3 >/dev/null 2>&1; then PY=python3
elif command -v python >/dev/null 2>&1; then PY=python
else echo "ERRO: Python 3 não encontrado."; exit 1; fi

echo "  [1/5] Instalando PyInstaller..."
"$PY" -m pip install --user --upgrade --quiet pyinstaller

echo "  [2/5] Instalando openpyxl (XLSX)..."
"$PY" -m pip install --user --upgrade --quiet openpyxl

echo "  [3/5] Instalando reportlab (PDF)..."
"$PY" -m pip install --user --upgrade --quiet reportlab

echo "  [4/5] Compilando (pode demorar 3-5 minutos)..."
"$PY" -m PyInstaller --onefile --windowed --name CalculadoraBCB --noconfirm \
  --collect-all openpyxl --collect-all reportlab \
  calculadora_bcb.py

echo "  [5/5] Limpando temporários..."
rm -rf build CalculadoraBCB.spec

echo ""
echo "  ===== SUCESSO! ====="
echo "  Executável: dist/CalculadoraBCB"
echo ""
