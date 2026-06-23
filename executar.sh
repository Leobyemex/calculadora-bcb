#!/usr/bin/env bash
# Calculadora do Cidadão (BCB) - Iniciar (Linux/Mac)
set -e
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

if command -v python3 >/dev/null 2>&1; then
    python3 "$SCRIPT_DIR/calculadora_bcb.py"
elif command -v python >/dev/null 2>&1; then
    python "$SCRIPT_DIR/calculadora_bcb.py"
else
    echo "ERRO: Python 3 não encontrado. Instale em https://python.org"
    exit 1
fi
