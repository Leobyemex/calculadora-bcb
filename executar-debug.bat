@echo off
chcp 65001 >nul
cd /d "%~dp0"

REM Apagar cache do Python sempre antes de rodar
if exist __pycache__ rmdir /s /q __pycache__

REM Rodar com console visivel (pra ver erros)
python -u calculadora_bcb.py
pause
