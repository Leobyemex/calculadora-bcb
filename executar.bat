@echo off
chcp 65001 >nul
cd /d "%~dp0"

REM Apagar cache do Python sempre antes de rodar (evita carregar versao antiga)
if exist __pycache__ rmdir /s /q __pycache__

REM Rodar app (pythonw = sem console preto)
start "" pythonw calculadora_bcb.py
