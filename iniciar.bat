@echo off
title Sistema de Fichas de Inventario SAP MI31 (A4 + Debug)
cd /d "%~dp0"

echo ========================================================
echo   Inicializando Emissor de Fichas SAP MI31 - Folha A4
echo ========================================================

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Python nao encontrado no PATH do Windows.
    pause
    exit /b
)

echo Verificando dependencias necessarias...
python -c "import flask, win32com, win32clipboard, pythoncom" >nul 2>&1
if errorlevel 1 (
    echo Instalando Flask e pywin32...
    pip install flask pywin32
)

echo Iniciando servidor Flask e abrindo navegador...
python app.py
pause