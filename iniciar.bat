@echo off
setlocal
title Sistema de Fichas de Inventario SAP MI31 (A4 + Debug)
cd /d "%~dp0"

echo ========================================================
echo   Inicializando Emissor de Fichas SAP MI31 - Folha A4
echo ========================================================

rem --- Localiza o Python (lancador "py" ou "python" no PATH) ---
set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY (
    python --version >nul 2>&1 && set "PY=python"
)
if not defined PY (
    echo [ERRO] Python nao encontrado. Instale o Python 3 em https://www.python.org
    echo        marcando a opcao "Add python.exe to PATH" e tente novamente.
    pause
    exit /b 1
)
echo Python encontrado:
%PY% --version

echo Verificando dependencias necessarias...
%PY% -c "import flask, win32com, win32clipboard, pythoncom, win32print, reportlab, pypdfium2, PIL" >nul 2>&1
if not errorlevel 1 goto iniciar

rem --- Garante o pip (usa "python -m pip", nao depende do pip no PATH) ---
%PY% -m pip --version >nul 2>&1
if errorlevel 1 (
    echo Instalando o pip...
    %PY% -m ensurepip --upgrade
)

echo Instalando Flask, pywin32, reportlab, pypdfium2 e Pillow...
%PY% -m pip install --disable-pip-version-check flask pywin32 reportlab pypdfium2 pillow
if errorlevel 1 (
    echo Tentando instalar apenas para o usuario atual...
    %PY% -m pip install --user --disable-pip-version-check flask pywin32 reportlab pypdfium2 pillow
)

%PY% -c "import flask, win32com, win32clipboard, pythoncom, win32print, reportlab, pypdfium2, PIL" >nul 2>&1
if errorlevel 1 (
    echo.
    echo [ERRO] Nao foi possivel instalar as dependencias.
    echo        Verifique a conexao com a internet ou o proxy da empresa e tente novamente.
    echo        Comando manual: %PY% -m pip install flask pywin32 reportlab pypdfium2 pillow
    pause
    exit /b 1
)

:iniciar
echo Iniciando servidor Flask e abrindo navegador...
%PY% app.py
pause
