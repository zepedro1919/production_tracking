@echo off
chcp 65001 >nul 2>&1
setlocal EnableDelayedExpansion
title Print Agent Nautilus
color 0A

set "BASE_DIR=%~dp0"
set "BASE_DIR=%BASE_DIR:~0,-1%"
set "CONFIG=%BASE_DIR%\config.ini"

:: Ler Python do config.ini
set "PYTHON_EXE="
if exist "%CONFIG%" (
    for /f "usebackq tokens=1,* delims==" %%A in ("%CONFIG%") do (
        if "%%A"=="PYTHON_EXE" set "PYTHON_EXE=%%B"
    )
)

:: Se nao encontrou no config, tentar python do PATH
if "!PYTHON_EXE!"=="" (
    where python >nul 2>&1
    if !errorlevel!==0 (
        set "PYTHON_EXE=python"
    ) else (
        echo.
        echo  [ERRO] Python nao encontrado!
        echo  Corra setup.bat primeiro.
        pause
        exit /b 1
    )
)

:: Verificar se existe
if not "!PYTHON_EXE!"=="python" (
    if not exist "!PYTHON_EXE!" (
        echo.
        echo  [ERRO] Python nao encontrado em: !PYTHON_EXE!
        echo  Corra setup.bat novamente.
        pause
        exit /b 1
    )
)

:: Arrancar tudo via launcher.py (Print Agent + LocalTunnel num so processo)
:: PYTHONIOENCODING forca UTF-8 com replace para chars invalidos no Win7
set "PYTHONIOENCODING=utf-8:replace"
set "PYTHONLEGACYWINDOWSSTDIO=1"
"!PYTHON_EXE!" -u "%BASE_DIR%\launcher.py"

pause
