@echo off
chcp 65001 >nul 2>&1
setlocal EnableDelayedExpansion
title Instalador Print Agent - Producao Nautilus
color 0A

echo.
echo  ============================================================
echo    INSTALADOR PRINT AGENT - Producao Nautilus
echo  ============================================================
echo.
echo  Este instalador vai configurar tudo automaticamente:
echo    1. Verificar/Instalar Python
echo    2. Instalar pacotes necessarios (Flask)
echo    3. Verificar/Instalar SumatraPDF
echo    4. Verificar/Instalar ngrok
echo    5. Detectar impressora
echo    6. Configurar print agent
echo    7. Criar atalho no Ambiente de Trabalho
echo.
echo  NOTA: Compativel com Windows 7+
echo  ============================================================
echo.
pause

:: =============================================================
:: PATHS
:: =============================================================
set "INSTALL_DIR=%~dp0"
set "INSTALL_DIR=%INSTALL_DIR:~0,-1%"
set "CONFIG_FILE=%INSTALL_DIR%\config.ini"
set "PYTHON_EXE="
set "SUMATRA_EXE="
set "NGROK_EXE="
set "PRINTER_NAME="

:: =============================================================
:: STEP 1: Find or Install Python
:: =============================================================
echo.
echo  [1/7] A verificar Python...
echo  --------------------------------------------------------

:: Try common Python locations
set "FOUND_PYTHON=0"

:: Check if python is in PATH
where python >nul 2>&1
if %errorlevel%==0 (
    for /f "delims=" %%i in ('where python') do (
        set "PYTHON_EXE=%%i"
        set "FOUND_PYTHON=1"
        goto :python_check_version
    )
)

:: Check common install locations
for %%P in (
    "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python39\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python38\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python37\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python36\python.exe"
    "C:\Python310\python.exe"
    "C:\Python39\python.exe"
    "C:\Python38\python.exe"
    "C:\Python37\python.exe"
    "C:\Python36\python.exe"
    "C:\Python27\python.exe"
    "%ProgramFiles%\Python310\python.exe"
    "%ProgramFiles%\Python39\python.exe"
    "%ProgramFiles%\Python38\python.exe"
    "%ProgramFiles(x86)%\Python310\python.exe"
) do (
    if exist %%P (
        set "PYTHON_EXE=%%~P"
        set "FOUND_PYTHON=1"
        goto :python_check_version
    )
)

:: Check the installer folder for portable python
if exist "%INSTALL_DIR%\python\python.exe" (
    set "PYTHON_EXE=%INSTALL_DIR%\python\python.exe"
    set "FOUND_PYTHON=1"
    goto :python_check_version
)

:python_not_found
echo.
echo  [!] Python NAO encontrado no sistema.
echo.
echo  Para Windows 7, precisa de Python 3.8 (ultima versao compativel).
echo  Para Windows 10/11, pode usar Python 3.10+.
echo.
echo  OPCOES:
echo    1. Abra o browser e va a: https://www.python.org/downloads/
echo    2. Descarregue Python 3.8.20 (para Windows 7)
echo       ou Python 3.10+ (para Windows 10/11)
echo    3. Na instalacao, MARQUE "Add Python to PATH"
echo    4. Depois volte a correr este instalador.
echo.
echo  OU se ja tem Python instalado, indique o caminho:
set /p "PYTHON_EXE=  Caminho completo para python.exe (ou ENTER para sair): "
if "!PYTHON_EXE!"=="" (
    echo  Instalacao cancelada.
    pause
    exit /b 1
)
if not exist "!PYTHON_EXE!" (
    echo  [ERRO] Ficheiro nao encontrado: !PYTHON_EXE!
    pause
    exit /b 1
)

:python_check_version
echo  [OK] Python encontrado: !PYTHON_EXE!
"!PYTHON_EXE!" --version 2>&1
echo.

:: =============================================================
:: STEP 2: Install pip packages
:: =============================================================
echo  [2/7] A instalar pacotes Python (Flask, requests)...
echo  --------------------------------------------------------
"!PYTHON_EXE!" -m pip install --upgrade pip 2>nul
"!PYTHON_EXE!" -m pip install flask requests
if %errorlevel% neq 0 (
    echo.
    echo  [AVISO] Falha ao instalar pacotes. A tentar com --user...
    "!PYTHON_EXE!" -m pip install --user flask requests
)
echo  [OK] Pacotes instalados.
echo.

:: =============================================================
:: STEP 3: Find or Install SumatraPDF
:: =============================================================
echo  [3/7] A verificar SumatraPDF...
echo  --------------------------------------------------------

set "FOUND_SUMATRA=0"

:: Check common locations
for %%S in (
    "%LOCALAPPDATA%\SumatraPDF\SumatraPDF.exe"
    "%ProgramFiles%\SumatraPDF\SumatraPDF.exe"
    "%ProgramFiles(x86)%\SumatraPDF\SumatraPDF.exe"
    "C:\SumatraPDF\SumatraPDF.exe"
    "%INSTALL_DIR%\SumatraPDF\SumatraPDF.exe"
    "%INSTALL_DIR%\SumatraPDF.exe"
) do (
    if exist %%S (
        set "SUMATRA_EXE=%%~S"
        set "FOUND_SUMATRA=1"
        goto :sumatra_found
    )
)

:: Check if included in installer folder
if exist "%INSTALL_DIR%\tools\SumatraPDF.exe" (
    set "SUMATRA_EXE=%INSTALL_DIR%\tools\SumatraPDF.exe"
    set "FOUND_SUMATRA=1"
    goto :sumatra_found
)

echo  [!] SumatraPDF NAO encontrado.
echo.
echo  SumatraPDF e necessario para impressao silenciosa.
echo  Descarregue de: https://www.sumatrapdfreader.org/download-free-pdf-viewer
echo  Escolha a versao PORTABLE (nao precisa instalar).
echo  Copie SumatraPDF.exe para a pasta: %INSTALL_DIR%\tools\
echo.
echo  OU indique o caminho se ja o tem instalado:
set /p "SUMATRA_EXE=  Caminho para SumatraPDF.exe (ou ENTER para tentar depois): "
if "!SUMATRA_EXE!"=="" (
    echo  [AVISO] SumatraPDF nao configurado. Vai usar metodo alternativo de impressao.
    set "SUMATRA_EXE=NAO_ENCONTRADO"
) else (
    if not exist "!SUMATRA_EXE!" (
        echo  [AVISO] Ficheiro nao encontrado. Vai usar metodo alternativo.
        set "SUMATRA_EXE=NAO_ENCONTRADO"
    )
)
goto :sumatra_done

:sumatra_found
echo  [OK] SumatraPDF encontrado: !SUMATRA_EXE!

:sumatra_done
echo.

:: =============================================================
:: STEP 4: Find or Install ngrok
:: =============================================================
echo  [4/7] A verificar ngrok...
echo  --------------------------------------------------------

set "FOUND_NGROK=0"

:: Check PATH
where ngrok >nul 2>&1
if %errorlevel%==0 (
    for /f "delims=" %%i in ('where ngrok') do (
        set "NGROK_EXE=%%i"
        set "FOUND_NGROK=1"
        goto :ngrok_found
    )
)

:: Check common locations
for %%N in (
    "%INSTALL_DIR%\ngrok.exe"
    "%INSTALL_DIR%\tools\ngrok.exe"
    "%USERPROFILE%\ngrok.exe"
    "%LOCALAPPDATA%\ngrok\ngrok.exe"
    "C:\ngrok\ngrok.exe"
) do (
    if exist %%N (
        set "NGROK_EXE=%%~N"
        set "FOUND_NGROK=1"
        goto :ngrok_found
    )
)

echo  [!] ngrok NAO encontrado.
echo.
echo  ngrok e necessario para criar o tunel internet.
echo  Descarregue de: https://ngrok.com/download
echo  Escolha: Windows (32-bit ou 64-bit conforme o PC)
echo  NOTA: Para Windows 7, pode precisar da versao 32-bit.
echo.
echo  Extraia ngrok.exe para: %INSTALL_DIR%\tools\
echo.
echo  OU indique o caminho se ja o tem:
set /p "NGROK_EXE=  Caminho para ngrok.exe (ou ENTER para configurar depois): "
if "!NGROK_EXE!"=="" (
    echo  [AVISO] ngrok nao configurado. Configure antes de iniciar.
    set "NGROK_EXE=NAO_ENCONTRADO"
) else (
    if not exist "!NGROK_EXE!" (
        echo  [AVISO] Ficheiro nao encontrado.
        set "NGROK_EXE=NAO_ENCONTRADO"
    )
)
goto :ngrok_done

:ngrok_found
echo  [OK] ngrok encontrado: !NGROK_EXE!

:ngrok_done
echo.

:: =============================================================
:: STEP 5: Detect Printer
:: =============================================================
echo  [5/7] A detectar impressoras...
echo  --------------------------------------------------------
echo.
echo  Impressoras encontradas:
echo  ---
wmic printer get name 2>nul | findstr /v "^$" | findstr /v "Name"
echo  ---
echo.
echo  Qual impressora quer usar? Copie o nome EXACTO de cima.
echo  (Se for a EPSON, provavelmente e: EPSON ET-M1170 Series)
echo.
set /p "PRINTER_NAME=  Nome da impressora: "
if "!PRINTER_NAME!"=="" (
    echo  [AVISO] Nenhuma impressora selecionada. Vai usar a impressora predefinida.
    set "PRINTER_NAME=DEFAULT"
)
echo  [OK] Impressora: !PRINTER_NAME!
echo.

:: =============================================================
:: STEP 6: Create config and print_agent.py
:: =============================================================
echo  [6/7] A criar ficheiros de configuracao...
echo  --------------------------------------------------------

:: Save config
(
echo [CONFIG]
echo PYTHON_EXE=!PYTHON_EXE!
echo SUMATRA_EXE=!SUMATRA_EXE!
echo NGROK_EXE=!NGROK_EXE!
echo PRINTER_NAME=!PRINTER_NAME!
echo PORT=5555
echo AUTH_TOKEN=producao2026
echo INSTALL_DIR=!INSTALL_DIR!
) > "!CONFIG_FILE!"

echo  [OK] Configuracao guardada em: !CONFIG_FILE!
echo.

:: Create tools dir if needed
if not exist "%INSTALL_DIR%\tools" mkdir "%INSTALL_DIR%\tools"

:: =============================================================
:: STEP 7: Create desktop shortcut & launcher
:: =============================================================
echo  [7/7] A criar atalhos...
echo  --------------------------------------------------------

:: Create the launcher VBS (to run without visible cmd window issues)
set "DESKTOP=%USERPROFILE%\Desktop"
if not exist "!DESKTOP!" set "DESKTOP=%USERPROFILE%\Ambiente de Trabalho"

:: Create a simple shortcut batch on desktop
(
echo @echo off
echo cd /d "%INSTALL_DIR%"
echo call iniciar.bat
) > "!DESKTOP!\Print Agent Nautilus.bat"

echo  [OK] Atalho criado no Ambiente de Trabalho.
echo.

:: =============================================================
:: DONE
:: =============================================================
echo.
echo  ============================================================
echo    INSTALACAO CONCLUIDA!
echo  ============================================================
echo.
echo  Resumo:
echo    Python:     !PYTHON_EXE!
echo    SumatraPDF: !SUMATRA_EXE!
echo    ngrok:      !NGROK_EXE!
echo    Impressora: !PRINTER_NAME!
echo    Pasta:      !INSTALL_DIR!
echo.
if "!SUMATRA_EXE!"=="NAO_ENCONTRADO" (
    echo  [!] FALTA: SumatraPDF - descarregue e coloque em %INSTALL_DIR%\tools\
)
if "!NGROK_EXE!"=="NAO_ENCONTRADO" (
    echo  [!] FALTA: ngrok - descarregue e coloque em %INSTALL_DIR%\tools\
)
echo.
echo  Para iniciar: corra "iniciar.bat" ou o atalho no Desktop.
echo  ============================================================
echo.
pause
