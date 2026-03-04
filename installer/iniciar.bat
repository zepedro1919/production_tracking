@echo off
chcp 65001 >nul 2>&1
setlocal EnableDelayedExpansion
title Print Agent Nautilus - A Iniciar...
color 0A

:: ============================================================
:: Carregar configuracao do config.ini
:: ============================================================
set "BASE_DIR=%~dp0"
set "BASE_DIR=%BASE_DIR:~0,-1%"
set "CONFIG=%BASE_DIR%\config.ini"

set "PYTHON_EXE="
set "NGROK_EXE="
set "PORT=5555"

if not exist "%CONFIG%" (
    echo.
    echo  [ERRO] Ficheiro config.ini nao encontrado!
    echo  Corra setup.bat primeiro.
    echo.
    pause
    exit /b 1
)

:: Ler config.ini
for /f "usebackq tokens=1,* delims==" %%A in ("%CONFIG%") do (
    if "%%A"=="PYTHON_EXE" set "PYTHON_EXE=%%B"
    if "%%A"=="NGROK_EXE" set "NGROK_EXE=%%B"
    if "%%A"=="PORT" set "PORT=%%B"
)

:: Verificar Python
if not exist "!PYTHON_EXE!" (
    echo  [ERRO] Python nao encontrado em: !PYTHON_EXE!
    echo  Corra setup.bat novamente.
    pause
    exit /b 1
)

:: Verificar ngrok
if "!NGROK_EXE!"=="NAO_ENCONTRADO" (
    echo.
    echo  [ERRO] ngrok nao esta configurado!
    echo  Descarregue ngrok.exe e coloque em: %BASE_DIR%\tools\
    echo  Depois corra setup.bat novamente.
    echo.
    pause
    exit /b 1
)
if not exist "!NGROK_EXE!" (
    :: Tentar novamente na pasta tools
    if exist "%BASE_DIR%\tools\ngrok.exe" (
        set "NGROK_EXE=%BASE_DIR%\tools\ngrok.exe"
    ) else (
        echo  [ERRO] ngrok nao encontrado em: !NGROK_EXE!
        echo  Coloque ngrok.exe em: %BASE_DIR%\tools\
        pause
        exit /b 1
    )
)

echo.
echo  ============================================================
echo    PRINT AGENT NAUTILUS - A INICIAR
echo  ============================================================
echo.

:: ============================================================
:: Iniciar Print Agent (em janela separada)
:: ============================================================
echo  [1/2] A iniciar Print Agent na porta !PORT!...
start "Print Agent Nautilus" cmd /k "title Print Agent Nautilus - Porta !PORT! && "!PYTHON_EXE!" "%BASE_DIR%\print_agent.py""

:: Esperar que arranque
echo  A aguardar que o servidor arranque...
timeout /t 3 /nobreak >nul

:: Verificar se esta a correr
powershell -Command "try { $r = Invoke-WebRequest -Uri 'http://localhost:!PORT!/health' -UseBasicParsing -TimeoutSec 5; Write-Host '  [OK] Print Agent a correr!' } catch { Write-Host '  [AVISO] Print Agent pode ainda estar a arrancar...' }" 2>nul

echo.

:: ============================================================
:: Iniciar ngrok (em janela separada)
:: ============================================================
echo  [2/2] A iniciar ngrok (tunel para porta !PORT!)...
start "ngrok" cmd /k "title ngrok - Tunel Activo && "!NGROK_EXE!" http !PORT!"

:: Esperar que ngrok arranque
echo  A aguardar que ngrok crie o tunel...
timeout /t 5 /nobreak >nul

:: ============================================================
:: Obter URL do ngrok
:: ============================================================
echo.
echo  A obter URL do ngrok...
echo.

set "NGROK_URL="
set "ATTEMPTS=0"

:get_url
set /a ATTEMPTS+=1
if !ATTEMPTS! gtr 10 goto :url_manual

:: Tentar obter via API do ngrok
for /f "delims=" %%U in ('powershell -Command "try { $r = Invoke-WebRequest -Uri 'http://localhost:4040/api/tunnels' -UseBasicParsing -TimeoutSec 5; $j = $r.Content | ConvertFrom-Json; $t = $j.tunnels | Where-Object { $_.proto -eq 'https' } | Select-Object -First 1; Write-Host $t.public_url } catch { Write-Host 'WAITING' }" 2^>nul') do (
    set "NGROK_URL=%%U"
)

if "!NGROK_URL!"=="WAITING" (
    timeout /t 2 /nobreak >nul
    goto :get_url
)
if "!NGROK_URL!"=="" (
    timeout /t 2 /nobreak >nul
    goto :get_url
)

goto :url_found

:url_manual
echo  [AVISO] Nao foi possivel obter URL automaticamente.
echo  Abra o browser em: http://localhost:4040
echo  E copie o URL HTTPS que aparece.
goto :done

:url_found
echo  ============================================================
echo.
echo    URL DO NGROK (copie este URL):
echo.
echo    !NGROK_URL!
echo.
echo  ============================================================
echo.
echo  PROXIMO PASSO:
echo    1. Copie o URL acima
echo    2. Abra o Google Apps Script do projecto
echo    3. No ficheiro Config.gs, altere:
echo       PRINT_AGENT_URL = "!NGROK_URL!"
echo    4. Guarde e EDITE a implementacao existente
echo       (NAO crie uma nova implementacao!)
echo.

:: Copiar para clipboard
echo !NGROK_URL!| clip
echo  [INFO] URL ja copiado para a area de transferencia!

:done
echo.
echo  ============================================================
echo  O sistema esta a correr. NAO feche estas janelas!
echo  Para parar, feche as janelas "Print Agent" e "ngrok".
echo  ============================================================
echo.
pause
