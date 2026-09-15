@echo off
setlocal
title Atualizar Dashboard - Indicadores Ambientais FG

rem ============================================================================
rem  Atualizar_Dashboard.bat   (versao do usuario — colocar na U:\)
rem  ───────────────────────────────────────────────────────────
rem  Fluxo:
rem    1) Chama o servico de refresh em http://10.1.1.151:8081/
rem       (o servidor regera o dashboard a partir do PBIX)
rem    2) Abre o dashboard em http://10.1.1.151:8080
rem
rem  NAO precisa de Python no PC do usuario.
rem  Funciona em qualquer maquina com acesso a rede.
rem
rem  Local recomendado:
rem    U:\Qualidade\AMBIENTAL\GESTAO AMBIENTAL\
rem    INDICADORES AMBIENTAIS\POWER BI\Atualizar_Dashboard.bat
rem ============================================================================

set "SRV=10.1.1.151"
set "REFRESH_URL=http://%SRV%:8081/"
set "DASHBOARD_URL=http://%SRV%:8080"

echo.
echo ============================================================
echo   DASHBOARD FG - INDICADORES AMBIENTAIS
echo   Solicitando atualizacao dos dados...
echo ============================================================
echo.
echo   Servidor: %SRV%
echo.

rem ── Chama o endpoint de refresh; espera ate 3 min pelo servidor ────────────
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "try { $r = Invoke-WebRequest -Uri '%REFRESH_URL%' -TimeoutSec 180 -UseBasicParsing -ErrorAction Stop; Write-Host $r.Content -ForegroundColor Green; exit 0 } catch { Write-Host ('[ERRO] ' + $_.Exception.Message) -ForegroundColor Red; exit 1 }"

if errorlevel 1 (
    echo.
    echo Nao foi possivel atualizar. Possiveis causas:
    echo   - Servidor %SRV% fora do ar
    echo   - Servico de refresh nao esta rodando na porta 8081
    echo   - Firewall bloqueando a porta 8081
    echo.
    echo Abrindo o dashboard mesmo assim (com os dados anteriores)...
    echo.
    timeout /t 3 >nul
    start "" "%DASHBOARD_URL%"
    pause
    exit /b 1
)

echo.
echo Abrindo dashboard atualizado no navegador...
timeout /t 2 >nul
start "" "%DASHBOARD_URL%"

echo.
echo Esta janela fechara em 3 segundos.
timeout /t 3 >nul
exit /b 0
