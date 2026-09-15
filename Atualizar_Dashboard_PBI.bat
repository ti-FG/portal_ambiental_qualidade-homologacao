@echo off
setlocal
title Dashboard FG - Indicadores Ambientais

rem ============================================================================
rem  Atualizar_Dashboard_PBI.bat
rem
rem  Fluxo:
rem   1) Le o PBIX local em C:\Apps\Dash_ambiental_qualidade\
rem   2) Gera dashboard_fg.html
rem   3) Copia o HTML para U:\...\POWER BI\dashboard_fg.html (se acessivel)
rem   4) Abre http://10.1.1.151:8080 no navegador (so em modo interativo)
rem
rem  Modo HEADLESS:
rem   Quando DASHFG_HEADLESS=1 esta definida, nao usa pause nem abre navegador.
rem ============================================================================

cd /d "%~dp0"

rem ---------------------------------------------------------------------------
rem  Define o comando de pausa: vira no-op em modo headless
rem ---------------------------------------------------------------------------
if defined DASHFG_HEADLESS (
    set "PAUSE_CMD=rem headless sem pause"
) else (
    set "PAUSE_CMD=pause"
)

echo.
echo ============================================================
echo   DASHBOARD FG - INDICADORES AMBIENTAIS
echo   Atualizando dados a partir do modelo Power BI...
echo ============================================================
echo.

rem ---------------------------------------------------------------------------
rem  1) Verifica Python 3.12
rem ---------------------------------------------------------------------------
py -3.12 --version >nul 2>&1
if errorlevel 1 goto err_python

rem ---------------------------------------------------------------------------
rem  2) Verifica script de geracao
rem ---------------------------------------------------------------------------
if not exist "gerar_dashboard_pbix.py" goto err_script

rem ---------------------------------------------------------------------------
rem  3) Gera o dashboard
rem ---------------------------------------------------------------------------
echo Lendo PBIX e gerando dashboard...
echo.
py -3.12 "gerar_dashboard_pbix.py"
if errorlevel 1 goto err_python_run

rem ---------------------------------------------------------------------------
rem  4) Verifica o HTML gerado
rem ---------------------------------------------------------------------------
if not exist "dashboard_fg.html" goto err_html_missing

rem ---------------------------------------------------------------------------
rem  5) Copia o HTML para a U:\
rem ---------------------------------------------------------------------------
set "DEST=U:\Qualidade\AMBIENTAL\GESTAO AMBIENTAL\INDICADORES AMBIENTAIS\POWER BI\dashboard_fg.html"
echo.
echo Publicando na U:\...
copy /Y "dashboard_fg.html" "%DEST%" >nul 2>&1
if errorlevel 1 (
    echo [AVISO] Nao foi possivel copiar para a U:\.
    echo O HTML local em C:\Apps\Dash_ambiental_qualidade esta OK.
) else (
    echo Copia para a U:\ concluida.
)

echo.
echo ============================================================
echo   Dashboard atualizado com sucesso.
echo ============================================================

rem ---------------------------------------------------------------------------
rem  6) Abre o site no navegador (so em modo interativo)
rem ---------------------------------------------------------------------------
if defined DASHFG_HEADLESS goto end_headless

echo Abrindo http://10.1.1.151:8080 no navegador padrao...
start "" "http://10.1.1.151:8080"
echo.
echo Esta janela fechara em 5 segundos.
timeout /t 5 >nul
exit /b 0

:end_headless
echo Modo headless: encerrando sem abrir navegador.
exit /b 0


rem ============================================================================
rem  Tratamento de erros (todos com goto, sem parenteses)
rem ============================================================================

:err_python
echo [ERRO] Python 3.12 nao encontrado no PATH.
echo Instale o Python 3.12 e tente novamente.
%PAUSE_CMD%
exit /b 1

:err_script
echo [ERRO] gerar_dashboard_pbix.py nao encontrado em:
echo   %CD%
echo Coloque este .bat na mesma pasta dos scripts Python.
%PAUSE_CMD%
exit /b 1

:err_python_run
echo.
echo [ERRO] Falha ao gerar o dashboard.
echo Causas possiveis:
echo   - Bibliotecas Python ausentes: pbixray, pandas, numpy
echo   - Arquivo .pbix corrompido ou em uso
echo   - Permissoes negadas na pasta de saida
%PAUSE_CMD%
exit /b 1

:err_html_missing
echo [ERRO] dashboard_fg.html nao foi criado pelo Python.
%PAUSE_CMD%
exit /b 1
