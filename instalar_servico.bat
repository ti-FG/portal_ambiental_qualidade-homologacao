@echo off
setlocal EnableDelayedExpansion
echo ============================================================
echo  Instalacao de Servico — Dashboard Ambiental (Python)
echo  Servidor: 10.1.1.151  Porta: 8081
echo ============================================================
echo.

set PASTA=C:\Apps\Dash_ambiental_qualidade

REM Localiza Python 3.12 (versao preferida)
set PYTHON_EXE=
if exist "C:\Users\julio.comin\AppData\Local\Programs\Python\Python312\python.exe" (
  set PYTHON_EXE=C:\Users\julio.comin\AppData\Local\Programs\Python\Python312\python.exe
) else (
  for /f "tokens=*" %%i in ('where python 2^>nul') do (
    if "!PYTHON_EXE!"=="" set PYTHON_EXE=%%i
  )
)

if "!PYTHON_EXE!"=="" (
  echo [ERRO] Python nao encontrado.
  exit /b 1
)
echo [OK] Python: !PYTHON_EXE!

where nssm >nul 2>&1
if %errorlevel% neq 0 (
  echo [ERRO] NSSM nao encontrado no PATH.
  exit /b 1
)
echo [OK] NSSM encontrado.

if not exist "%PASTA%\refresh_service.py" (
  echo [ERRO] refresh_service.py nao encontrado em %PASTA%
  exit /b 1
)

mkdir "%PASTA%\logs" 2>nul

echo.
echo ── Instalando servico FGDashAmbiental (porta 8081) ─────────────────────
sc query FGDashAmbiental >nul 2>&1
if %errorlevel% equ 0 (
  echo [INFO] Servico FGDashAmbiental ja existe. Removendo para reinstalar...
  sc stop FGDashAmbiental >nul 2>&1
  timeout /t 3 /nobreak >nul
  nssm remove FGDashAmbiental confirm
)

nssm install FGDashAmbiental "!PYTHON_EXE!" "refresh_service.py"
nssm set FGDashAmbiental AppDirectory        "%PASTA%"
nssm set FGDashAmbiental DisplayName         "FG - Dashboard Ambiental (Python)"
nssm set FGDashAmbiental Description         "Refresh service Python — Dashboard Ambiental, porta 8081"
nssm set FGDashAmbiental Start               SERVICE_AUTO_START
nssm set FGDashAmbiental ObjectName          LocalSystem
nssm set FGDashAmbiental AppStdout           "%PASTA%\logs\service.log"
nssm set FGDashAmbiental AppStderr           "%PASTA%\logs\service-error.log"
nssm set FGDashAmbiental AppRotateFiles      1
nssm set FGDashAmbiental AppRotateSeconds    86400
nssm set FGDashAmbiental AppRotateBytes      5242880
nssm set FGDashAmbiental AppRestartDelay     5000
echo [OK] FGDashAmbiental instalado.

echo.
echo ── Iniciando servico ────────────────────────────────────────────────────
sc start FGDashAmbiental
timeout /t 5 /nobreak >nul
sc query FGDashAmbiental | findstr /i "state"

echo.
echo  Dashboard Refresh: http://10.1.1.151:8081/status
echo  Logs:              %PASTA%\logs\
echo.
endlocal
