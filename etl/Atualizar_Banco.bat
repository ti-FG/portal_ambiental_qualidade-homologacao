@echo off
rem ===========================================================================
rem  Atualizar_Banco.bat
rem  Recarrega o banco a partir da planilha da rede. Feito para rodar como
rem  TAREFA AGENDADA do Windows.
rem
rem  Antes de agendar, ajuste as tres variaveis abaixo.
rem
rem  IMPORTANTE - o erro que mais derruba esse tipo de tarefa:
rem  PLANILHA tem que ser caminho UNC (\\servidor\share\...), NUNCA "U:\...".
rem  Unidade mapeada so existe dentro da sessao de quem a mapeou; a tarefa
rem  agendada nao enxerga. E a conta que roda a tarefa precisa ter leitura
rem  no share.
rem
rem  Codigos de saida:  0 = carregou (ou nada mudou)   1 = falha na carga
rem                     2 = erro de estrutura no modelo
rem                     3 = planilha inacessivel
rem ===========================================================================
setlocal

rem --- 1) Onde esta a planilha (UNC!) ---------------------------------------
set "PLANILHA=\\SERVIDOR\Qualidade\AMBIENTAL\GESTAO AMBIENTAL\INDICADORES AMBIENTAIS\POWER BI\Controle Ambiental 2025_oficial_db.xlsx"

rem --- 2) Banco de destino ---------------------------------------------------
rem  A senha sai daqui e vai para a variavel PGPASSWORD, para nao ficar na
rem  linha de comando (onde apareceria na lista de processos).
set "PGHOST=localhost"
set "PGPORT=5432"
set "PGBASE=ambiental"
set "PGUSER=etl_ambiental"
set "PGPASSWORD=TROQUE_AQUI"
set "DB_URL=postgresql+psycopg2://%PGUSER%@%PGHOST%:%PGPORT%/%PGBASE%"
set "SCHEMA=ambiental"

rem --- 3) Python e pastas ----------------------------------------------------
set "PY=C:\Users\julio.comin\AppData\Local\Programs\Python\Python312\python.exe"
set "ETL=%~dp0"
set "LOG=%ETL%logs\etl_banco.log"

if not exist "%PY%" (
    echo [ERRO] Python nao encontrado em %PY%
    exit /b 3
)

echo Atualizando banco a partir de: %PLANILHA%
"%PY%" "%ETL%etl_excel_para_db.py" ^
    --excel "%PLANILHA%" ^
    --db-url "%DB_URL%" ^
    --schema "%SCHEMA%" ^
    --se-mudou ^
    --log "%LOG%"

set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" (
    echo [ERRO] ETL terminou com codigo %RC%. Veja %LOG%
)
endlocal & exit /b %RC%
