@echo off
rem ===========================================================================
rem  Atualizar_Banco.bat
rem  Recarrega o banco a partir da planilha. Feito para rodar como TAREFA
rem  AGENDADA do Windows.
rem
rem  DESTINO DO BANCO: nao fica mais aqui.
rem  Ate 18/09/2026 este arquivo trazia host/base/usuario/senha chumbados, e
rem  apontavam para um PostgreSQL "localhost" com base propria (ambiental) e
rem  role etl_ambiental -- configuracao anterior a migracao. O destino real e
rem  o PostgreSQL consolidado 10.1.1.152, base fg_apps compartilhada, com
rem  isolamento POR SCHEMA e um role dedicado por app.
rem
rem  Agora tudo isso vem do .env na raiz do projeto, que o proprio
rem  etl_excel_para_db.py le:
rem      DATABASE_URL       destino (10.1.1.152/fg_apps, role app_ambiental)
rem      DASHFG_DB_SCHEMA   schema do app (bi_ambiental_qualidade)
rem      PGPASSWORD         senha do role
rem  Uma copia so da configuracao, num arquivo so. Nao duplicar aqui.
rem
rem  ANTES DE AGENDAR, rodar uma vez:
rem      py -3.12 etl\validar_conexao.py
rem  Ele conecta, confere search_path/permissoes/registro em meta.schemas e
rem  so entao voce sabe que a tarefa vai ter onde gravar.
rem
rem  Codigos de saida:  0 = carregou (ou nada mudou)   1 = falha na carga
rem                     2 = erro de estrutura no modelo
rem                     3 = planilha inacessivel
rem ===========================================================================
setlocal

rem --- 1) Onde esta a planilha ----------------------------------------------
rem  Sem --excel, o ETL procura nos caminhos internos: primeiro a copia local
rem  na pasta do projeto, depois o caminho de rede. E o que vinha funcionando.
rem
rem  Para fixar a origem, descomente as duas linhas abaixo com o caminho UNC
rem  COMPLETO ATE O .xlsx. Tem que ser UNC (\\servidor\share\...), NUNCA
rem  "U:\..." nem "S:\...": unidade mapeada so existe dentro da sessao de quem
rem  a mapeou, e a tarefa agendada nao enxerga. A conta que roda a tarefa
rem  precisa ter leitura no share. Falta confirmar o nome real do servidor.
rem set "PLANILHA=\\<servidor>\Qualidade\AMBIENTAL\GESTAO AMBIENTAL\INDICADORES AMBIENTAIS\POWER BI\Controle Ambiental 2025_oficial_db.xlsx"
rem set "ARG_EXCEL=--excel "%PLANILHA%""

rem --- 2) Python e pastas ----------------------------------------------------
set "PY=C:\Users\julio.comin\AppData\Local\Programs\Python\Python312\python.exe"
set "ETL=%~dp0"
set "LOG=%ETL%logs\etl_banco.log"

if not exist "%PY%" (
    echo [ERRO] Python nao encontrado em %PY%
    exit /b 3
)

rem --- 3) Carga --------------------------------------------------------------
rem  --se-mudou: so recarrega se a planilha mudou desde a ultima carga OK.
rem  Sem --db-url e sem --schema: os dois vem do .env.
echo Atualizando o banco a partir da planilha...
"%PY%" "%ETL%etl_excel_para_db.py" ^
    %ARG_EXCEL% ^
    --se-mudou ^
    --log "%LOG%"

set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" (
    echo [ERRO] ETL terminou com codigo %RC%. Veja %LOG%
)
endlocal & exit /b %RC%
