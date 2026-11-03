# Deploy do sistema de lançamento em produção

Este guia sobe o **sistema de lançamento** (`sistema_lancamento/`, Node + PostgreSQL),
que substitui a planilha. Ele **não** é o dashboard de `http://10.1.1.151:8080/`:
aquele é o HTML gerado pelo `gerar_dashboard_pbix.py` na raiz do repositório e continua
como está (ver `../DEPLOY_PRODUCAO.md`).

Testado em 03/11/2026 em clone limpo no servidor de homologação (155), em portas
isoladas (8097 e 15450), sem tocar no 8095 nem no DW. Resultado no fim deste arquivo.

## 0. Decisões que têm de estar tomadas antes de começar

| # | Decisão | Por quê |
|---|---|---|
| 1 | **Qual banco de produção?** O 152 (`fg_apps`, schema e role `app_ambiental` pedidos à TI) ou o schema `portal_ambiental_qualidade` que já existe no DW (`dw_fg`, criado à mão em 30/10/2026) | Hoje a HML roda num PGlite local; **nada aponta para o DW**. Os schemas `portal_*` do DW são réplicas do 152, então o criado à mão pode colidir quando o do 152 for promovido |
| 2 | **Usuário de serviço só leitura no DW** (`DWC_PG*`) para Salesforce e Prevision | A credencial atual é pessoal (`natalia_costa`) |
| 3 | **Conta de bind do AD** (`LDAP_BIND_*`) | A de HML é `sa_automacao.users@FG.local`; confirmar que vale em produção |
| 4 | **Porta e liberação de firewall** | Proposta: 8095. Está livre no 151 em 03/11/2026. O alerta HmlGov da 8095 na HML nunca foi aprovado |

Sem as decisões 1 e 2 o sistema sobe, mas sem dado ou sem a busca no Salesforce/Prevision.

## 1. Onde instalar (importante)

**Não clonar dentro de `C:\Apps\Dash_ambiental_qualidade` no 151.** Essa pasta é o site IIS da
porta 8080, que serve a pasta inteira: `server.js`, `dados.js` e o `.env` ficariam baixáveis.

Clonar em outra pasta, por exemplo:

```powershell
git clone https://github.com/ti-FG/portal_ambiental_qualidade-homologacao.git C:\Apps\portal_ambiental_qualidade
cd C:\Apps\portal_ambiental_qualidade\sistema_lancamento\www
npm ci --omit=dev
```

- Instalar só `www/`. A pasta `hml/` (PGlite) é de homologação e **não** vai para produção.
- Pré-requisitos no 151: Node (já instalado em `C:\Program Files\nodejs`), NSSM (já em
  `C:\Windows\System32\nssm.exe`) e o Python com `pandas`, `openpyxl` e `psycopg2` (o
  `requirements.txt` da raiz já lista as versões). Conferido só pela existência das pastas e
  dos arquivos, em 03/11/2026; as versões do Node não foram verificadas.

## 2. Banco

O schema é idempotente: pode ser reaplicado.

```powershell
psql "host=<host> port=<porta> dbname=<banco> user=<dono_do_schema>" -f ..\db\schema.sql
```

- **Schema novo e vazio:** depois do `schema.sql`, aplicar `db/carga_inicial.sql`. É o histórico
  da planilha de 21/10/2026, **sem** as escolhas da Qualidade (Salesforce, Prevision, Sede).
- **Schema do DW já carregado (decisão 1):** aplicar só o `schema.sql`. **Nunca** rodar
  `carga_inicial.sql` nele: começa com `TRUNCATE` e apaga as escolhas de Reclamações e Avanço.

## 3. `.env` de produção

Criar `sistema_lancamento\www\.env` (nunca versionar; o `.gitignore` já ignora):

```
PORT=8095
AMBIENTE=producao

PGHOST=...
PGPORT=...
PGDATABASE=...
PGUSER=...
PGPASSWORD=...
APP_SCHEMA=portal_ambiental_qualidade
DB_POOL_MAX=4

LDAP_URL_PRIMARY=ldap://vSRV-2022-AD01.FG.local:389
LDAP_URL_FALLBACK=ldap://FGVATD-01.FG.local:389
LDAP_BIND_UPN=...
LDAP_BIND_PASSWORD=...
LDAP_BASE_DN=DC=FG,DC=local
LDAP_DOMAIN=FG
ADMIN_LOGINS=natalia.costa

DWC_PGHOST=...      # usuário de serviço só leitura (decisão 2)
DWC_PGPORT=...
DWC_PGDATABASE=dw_fg
DWC_PGUSER=...
DWC_PGPASSWORD=...

PYTHON=C:\caminho\do\python.exe   # com pandas, openpyxl e psycopg2 (gera o painel)
```

- `AMBIENTE` diferente de `homologacao` desliga a "recarga da planilha" na tela Início (é só de HML).
- `SYNC_AUTO=0` desliga a sincronização automática (Salesforce/Prevision, 09/13/17h).
  Na primeira subida vale começar com `SYNC_AUTO=0` e ligar depois de conferir o painel.
- O `PYTHON` da HML aponta para o perfil de `julio.comin`. Em produção, preferir um Python
  instalado para todos os usuários (o refresh do dashboard na 8080 tem o mesmo risco).

## 4. Serviço

Um serviço só (sem o memdb). **Não usar `servico_hml.ps1`**: ele cria os serviços de HML e o PGlite.

```powershell
$node = (Get-Command node).Source
nssm install FGAmbientalQualidade $node server.js
nssm set FGAmbientalQualidade AppDirectory C:\Apps\portal_ambiental_qualidade\sistema_lancamento\www
nssm set FGAmbientalQualidade DisplayName "FG Lancamentos Ambientais"
nssm set FGAmbientalQualidade Start SERVICE_AUTO_START
nssm set FGAmbientalQualidade AppStdout C:\Apps\portal_ambiental_qualidade\sistema_lancamento\logs\servidor.log
nssm set FGAmbientalQualidade AppStderr C:\Apps\portal_ambiental_qualidade\sistema_lancamento\logs\servidor.log
nssm set FGAmbientalQualidade AppRotateFiles 1
nssm set FGAmbientalQualidade AppRotateBytes 10485760
nssm start FGAmbientalQualidade
```

Criar antes a pasta `logs` (está no `.gitignore`).

## 5. Conferência depois de subir

1. `http://10.1.1.151:8095/api/saude` deve responder `"ok":true` e `"banco":"ok"`; `"ad":true` se o
   `LDAP_*` estiver preenchido.
2. `http://10.1.1.151:8095/api/versao` mostra `"ambiente":"producao"`.
3. Entrar com o login do AD (`natalia.costa` é admin pelo `ADMIN_LOGINS`).
4. Tela **Início** abre o mês; tela **Painel** (ou `/painel`) abre o dashboard gerado do banco.
5. Sem login, `/api/cadastros` deve dar 401, e `/.env`, `/server.js` e `/dados.js` devem dar 404.
6. Só depois: tela **Integrações** → rodar uma sincronização manual (precisa da decisão 2).

## 6. Voltar atrás

`nssm stop FGAmbientalQualidade`. O banco não é alterado por parar o serviço. Se o schema novo
foi aplicado em banco compartilhado, o `schema.sql` só cria e ajusta objetos do próprio schema.

## 7. O que foi testado e o que não foi

**Testado (clone limpo, 03/11/2026, banco PGlite novo em 15450, servidor na 8097):**
- `npm ci` em `www/` e `hml/` sem erro;
- `schema.sql` aplicado (32 tabelas + 4 views) e `carga_inicial.sql` (14 obras, 1.832 resíduos,
  687 consumos);
- servidor sobe, `/api/saude` = ok/banco ok, `/api/versao` ok, tela inicial 200;
- `/api/cadastros` sem login = 401; `/.env`, `/server.js`, `/dados.js` = 404;
- painel gerado 100% do banco (19 tabelas do banco, nenhuma da planilha), em 4,8 s, **a partir
  dos caminhos relativos do repositório** (corrigido no commit `f20a8ba`: antes apontava para
  `C:\Apps\Dash_ambiental_qualidade` escrito no código).

**Não testado:**
- login pelo AD (o clone de teste subiu sem `LDAP_*`);
- gravação de lançamentos e fechamento de mês;
- sincronização com Salesforce e Prevision (sem `DWC_*` no teste, de propósito);
- Postgres real (o teste usou PGlite; o `schema.sql` já foi aplicado no DW em 30/10/2026);
- o serviço NSSM de produção e o acesso pela rede ao 151.
