# Deploy do sistema de lançamento em produção

Este guia sobe o **sistema de lançamento** (`sistema_lancamento/`, Node + PostgreSQL),
que substitui a planilha. Ele **não** é o dashboard de `http://10.1.1.151:8080/`:
aquele é o HTML gerado pelo `gerar_dashboard_pbix.py` na raiz do repositório e continua
como está (ver `../DEPLOY_PRODUCAO.md`).

Testado em 03/11/2026 em clone limpo no servidor de homologação (155), em portas
isoladas (8097 e 15450), sem tocar no 8095 nem no DW. Resultado no fim deste arquivo.

## 0. O que já está decidido e o que falta

**Banco de produção: já decidido.** É o schema `portal_ambiental_qualidade` no DW
(`dw_fg`, 192.168.100.104:5433), criado e carregado em 30/10/2026 para este sistema. Ele já tem
as tabelas e os dados. A HML roda num PGlite local só porque era homologação; produção aponta
para o DW.

O que falta, e é de quem tem acesso:

| # | O que falta | Detalhe |
|---|---|---|
| 1 | **Usuário do app no DW** (`PG*`) | Com escrita no schema `portal_ambiental_qualidade`. A credencial atual é pessoal (`natalia_costa`, dona do schema) |
| 2 | **Usuário de serviço só leitura no DW** (`DWC_PG*`) para Salesforce e Prevision (`silver.*`) | Hoje é a mesma credencial pessoal |
| 3 | **Conta de bind do AD** (`LDAP_BIND_*`) | A de HML é `sa_automacao.users@FG.local`; confirmar que vale em produção |
| 4 | **Porta e liberação de firewall** | Proposta: 8095. Está livre no 151 em 03/11/2026. O alerta HmlGov da 8095 na HML nunca foi aprovado |

Sem o item 1 o sistema não grava; sem o 2 não busca no Salesforce/Prevision.

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

**Produção = schema do DW, que já existe com dados.** Como esse schema já tem dados, o que se
aplica nele é:

1. `db/schema.sql` (idempotente, só cria/ajusta objetos). O DW já o recebeu em 30/10; reaplicar é
   seguro.
2. `db/atualizar_dw_para_homologacao.sql`, **uma vez**, com o dono do schema. Leva o DW ao mesmo
   estado do sistema de homologação (8095) em 03/11/2026, sem o que foi teste. Só **acrescenta** o
   que a HML tem e o DW não, e atualiza 80 valores de avanço. Não apaga nada. Tem guard: aborta se
   já foi aplicado.

Nunca rodar `carga_inicial.sql` no DW (começa com `TRUNCATE` e apaga as escolhas de Reclamações e
Avanço) nem `carga_homologacao.sql` (aborta pelo guard: o schema já tem obras).

**O que o `atualizar_dw_para_homologacao.sql` faz** (comparado em 03/11/2026, tabela por tabela,
DW × HML sem teste; as outras 24 tabelas já eram iguais):

| Tabela | O que muda no DW |
|---|---|
| `avanco_fisico` | 80 meses passam ao valor do Prevision (6 casas) e `origem='prevision'`; no DW estavam com o valor da planilha arredondado em 4 casas (ex.: 0,024700 → 0,024456). Mais 10 meses novos, um por obra, trazidos pela sincronização de 01/11 |
| `impressoes`, `impressoes_setor`, `impressoes_pessoa` | 1 + 53 + 278 linhas do relatório de impressão de outubro (parcial, 01 a 29/10), subido em 30/10 |
| `reclamacao` | 3 chamados do Salesforce importados em 01/11 |
| `salesforce_conta` | 1 conta ligada a obra em 01/11 |

Fica de fora, por ser teste: obra, unidade consumidora, assunto de NC e tipo de resíduo chamados
"Teste" (nenhum lançamento estava ligado a eles) e o GESSO / Blue View / set-2026 (1 m³ de teste na
HML; no DW já está 0, o valor da planilha).
Fica de fora, por não ser dado: `usuario`, `sessao`, `auditoria` (acessos), `log_alteracao` e
`sincronizacao`. Em produção cada pessoa pede acesso no primeiro login e o administrador aprova
(`ADMIN_LOGINS` já cobre a administradora).

`db/carga_homologacao.sql` (a HML inteira, sem teste, para schema **vazio**) serve para subir um
banco novo, não para o DW. Os dois arquivos foram gerados por leitura da HML
(`db/exportar_homologacao.js` gera a carga completa).

## 3. `.env` de produção

Criar `sistema_lancamento\www\.env` (nunca versionar; o `.gitignore` já ignora):

```
PORT=8095
AMBIENTE=producao

PGHOST=192.168.100.104
PGPORT=5433
PGDATABASE=dw_fg
PGUSER=...          # usuário do app, com escrita no schema (item 1 da seção 0)
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

**Testado (carga de dados, 03/11/2026):** `carga_homologacao.sql` aplicada sobre o `schema.sql` num
banco vazio, comparada com a HML tabela por tabela (contagem e hash do conteúdo, excluindo o teste):
as 27 tabelas conferem; o guard aborta numa segunda aplicação; as sequências de id ficam acima do
maior id. Única diferença esperada: `parametro.atualizado_em`, que o `schema.sql` preenche na
semeadura.

**Testado (atualização do DW, 03/11/2026):** o `atualizar_dw_para_homologacao.sql` **não** foi aplicado
no DW. Foi testado por simulação: banco PGlite no estado do DW (HML sem as linhas que o script traz,
com os 80 valores de avanço de volta ao da planilha), script aplicado sem erro, guard aborta numa
segunda aplicação, e o resultado fica igual à HML sem teste em todas as tabelas comparadas.

**Não testado:**
- login pelo AD (o clone de teste subiu sem `LDAP_*`);
- gravação de lançamentos e fechamento de mês;
- sincronização com Salesforce e Prevision (sem `DWC_*` no teste, de propósito);
- Postgres real (o teste usou PGlite; o `schema.sql` já foi aplicado no DW em 30/10/2026);
- o serviço NSSM de produção e o acesso pela rede ao 151.
