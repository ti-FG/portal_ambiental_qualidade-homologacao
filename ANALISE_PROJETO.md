# Análise do projeto — Dash Ambiental Qualidade

Local: `C:\Apps\Dash_ambiental_qualidade` (servidor 10.1.1.151)
Análise feita em: 2026-09-10
Observação: **não é um repositório git** — não há histórico de versões nem remoto no GitHub para consultar. Todo o levantamento abaixo veio da leitura direta dos arquivos no servidor.

## O que o projeto faz

Gera um dashboard HTML standalone (`dashboard_fg.html`) com indicadores de gestão ambiental (conformidade LAI/EIV, resíduos, ruído, água/energia, plano de ação, reclamações de vizinhos, IPA, SGA, impressões) para obras da FG, servido por IIS na porta 8080, com refresh automático diário.

## Arquivos e papel de cada um

| Arquivo | Papel |
|---|---|
| `gerar_dashboard.py` (3506 linhas) | Versão **original**: lê dados de um Excel na rede (`U:\Qualidade\AMBIENTAL\...\Controle Ambiental 2025_oficial_db.xlsx`), monta o HTML final via string Python gigante com Chart.js. Hoje sobrevive principalmente como "template" — `gerar_dashboard_pbix.py` extrai o bloco `HTML = r"""..."""` de dentro dele via parsing de texto. |
| `gerar_dashboard_pbix.py` (1119 linhas) | Versão **em produção**: usa a lib `pbixray` para ler as tabelas do modelo semântico do `.pbix` (`Dim_Obras`, `Fat_Conformidade_LAI_EIV`, `Fat_soma_residuos`, `Fat_Ruido`, `Fat_consumo_agua_luz`, `Dim_Auditoria_Ambientais`, `Dim_Plano_acao`, `Dim_NC_LAI_EIV`, `Dim_reclamacoes_vizinhos`, `Dim_emergencias_ambientais`, `Dim_Acoes_SGA`, `Fat_Impressoes`, `Fat_avanco_fisico`, `Fat_Metas_LAI_EIV`), reaproveita o template HTML do outro script e gera o `dashboard_fg.html` final. Suporta empacotamento via PyInstaller. |
| `refresh_service.py` (132 linhas) | Servidor HTTP simples (porta 8081, `0.0.0.0`, sem autenticação/TLS) que dispara `gerar_dashboard_pbix.py` sob demanda (`GET /` ou `/refresh`) e expõe `GET /status`. Roda como serviço Windows (`FGDashAmbiental`, via NSSM, conta `LocalSystem`). |
| `instalar_servico.bat` | Instala o `refresh_service.py` como serviço Windows via NSSM. |
| `Atualizar_Dashboard.bat` | Para usuários finais (deve ficar na `U:\`): chama `http://10.1.1.151:8081/` e abre `http://10.1.1.151:8080` no navegador. |
| `Atualizar_Dashboard_PBI.bat` | Roda no próprio servidor: executa `py -3.12 gerar_dashboard_pbix.py`, copia o HTML gerado para a `U:\...\POWER BI\`, abre navegador (a menos que `DASHFG_HEADLESS=1`). |
| `web.config` | Config mínima do IIS — só define `dashboard_fg.html` como documento padrão. |
| `AMBIENTAL_QUALIDADE_202602_v1_DEV.pbix` | Modelo Power BI, fonte de dados atual. |
| `logo-fg.png` | Logo, embutido em base64 no HTML final. |
| `logs/` | `service.log` (refreshes bem-sucedidos) e `service-error.log` (263KB, cheio de erros de encoding — ver abaixo). |

## Mecanismo de atualização

1. Tarefa agendada do Windows dispara `http://127.0.0.1:8081/` todo dia às 03:00.
2. `refresh_service.py` roda `gerar_dashboard_pbix.py`, que lê o `.pbix` e regrava `dashboard_fg.html`.
3. IIS serve esse HTML estático na porta 8080.
4. Usuários também podem forçar refresh manual rodando `Atualizar_Dashboard.bat`.

## Problemas/riscos identificados (por prioridade)

1. **Bug de encoding real em produção**: `refresh_service.py` usa `print()` com caractere `→` (U+2192); como o processo roda sob NSSM com stdout em cp1252, isso gera `UnicodeEncodeError` repetido em `logs/service-error.log` (262KB de tracebacks) e pode causar respostas HTTP 500 espúrias ao cliente mesmo quando o refresh deu certo.
2. **Caminhos hardcoded pessoais/frágeis**:
   - `U:\Qualidade\AMBIENTAL\...` (drive de rede, depende de mapeamento manual).
   - `C:\Users\julio.comin\AppData\Local\Programs\Python\Python312\python.exe` em `instalar_servico.bat` — depende de essa conta/usuário existir na máquina.
   - `C:\Users\edueslei.souza\OneDrive - FG BRAZIL HOLDING LTDA\...` em `gerar_dashboard_pbix.py` — caminho de desenvolvimento pessoal esquecido em produção.
3. **Serviço HTTP sem autenticação/TLS** na porta 8081, acessível a qualquer host da rede interna (qualquer um pode forçar refresh; scraping de `/status` também é livre).
4. **Acoplamento frágil**: `gerar_dashboard_pbix.py` extrai o template HTML de `gerar_dashboard.py` via `src.find('HTML = r"""')` — qualquer mudança de formatação nesse trecho do arquivo "template" quebra a geração silenciosamente.
5. **Falhas silenciosas de parsing**: matching de abas/colunas por substring (`sheet('gua')`, `if 'kg' in c.lower()` etc.) não falha quando não encontra — segue com DataFrame vazio, podendo gerar indicadores zerados sem alertar ninguém. `warnings.filterwarnings('ignore')` mascara ainda mais isso.
6. **Cópia para a pasta pública (`U:\...POWER BI\`) tolera falha silenciosamente** em `Atualizar_Dashboard_PBI.bat` — se o `copy` falhar, o dashboard fica desatualizado sem aviso.
7. **Sem controle de versão**: não há repositório git — nenhum histórico de mudanças, dificultando auditoria/rollback de um sistema com vários scripts interdependentes.
8. **Serviço roda como `LocalSystem`** — privilégio maior do que o necessário para ler um `.pbix` e escrever um `.html`.

## Como retomar em outra conversa

Sugestão de ordem de prioridade para quem for continuar este trabalho:

1. **Corrigir o bug de encoding** em `refresh_service.py` (trocar `→` por `->` nos `print()`, ou forçar `sys.stdout.reconfigure(encoding='utf-8')` no início do serviço) — resolve o inchaço do log de erro e possíveis 500 falsos.
2. **Iniciar um repositório git local** para o projeto (`git init` em `C:\Apps\Dash_ambiental_qualidade`), commitar o estado atual, e a partir daí versionar toda mudança — hoje não há como comparar versões nem reverter.
   - Nota: no servidor 151, `git` e `Get-Service`/`python` direto por linha de comando têm restrições conhecidas (ver memória "Restrições do servidor 151") — testar `git` com cuidado e validar liveness pelo log em vez de assumir que comandos padrão funcionam.
3. **Remover caminhos pessoais hardcoded** (`julio.comin`, `edueslei.souza\OneDrive`) e substituir por variáveis de ambiente/config, para não depender de contas de usuários específicos existirem na máquina.
4. **Adicionar autenticação mínima** (token simples via header, ou restringir por IP de origem) ao endpoint HTTP 8081 do `refresh_service.py`.
5. **Tornar falhas de parsing visíveis**: quando uma aba/tabela/coluna esperada não é encontrada, logar isso de forma destacada (ex: gravar num arquivo `logs/avisos_dados.log` separado) em vez de só seguir com dado vazio — hoje um problema na fonte de dados pode passar despercebido por semanas.
6. **Avaliar unificar os dois scripts geradores**: hoje `gerar_dashboard.py` (fonte Excel) só existe para fornecer o template HTML ao `gerar_dashboard_pbix.py` (fonte PBIX) via parsing de texto frágil. Extrair o template para um arquivo `.html` próprio (com placeholders) elimina esse acoplamento e deixa claro que a fonte de dados oficial hoje é o `.pbix`.
7. **Revisar se o serviço Windows precisa rodar como `LocalSystem`** ou se uma conta de serviço com menos privilégio é suficiente.

Nenhum código foi alterado durante esta análise — é somente um levantamento de estado.
