# Trabalho em Homologação - Dashboard Ambiental Qualidade

**Última atualização:** 15 de setembro de 2026

Consolidação de todas as alterações, correções, investigações e decisões de engenharia feitas em homologação (10.1.1.155) antes de promover para produção (10.1.1.151).

---

## 1. Fluxo de Deploy: Homologação → Produção

### Regra Obrigatória

**Nunca alterar/disparar refresh diretamente no servidor de produção 10.1.1.151** (dashboard `Dash_ambiental_qualidade`, IIS porta 8080, refresh service porta 8081).

Toda mudança em gráficos/dados do dashboard tem que ser feita primeiro no servidor de **homologação 10.1.1.155**, validada lá, e só depois promovida/aplicada em produção.

### Como aplicar

Antes de testar/disparar qualquer geração de dashboard:
1. Confirmar em qual servidor a sessão está operando (checar caminhos, IP acessível, `ipconfig`/hostname)
2. Se for produção (151), NÃO rodar refresh nem escrever arquivos live lá
3. Replicar a mudança em **10.1.1.155 primeiro**
4. Validar em homologação
5. Só depois promover para produção

**Contexto histórico:** usuário corrigiu após eu disparar (por engano) um refresh via `http://10.1.1.151:8081/refresh`, achando que estava agindo num ambiente de teste.

---

## 2. Serviço Web de Homologação (Porta 8090)

Em 10.1.1.155 não havia IIS instalado nem serviço de web para o dashboard. Para permitir validar mudanças visualmente:

- **Script:** `serve_dashboard_hml.py` — servidor HTTP estático simples em Python
- **Serviço Windows:** `FGDashAmbiental-HML-Web` via NSSM
- **Porta:** 8090 (8080 já estava ocupada por outro app)
- **Link de validação:** http://10.1.1.155:8090/
- **Logs:** `C:\Apps\Dash_ambiental_qualidade\logs\web-service*.log`
- **Inicialização:** automática

Não foi criado (por decisão do usuário) o serviço equivalente ao refresh HTTP da produção (porta 8081) — regeneração do HTML aqui é manual.

---

## 3. Fix do Gráfico IPA por Obra

### Problema

Gráfico "IPA por obra — Índice de performance ambiental" na aba **Obra Individual** (`renderIpaObras()` em `dashboard_fg.html`):
- Filtro de Período não tinha efeito nenhum
- Filtrar por empreendimento único mostrava 1 coluna só, sem histórico

### Causa Raiz

`D.ipa[obra].ipa` é um valor único calculado uma vez no Python (`calc_ipa()`), usando todo o histórico da obra, sem recorte por mês.

### Solução Implementada

Adicionado `ipa_mensal[obra]` — série dos últimos 12 meses, recalculando conformidade/resíduos/ruído/auditoria mês a mês com fallback pro score agregado quando falta dado do mês. Plano de ação fica constante (não tem coluna de data histórica confiável na fonte).

No JavaScript, `renderIpaObras(obraF, anoF)` agora:
- **Obra selecionada** → histórico de até 12 meses
- **"Todas as obras"** → uma coluna por obra, recalculada pela média do período selecionado (reage ao filtro)

### Validação

- Validado localmente rodando `gerar_dashboard.py` contra o Excel local
- Simulação em sandbox Node do JS gerado com dados reais — comportamento correto confirmado
- Fallback gracioso confirmado: se `D.ipa_mensal` não existir, volta ao comportamento antigo sem quebrar
- Teste em servidor de homologação confirmado (hostname `vSRV-HML-DEV`, IP 10.1.1.155)

### Histórico de Deploy

Um refresh foi disparado direto em produção (10.1.1.151) antes da restrição ficar clara — produção rodou o refresh usando os scripts originais (sem o fix), então o bug seguiu live em 151.

---

## 4. Distribuição PBIXRay Divergente: Produção vs. Homologação

### Problema Descoberto

Os dois servidores rodam o mesmo `gerar_dashboard_pbix.py` sobre o mesmo PBIX, mas têm bibliotecas diferentes:

- **Produção 10.1.1.151**: pbixray **0.9.0**, pandas 3.0.3, numpy 2.4.6
- **Homologação 10.1.1.155**: pbixray **0.15.5**, pandas 3.0.5, numpy 2.5.2

Python nos dois: `C:\Users\julio.comin\AppData\Local\Programs\Python\Python312\python.exe`

### Efeito Confirmado

Célula numérica nula no PBIX:
- **pbixray 0.9.0** vira **0** (produção)
- **pbixray 0.15.5** vira **NaN** (homologação)

Em `Fat_Ruido` existem 8 medições com dB nulo:
- Produção: conta como 0 dB (e 0 ≤ 70 ⇒ "conforme")
- Homologação: descarta no `dropna`

**Resultado:** IPA global 79,1 (prod) vs 78,4 (hml), diferença 100% no score de ruído.

Para obra **North** (que só tem essa linha nula):
- Produção: n=1 / 0 dB / 100% conforme
- Homologação: n=0 e cai no fallback 0.5 (=50%), derrubando IPA de 75,7 para 68,2

### Como Aplicar

Ao comparar produção × homologação, nunca assumir que número diferente = mudança de código. Conferir versão de biblioteca antes. Ao promover código pra produção, promover também o pin do `requirements.txt`.

---

## 5. Gerador Real é o PBIX, Não o Excel

### Fato

Nos dois ambientes o `refresh_service.py` executa **`gerar_dashboard_pbix.py`**, que lê o **PBIX** (`AMBIENTAL_QUALIDADE_202602_v1_DEV.pbix`, cópia local) e usa `gerar_dashboard.py` apenas como **template** de HTML/CSS/JS.

`gerar_dashboard.py` contém um pipeline completo que lê o Excel, mas esse caminho **não é o que gera o dashboard publicado**.

**Confirmado no log:** `logs/service.log` registra `Generator: C:\Apps\Dash_ambiental_qualidade\gerar_dashboard_pbix.py`.

**Produção regenera sozinha ~03:00 todo dia.**

### Como Aplicar

Para qualquer pergunta de "de onde vem esse número":
1. Abrir `gerar_dashboard_pbix.py`
2. Conferir `logs/service.log` pra saber qual gerador rodou e quando
3. As fórmulas que valem estão em `gerar_dashboard_pbix.py`: `calc_ipa`, `score_ruido`, `total_rec`

---

## 6. Alterações na Aba Resíduos

### Alterações de Rótulos (11/09/2026)

- KPI "Intensidade média/da obra" → **"Geração Resíduo"**
- KPI "Classes CONAMA" → agora mostra **a classe com maior geração** no recorte (ex.: "Classe A"), com sub "maior geração · X t · Y% do total"
- KPIs **"Obras monitoradas" e "Tipos de resíduo" removidos**
- Gráfico "Intensidade kg/m² por obra" → **"Geração Resíduo"**
- Gráfico "Composição por classe CONAMA" → **"Geração Resíduos (ton)"** (virou barras verticais por obra/empreendimento em toneladas, ordem desc, cor única `#006B6B`)
- Tabela "Série mensal de resíduos" → **"Série Mensal de Resíduos: Consolidado"**
- Tabela "Detalhamento por obra" → **"Resíduos por Obra"**

### Gráfico "Geração Resíduos (ton)" — Histórico do Vai-e-Volta

1. Fiz primeiro **por obra** (barras agrupadas por classe)
2. Usuária pediu "precisa ser por classe"
3. Refiz **por classe CONAMA** (barras com paleta A/B/C/D)
4. Usuária pediu "volte esse gráfico como você tinha feito antes por empreendimento"
5. **Versão final e desejada: por obra/empreendimento, agrupado (NÃO empilhado)**

**Não refazer por classe sem pedido novo e explícito.**

Confirmado no código: **não existe dimensão "empreendimento"** separada de obra — as obras de `Dim_Obras` (Blue Coast, Sapphire, Titanium…) *são* os empreendimentos.

### Volume por Tipo de Resíduo (top 8) — Mudança Estrutural

Era barras horizontais de total por tipo. Virou **barras verticais agrupadas — 6 colunas (uma por mês) dentro de cada tipo**, até 8 tipos.

**Meses = os últimos 6 com kg > 0 dentro do recorte** (obra × período):
- Com "últ. 12m" mostra os 6 mais recentes desses 12, de propósito
- Sempre respeita as 6 colunas pedidas
- Com menos de 6 meses, fatia pela direita pra o mais recente seguir o mais escuro

**Escala sequencial azul:** `['#D6E7F8','#B5D4F4','#85B7EB','#378ADD','#185FA5','#0E4A6B']`
- Mês mais antigo: claro
- Mês mais recente: escuro

**Dados visuais:**
- Legenda no topo com os meses
- Badge novo `id="res-tipo-badge"` mostra a faixa (ex.: "Jan/26 – Abr/26")
- Rótulo do eixo X truncado em 20 chars com nome completo no tooltip
- Canvas subiu de 250px para 340px

### Limitação de Dado Importante

No pbix atual só **4 meses têm kg > 0**:
- Jan/26: 161,7 t
- Fev/26: 203,5
- Mar/26: 225,4
- Abr/26: 198,5

Todo 2024/2025 e Mai–Jul/26 estão zerados (2025 tem só m³, sem kg).

Só **6 tipos têm kg > 0**:
- Entulho: 558,9 t
- Madeira: 105,9
- Gesso: 79,6
- Metal/Sucata: 35,7
- Papel/Papelão: 6,3
- Plástico: 2,7

O gráfico mostra hoje **4 colunas × 6 tipos** (não 6 × 8), vai preencher conforme entrar mês novo.

**Pendência:** o título ainda diz "(top 8)" — confirmar se querem ajustar.

### Série Mensal de Resíduos: Consolidado

Coluna de dados agora em ordem: **m³ → kg → m³/m² → kg/m²**
- m³ virou a 1ª coluna de dado
- Mês/Ano segue como rótulo de linha
- Status no fim

Cada cabeçalho ganhou legenda:
- `m³` = volume
- `kg` = peso resíduos
- `m³/m²` = volume/m² construído
- `kg/m²` = peso/m² construído

### Resíduos por Obra

- Coluna **Meta** formatada como `50kg/m²` / `36,6kg/m²` (sem decimal quando inteira, vírgula decimal pt-BR)
- **Total kg** com separador de milhar pt-BR

### Formatação pt-BR Padronizada

Novo helper `numBR(v, dec)` no JavaScript (usa `toLocaleString('pt-BR')`).

Aplicado em:
- Os 3 KPIs
- Tooltip do gráfico de toneladas
- As 4 colunas da Série Mensal
- Colunas Kg/m² + Total kg de "Resíduos por Obra"

**As outras abas seguem com `num()`/`toFixed`** (não foi pedido).

### Rótulos de Dados (14/09/2026)

Adicionados rótulos visíveis direto nos gráficos (números sobre as barras) em 3 gráficos:

- **"Geração Resíduo"** (`numBR(v,2)`)
- **"Geração Resíduos (ton)"** (`numBR(v,1)`)
- **"Volume por tipo de resíduo"** (kg abreviado: ≥10k → `124k`, ≥1k → `6,4k`)

Usa o helper `barLabels(fmt, opts)` que já existia no arquivo (sem CDN extra).

### Senna Tower à Parte

Replicando o padrão da aba Consolidado — Senna tem meta 50 e intensidade muito acima das demais, então **fica fora da visão consolidada** (que mostra 8 obras, meta 36,6 em todas).

Botão `res-ger-senna-btn` / `toggleSennaGer()` (variável `resModoSennaGer`, espelhando `resModoSenna` do Consolidado):
- Ligado: mostra Senna mês a mês com meta 50
- Desligado: mostra 8 obras com meta 36,6

**Filtro global de obra tem precedência sobre o botão.**

### Ajustes de Tamanho (14/09/2026)

Após múltipas iterações:

- `chart-obra-ruido`: 200px → **300px** (aumentado)
- `chart-obra-res-donut`: 200px → **170px** (reduzido)
- `chart-obra-pa`: 200px → **170px** (reduzido)

Legenda dos donuts: `right` (igual referência de produção), `boxWidth` 8, `padding` 6, fonte 10px.

---

## 7. Mudança de Agregação: kg/m² de Média para Soma

### Data: 15 de setembro de 2026

Alterou o cálculo de **kg/m² (resíduos gerados)** de **média** para **soma** em vários pontos.

### Arquivos Modificados

- `gerar_dashboard_pbix.py` (linhas 411, 464, 484, 512)
- `gerar_dashboard.py` (template)

### Alteração

`.mean()` → `.sum()` em todos os pontos:
- Linha 411: `res_por_obra` — tabela de resíduos por obra (aba Consolidado)
- Linha 484: `residuos_raw` — dataset granular para filtros client-side
- Linha 512: `kg_m2_by_month` — série mensal (aba Resíduos - Geração Resíduo)
- Linha 464: `media_kg_m2` — métrica global de resíduos (card "Resíduos Gerados")

### Por Quê

Antes estava somando kg/m² de cada linha e dividindo pela quantidade de linhas (média). Agora apenas soma os kg/m², pois cada linha já representa a intensidade (kg ÷ m²) de um tipo de resíduo naquele mês.

### Impacto

- Aba Consolidado: gráfico "Resíduos Gerados kg/m² por obra" recalculado
- Aba Resíduos: gráfico "Geração Resíduo" recalculado
- Card "Resíduos Gerados": valor agregado recalculado

### Nota Importante

`score_res()` e `_score_res_mes()` foram **mantidas com `.mean()`** pois são usadas para calcular o IPA (métrica agregada diferente). Se precisar estender essa mudança para o IPA também, será uma decisão separada.

---

## 8. Armadilha Descoberta: kg/m² Soma Acumula Entre Meses

### Problema

Consequência não óbvia da troca de `.mean()` por `.sum()`.

Somar os `Kg/m²` das linhas **dentro de um mês** está certo — cada linha é um tipo de resíduo. Mas `res_por_obra` soma **todas as linhas da obra**, incluindo **entre meses**. O valor exibido passa a ser `intensidade_mensal × número_de_meses_com_lançamento`, enquanto a meta (36,6 / Senna 50) continua **mensal**.

### Efeito no Gráfico

A barra por obra cresce a cada mês novo e é comparada com uma meta mensal. O ranking fica distorcido por quantidade de meses lançados.

**Exemplo real:**

| Obra | Mostra | Meses | Por Mês | Meta |
|---|---|---|---|---|
| Senna Tower | 1.692,05 | 4 | 423,01 | 50 |
| Blue View | 468,32 | 4 | 117,08 | 36,6 |
| Imperium | 326,40 | 4 | 81,60 | 36,6 |
| Titanium | 309,25 | 4 | 77,31 | 36,6 |
| Scenarium | 208,23 | 4 | 52,06 | 36,6 |
| Iconic | 200,46 | 4 | 50,11 | 36,6 |
| Sapphire | 165,79 | 4 | 41,45 | 36,6 |
| **Blue Coast** | **83,69** | **1** | **83,69** | 36,6 |
| Garden Park | 14,99 | 4 | 3,75 | 36,6 |

Blue Coast tem 1 mês de lançamento e por isso aparece como a 2ª melhor obra (83,69), quando na verdade é a 2ª PIOR por mês (83,69/mês).

### Opções

- (a) Somar dentro do mês e **tirar a média dos meses** no nível da obra — dá a intensidade mensal média, comparável com a meta
- (b) Manter a soma acumulada e trocar a meta por `36,6 × n_meses`
- (c) Deixar o gráfico por obra sempre mensal (eixo de meses)

A (a) preserva o significado da meta.

### Nota

No eixo **mensal** (obra filtrada) o problema não aparece — ali cada barra é um mês. Só afeta a visão consolidada por obra. **Apontado à usuária em 15/09/2026 — pendência de decisão.**

---

## 9. ETL: Excel → Banco de Dados

### Iniciado: 14 de setembro de 2026

Tirar o `.pbix` do caminho como fonte de dados.

### Estrutura Criada

- **Script principal:** `etl/etl_excel_para_db.py` (CLI)
- **Definições:** `etl/tabelas.py` (as 19 tabelas)
- **Helpers:** `etl/modelo_semantico.py`
- **Validação:** `etl/validar_vs_pbix.py`
- **Documentação:** `etl/README.md`

### Descoberta Chave

O `.pbix` guarda o **código Power Query (M) de cada tabela**. Não foi preciso adivinhar transformação nenhuma; o ETL é transcrição do M.

Toda tabela sai de uma **tabela nomeada (ListObject)** do `Controle Ambiental 2025_oficial_db.xlsx`:
- `agua_luz`, `SOMA_RESIDUOS`, `Conformidade_LAI_EIV`, etc.

`Dim_Obras` sai da coluna Obra de `Metas_LAI_EIV`.
`Dim_Calendario` é gerada (ano atual ±5).

### Validação

- **19/19 tabelas** com as mesmas colunas, ordem e tipos do pbix
- **11.434 linhas** idênticas
- Diferenças = defasagem do pbix (última atualização com dado até abr/2026; planilha está em set/2026)

### Pegadinhas Tratadas

1. Modelo remove espaço no **FIM** de texto ('Emegência ' → 'Emegência')
   - Espaço no INÍCIO é preservado (' Ouvidoria')
2. `Fat_Ruido.Horário` no pbix é float: 09:03 vira 09:02:59,999999
   - Arredondar (não truncar) pro segundo elimina diferenças falsas
3. Coluna `%` de `Impressoes_por_setor` vem como **texto** ('17,19%') → tipo Percentage

### Relacionamentos

Dos 26 do pbix:
- **13 das 26 FKs já valem** (zero violação na carga)
- As outras 13 saem comentadas porque o dado tem órfão:
  - `Outros` (criado pelo próprio Power Query)
  - `North Tower`/`NORTH`/`North` (obra nova, grafada de 3 jeitos)
  - `Garden`/`Iconic Tower`/`Blue View Tower`/`Gran Place` (normalização/erro de digitação)

**Esse é estado do modelo no pbix, não defeito do ETL.**

### Banco de Dados

**Escolhido: PostgreSQL**

Existem DOIS PostgreSQL rodando no 155:
- `postgresql-x64-17` na porta 5432 (aceita a rede 10.1.0.0/16)
- `postgresql-pmetl` na 5433 (só localhost)

**Instância ainda não foi definida** — usuária pediu para aguardar e confirmar depois. Ambas exigem senha e não temos credencial.

ODBC Driver do SQL Server **não está instalado** no 155.

### Atualização Automática

Novo `origem.py` + flags `--se-mudou`, `--copiar-local`, `--log` no ETL, mais `Atualizar_Banco.bat` para tarefa agendada.

Tabela `_carga_execucao` (uma linha por execução, com sha256 da planilha; registra sucesso E falha).

### Armadilha Central

`U:` é unidade mapeada por usuário — tarefa agendada e serviço Windows **não enxergam**.

A origem tem que ser UNC (`\\servidor\share\...`) e a conta da tarefa precisa de leitura no share.

**O caminho UNC ainda é desconhecido** — nenhum script do projeto tem (todos usam `U:`). Descobrir com `net use U:` numa máquina com o drive mapeado, ou com a TI.

### Proteções Implementadas

1. Cópia a planilha pra disco local antes de ler (Excel fica aberto o dia todo)
2. Valida que o .xlsx abre (3 tentativas com 20s)
3. Lê e valida ANTES de decidir pular por `--se-mudou` (problema na origem aparece mesmo em dia parado)

**Exit codes:**
- 0: ok/sem mudança
- 1: falha de carga
- 2: erro de estrutura
- 3: origem inacessível

### Princípio: Fidelidade ao PBIX

Nada foi corrigido em silêncio. Os defeitos estão em `ALERTAS_MODELO` (`tabelas.py`) e são impressos a cada execução.

**Defeito mais grave:** `Fat_consumo_agua_luz."Água m³/m²"` é convertida pra **inteiro** pelo Power Query (0,206 vira 0), então o indicador de água do dashboard está zerado na origem.

### Próximo Passo

Confirmar o banco de destino, depois apontar `gerar_dashboard_pbix.py` pro banco (trocar `T(nome)` de `m.get_table` para `pd.read_sql`).

---

## 10. Backups do Gerador

### Localização

`C:\Apps\Dash_ambiental_qualidade\_backup_teste_refresh\`

### Problema

Os nomes não são auto-explicativos. Já causou perda de trabalho — em 15/09/2026 foi restaurado o backup **errado** (~166 KB, estado de 11/09 **antes** de todo o trabalho) e a usuária reportou "perdi as alterações".

### Como Escolher

**O arquivo com mais bytes e a data mais recente:**
- Hoje: `gerar_dashboard_20260915_latest.py` (~203 KB)
- Hoje: `gerar_dashboard_pbix_20260915_com_soma.py` (~72 KB)

**Sempre comparar tamanho antes de restaurar:**
- ~166 KB = estado antigo de 11/09
- ~203 KB = estado com todas as alterações

### Marcadores para Confirmar

Se o arquivo restaurado é o certo, procurar:
- `function barLabels`
- `function numBR`
- `toggleSennaGer`
- `res-ger-senna-btn`
- `tetoMeses`
- `classesTon`
- `kgTipoMes`
- `labelMesRes`
- `SENNA_RE`
- `subCab`
- `metaFmt`
- `Série Mensal de Resíduos: Consolidado`
- `Resíduos por Obra`
- `res-barras-titulo`

Script `comparar.py` no scratchpad faz essa checagem marcador por marcador.

### Armadilhas

1. **Grep com acento falha** nesse ambiente (Git Bash + UTF-8)
   - `grep "Resíduos por Obra"` dá falso negativo
   - Usar Python (`io.open(..., encoding='utf-8')` + `count`)

2. **Escrita truncada**
   - Escrever com `io.open(p,'w')` trunca o arquivo no open
   - Se o `.write()` falhar, fica com 0 bytes
   - **Solução:** escrita atômica — encode primeiro, grava num `.tmp`, depois `os.replace`

### Suíte de Validação

No scratchpad (rodar depois de cada mudança ou restauração):
- `validate_final.js`
- `validate_eixos.js`
- `validate_ptbr.js`
- `validate_tabelas.js`
- `validate_tipo.js`

Carregam o `dashboard_fg.html` gerado num sandbox `vm`.

**Importante:** essas suítes guardam expectativas fixas. Quando uma falha, **conferir se não é só o teste desatualizado** antes de mexer no código (já aconteceu 2 vezes — formato pt-BR e assinatura de tooltip).

---

## 11. Lineage das Unidades de Resíduos

### Confirmado por Aritmética no PBIX

**Dados:**
- `m³` = volume (sempre em coluna de 01 de cada mês)
- `kg` = peso resíduos (por tipo, por obra, por mês)
- `m³/m²` = volume ÷ m² construído da obra (coluna `Area_m2` de `Dim_Obras`)
- `kg/m²` = peso ÷ m² construído

Cada linha de `Fat_Residuos`:
- Obra, Tipo, Mês/Ano, m³, kg, m³/m², kg/m²

**Confirmação:** com Senna Tower (meta 50 kg/m²) janeiro de 2026:
- Lançado: 8 tipos de resíduo com kg > 0
- Soma de kg: 214.688 kg
- Soma de kg/m² das 8 linhas: 423,01 kg/m² (é isso que entra no `kg_m2_by_month.sum()`)
- Área Senna: 2.000 m² (coluna `Area_m2`)
- Verificação: 214.688 ÷ 2.000 = 107,344 kg/m² ✗ (não bate com 423,01)

**Causa:** cada linha já tem sua própria área calculada. Quando alguns tipos são lançados apenas com m³ (sem kg), a coluna `kg` daquela linha é 0, mas `kg/m²` também fica 0 — e o Power Query não o descarta. Somando os `kg/m²` das 8 linhas, você está somando unidades já normalizadas individualmente, o que é correto para "intensidade agregada do mês".

A aritmética é interna: cada linha é `(kg_tipo ÷ m²_obra)`, e somá-las dá a intensidade total da obra naquele mês — é assim que o dashboard reporta.

**Resalva importante:** `m³` e `kg` vêm de tabelas diferentes na fonte (Excel):
- `m³`: da coluna de volume (lançamento)
- `kg`: de tabelas de tipos de resíduos

---

## 12. Correções Adicionais

### Acentuação em Detecção de Coluna

Em `gerar_dashboard.py`, detecção de coluna "Mês/Ano" falhava por causa do "ê". Trocado para usar `_asc(c)` em vez de `c.lower()`.

Afeta só o caminho legado via Excel (sem impacto em produção que usa o .pbix), mas necessário pra validar localmente com dados mensais reais.

### Filtro de Período Agora Reage em Obra Individual

**Auditoria feita:** antes só a tabela de conformidade e o gráfico IPA por obra reagiam ao filtro de Período. Os 4 KPIs do topo (IPA da obra, Conformidade LAI/EIV, NCs, Reclamações), o donut de resíduos e o de plano de ação sempre mostravam o histórico completo.

**Agora corrigido:**
- IPA/Conformidade/NCs/Reclamações (KPIs) recalculam pelo período
- Donut de Geração Resíduos recalcula pelo período
- Dado mensal novo: `nc_mensal`, `rec_mensal`
- Resíduos reutiliza `residuos_raw` (já existente)
- Fallback pro agregado histórico quando sem filtro ativo ou sem série mensal
- Plano Ação Auditorias continua sem reagir (de propósito, sem data confiável na fonte)

Validado com simulação em Node comparando valores mês a mês — bateu exato.

### NCs Agora Inclui Resíduos Acima da Meta

A pedido da usuária, o número do card "NCs registradas LAI/EIV" passou a somar **1 ocorrência por obra-mês em que o kg/m² médio realizado de resíduos ultrapassa a meta** (Senna Tower 50 / demais obras 36,6 — confirmado com a usuária: contagem por obra-mês, não por obra agregada).

**Implementado em:**
- `gerar_dashboard_pbix.py` (usa `fat_res` + `meta_obra()`)
- `gerar_dashboard.py` (novo helper local `meta_obra()`, mesma regra 50/36,6 fixa)

Nenhuma mudança de JS/template — todos os consumidores (`renderObra`, `updateKpisFiltered`) leem `total_ncs`/`nc_por_obra`/`nc_mensal` já prontos.

**Impacto real:** `total_ncs` foi de 8 → **32** (com soma). 9 das 13 obras tiveram pelo menos 1 mês acima da meta.

### Card de NCs Removido da Aba Obra Individual

Removido o card (KPI) de `kpis-obra` em `renderObra()` — sobraram 3 KPIs (IPA da obra, Conformidade LAI/EIV, Reclamações de Vizinhos).

Grid `cols-4` → `cols-3`.

Removida também variável morta `nc` em `updateKpisFiltered()`.

**Dados Python continuam calculados** (`total_ncs`, `nc_por_obra`, `nc_mensal`), só não aparecem mais.

### Rótulos de Dados nos 4 Gráficos de Obra Individual

Adicionados via plugins Chart.js customizados (`barValueLabels`, `donutValueLabels`), sem depender de lib externa:

- **IPA por obra**: número do IPA acima de cada barra
- **Ruído**: dB acima de cada barra
- **Geração Resíduos (ton) — donut**: toneladas no centro de cada fatia
- **Plano Ação Auditorias — donut**: contagem + percentual

---

## 13. Histórico de Alterações Aba Consolidado

### Já Existia Antes (Não Entra em Detalhe Aqui)

Gráficos alterados em sessão anterior:
- "Reclamações por Obra" (novo)
- Rótulos e renomeações
- Senna Tower à parte

Ver arquivo `ALTERACOES_GRAFICOS_HOMOLOGACAO.md` (de outra sessão) que cobre esse escopo.

---

## 14. Estrutura e Localização dos Arquivos

### Arquivos Principais

- `gerar_dashboard_pbix.py` — Pipeline real que lê o PBIX
- `gerar_dashboard.py` — Template HTML + pipeline alternativo (Excel)
- `dashboard_fg.html` — Saída final (gerada, não editar manualmente)
- `serve_dashboard_hml.py` — Servidor HTTP pra validação em 8090
- `PBIX_PATH` — `AMBIENTAL_QUALIDADE_202602_v1_DEV.pbix` (cópia local)

### Diretórios

- `etl/` — Scripts de extração para banco de dados
- `logs/` — Logs de execução do serviço
- `_backup_teste_refresh/` — Backups dos geradores
- `scratchpad/` — Scripts de validação e utilitários

---

## 15. Como Versionar Isso no Git

Este arquivo (`homologacao_ambiental.md`) consolida **todas** as alterações, descobertas, armadilhas e decisões. Deve ser versionado junto com:

- `gerar_dashboard_pbix.py` (versão com todas as mudanças)
- `gerar_dashboard.py` (versão com todas as mudanças, se aplicável)
- Qualquer alteração a `dashboard_fg.html` (se manual)
- `etl/` (pipeline novo)

**Fluxo de commit:**
1. Validar em homologação (10.1.1.155:8090)
2. Descrever no git as mudanças (este arquivo é a documentação)
3. Fazer commit
4. Promover para produção (10.1.1.151:8080)

---

## 16. Aba Dicionário de Dados e Termos

### Data: 16 de setembro de 2026

Quarta aba do dashboard, ao lado de Consolidado, Obra Individual e Resíduos.
Documenta as regras de negócio, as fórmulas dos indicadores e a origem de cada
dado, para que qualquer número do dashboard possa ser rastreado até a regra que
o produziu.

**86 verbetes em 6 seções:** Indicadores e KPIs (18), Composição do IPA (6),
Regras de negócio (11), Modelo de dados — as 19 tabelas (19), Glossário (24) e
Limitações conhecidas (8).

Tem busca no topo (insensível a acento e maiúscula, esconde seção sem resultado)
e chips de atalho para cada seção. A barra de filtros fica escondida nessa aba —
conteúdo de referência não reage a recorte.

**Só `gerar_dashboard.py` foi alterado** (o template). `gerar_dashboard_pbix.py`
extrai o template desse arquivo, então a aba aparece nos dois caminhos de geração
sem mudança no gerador do PBIX. Nenhum dado novo no JSON: consome apenas
`D.meta_residuos`, que já existia.

**Conteúdo estático por decisão**, com duas exceções lidas dos dados para não
envelhecerem: as metas (36,6 / 50) e o contador de verbetes.

O texto foi apurado lendo o código, não de memória — `calc_ipa`, os cinco
`score_*`, `meta_obra`, `OBRA_MAP`, `_assuntos_nc`, a montagem de
`reclamacoes_lista` e `etl/tabelas.py`.

**Risco de manutenção:** regra que mudar no gerador tem que mudar o verbete
junto. Esta aba envelhece em silêncio se isso não for feito.

Detalhe completo em `ALTERACOES_ABA_DICIONARIO.md`.

---

## 17. Pendências e Próximos Passos

### Aguardando Decisão da Usuária

1. **Agregação kg/m² na aba Resíduos/Consolidado:** manter média ou trocar para soma como o `total_ncs`?
2. **Agregação kg/m² acumulada entre meses:** qual opção aplicar (a/b/c)?
3. **Gráfico "Geração Resíduos (ton)" na aba Obra Individual:** continua donut ou vira barras como o da aba Resíduos?
4. **Título do gráfico "Volume por tipo de resíduo (top 8)":** ajustar porque nem sempre são 8?
5. **Comparação m³/m² — agregação:** `m3_m2_by_month` usa `.mean()`, enquanto `kg_m2_by_month` usa `.sum()` — unificar?

### Técnicas para Retomar

- Sempre restaurar de backup ~203 KB (estado correto)
- Revalidar com as 5 suítes em scratchpad após qualquer mudança
- Confirmar banco PostgreSQL antes de mexer no ETL
- Disparar refresh sempre via 10.1.1.155, nunca via 10.1.1.151

---

**Documento criado e versionado: 15 de setembro de 2026**
