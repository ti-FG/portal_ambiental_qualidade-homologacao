# Alterações na ABA RESÍDUOS — Dashboard Ambiental Qualidade

Período: **11/09/2026 – 15/09/2026** · Ambiente de trabalho: **homologação 10.1.1.155** (`vSRV-HML-DEV`)
Validação visual: **http://10.1.1.155:8090/** · Produção: **10.1.1.151:8080** (só após aprovação)

> Documento complementar a [`ALTERACOES_GRAFICOS_HOMOLOGACAO.md`](ALTERACOES_GRAFICOS_HOMOLOGACAO.md),
> que cobre a aba **Obra Individual**. Este cobre exclusivamente a aba **Resíduos**.

---

## Onde cada coisa mora

| Arquivo | Papel |
|---|---|
| `gerar_dashboard.py` | **Template HTML/CSS/JS.** Todo o visual e a lógica client-side da aba vivem aqui (4.531 linhas). O `gerar_dashboard_pbix.py` lê este arquivo como template. |
| `gerar_dashboard_pbix.py` | **Gerador de produção** (1.408 linhas). Extrai os dados do `.pbix` via `pbixray` e injeta no template. É o que o refresh executa. |
| `dashboard_fg.html` | Saída gerada. Servida em `155:8090` (HML) e `151:8080` (produção). |

Regenerar em homologação:

```bash
"C:\Users\julio.comin\AppData\Local\Programs\Python\Python312\python.exe" gerar_dashboard_pbix.py
```

Esse é o interpretador que tem `pbixray` instalado (0.15.5 em HML) e é o mesmo usado em produção.

---

## 1. KPIs da aba: de 5 para 3 cards

**Arquivo:** `gerar_dashboard.py` — HTML linha ~1335, JS em `renderResiduos()` (linha ~3703)

| Antes | Depois |
|---|---|
| Volume total | Volume total *(mantido)* |
| Intensidade média / da obra | **Geração Resíduo** *(só o rótulo; valor segue kg/m²)* |
| Classes CONAMA = contagem de classes | **Classes CONAMA = a classe com maior geração** |
| Obras monitoradas | *removido* |
| Tipos de resíduo | *removido* |

- `<div class="kpi-row cols-5" id="kpis-residuos">` → `cols-3`
- O card de classe passou a mostrar `Classe A` como valor e, no subtítulo,
  `maior geração · 558,9 t · 70,8% do total`. Reaproveita o agregado `classeAgg`, que
  também alimenta o gráfico de toneladas — a 1ª barra do gráfico é sempre a classe do KPI.

---

## 2. Gráfico "Geração Resíduo" (kg/m²)

**Antes:** `Intensidade kg/m² por obra` — uma barra por obra, todas as obras juntas, badge estático.

**Depois:**

- Título renomeado para **Geração Resíduo** (span dinâmico `res-ger-titulo`).
- **Senna Tower separada das demais obras**, replicando a mesma régua do card
  `Resíduos Gerados kg/m² por obra` da aba Consolidado: a Senna tem meta 50 (as demais 36,6)
  e intensidade muito acima, então **fica fora da visão consolidada** para não achatar as
  outras barras.
  - Visão consolidada: **8 obras**, linha de meta 36,6 em todas.
  - Botão **"Ver Senna Tower"** no cabeçalho do card (`res-ger-senna-btn` →
    `toggleSennaGer()`, linha ~1814) mostra a Senna mês a mês com meta 50.
  - Reaproveita `SENNA_RE` (linha ~1810) e o CSS `.res-toggle` já existentes.
  - O filtro global de obra tem **precedência** sobre o botão.
- Badge dinâmico (`res-ger-badge`): `Jan/26 – Abr/26 · meta 36,6 kg/m²` no consolidado;
  `4 meses · meta 50 kg/m²` no modo Senna.
- **Com uma obra filtrada, o eixo passa a ser de meses** (antes mostrava 1 coluna só).
- **Rótulo de dados** em cada barra.
- Canvas com `max-height:320px`.

---

## 3. Gráfico "Geração Resíduos (ton)"

**Antes:** `Composição por classe CONAMA` — donut de composição percentual.

**Depois:** barras verticais **agrupadas por classe CONAMA**.

- Eixo = obra/empreendimento (ou **meses**, quando uma obra está filtrada).
- Dentro de cada obra/mês, **uma barra separada por classe, lado a lado** — obra com duas
  classes mostra duas barras. Hoje **8 das 9 obras** exibem duas barras (A e B).
- Paleta CONAMA 307: `A #006B6B` · `B #336B87` · `C #C5C639` · `D #CE8686`.
- Legenda no topo; tooltip com o valor da classe e rodapé `Total` do grupo.
- **Rótulo de dados** em cada barra.
- O `canvas id="chart-res-conama"` foi **mantido** (não renomeado), apesar do nome não ter
  mais relação com CONAMA — renomear exigiria mexer em mais pontos sem ganho.

### ⚠️ Histórico de iterações deste gráfico (não reverter sem pedido)

Este gráfico mudou de dimensão três vezes. O estado atual é o desejado:

1. Primeira versão: barras **por obra** em toneladas.
2. Pedido *"ainda precisa ser por classe"* → refeito **só por classe** (perdeu a obra).
3. Pedido *"volte como estava antes, por empreendimento"* → voltou **só por obra** (perdeu a classe).
4. Pedido *"se tiver mais de uma classe precisa aparecer"* → **empilhado** por classe dentro da obra.
5. Pedido *"cada barra para cada classe, se o Scenarium tem duas classes ficam duas barras"*
   → **agrupado** (estado final). `scales.x/y.stacked` foi removido.

---

## 4. Gráfico "Volume por tipo de resíduo (top 8)"

**Antes:** barras horizontais, uma por tipo, total do período inteiro, sem reagir ao filtro de Período.

**Depois:** barras verticais **agrupadas — até 6 colunas (uma por mês) dentro de cada tipo**, até 8 tipos.

- Cada mês é um `dataset` do Chart.js (é assim que ele agrupa).
- Meses = os **últimos com lançamento dentro do recorte** (obra × período).
- Tipos = **top 8 por kg dentro desses meses** (antes o top 8 vinha pré-calculado no Python
  sem descartar zerados, então o gráfico exibia barras vazias).
- Escala sequencial azul, mês mais antigo claro → mais recente escuro:
  `#D6E7F8 #B5D4F4 #85B7EB #378ADD #185FA5 #0E4A6B`. Com menos de 6 meses a rampa é
  fatiada pela direita, para o mês mais recente continuar sendo o mais escuro.
- Badge novo (`res-tipo-badge`) com a faixa de meses: `Jan/26 – Abr/26`.
- Rótulos do eixo X truncados em 20 caracteres (nomes de tipo são longos, ex.:
  *"Não reciclado (fitas adesivas, lixas entre outros)"*); o nome completo aparece no
  título do tooltip.
- Canvas de `250px` → `340px`.
- **Rótulo de dados** com kg abreviado (`124k`, `6,4k`, `602`) — são até 48 barras estreitas.
- Passou a **reagir ao filtro de Período** (antes só reagia a obra).

**Bug corrigido no caminho:** quando o recorte de período não tinha dado (ex.: 2025),
o gráfico caía no agregado histórico e exibia dado de outro período. Agora: existe
`residuos_raw` + recorte vazio → placeholder *"Sem dados"*; **só** na ausência de
`residuos_raw` (caminho legado via Excel) volta às barras horizontais antigas.

---

## 5. Regra do eixo de meses (vale para os gráficos 2 e 3)

`gerar_dashboard.py`, `renderResiduos()` linha ~3826:

```js
const tetoMeses = (anoF === '12m') ? 12 : 6;
```

| Filtro de Período | Meses no eixo |
|---|---|
| todo o período | **6** |
| últ. 6 meses | 6 |
| últ. 12 meses | 12 |
| ano específico (2025 / 2026) | 6 |

Meses **sem lançamento não entram** — por isso hoje aparecem sempre 4 colunas
(Jan–Abr/26), que são os únicos meses com kg no `.pbix` atual.

---

## 6. Tabela "Série Mensal de Resíduos: Consolidado"

**Antes:** `Série mensal de resíduos`, colunas `kg · kg/m² · m³ · m³/m²`, badge `Meta: 50 kg/m²`.

**Depois:**

- Título: **Série Mensal de Resíduos: Consolidado**.
- Ordem das colunas invertida para **m³ → kg → m³/m² → kg/m²** (m³ como primeira coluna de dado).
- Cada cabeçalho ganhou a legenda da unidade (helper `subCab`, linha ~4169):

| Mês/Ano | m³<br>*volume* | kg<br>*peso resíduos* | m³/m²<br>*volume/m² construído* | kg/m²<br>*peso/m² construído* | Status |
|---|---|---|---|---|---|
| mai/25 | 2.148,00 | — | 0,673 | — | ● |

- Badge trocado para **`Senna: 50 · demais: 36,6 kg/m²`**, igual ao card da aba Consolidado
  (o anterior dizia "Meta: 50" para todas as obras, o que estava errado — a meta de 50 é
  só da Senna, e o próprio rodapé do card já informava 36,6 corretamente).

**Onde o card fica na tela:** é o **5º bloco** da aba, metade esquerda de uma linha de duas
colunas, logo abaixo do gráfico "Volume por tipo de resíduo", dividindo a linha com
"% Resíduos perigosos com destinação adequada".

---

## 7. Tabela "Resíduos por Obra"

**Antes:** `Detalhamento por obra`, meta como `36.6`, total como `190221`.

**Depois:**

- Título: **Resíduos por Obra**.
- Coluna **Meta** formatada como `36,6kg/m²` / `50kg/m²` — helper `metaFmt` (linha ~4243),
  sem decimal quando o valor é inteiro.
- Coluna **Total kg** com separador de milhar pt-BR: `190.221`.

---

## 8. Formatação pt-BR em toda a aba

Novo helper ao lado do `num()` existente (`gerar_dashboard.py` linha ~1456):

```js
// Número no padrão pt-BR: milhar "." e decimal "," (ex.: 1.234,56)
function numBR(v, dec = 0) {
  return v != null
    ? Number(v).toLocaleString('pt-BR', { minimumFractionDigits: dec, maximumFractionDigits: dec })
    : '—';
}
```

Aplicado em toda a `renderResiduos()`: os 3 KPIs, os tooltips dos gráficos, as 4 colunas da
Série Mensal e as colunas Kg/m² + Total kg da tabela por obra.

O `num()` antigo (`toFixed`) **continua existindo** e é usado pelas outras abas — elas não
foram padronizadas porque não foi pedido. Para estender, basta trocar `num` por `numBR` nas
demais funções de render.

---

## 9. Reatividade aos filtros (obra × período)

Toda a aba passou a respeitar os dois filtros, seguindo o **padrão híbrido** já usado no projeto:

- **"Todo o período" + "Todas as obras"** (view padrão) → usa os agregados estáticos
  pré-calculados no Python, para os números da tela inicial não mudarem.
- **Com qualquer filtro ativo** → recalcula a partir de `residuos_raw` (dataset granular
  obra × mês × classe × tipo), que já existia no `gerar_dashboard_pbix.py`.
- **Sem `residuos_raw`** (caminho legado via Excel) → fallback gracioso para o
  comportamento antigo, sem quebrar.

---

## 10. Rótulo de dados — reaproveitar `barLabels`, não instalar plugin

`gerar_dashboard.py` linha ~1762. Já existia no projeto, feito à mão **de propósito**:
o dashboard roda em rede interna e não se quis acrescentar dependência de CDN
(`chartjs-plugin-datalabels`). **Não criar plugin novo — usar este.**

```js
plugins: [barLabels(v => numBR(v, 2), { id: 'resger' })]
```

| Gráfico | Formato | id |
|---|---|---|
| Geração Resíduo | `numBR(v, 2)` → `33,45` | `resger` |
| Geração Resíduos (ton) | `numBR(v, 1)` → `118,0` | `reston` |
| Volume por tipo | kg abreviado → `124k`, `6,4k`, `602` | `restipo` |

Todos com `layout.padding.top = 18` para o rótulo caber. O `barLabels` já pula valores
`0`/`null` e **não rotula dataset `type: 'line'`** (a linha de meta).

---

## Layout final da aba Resíduos

| # | Bloco |
|---|---|
| 1 | Linha de 3 KPIs — Volume total · Geração Resíduo · Classes CONAMA |
| 2 | **Geração Resíduo** (kg/m², Senna à parte, botão de toggle) |
| 3 | **Geração Resíduos (ton)** (barras por classe CONAMA) |
| 4 | **Volume por tipo de resíduo (top 8)** (6 meses × 8 tipos) |
| 5 | **Série Mensal de Resíduos: Consolidado** |
| 6 | % Resíduos perigosos com destinação adequada |
| 7 | **Resíduos por Obra** |
| 8 | Auditorias ambientais |

---

## Significado das unidades — confirmado por aritmética no `.pbix`

Verificado nas tabelas do modelo, não apenas pelo nome das colunas. **Erro relativo mediano
0,000000% em todas as três relações:**

| Relação verificada | Resultado |
|---|---|
| `Kg` = `m³` × `Densidade` | **134/134 linhas exatas** |
| `Kg/m²` = `Kg` ÷ `m²` | **134/134 linhas exatas** |
| `Resíduos m³/m²` = `Resíduos m³` ÷ `Avanço Físico m²` | **128/128 linhas exatas** |

A primeira é a decisiva: existe uma coluna `Densidade` em `Fat_soma_residuos` e o **kg é
literalmente calculado como volume × densidade** (densidade de 20,7 a 483,4 kg/m³, variando
por tipo). Isso prova que m³ é volume e kg é peso, não duas medições independentes.
A área dos indicadores "por m²" é o **`Avanço Físico m²`**.

### De onde vem cada coluna da Série Mensal

| Coluna | Origem | Agregação do mês | Linha (`_pbix.py`) |
|---|---|---|---|
| m³ | `Fat_Residuos_m³_m²`.`Resíduos m³` | `.sum()` | 553 |
| kg | `Fat_soma_residuos`.`Kg` | `.sum()` | 540 |
| m³/m² | `Fat_Residuos_m³_m²`.`Resíduos m³/m²` | **`.mean()`** | 557 |
| kg/m² | `Fat_soma_residuos`.`Kg/m²` | **`.sum()`** | 543 |

> ⚠️ As duas colunas de intensidade usam agregações **diferentes** lado a lado na mesma
> tabela: `kg/m²` passou para soma em 15/09, mas `m³/m²` continuou como média. Ver
> pendência 5.

**Ressalva:** m³ e kg vêm de **tabelas diferentes**, e `Fat_soma_residuos` tem um m³ próprio
(o que gera o kg). Cruzando obra+mês, os dois m³ batem em **128/137 combinações (93,4%)** e
batem **100% nos 4 meses que têm kg**. As 9 divergências são todas de 2025
(a maior: Blue Coast jun/25, 28 vs 56 m³).

### Limitação de dado no `.pbix` atual

- Apenas **4 meses têm kg**: Jan/26 (161,7 t) · Fev/26 (203,5 t) · Mar/26 (225,4 t) · Abr/26 (198,5 t).
  Todo 2024/2025 está zerado em kg (2025 tem só m³) e Mai–Jul/26 ainda não têm lançamento.
- Apenas **6 tipos têm kg**: Entulho 558,9 t · Madeira 105,9 · Gesso 79,6 · Metal/Sucata 35,7 ·
  Papel/Papelão 6,3 · Plástico 2,7.
- Apenas **2 classes CONAMA**: A (558,9 t) e B (230,2 t). C e D sem registro.

Por isso o gráfico de tipo mostra hoje **4 colunas × 6 tipos**, não 6 × 8 — vai preencher
sozinho conforme entrarem meses novos.

---

## Pendências conhecidas

### 1. `kg/m²` somado acumula entre meses — distorce o ranking

Em 15/09 a agregação do kg/m² foi trocada de `.mean()` para `.sum()`
(`gerar_dashboard_pbix.py` linhas 442, 495, 515, 543). Somar **dentro do mês** está correto:
cada linha é um tipo de resíduo e a soma dá `Σkg ÷ área`, a intensidade do mês.

Mas `res_por_obra` soma **todas as linhas da obra**, o que inclui somar **entre meses** —
então o valor exibido é `intensidade_mensal × nº de meses com lançamento`, enquanto a meta
(36,6 / Senna 50) continua **mensal**:

| Obra | Mostra | Meses | Por mês | Meta |
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

**Blue Coast tem 1 mês lançado e por isso aparece como a 2ª melhor obra do gráfico, quando
por mês é a 2ª pior.** O número também cresce sozinho a cada mês novo.

Opções: (a) somar dentro do mês e tirar a **média dos meses** no nível da obra — preserva o
significado da meta; (b) manter a soma e usar meta `36,6 × n_meses`; (c) deixar o gráfico por
obra sempre mensal. No eixo mensal (obra filtrada ou modo Senna) o problema **não** aparece.

### 2. IPA continua com `.mean()`

`score_res()` (linha 264) e `_score_res_mes()` (linha 364) do `gerar_dashboard_pbix.py` foram
**mantidas com média** de propósito. Resíduos pesa 20% no IPA, então estender a soma para lá é
decisão separada. Impacto estimado se mudar: 7 das 9 obras perderiam pontos
(ex.: Blue View de 18,7 para ~6,3 de 20).

### 3. Título "(top 8)" no gráfico de tipo

O título continua `Volume por tipo de resíduo (top 8)`, mas o gráfico agora também é por mês e
hoje mostra 6 tipos (os que têm kg > 0). Pode valer ajustar para algo como
*"Volume por tipo e mês"*.

### 4. `m³/m²` ainda é média enquanto `kg/m²` virou soma

Na tabela Série Mensal, as duas colunas de intensidade usam agregações diferentes:
`kg_m2_by_month` usa `.sum()` (linha 543) e `m3_m2_by_month` usa `.mean()` (linha 557).
A troca de 15/09 alcançou só o kg/m². Pelo mesmo raciocínio — cada linha é a intensidade de
um tipo naquele mês — o m³/m² provavelmente também deveria ser soma, mas isso **não foi
verificado** (a checagem aritmética confirmou a fórmula da coluna, não qual agregação a área
espera ver). Confirmar com a área antes de mexer.

### 5. Erro na fonte — Senna Tower Abr/26

`Fat_soma_residuos` tem **área = 1,00 m²** para Senna Tower em Abr/26, gerando 1.686,70 kg/m²
naquele mês e distorcendo a obra. Conferir o lançamento na planilha de origem.

---

## Validação

Suíte em Node (sandbox `vm` com stubs de DOM e Chart.js) que carrega o `dashboard_fg.html`
real gerado do `.pbix` e confere os valores renderizados contra `residuos_raw` recalculado à
parte. **Rodar as 5 depois de qualquer alteração ou restauração:**

| Script | Cobre |
|---|---|
| `validate_final.js` | Rótulos de dados, toggle da Senna, eixo de meses base 6 |
| `validate_eixos.js` | Ton agrupado por classe, paleta CONAMA, célula por célula |
| `validate_ptbr.js` | Varredura por decimal com ponto em todo o texto renderizado |
| `validate_tabelas.js` | Ordem/legendas da Série Mensal, formato de Meta e Total kg |
| `validate_tipo.js` | Matriz tipo × mês, top 8, rampa de cor, placeholder |

Os scripts vivem no diretório de scratchpad da sessão. **Vale movê-los para
`tests/` no repositório** — hoje eles se perdem junto com a sessão.

> ⚠️ Os scripts guardam expectativas fixas. Quando um falha, **conferir primeiro se não é só o
> teste desatualizado** — já aconteceu duas vezes (formato pt-BR e assinatura de tooltip).

---

## Armadilhas encontradas no processo

**1. Escrita não atômica truncou o gerador.** `io.open(p, 'w')` **trunca o arquivo no open** —
se o `.write()` falhar (no caso, `UnicodeEncodeError` por usar pares surrogados
`\ud83d\udd34` em vez do codepoint `\U0001F534` para emoji), o `gerar_dashboard.py` fica com
**0 bytes**. Sempre: codificar primeiro, gravar num `.tmp` e `os.replace`.

**2. `grep` com acento falha** neste ambiente (Git Bash + UTF-8): `grep "Resíduos por Obra"`
dá falso negativo. Usar Python (`io.open(..., encoding='utf-8')` + `count`) para qualquer
verificação com acento.

**3. Backup errado restaurado.** Em 15/09 o `_backup_teste_refresh\gerar_dashboard_20260911_104920_antes_consolidado.py`
(**166 KB**, estado de 11/09 *antes* de todo o trabalho) foi restaurado em cima do bom,
apagando as alterações. **Referência de tamanho: ~203 KB = estado com tudo; ~166 KB = estado
antigo de 11/09.** Os nomes dos backups não deixam isso claro — vale renomeá-los.

Marcadores para conferir se o arquivo restaurado é o certo (todos têm de existir):

```
function barLabels · function numBR · toggleSennaGer · res-ger-senna-btn · tetoMeses
classesTon · kgTipoMes · labelMesRes · SENNA_RE · subCab · metaFmt
"Série Mensal de Resíduos: Consolidado" · "Resíduos por Obra" · res-barras-titulo
```

---

## Fluxo de deploy

1. Editar `gerar_dashboard.py` (visual/lógica) e/ou `gerar_dashboard_pbix.py` (dados)
   — **sempre em homologação 10.1.1.155, nunca em produção**.
2. Regenerar: `python gerar_dashboard_pbix.py` (com o interpretador que tem `pbixray`).
3. Rodar a suíte de validação.
4. Validar visualmente em **http://10.1.1.155:8090/** (Ctrl+F5 — o serviço não manda
   header de no-cache).
5. Só depois de aprovado, promover para produção (10.1.1.151:8080).

> **Nota de ambiente:** `pbixray` está em versões diferentes nos dois servidores
> (0.9.0 em produção, 0.15.5 em HML). A diferença muda nulo→0 vs NaN e, sozinha, explica
> divergências de IPA (79,1 vs 78,4). Considerar alinhar as versões.
