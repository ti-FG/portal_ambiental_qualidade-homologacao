# Auditoria: filtro de Período em todos os gráficos e cards

**Data:** 18 de setembro de 2026
**Arquivo alterado:** `gerar_dashboard.py`
**Motivo:** depois do lote anterior (`a244fce`) que ligou o filtro de Período à
maioria dos gráficos, a área reportou que "nem todos os gráficos estão com
ação atribuída ao filtro". Este documento é o resultado de uma auditoria
função por função de **todo canvas e todo card** das 3 abas com filtro
(Consolidado, Obra Individual, Resíduos), e as 4 correções que saíram dela.

---

## Método

Levantei os 20 `<canvas>` e todos os cards/tabelas das 3 abas, e para cada um
rastreei a função de render até a origem do dado, perguntando: "esta função
recebe `anoF`? E se recebe, ela realmente usa pra filtrar, ou só decora o
parâmetro?" A maioria (17 dos 20 gráficos, mais heatmap/Top NCs/reclamações/
tabelas de resíduos) já estava correta desde o commit anterior. Quatro pontos
não estavam.

---

## 1. "Obras em risco crítico" (Consolidado) — nunca recalculava

### Causa

O card era montado **uma única vez**, dentro de `renderConsolidado()`, a partir
de `D.obras_risco` — uma lista estática calculada no Python sobre o **histórico
completo** (`conf_atual[o].perc < 85`). `applyFilters()` nunca chamava essa
lógica de novo, então trocar Obra ou Período não tinha efeito nenhum ali —
único card do Consolidado nesse caso.

### Correção

Nova função `renderObrasRisco(obraF, anoF)`, chamada tanto em
`renderConsolidado()` (carga inicial) quanto em `applyFilters()`. Recalcula a
conformidade média de cada obra dentro do recorte de período usando
`D.conf_heatmap` (mesma fonte do heatmap); sem dado no recorte, cai no
agregado histórico da obra — mesmo padrão de fallback já usado no IPA e no
ruído por obra.

### Evidência

| Período | Obras em risco |
|---|---|
| Todo o período | Blue Coast (76,8%) · Blue View (83,4%) |
| 2025 | Grand Place (69,9%) · Blue Coast (71,2%) |
| 2026 | Blue Coast (67,0%) · Blue View (71,5%) · Scenarium (78,1%) · Garden Park (84,2%) |
| Últimos 6 meses | 4 obras, percentuais diferentes de 2026 |
| Últimos 12 meses | 3 obras, percentuais diferentes de todos os outros |

---

## 2. "Plano Ação Auditorias" na Obra Individual — sem aviso

### Causa

O donut `chart-obra-pa` sempre usa `D.pa_por_obra[obra]`, o agregado histórico
completo — igual ao card equivalente do Consolidado. A diferença é que o
Consolidado **avisa** no badge ("histórico completo (sem data na fonte)")
desde o lote anterior, e o card da Obra Individual **não tinha badge nenhum**
— parecia um bug silencioso em vez de uma limitação conhecida e documentada.

### Por que não dá pra filtrar de verdade

`Dim_Plano_acao` não tem data de registro da ação, só `Prazo` (data futura de
conclusão). Filtrar por prazo responderia "ações que vencem no período", uma
pergunta diferente. Ver `ALTERACOES_FILTRO_PERIODO_E_SETOR.md` §4 para o
histórico completo dessa decisão.

### Correção

Adicionado `<span class="card-badge" id="obra-pa-badge">` ao card, preenchido
em `renderObra()` com o mesmo texto e a mesma lógica do card do Consolidado.

---

## 3. "Emergências ambientais — analítica" na Obra Individual — não filtrava

### Causa

`renderObraEmergencias(obra)` filtrava só por obra; o parâmetro de período
nunca existiu na função. Diferente do card equivalente do Consolidado
(`renderEmergencias`, já correto), que usa a mesma lista `emergencias_lista`.

### Correção

Função passou a receber `anoF` e filtrar por `mesesDoPeriodo()`, no mesmo
padrão de `renderEmergencias`. O rótulo de escopo no cabeçalho da tabela e o
badge agora mostram o período quando há um selecionado.

### Evidência (obra Senna Tower)

| Período | Registros |
|---|---|
| Todo o período | 5 |
| 2025 | 0 |
| 2026 | 5 |
| Últimos 6 meses | 2 |
| Últimos 12 meses | 5 |

---

## 4. "Resíduos por Obra" (aba Resíduos) — só filtrava por obra

### Causa

A tabela usava sempre `D.res_por_obra[o]`, o agregado histórico pronto — o
filtro de obra funcionava (linha some/aparece), mas o **valor de kg/m² e o
total em kg mostrados nunca mudavam** com o período.

### Correção

Quando há período ativo (e existe o dataset granular), a tabela recalcula
`kg_m2` (média) e `total_kg` (soma) a partir de `filtRes` — o mesmo recorte
obra × período já usado pelos outros gráficos desta função. Em "todo o
período" continua usando o agregado pronto, sem custo extra de recomputar.

### Evidência (obra Blue Coast)

| Período | Kg/m² | Total kg |
|---|---|---|
| Todo o período | 83,69 | 14.397 |
| 2025 | — | 0 |
| 2026 | 16,74 | 14.397 |
| Últimos 12 meses | 16,74 | 14.397 |

O salto de 83,69 (agregado) para 16,74 (2026) é esperado e é justamente a
pendência já documentada em `homologacao_ambiental.md` §8: o agregado soma
entre meses, então uma obra com poucos meses lançados parecia "melhor" do
que era por mês. Este fix não resolve aquela pendência — só faz a tabela
reagir ao período, que é o que foi pedido agora.

---

## Limpeza feita durante a auditoria

`renderRuidoBarras()` era uma função inteira (~85 linhas) que não era chamada
de lugar nenhum — resquício de uma versão anterior do gráfico de ruído,
substituída por `renderRuidoObra()`. Referenciava um `<canvas id="chart-
ruido-barras">` que também não existe mais no HTML. Removida por ser código
morto, sem relação com o bug relatado.

---

## O que foi conferido e está correto (sem mudança)

Auditados e confirmados como já reagindo corretamente a obra × período:
heatmap de conformidade, Top NCs recorrentes, resíduos por obra (gráfico kg/m²
e Senna Tower), reclamações (gráfico + ranking + tabela detalhe), ruído
(Consolidado e Obra Individual), IPA por obra, auditorias, água, energia,
impressões, geração de resíduos (kg/m² e ton, nas duas abas), volume por tipo
de resíduo, série mensal de resíduos, % resíduos perigosos, KPIs do
Consolidado e da Obra Individual.

**Dois casos ficam de propósito fora do filtro, e agora os dois avisam:**
Plano de Ação (Consolidado e Obra Individual) — sem data de registro na fonte
— e SGA, que só reage às opções de ano porque a fonte é anual, sem mês.

---

## Validação feita

- [x] Geração sem erro — IPA global 78,4 mantido (nenhuma mudança de dado)
- [x] "Obras em risco crítico" muda de obra e de quantidade nas 5 opções de período
- [x] Badge do Plano de Ação (Obra Individual) mostra o aviso quando há período
- [x] Emergências (Obra Individual): 5 → 0 → 5 → 2 → 5 registros nas 5 opções
- [x] Resíduos por Obra: kg/m² e total kg mudam com o período
- [x] Varredura completa: 2 obras × 5 períodos × 4 abas (40 combinações) sem
      erro no console
- [x] `renderRuidoBarras` removida sem quebrar nada (não era chamada)
