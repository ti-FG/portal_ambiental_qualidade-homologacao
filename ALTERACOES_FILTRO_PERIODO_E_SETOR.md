# Filtro de Período em todos os gráficos · filtro de Setor em água e energia

**Data:** 16 de setembro de 2026
**Arquivos alterados:** `gerar_dashboard.py` (template) e `gerar_dashboard_pbix.py` (grão do dataset de água/luz)

Quatro alterações pedidas pela área, documentadas abaixo na ordem em que foram feitas.

---

## 1. Fonte do Dicionário +10%

Todas as oito regras de tamanho de fonte da aba Dicionário subiram exatamente 10%:

| Elemento | Antes | Depois |
|---|---|---|
| `.dic-table` (corpo das tabelas) | 11,5px | 12,65px |
| `.dic-intro` / `.dic-vazio` | 12px | 13,2px |
| `.dic-search` | 12px | 13,2px |
| `.dic-chip` / `.dic-code` / `.dic-peso-num` | 11px | 12,1px |
| `.dic-nota` | 10,5px | 11,55px |

Nenhuma outra aba foi afetada — todas as regras alteradas são do prefixo `.dic-`.

---

## 2. Filtro de Setor nos blocos de água e energia

Novo `<select>` "Setor" no cabeçalho de cada bloco, ao lado do filtro de período
que já existia.

### Precisou mudar o gerador

O dataset granular (`consumo_agua_raw` / `consumo_energia_raw`) era agregado por
**obra × mês**, descartando o setor. Passou a ser **obra × setor × mês**, e um
novo `consumo_setores_list` alimenta os dois selects.

### Por que o setor é uma dimensão de verdade aqui

A coluna `nome_setor` de `Fat_consumo_agua_luz` mistura duas coisas:

- **Obras**, em caixa alta: `BLUE COAST`, `SENNA`, `TITANIUM`…
- **Locais que não são obra**, com `Obra = 'Outros'`: `Sede`, `Depósito`,
  `Sala 02`…`Sala 07`, `Plantão brava`, `Plantão norte`, `Plantão Senna`,
  `ITC Busness SL 612/613`, `Tour Royalle`, `Assistência técnica`

Sem o setor no grão não havia como separar o consumo da Sede do consumo das
obras. O novo helper `setor_col()` normaliza a grafia com o mesmo `norm_obra()`
das demais tabelas (para `BLUE COAST` casar com `Blue Coast`) e devolve
`Sem setor` quando a célula está vazia, para o filtro nunca perder linha.

**25 setores** entraram na lista. Exemplo de efeito: energia total de
1.338.374 kWh cai para 140.488 kWh ao filtrar `Sede`.

---

## 3. "Volume por tipo de resíduo" na aba Obra Individual

O gráfico que já existia na aba Resíduos passou a aparecer também na Obra
Individual, no recorte da obra selecionada.

**Não foi copiado e colado.** O bloco foi extraído de dentro de
`renderResiduos()` para uma função nova, `renderResTipo(canvasId, badgeId,
filtRes, tipoSrc, hasRaw)`, chamada pelas duas abas. Assim o critério é
literalmente o mesmo nos dois lugares — se a regra mudar, muda nos dois.

A função recebe o recorte **já filtrado** em vez de refazer o filtro por dentro,
porque cada aba monta o recorte de um jeito (a aba Resíduos tem o conceito de
"todas as obras", a Obra Individual sempre tem uma obra em foco).

---

## 4. Filtro de Período ligado a todos os gráficos

Antes, o filtro de Período do topo só afetava o heatmap de conformidade, o Top
NCs, o IPA da aba Obra Individual e a aba Resíduos. Os demais gráficos mostravam
sempre o histórico completo, sem avisar.

### Helper central

```js
mesesDoPeriodo(anoF, universo)   // → Set de meses 'YYYY-MM'
periodoGlobal()                  // valor corrente do filtro do topo
periodoEfetivo(idSelectCard, padraoCard)  // precedência card > global
```

O recorte é calculado **sobre o universo de meses da própria fonte**, não sobre
um calendário fixo: ruído tem meses que resíduos não têm, e vice-versa. Então
"últimos 6 meses" significa os 6 últimos meses **com dado naquela fonte** — que
é como o dashboard já se comportava nos gráficos que filtravam.

### Bug corrigido de passagem

`_filtraPeriodo()` (água e energia) tratava só `6m` e `12m`. Escolher **2025** ou
**2026** devolvia a série inteira, silenciosamente. Agora recorta pelo ano do
próprio mês.

### O que passou a reagir

| Gráfico / KPI | Fonte usada para o recorte |
|---|---|
| Resíduos Gerados kg/m² por obra | `residuos_raw` |
| Resíduos Gerados — Senna Tower | `residuos_raw` |
| Volume por tipo de resíduo (2 abas) | `residuos_raw` |
| Reclamações por obra + detalhamento | `reclamacoes_lista.mes_ano` |
| Monitoramento de ruído | `ruido_obra_mes` |
| Emergências ambientais | `emergencias_lista.mes_ano` |
| Consumo de água (m³ e m³/m²) | `consumo_agua_raw` |
| Consumo de energia (kWh e kWh/m²) | `consumo_energia_raw` |
| Impressões | `imp_setor_mes` |
| Auditorias | `auditorias_obra_mes` |
| KPIs do Consolidado | ver abaixo |

### KPIs do Consolidado

Nova função `kpisDoPeriodo(obraF, anoF)`, que devolve `null` quando não há
período selecionado (aí valem os agregados prontos, sem recomputar nada):

- **IPA** — média dos meses do recorte em `ipa_mensal`
- **Conformidade** — média das células do heatmap no recorte
- **Resíduos kg/m²** — soma das linhas do recorte (mesmo critério da aba Resíduos)
- **Reclamações** — contagem real no recorte
- **Plano de ação** — permanece o agregado (ver abaixo)

### Precedência dos filtros de card

Água, energia, impressões e auditorias têm filtro de período próprio. A regra
aplicada: **o filtro do card vence quando o usuário escolheu algo nele**; quando
está no valor padrão, o card obedece ao filtro global do topo. Ninguém perde o
controle fino que já tinha, e o filtro global passou a ter efeito onde antes não
tinha.

---

## Duas limitações que permanecem — de propósito

### Plano de Ação não reage ao período

É o **único** gráfico que continua fora do recorte. `Dim_Plano_acao` não tem data
de registro da ação; tem só `Prazo`, que é data futura de conclusão. Filtrar por
prazo responderia *"ações que vencem no período"*, que é outra pergunta.

Em vez de filtrar errado em silêncio, **o card avisa**: com período selecionado,
o badge passa a mostrar `histórico completo (sem data na fonte)`.

Para o plano de ação reagir ao período de verdade, a planilha precisa ganhar uma
coluna de data de abertura da ação. É decisão da área.

### SGA só reage às opções de ano

`Dim_Acoes_SGA` é agregada por ano, sem mês. Com **2025** ou **2026** o gráfico
recorta; com "últimos 6/12 meses" mantém a série completa, porque recortar um ano
inteiro por causa de meses seria inventar precisão que o dado não tem.

---

## Validação feita

Dashboard gerado neste repositório e testado no navegador:

- [x] Geração sem erro — **IPA global 78,4** (valor esperado de homologação)
- [x] Alternando o período de "todo o período" para "2025", **todos** os gráficos
      mudaram, exceto o Plano de Ação (esperado e avisado no badge)
- [x] As cinco opções de período (all / 6m / 12m / 2025 / 2026) testadas: água
      vai de 12 → 6 → 12 → 12 → 3 pontos; IPA de 78,4 → 78,3 → 78 → 78 → 78
- [x] Filtro de setor: 26 opções nos dois blocos; `Sede` reduz energia de
      1.338.374 para 140.488 kWh; `Sala 02` dá 28 m³ de água
- [x] "Volume por tipo de resíduo" na Obra Individual com Senna Tower: 5 tipos ×
      4 meses, badge `Jan/26 – Abr/26`
- [x] Fontes do dicionário conferidas no navegador: 12,65 / 13,2 / 12,1 / 11,55px
- [x] As quatro abas abrem sem erro, com e sem obra selecionada
- [x] Console do navegador sem erro

---

## Pendências

1. **Coluna de data em `Dim_Plano_acao`** — é o que falta para o Plano de Ação
   entrar no filtro de período. Depende da área.
2. **Revisar os 25 setores** com a área: alguns nomes vêm da planilha com grafia
   inconsistente (`ITC Busness` está assim na origem, com o erro de digitação).
3. O gráfico de tipo ainda se chama "Volume por tipo de resíduo (top 8)" na aba
   Resíduos, mas nem sempre são 8 tipos — pendência antiga, não tratada aqui.
