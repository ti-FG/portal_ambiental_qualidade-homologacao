# Fonte de títulos e subtítulos dos cards e gráficos +10%

**Data:** 17 de setembro de 2026
**Arquivo alterado:** `gerar_dashboard.py`
**Abas afetadas:** Consolidado, Obra individual, Resíduos — **Dicionário ficou de fora**, por pedido explícito.

## O que subiu 10%

| Elemento | Onde | Antes | Depois |
|---|---|---|---|
| Título do card (`.card-title`) | Consolidado · Obra · Resíduos | 14px | 15,4px |
| Badge/subtítulo do card (`.card-badge`) | Consolidado · Obra · Resíduos | 12px | 13,2px |
| Título de bloco (`.imp-section-title`, ex.: "Consumo de energia") | Consolidado | 11px | 12,1px |
| Título de eixo dos gráficos Chart.js (kg/m², dB, kg, ton, Pontuação, Pessoas treinadas) | Todos os 20 gráficos | 11px (herdado de `Chart.defaults.font.size`) | 12,1px explícito |

**O que não mudou de propósito:** legendas, rótulos de eixo (ticks), tooltips e
rótulos de dados sobre as barras — o pedido foi só títulos e subtítulos, não a
fonte inteira dos gráficos.

## Como foi escopado

`.card-title`, `.card-badge` e `.imp-section-title` são classes **compartilhadas**
com a aba Dicionário (o cabeçalho "Dicionário de dados e termos" também usa
`.card-title`). Para não vazar a mudança para lá, a regra base ficou intacta e
três seletores por ID foram adicionados por cima:

```css
#page-consolidado .card-title, #page-obra .card-title, #page-residuos .card-title { font-size: 15.4px; }
#page-consolidado .card-badge, #page-obra .card-badge, #page-residuos .card-badge { font-size: 13.2px; }
#page-consolidado .imp-section-title, #page-obra .imp-section-title, #page-residuos .imp-section-title { font-size: 12.1px; }
```

Os títulos de eixo do Chart.js não tinham `font.size` próprio — herdavam o
padrão global `Chart.defaults.font.size = 11`. Como não há gráfico na aba
Dicionário, bastou adicionar `font: { size: 12.1 }` explícito nos 14 pontos do
código que declaram `title:` ou `subtitle:` de eixo (regex único, sem risco de
pegar texto errado): 11 títulos de eixo (kg/m², dB, kg, ton, Pontuação, Pessoas
treinadas) e os 3 placeholders "Sem dados" que aparecem quando um gráfico fica
vazio no recorte.

## Complemento — subtítulo dos KPIs (17/09/2026)

Pedido de acompanhamento: os KPIs no topo (IPA Global, Conformidade LAI/EIV,
Resíduos Gerados etc.) são um componente **diferente** de `.card-title`/
`.card-badge` — usam `.kpi-label` (título em caixa alta) e `.kpi-sub` (texto
menor abaixo do valor, ex.: "Índice de Performance Ambiental", "média LAI/EIV").
O pedido original não cobria esse componente.

`.kpi-sub` subiu de 10px para **11px**, mesmo escopo por página
(`#page-consolidado`, `#page-obra`, `#page-residuos`). Esses KPIs são montados
por JS (`innerHTML` de `updateKpisFiltered()` e `renderObra()`), então o
seletor por `#page-id` precisa alcançar tanto o HTML estático quanto o gerado
em runtime — confirmado filtrando por obra e checando que o subtítulo
continua em 11px depois da reescrita do DOM.

`.kpi-label` (o título em caixa alta acima do valor) não foi alterado — o
pedido foi só subtítulo.

## Validação feita

- [x] `.card-title` = 15,4px e `.card-badge` = 13,2px nas 3 abas
- [x] `.card-title` da aba Dicionário continua em **14px** (não mudou)
- [x] `.imp-section-title` = 12,1px
- [x] Os 11 gráficos com título de eixo confirmados em 12,1px via
      `Chart.instances[...].options.scales.y.title.font`
- [x] IPA global 78,4 mantido (nenhuma mudança de dado/regra)
- [x] Console do navegador sem erro
- [x] `.kpi-sub` = 11px nas 3 abas, inclusive após reescrita via JS (testado
      filtrando por obra); aba Dicionário não tem KPIs (sem risco de vazamento)
