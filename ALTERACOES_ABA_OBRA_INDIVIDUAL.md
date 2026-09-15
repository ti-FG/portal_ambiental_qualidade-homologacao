# Alterações nos Gráficos — Dashboard Ambiental Qualidade

## Resumo Geral

Série de melhorias visuais e funcionais nos gráficos da aba **Obra Individual** do dashboard, implementadas e validadas em homologação (servidor 10.1.1.155) entre **11/09 e 15/09/2026** antes de promoção para produção (10.1.1.151).

---

## Alterações Aplicadas

### 1. IPA por Obra — Período Reativo & Histórico de 12 Meses

**Arquivo:** `gerar_dashboard.py`, `gerar_dashboard_pbix.py`

**Problema:** 
- Filtro de Período não afetava o gráfico IPA (mostrava valor único agregado)
- Obra individual mostrava apenas 1 coluna (sem histórico)

**Solução:**
- Adicionado `ipa_mensal[obra]` — série de até 12 meses calculada mês a mês
- Conformidade, Resíduos, Ruído e Auditoria recalculados por mês com fallback para agregado
- Plano de Ação mantém valor constante (sem histórico confiável na fonte)
- `renderIpaObras()` agora:
  - **Obra selecionada** → histórico de 12 meses
  - **"Todas as obras"** → média por período, reage ao filtro
  - Fallback gracioso se `D.ipa_mensal` não existir

**Impacto visual:** Gráfico de barras agora exibe histórico temporal ou comparativo entre obras conforme período.

---

### 2. Rótulos de Dados (Data Labels) nos 4 Gráficos

**Arquivo:** `gerar_dashboard.py` (plugin `barValueLabels()` linhas 1495–1501, plugin `donutValueLabels()` linhas 1529–1552)

**Gráficos afetados:**
- IPA por obra — Índice de Performance Ambiental
- Ruído — medições
- Geração Resíduos (ton)
- Plano Ação Auditorias

**Implementação:**
- **`barValueLabels()`** — renderiza valor numérico acima de cada barra
  - Fonte: 600 11px Segoe UI
  - Cor: #3A3833 (escuro)
  - Padding-top de 18px para evitar sobreposição
  
- **`donutValueLabels(formatFn)`** — renderiza valor no centro de cada fatia
  - Fonte: 600 11px Segoe UI, branco
  - Centralmente alinhado
  - Parâmetro `formatFn` permite formatação customizada:
    - Geração Resíduos: `(v) => numBR(v / 1000, 1) + ' t'` (converte kg → toneladas)
    - Plano Ação: `(v, total) => \`${v} (${total > 0 ? Math.round(v / total * 100) : 0}%)\`` (valor + percentual)

**Impacto visual:** Valores agora legíveis diretamente no gráfico sem precisar hover.

---

### 3. Remoção do Card NC Registradas LAI/EIV

**Arquivo:** `gerar_dashboard.py`, linhas 3414–3440 (função `renderObra()`)

**Mudanças:**
- Card "NC registradas LAI/EIV" removido da seção de KPIs
- Grid de KPI redimensionado de **4 colunas** para **3 colunas**
- KPIs mantidos: IPA, Conformidade, Reclamações
- Variável `nc` removida do escopo

**Impacto visual:** Layout mais limpo; dados de NC agora integrados em outro KPI baseado em resíduos.

---

### 4. NC Contados por Soma de kg/m² (Não Média)

**Arquivo:** `gerar_dashboard.py` (linhas 529–575), `gerar_dashboard_pbix.py` (linhas 917–953)

**Problema:** 
- Agregação por **média** subestimava NCs; 7 de 9 obras apareciam dentro da meta quando realmente estavam acima

**Solução:**
- Mudado para agregação por **soma** simples
- NC = 1 por obra-mês onde Σ(kg/m²) > meta
- Meta: Senna=50 kg/m², demais=36,6 kg/m²
- Função `_residuos_kgm2_mes()` agora: `d['kg_m2'].sum(skipna=True)`

**Impacto numérico:** total_ncs: 8 (sem feature) → 10 (com média) → **32 (com soma)** — contagem acurada.

---

### 5. Aumento do Gráfico Ruído & Redimensionamento dos Donuts

**Arquivo:** `gerar_dashboard.py`, linhas 1301, 1307, 1313 (canvas styling), linhas 1284, 1305, 1307, 1313 (grid)

**Mudanças:**
- Grid alterado de **4 colunas** para **3 colunas** (cards mais largos)
- `chart-obra-ruido` (Ruído — medições):
  - **200px** → **240px** → **300px** (final)
  - Mais altura para legibilidade
  
- `chart-obra-res-donut` (Geração Resíduos ton):
  - **200px** → **140px** (intermediário) → **170px** (final)
  - 170px mantido como compromisso visual
  
- `chart-obra-pa` (Plano Ação Auditorias):
  - **200px** → **140px** → **170px** (final)

**Impacto visual:** Gráfico de ruído mais destacado; donuts com tamanho adequado para renderizar rótulos internos.

---

### 6. Legenda Compacta nos Donuts

**Arquivo:** `gerar_dashboard.py`, linhas 3606 e 3632 (opções dos donuts)

**Mudanças:**
- Legenda posicionada à **direita** (`position: 'right'`)
- Dimensões reduzidas:
  - `boxWidth`: 12 → **8** (caixinha de cor menor)
  - `padding`: 12 → **6** (espaçamento reduzido)
  - `font.size`: 11 → **10** (fonte menor, explícita)
- **Sem** `layout.padding` grande (removido o padding assimétrico de 58px)

**Impacto visual:** Legenda ocupa menos espaço; anel do donut fica mais centralizado e aparentemente maior no card estreito (3 colunas).

---

## Estrutura de Arquivos Alterados

```
C:\Apps\Dash_ambiental_qualidade\
├── gerar_dashboard.py              (template HTML master)
│   ├── Linhas 529–575    : Cálculo NC de resíduos (soma)
│   ├── Linhas 1284–1319  : Layout de grid + canvas heights
│   ├── Linhas 1495–1552  : Plugins barValueLabels + donutValueLabels
│   ├── Linhas 2684–2691  : IPA chart com barValueLabels
│   ├── Linhas 2755–2777  : Ruído chart com barValueLabels
│   ├── Linhas 3414–3440  : renderObra() sem NC card
│   ├── Linhas 3600–3608  : chart-obra-res-donut com donutValueLabels + legenda compacta
│   ├── Linhas 3626–3635  : chart-obra-pa com donutValueLabels % + legenda compacta
│   └── Linhas 3637–3639  : renderIpaObras() com histórico mensal
│
├── gerar_dashboard_pbix.py         (gerador de produção, extrai do .pbix)
│   ├── Linhas 917–953    : Cálculo NC de resíduos (mirrored, usa sum())
│   └── Demais dados vêm do template
│
└── dashboard_fg.html               (saída regenerada, deployado em 155:8090)
    └── Contém todos os gráficos + estilos + plugins inline
```

---

## Processo de Validação

### Staging (10.1.1.155)
1. Modifica `gerar_dashboard.py`
2. Executa `py_compile gerar_dashboard.py` (syntax check)
3. Executa `gerar_dashboard_pbix.py` (regenera `dashboard_fg.html`)
4. Acessa **http://10.1.1.155:8090/** para validação visual
5. Usuária confirma no navegador

### Produção (10.1.1.151)
- **Nunca** alteração direta
- Após aprovação em homologação, mesmos arquivos copiados para produção e refresh disparado

---

## Notas Técnicas

### donutValueLabels Plugin
- Desenha rótulo no **centro de cada fatia** usando `arc.getCenterPoint()`
- Versão anterior (descartada) tentava rótulos **externos** com linha-guia, mas ocupava muito espaço
- Formato customizável via `formatFn(valor, total)`

### barValueLabels Plugin
- Desenha valor **acima** de cada barra
- Usa `ctx.fillText()` com `textAlign: 'center'`
- Layout padding-top ajustado para não colidir com o gráfico

### IPA Mensal
- Calcula 12 meses retroativos de conformidade/resíduos/ruído/auditoria
- Fallback para agregado quando dado mensal não disponível
- Plano de Ação mantém valor anual (sem coluna de data confiável)

### NC de Resíduos
- Agregação por **soma** de kg/m² por obra-mês
- 1 NC = 1 obra-mês com Σ(kg/m²) > meta_obra
- Meta: Senna Tower = 50, demais = 36,6
- Aplicado em ambos os geradores (Excel template + PBIX)

---

## Data & Autoria

- **Período:** 11/09/2026 — 15/09/2026
- **Servidor alvo:** Homologação 10.1.1.155 (validação), depois produção 10.1.1.151
- **Arquivos:** Gerador template + PBIX + outputs HTML
- **Responsável validação:** Natália (usuária final, visual feedback)

---

## Próximas Ações (Se Necessário)

- Monitorar feedback visual em produção após deploy
- Se ajustes forem necessários, sempre aplicar em HML primeiro (10.1.1.155)
- Considerar consolidar `donutValueLabels` e `barValueLabels` em um módulo compartilhado se mais gráficos forem adicionados
