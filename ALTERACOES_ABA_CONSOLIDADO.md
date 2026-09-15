# Changelog — Dashboard Indicadores Ambientais

Registro das alterações feitas no gerador do dashboard. Todo o trabalho foi
feito e validado em **homologação (10.1.1.155)**, servido em
<http://10.1.1.155:8090/>, antes de qualquer promoção para produção
(10.1.1.151, porta 8080).

---

## 2026-09-11 a 2026-09-15 — Revisão da aba Consolidado

Escopo: renomear e recalcular indicadores da aba Consolidado, corrigir defeitos
de dados encontrados no caminho e reorganizar o layout.

### Arquivos alterados

| Arquivo | Antes | Depois | Papel |
|---|---|---|---|
| `gerar_dashboard_pbix.py` | 62.655 B | 72.251 B (1.407 linhas) | Lê o `.pbix` e calcula os indicadores |
| `gerar_dashboard.py` | 166.201 B | 203.247 B (4.530 linhas) | Template HTML/JS (não lê dados) |
| `requirements.txt` | — | novo, 31 linhas | Fixa versões das libs |

> **Atenção ao editar:** `gerar_dashboard_pbix.py` extrai o bloco
> `HTML = r"""..."""` de dentro de `gerar_dashboard.py` por busca de texto.
> Mudar a formatação desse trecho quebra a geração silenciosamente.

---

## 1. Medidas recalculadas

### 1.1 NC registradas LAI/EIV — de 32 para 20

**Antes:** somava duas coisas diferentes — 8 linhas da tabela `Dim_NC_LAI_EIV`
mais 24 "NCs de resíduos" derivadas de `Fat_soma_residuos` (obra-mês acima da
meta de kg/m²). O card dizia "LAI/EIV", mas 75% do valor não era LAI/EIV.

**Agora:** conta **assuntos** das colunas `NC LAI` e `NC EIV` de
`Dim_NC_LAI_EIV`, separados por `|`. Fonte única.

```
"Armazenamento e separação de resíduos | Ruído | Placa da LAI"  →  3 NCs (era 1)
```

Resultado: **14 assuntos LAI + 6 assuntos EIV = 20**. A coluna `NC EIV`, que era
totalmente ignorada, passou a contar.

O cálculo das NCs de resíduos **continua no código** e é exportado em
`res_nc_por_obra` / `res_nc_total` (= 24), apenas não soma mais neste card.

**Efeito no gráfico "Top não conformidades recorrentes":** ele contava a célula
inteira como um item, então "Ruído" aparecia 3x (só as células com exatamente
essa palavra) e os outros 5 ruídos ficavam escondidos dentro de textos
compostos. Agora **Ruído = 8** — está apontado em 100% das obras auditadas.

### 1.2 Reclamações de Vizinhos — de 16 para 11

`total_rec` era `len(dim_rec)`, que contava **5 linhas totalmente em branco** da
grade da planilha. Passou a ser `len(reclamacoes_lista)`, que descarta linhas
vazias — a mesma base do gráfico.

Distribuição: Titanium 3 · Garden Park 2 · Imperium 2 · Senna Tower 2 ·
Blue Coast 1 · Horizon 1.

### 1.3 Obra "Garden" não casava com "Garden Park"

`Dim_reclamacoes_vizinhos` grava **"Garden"**; todas as outras tabelas usam
**"Garden Park"**. Sem mapeamento, 2 reclamações dessa obra sumiam do gráfico
(que somava 9 em vez de 11).

Corrigido em `OBRA_MAP`:

```python
'GARDEN': 'Garden Park',
```

Verificado que "Garden" só existe nessa tabela — o mapeamento é seguro.

### 1.4 Agregação do kg/m² — média → soma

Alterado em 15/09 pela equipe, em 4 pontos de `gerar_dashboard_pbix.py`
(linhas 442, 495, 515, 543). O `score_res` do IPA (linhas 264 e 364) continua
com média.

**Por que a média estava errada:** cada linha da fonte é um tipo de resíduo e já
vem como `kg do tipo ÷ área do mês`, com a **mesma área** em todas as linhas do
mês. Somar devolve `kg total ÷ área` (a intensidade real); a média divide isso
pelo **número de tipos lançados**.

Exemplo — Blue Coast, jan/2026:

| Tipo | Kg | m² | Kg/m² |
|---|---|---|---|
| Entulho | 11.602,80 | 172,03 | 67,45 |
| Madeira | 2.001,40 | 172,03 | 11,63 |
| Metal, sucata | 532,88 | 172,03 | 3,10 |
| Papel, papelão | 176,56 | 172,03 | 1,03 |
| Plástico | 82,88 | 172,03 | 0,48 |
| **Total** | **14.396,50** | **172,03** | **83,69** |

- Soma = 83,69 kg/m² (= 14.396,50 ÷ 172,03)
- Média = 16,74 kg/m² (= 83,69 ÷ 5 tipos)

O indicador antigo **melhorava quando a obra separava melhor o resíduo**: os
mesmos 14.396 kg lançados em 10 tipos dariam 8,37 em vez de 16,74. E o fator de
distorção variava por obra (3,0x a 5,3x), invalidando a comparação entre elas.

Aritmética conferida: `Kg/m² = Kg ÷ m²` bate em **134 de 134 linhas**; a área é
única por obra-mês em **33 de 33** casos. A coluna `m²` bate com `Avanço m²` de
`Fat_avanco_fisico` — é o avanço físico **do mês**.

---

## 2. Cards da aba Consolidado

A linha de KPIs foi de 6 para 5 colunas.

| Antes | Depois |
|---|---|
| Conformidade | **Conformidade LAI/EIV** |
| NCs registradas LAI/EIV | **removido** do Consolidado (mantido na aba Obra individual) |
| Resíduos | **Resíduos gerados** + cor verde abaixo da meta |
| Reclamações de Vizinhos | (já estava correto) |
| Plano no prazo | **Plano Ação Auditorias** |

Cards do rodapé **Ações SGA**, **Auditorias** e **Emergências** removidos. A
tabela de Emergências Ambientais permanece.

Nova função `resClass(v, meta)`: verde abaixo da meta da obra, vermelho a partir
dela. Meta 36,6 kg/m² em todas as obras, exceto Senna Tower (50).

---

## 3. Gráficos

### Removido
- **Conformidade — evolução** (card, função `renderConfEvolucao` e chamadas)

### Resíduos Gerados kg/m² por obra
- Renomeado de "Resíduos kg/m² por obra"
- **Senna Tower separado** — tem meta 50, as demais 36,6; misturar as réguas
  distorcia a leitura. Botão **"Ver Senna Tower"** alterna para o modo dela
- **Dois modos:** sem filtro → uma barra por obra (últimos 6 meses); com obra
  filtrada ou botão Senna → uma barra por mês
- Card do acumulado aparece no modo Senna
- Linha de meta em degraus (50 / 36,6) no modo consolidado; some quando
  filtrado por empreendimento
- Rótulos de dados nas barras
- Obras sem lançamento não ocupam espaço no eixo — voltam sozinhas quando
  houver dado

### Monitoramento de ruído — dB por obra × mês
- Eixo X em **ordem alfabética** (era por dB médio decrescente)
- Blue Coast oculto via `RUIDO_OBRAS_OCULTAS`
- Só obras com medição no recorte

### Plano de ação Auditorias distribuição de status
- Renomeado de "Plano de ação — distribuição de status"
- Rótulos com **quantidade e percentual**, desenhados **fora do arco** com
  traço ligando à fatia (dentro do anel eram cortados)
- **Abertura por obra no tooltip**: cada fatia mostra as obras e suas ações
  (indicador · ação · responsável · prazo), vindas de `plano_acao_lista`
- Legenda embaixo, liberando os dois lados para os rótulos

### Reclamações de Vizinhos por Obra
- Renomeado de "Reclamações por obra"
- Base trocada para `reclamacoes_lista` (mesma do card)
- Tooltip com **quantidade e assunto**
- Eixo X removido (os valores estão nos rótulos das barras)
- Ao lado do Plano de Ação, em grade `cols-pa-rec` (1 : 1,4)

### Reclamações — detalhe, dividido em dois
- **Maiores incidências por assunto** — gráfico de barras horizontais
- **Assunto e ação por obra** — tabela ordenada por obra (alfabética) e, dentro
  da obra, por data; nome da obra só na primeira linha do grupo

A fonte não tem coluna de assunto preenchida (`Categoria` está 100% vazia), então
o assunto é classificado por palavra-chave da ocorrência, em `REC_ASSUNTOS`:
Ruído, Poeira, Sujeira, Dano a terceiros, Águas pluviais, Queda de objeto.

### Top não conformidades recorrentes
- Passou a **respeitar os filtros de obra e período**, via `nc_assuntos_lista`
  (obra × mês × assunto). Antes usava o agregado estático e não reagia a nada
- Badge mostra o recorte ativo

### Histórico mensal de páginas impressas
- Eixo Y removido (título "Total de páginas" e os valores 50.000, 40.000…)
- Rótulos de dados sobre cada ponto
- Altura de 220px para 280px

### Renomeações
| Antes | Depois |
|---|---|
| Ações SGA — pessoas por ano | Eventos Ambientais |
| Consumo de energia (kWh) | Consumo de Energia Obras (kWh) |
| Intensidade energética (kWh/m²) | Intensidade Energética Obras (kWh/m²) |
| Consumo de água (m³) | Consumo de Água Obras (m³) |
| Intensidade hídrica (m³/m²) | Intensidade Hídrica Obras (m³/m²) |

---

## 4. Infraestrutura

### `requirements.txt` (novo)

```
pbixray==0.15.5
pandas==3.0.5
numpy==2.5.2
openpyxl==3.1.5
```

**Versões fixas por necessidade, não por preferência.** Produção roda
`pbixray 0.9.0` e homologação `0.15.5`. Com o **mesmo `.pbix`** (conferido por
MD5), as duas versões devolvem valores diferentes para células vazias da coluna
`dB` de `Fat_Ruido`:

- **0.9.0** devolve um resíduo positivo minúsculo no lugar de vazio. Esse valor
  passa no filtro `db > 0` e vira uma "medição" fantasma
- **0.15.5** lê vazio como vazio

Efeitos observados em 14/09:

| | Produção (0.9.0) | Homologação (0.15.5) |
|---|---|---|
| IPA global | 79,1 | **78,4** |
| Gráfico de ruído | exibia abr/2026, mês **sem medição**, e por isso perdia nov/2025 | janela correta |
| North (1 linha de ruído em branco) | 100% de conformidade de ruído | sem dado → fallback 50% |

Das 258 células obra×mês, 243 são idênticas; 9 divergem por esse efeito e 6 por
arredondamento de 1 dB (cosmético).

**Promover o código sem igualar a versão da lib não reproduz os números.**

### Utilitários de plugin (template)

- `barLabels(fmt, opts)` — rótulos de dados em barras e pontos de linha
- `paDonutLabels` — rótulos externos da rosca

Ambos escritos à mão em vez de `chartjs-plugin-datalabels`, para **não
acrescentar dependência de CDN** — o dashboard roda em rede interna.

---

## 5. Novas estruturas no JSON de dados

| Chave | Conteúdo | Usada por |
|---|---|---|
| `nc_assuntos_lista` | obra × mês × assunto (LAI/EIV) | Top NCs filtrável |
| `reclamacoes_lista` | ocorrência, status, canal, categoria | Gráfico, tooltip e tabela de reclamações |
| `plano_acao_lista` | obra, status, indicador, ação, responsável, prazo | Tooltip do Plano de Ação |
| `res_nc_por_obra` / `res_nc_total` | NCs de resíduos (24) | Exportado para uso futuro |

---

## 6. Defeitos de dados encontrados (fonte, não código)

### 6.1 `m² = 1` em obra-mês não fechado

Quando a área não é lançada, a planilha grava **1,0**. Dividir por 1 não divide:
o `Kg/m²` sai **igual ao Kg**.

Caso concreto — Senna Tower, abr/2026:

```
MADEIRA   1.601,12 kg ÷ 1,0 = 1.601,12 kg/m²   (em mar/26 o mesmo peso deu 0,38)
```

`Fat_avanco_fisico` confirma: em abr/2026 a obra teve **avanço 0,00%**. O 1,0 foi
um conserto manual para evitar divisão por zero.

**São 1.224 das 1.608 linhas** com `m² = 1` — é o padrão para mês em aberto.
Enquanto o Kg fica 0 é inofensivo; o problema aparece quando o Kg é lançado
**antes** do m².

O gerador **detecta e avisa no log**, mas **não descarta** (decisão da área, para
o acumulado do Senna Tower mostrar 130,16 kg/m²):

```
AVISO: 3 linha(s) de residuos com area m2 <= 1 (nao lancada).
       MANTIDAS na base a pedido da area; o Kg/m2 delas sai igual ao Kg
```

Para voltar a descartar, acrescentar `fat_res = fat_res[~_area_ruim].copy()`.
Com o descarte: card de resíduos 25,89 → 13,61, IPA 78,4 → 79,3, Senna 130,16 → 0,53.

### 6.2 `.pbix` desatualizado em relação à planilha

| | Planilha (`SOMA RESÍDUOS`) | `.pbix` |
|---|---|---|
| Linhas | 1.860 | 1.608 |
| Último mês | **set/2026** | jul/2026 |
| Obras | 13 (tem `North Tower`) | 12 |
| Grafia | "Senna" | "Senna Tower" |

O `.pbix` é de **28/05/2026**. Tudo lançado depois não chega ao dashboard — é a
razão de o **North** não aparecer nos gráficos de resíduos.

### 6.3 Outras inconsistências

- `Dim_Plano_acao`: 42 linhas, **39 em branco**. `pa_no_prazo_pct` divide 3 por
  42 → **7,1%** em vez de 100%
- `Fat_Ruido`: 9 medições gravadas como **0 dB** (fisicamente impossível em
  canteiro). `score_ruido` não filtra `db > 0`, então contam como **conformes**
- `Fat_Impressoes_por_setor`: setor **"Unknown"** com 36.596 páginas (2º do
  ranking), 12 linhas, só em 2025, todas com `Obra = "Outros"`. Não é setor real
   — são impressões sem usuário mapeado. Existe também `(Usuários sem conta)`
- `Fat_soma_residuos`: linha de ENTULHO de fev/2026 do Senna com `m² = 4.182,98`
  enquanto as outras 11 do mesmo mês têm 796,07 (inofensivo, Kg = 0)

---

## 7. Valores atuais em homologação

```
IPA global ............... 78,4
Conformidade média ....... 88,1%
NC registradas LAI/EIV ... 20   (14 LAI + 6 EIV)
Reclamações .............. 11
Plano de Ação ............ 7,1%   ← ver 6.3
media_kg_m2 .............. 3.469,18   ← ver pendência 1
Top NC ................... Ruído (8)
```

---

## 8. Pendências

1. **Alinhar gráfico e card à soma.** O gerador usa soma, mas o template ainda
   usa `mediaLinhas` no gráfico do Consolidado. E `media_kg_m2` virou **3.469,18**
   — soma de todas as linhas de todas as obras e meses, sem significado físico.
   Ver também a memória `kgm2-soma-acumula-entre-meses`: a soma acumula entre
   meses e é comparada com meta mensal
2. **Atualizar o `.pbix`** com os dados até set/2026 (traz o North)
3. **`score_ruido` ignorar `dB = 0`** — IPA 78,4 → ~78,2
4. **Filtro de período no Plano de Ação** — `Dim_Plano_acao` só tem a coluna
   `Prazo` (prazo da ação, não data de abertura). Filtrar por ela ou deixar o
   card fora do filtro?
5. **Plano de Ação 7,1% → 100%** — descartar as 39 linhas vazias do denominador
6. **"Unknown" nas impressões** — excluir do ranking ou renomear para
   "Não identificado"?
7. **Igualar `pbixray` em produção** antes de promover (ver seção 4)

---

## 9. Backups e restauração

Backups em `_backup_teste_refresh/`. **O tamanho identifica a versão:**

| Arquivo | Tamanho | Estado |
|---|---|---|
| `gerar_dashboard_*.py` | **~203 KB** | com as alterações |
| `gerar_dashboard_*.py` | ~166 KB | estado antigo de 11/09 |
| `gerar_dashboard_pbix_*.py` | **~72 KB** | com as alterações |
| `gerar_dashboard_pbix_*.py` | ~62 KB | estado antigo de 11/09 |

Em 15/09 os arquivos de trabalho foram sobrescritos pelas versões de 11/09 e as
alterações precisaram ser restauradas de
`gerar_dashboard_20260915_latest.py` e `gerar_dashboard_pbix_20260915_com_soma.py`.
**Conferir o tamanho antes de restaurar.**

---

## 10. Como validar

```bash
# 1. Gerar (em homologação, 10.1.1.155)
py -3.12 gerar_dashboard_pbix.py

# 2. Conferir no navegador
#    http://10.1.1.155:8090/

# 3. Checagem de sintaxe do JS gerado (opcional)
#    extrair o bloco <script> do dashboard_fg.html e rodar:
node --check app.js
```

O gerador imprime avisos de dados inconsistentes no stdout (ver 6.1) — vale ler
a saída em vez de só conferir o "Dashboard gerado com sucesso".
