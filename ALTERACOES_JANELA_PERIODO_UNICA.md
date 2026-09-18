# "Últimos 6/12 meses": uma janela de calendário única para todos os gráficos

**Data:** 18 de setembro de 2026
**Arquivos alterados:** `gerar_dashboard.py` e `gerar_dashboard_pbix.py`
**Relatado pela área:** ao escolher "Últimos 6 meses" ou "Últimos 12 meses", os
gráficos continuavam mostrando datas de períodos diferentes do selecionado.

---

## O que estava acontecendo

"Últimos 6 meses" era calculado como **os 6 últimos meses que aquela fonte
tem** (`slice(-6)` sobre o universo de meses de cada tabela). Como cada tabela
do modelo termina num mês diferente, a mesma seleção produzia **sete janelas
diferentes na mesma tela**:

| Gráfico | O que "últimos 6 meses" mostrava |
|---|---|
| Conformidade / IPA / NCs | Dez/25 – Mai/26 |
| Resíduos | **Fev/26 – Jul/26** |
| Ruído | Nov/25 – Mai/26 |
| Água e energia | **Out/25 – Mar/26** |
| Impressões | Nov/25 – Abr/26 |
| Auditorias | **Set/25 – Out/26** |
| Reclamações | Dez/25 – Mar/26 |

Ou seja: a mesma tela comparava a água de Out/25–Mar/26 com os resíduos de
Fev/26–Jul/26 e chamava as duas coisas de "últimos 6 meses". É exatamente o que
foi relatado.

### Por que as fontes terminam em meses diferentes

Três motivos distintos, todos legítimos do dado:

1. **Fim real do lançamento** — água e energia têm dado até Mar/26;
   conformidade e ruído até Mai/26.
2. **Grade zerada** — `Fat_soma_residuos` tem linhas para Mai, Jun e Jul/26,
   mas com **kg = 0** (linhas em branco da planilha). O `slice(-6)` contava
   esses meses vazios como se fossem dado, empurrando a janela para frente.
3. **Planejamento futuro** — `Dim_Auditoria_Ambientais` tem linhas até Out/26
   (auditorias planejadas), meses que ainda nem aconteceram.

---

## A correção

### 1. Uma âncora única, calculada no gerador

Novo `PERIODO_REF` em `gerar_dashboard_pbix.py`, exportado como
`D.periodo_ref`: o **mês mais recente com dado real**, o que exclui de
propósito os três casos acima —

- meses de grade zerada (só entra mês com valor diferente de zero);
- meses de planejamento futuro (teto no mês de geração do dashboard);
- fontes que acabaram antes não puxam a âncora para trás (é o **máximo** entre
  as fontes).

Nesta base o valor é **2026-05**, e o gerador imprime isso na execução:

```
Mes de referencia do filtro de Periodo: 2026-05
```

### 2. A janela passou a ser de calendário, não de posição

Novos helpers no template:

```js
janelaPeriodo(anoF)   // → {de: '2025-12', ate: '2026-05'}
noPeriodo(mes, anoF)  // → true/false, aceita 'YYYY-MM' e 'YYYY-MM-DD'
mesesDoPeriodo(...)   // agora recorta pela janela, não por slice(-N)
faixaPeriodo(anoF)    // → "Dez/25 – Mai/26", para os badges
```

Com isso:

- **Últimos 6 meses** = Dez/25 – Mai/26, em todos os gráficos
- **Últimos 12 meses** = Jun/25 – Mai/26, em todos os gráficos

Todos os 11 pontos que faziam `slice(-6)`/`slice(-12)` por conta própria foram
trocados pela janela: heatmap, Top NCs, IPA por obra, ruído por obra, tabela de
conformidade da Obra Individual, resíduos, auditorias, água, energia,
impressões e a tabela de conformidade do Consolidado.

**Um gráfico mostrar menos barras que o teto agora é o comportamento certo:**
água com "últimos 6 meses" mostra Dez/25–Mar/26 (4 barras) porque o dado de
água acaba em Mar/26. Antes ela completava 6 barras indo buscar Out e Nov/25 —
meses fora do período que o usuário tinha pedido.

### 3. Badge mostra a janela

O badge do heatmap passou a exibir a faixa real:
`últimos 6 meses · Dez/25 – Mai/26`. Sem isso não há como o usuário conferir
que os gráficos estão de fato falando do mesmo período.

### 4. Bug separado: emergências filtravam por DIA

`emergencias_lista.mes_ano` guarda a data completa (`2026-04-14`), não o mês.
O filtro comparava essa string com meses, então "últimos 6 meses" recortava
pelos **6 últimos dias distintos**. Novo campo `mes` (`2026-04`) no gerador, e
os dois cards de emergências passaram a usar ele.

Efeito: com "últimos 6 meses" o card ia para 2 registros; agora vai para os 10
registros que realmente estão na janela.

---

## Validação feita

Com o filtro em **últimos 6 meses** (janela Dez/25 – Mai/26):

| Gráfico | Meses exibidos | Correto? |
|---|---|---|
| Heatmap conformidade | Dez/25 … Mai/26 (6) | ✓ janela cheia |
| Ruído | Dez/25 … Mai/26 | ✓ |
| Água | Dez/25 … Mar/26 (4) | ✓ dado acaba em Mar/26 |
| Energia | Dez/25 … Mar/26 (4) | ✓ idem |
| Impressões | Dez/25 … Abr/26 (5) | ✓ dado acaba em Abr/26 |
| Volume por tipo de resíduo | Jan/26 … Abr/26 (4) | ✓ meses com kg > 0 na janela |

Com **últimos 12 meses** (Jun/25 – Mai/26): heatmap com os 12 meses exatos,
água Jun/25–Mar/26, impressões Jun/25–Abr/26.

Outros testes:

- [x] Varredura automática em **todos** os gráficos: nenhum rótulo de mês fora
      da janela (`mesesForaDaJanela: []`)
- [x] 2 obras × 5 períodos × 4 abas = 40 combinações, sem erro
- [x] Emergências: 42 (todo o período) → 10 (6m, 12m e 2026) → 0 (2025)
- [x] IPA global 78,4 mantido — nenhuma regra de cálculo mudou, só o recorte
- [x] Console sem erro

---

## Pendência que isto expõe

O `.pbix` é de maio/2026 e o dashboard é gerado em setembro/2026. Como a âncora
é o último mês com dado real (2026-05), "últimos 6 meses" significa
Dez/25–Mai/26, e **não** Abr/26–Set/26 como o calendário sugeriria. Está certo
para o dado que existe, mas some quando o modelo for atualizado — a âncora se
move sozinha, sem precisar mexer em código. Vale a área saber que a janela
acompanha o dado, não o relógio.
