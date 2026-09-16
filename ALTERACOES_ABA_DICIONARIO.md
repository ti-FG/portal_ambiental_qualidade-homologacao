# Aba Dicionário de dados e termos

**Criada em:** 16 de setembro de 2026
**Arquivo alterado:** `gerar_dashboard.py` (template HTML/CSS/JS — único arquivo tocado)

Quarta aba do dashboard, ao lado de Consolidado, Obra individual e Resíduos.
Reúne num só lugar as regras de negócio, as fórmulas dos indicadores e a origem
de cada dado, para que qualquer número visto nas outras abas possa ser rastreado
até a regra que o produziu e até a tabela do modelo semântico de onde ele saiu.

---

## 1. O que a aba tem

São **86 verbetes** em 6 seções:

| Seção | Verbetes | Conteúdo |
|---|---|---|
| Indicadores e KPIs | 18 | Cada indicador do dashboard: onde aparece, como é calculado, de que tabela vem |
| Composição do IPA | 6 | Os cinco componentes com peso, fórmula do score e a regra de retaguarda 0,5 |
| Regras de negócio | 11 | Metas, limite de ruído, normalização de obra, precedência de filtros, "sem dado não é zero" |
| Modelo de dados | 19 | As 19 tabelas do modelo semântico: origem na planilha, grão e conteúdo |
| Glossário de termos | 24 | LAI, EIV, NC, IPA, SGA, classes CONAMA, kg/m², dB, obra-mês, recorte, e demais |
| Limitações conhecidas | 8 | As pendências abertas, explicadas para quem lê o dashboard |

Recursos de navegação:

- **Busca** no topo, que filtra os verbetes de todas as seções ao mesmo tempo.
  É insensível a acento e a maiúscula — digitar `residuo` acha "Resíduos".
  Seção sem nenhum verbete correspondente some da tela, e o contador do cabeçalho
  passa a mostrar "N de 86 verbetes".
- **Chips de atalho** logo abaixo da busca, que rolam até cada seção.
- A **barra de filtros** (obra e período) fica escondida nesta aba: o conteúdo é
  de referência e não reage a recorte. Ela volta ao sair para qualquer outra aba.

---

## 2. Onde o conteúdo foi apurado

O texto não foi escrito de memória — cada regra veio da leitura do código que
gera os números:

| Conteúdo | Origem |
|---|---|
| Pesos e faixas do IPA | `calc_ipa()` e `ipa_class()` de `gerar_dashboard_pbix.py` |
| Fórmula de cada score | `score_conf`, `score_res`, `score_ruido`, `score_aud`, `score_pa` |
| Metas de resíduos | `META_RESIDUOS_DEFAULT` / `META_RESIDUOS_SENNA` e `meta_obra()` |
| Normalização de obra | `OBRA_MAP` e `norm_obra()` |
| Contagem de NC por assunto | `_assuntos_nc()` / `_assuntos_linha()` |
| Reclamações sem linha vazia | montagem de `reclamacoes_lista` e `total_rec` |
| As 19 tabelas do modelo | `etl/tabelas.py` (nome, fonte na planilha, chave e descrição) |
| Limitações conhecidas | seção 6 de `DEPLOY_PRODUCAO.md` e seção 8 de `homologacao_ambiental.md` |

As fórmulas dos cinco scores são **idênticas** entre produção e homologação —
o que difere entre os dois ambientes é a agregação do kg/m² e a contagem de NC,
e ambas estão descritas no texto como estão hoje no gerador deste repositório.

---

## 3. Como foi implementado

Seis inserções no template dentro de `gerar_dashboard.py`:

1. **CSS** — bloco `.dic-*` inserido antes de `/* Footer */`.
2. **Botão da aba** — `showPage('dicionario')` com ícone `ti ti-book`, depois do
   botão de Resíduos.
3. **Página** — `<div id="page-dicionario" class="page">`, inserida antes de
   `<div class="dash-footer">`.
4. **`const PAGES`** — passou a incluir `'dicionario'`.
5. **`showPage()`** — chama `renderDicionario()` ao entrar na aba e esconde a
   barra de filtros só nela.
6. **Funções JS** — `renderDicionario`, `filtrarDicionario`, `irParaDic`,
   `dicNorm` e `dicMeta`, mais a chamada no `DOMContentLoaded`.

Nenhuma mudança em `gerar_dashboard_pbix.py`: ele extrai o template deste arquivo
e por isso a aba aparece igual nos dois caminhos de geração (PBIX e Excel).
Nenhum dado novo foi adicionado ao JSON — a aba consome apenas `D.meta_residuos`,
que já existia.

### Conteúdo estático com duas exceções

O texto é estático de propósito: é documentação, não indicador. As duas exceções
são lidas de `D` para não envelhecerem sozinhas:

- as **metas de resíduos** (36,6 e 50), que vêm de `D.meta_residuos`;
- o **contador de verbetes**, contado no DOM em vez de escrito à mão.

---

## 4. Validação feita

Gerado com `py -3.12 gerar_dashboard_pbix.py` neste repositório e aberto no
navegador:

- [x] Geração completa sem erro — **IPA global 78,4**, o número esperado de
      homologação (`pbixray` 0.15.5)
- [x] As quatro abas alternam corretamente; Consolidado, Obra individual e
      Resíduos seguem funcionando
- [x] Barra de filtros escondida só no Dicionário e restaurada ao sair
- [x] 86 verbetes em 6 seções
- [x] Busca por `conama` devolve 3 verbetes em 3 seções; busca por `ruido`
      devolve 18; termo inexistente mostra o aviso de nenhum resultado
- [x] Busca sem acento encontra verbete com acento (`residuo` → "Resíduos")
- [x] Limpar a busca restaura os 86 verbetes e as 6 seções
- [x] Metas renderizadas a partir dos dados: 36,6 e 50
- [x] Console do navegador sem erro

---

## 5. Pendências

1. **Revisão da área.** O glossário descreve LAI, EIV, classes CONAMA e demais
   termos em linguagem corrente. A área de Qualidade / Sustentabilidade precisa
   confirmar se as definições estão como querem publicar.
2. **Limitações conhecidas expostas ao usuário final.** A seção descreve
   abertamente os pontos em aberto (kg/m² acumulado, plano de ação com
   denominador inflado, água zerada na origem). Foi escrita assim de propósito,
   para que ninguém leia um número fora de contexto — mas é decisão da área
   manter ou enxugar.
3. **Manutenção.** Quando uma regra mudar no gerador, o verbete correspondente
   precisa mudar junto. O risco real desta aba é envelhecer em silêncio.
