# Deploy para produção (10.1.1.151) — pacote de 15/09/2026

Guia de promoção das alterações feitas e validadas em homologação (10.1.1.155,
`http://10.1.1.155:8090/`) entre 11/09 e 15/09/2026.

Documentos relacionados, todos neste repositório:

| Documento | O que cobre |
|---|---|
| `homologacao_ambiental.md` | Fluxo de deploy, histórico e decisões do período |
| `ALTERACOES_ABA_CONSOLIDADO.md` | Detalhe de cada mudança da aba Consolidado |
| `ALTERACOES_ABA_OBRA_INDIVIDUAL.md` | Detalhe da aba Obra Individual |
| `ALTERACOES_ABA_RESIDUOS.md` | Detalhe da aba Resíduos |
| `etl/README.md` | **Modelo semântico e carga do banco** (ver seção 5) |

---

## 1. O que muda

| Frente | Arquivo | Resumo |
|---|---|---|
| Aba Consolidado | `gerar_dashboard_pbix.py` | NC LAI/EIV 32 → 20 (conta assunto, não linha); reclamações 16 → 11 (descarta linha vazia); `kg/m²` de média para soma; KPIs de 6 para 5 |
| Aba Obra Individual | `gerar_dashboard.py` | IPA com histórico de 12 meses e reagindo ao filtro de período; rótulos de dados em 4 gráficos; grid de KPI 4 → 3 |
| Aba Resíduos | `gerar_dashboard.py` | Formatação pt-BR; KPIs 5 → 3; Senna Tower separada; gráfico de tipo por mês; tabelas renomeadas |
| Infraestrutura | `requirements.txt` (novo) | Versões fixas — `pbixray` muda o resultado (seção 2) |

Tamanhos de referência dos arquivos corretos:

```
gerar_dashboard.py         204.438 bytes   (~203 KB)
gerar_dashboard_pbix.py     72.251 bytes   (~72 KB)
dashboard_fg.html          581.899 bytes   (gerado)
```

Um `gerar_dashboard.py` de ~166 KB é o estado antigo de 11/09, anterior a todo o
trabalho — se aparecer esse tamanho, o arquivo errado foi restaurado.

---

## 2. Pré-requisito crítico: versão do `pbixray`

Produção roda **0.9.0**, homologação **0.15.5**. Com o **mesmo `.pbix`**
(conferido por MD5), as duas versões leem células vazias da coluna `dB` de
`Fat_Ruido` de formas diferentes: a 0.9.0 devolve um resíduo positivo minúsculo
que passa no filtro `db > 0` e vira medição fantasma.

| | Produção (0.9.0) | Homologação (0.15.5) |
|---|---|---|
| IPA global | 79,1 | **78,4** |
| Gráfico de ruído | exibe abr/2026, mês sem medição, e perde nov/2025 | janela correta |
| North (1 linha em branco) | 100% de conformidade de ruído | sem dado → fallback 50% |

**Promover o código sem igualar a versão não reproduz os números.** Instalar
antes de qualquer outra coisa:

```powershell
py -3.12 -m pip install -r requirements.txt
py -3.12 -c "import pbixray; print(pbixray.__version__)"   # deve imprimir 0.15.5
```

---

## 3. Passo a passo

Rodar no servidor **10.1.1.151**, em `C:\Apps\Dash_ambiental_qualidade`.

**3.1 — Parar o refresh agendado** para não gerar no meio da troca de arquivos.

**3.2 — Guardar o estado atual** (permite voltar atrás):

```powershell
$hoje = Get-Date -Format 'yyyyMMdd_HHmmss'
Copy-Item gerar_dashboard.py      "_backup_pre_deploy_$hoje.py"
Copy-Item gerar_dashboard_pbix.py "_backup_pre_deploy_pbix_$hoje.py"
Copy-Item dashboard_fg.html       "_backup_pre_deploy_$hoje.html"
```

**3.3 — Atualizar dependências** (seção 2).

**3.4 — Trazer o código:**

```powershell
git pull origin master
```

**3.5 — Conferir o `.pbix`.** O gerador procura nesta ordem (linhas 44–49 de
`gerar_dashboard_pbix.py`): cópia local na pasta do script, depois `U:\...`,
depois o OneDrive do desenvolvedor. Em produção a cópia local é a que vale:

```powershell
Get-Item AMBIENTAL_QUALIDADE_202602_v1_DEV.pbix | Select-Object Length, LastWriteTime
```

**3.6 — Gerar:**

```powershell
py -3.12 gerar_dashboard_pbix.py
```

**Ler a saída do comando**, não só o "Dashboard gerado com sucesso" — o gerador
imprime avisos de dados inconsistentes. Este aviso é **esperado** e não indica
falha (decisão da área, registrada em `ALTERACOES_ABA_CONSOLIDADO.md` §6.1):

```
AVISO: 3 linha(s) de residuos com area m2 <= 1 (nao lancada).
       MANTIDAS na base a pedido da area; o Kg/m2 delas sai igual ao Kg
```

**3.7 — Validar** em `http://10.1.1.151:8080/` com **Ctrl+F5** (o serviço não
manda header de no-cache).

**3.8 — Religar o refresh agendado.**

---

## 4. Checklist de validação

Números esperados, iguais aos de homologação:

- [ ] `pbixray` em **0.15.5**
- [ ] `dashboard_fg.html` com data de modificação de agora
- [ ] **IPA global = 78,4** (se vier 79,1, a versão do `pbixray` não foi trocada)
- [ ] Conformidade média ≈ 88,1%
- [ ] NC registradas LAI/EIV = **20**
- [ ] Reclamações = **11**
- [ ] Top NC = **Ruído (8)**

Conferências visuais:

- [ ] Consolidado com **5 KPIs**: IPA global · Conformidade LAI/EIV · Resíduos
      gerados · Reclamações de Vizinhos · Plano Ação Auditorias
- [ ] Consolidado **sem** os cards de rodapé Ações SGA, Auditorias e Emergências
      (a tabela de Emergências Ambientais permanece)
- [ ] Resíduos com **3 KPIs**: Volume total · Geração Resíduo · Classes CONAMA
- [ ] Gráfico "Resíduos Gerados kg/m² por obra" com o botão **"Ver Senna Tower"**
- [ ] Obra Individual: trocar o filtro de Período **muda** o gráfico de IPA
- [ ] Rótulos de dados visíveis nas barras (IPA, Ruído, Resíduos, Plano de Ação)

> O valor do card "Resíduos gerados" é hoje **3.469,18** e não tem significado
> físico — é a soma de todas as linhas de todas as obras e meses. É uma pendência
> conhecida, não um defeito do deploy. Ver seção 6.

---

## 5. Banco de dados e modelo semântico

A migração do `.pbix` para banco relacional **já está construída e validada**, em
`etl/`. Ela não faz parte deste deploy do dashboard — o dashboard continua lendo
o `.pbix` —, mas é o que o Eduardo precisa para preparar o banco.

O que existe hoje:

| Item | Onde |
|---|---|
| Documentação completa (como rodar, agendar, monitorar) | `etl/README.md` |
| As 19 tabelas do modelo: origem, colunas, tipos, chave | `etl/tabelas.py` |
| Transformações equivalentes às do Power Query | `etl/modelo_semantico.py` |
| **DDL pronto para PostgreSQL, SQL Server e SQLite** | `etl/ddl/modelo_*.sql` |
| As 26 FOREIGN KEYs (13 aplicáveis, 13 comentadas por órfãos) | `etl/ddl/modelo_*_fk.sql` |
| Comparação ETL × `.pbix`, tabela a tabela | `etl/validar_vs_pbix.py` |
| Carga local em SQLite, para conferência | `etl/banco/ambiental.db` |

A receita não foi inventada: o código Power Query (M) de cada tabela foi extraído
de dentro do próprio `.pbix` com `PBIXRay(...).power_query`, e cada função em
`tabelas.py` reproduz os passos do M da sua tabela. A validação de 14/09/2026
mostra **as 19 tabelas com as mesmas colunas, na mesma ordem e com os mesmos
tipos**; as diferenças de linha são defasagem do `.pbix` (atualizado até abr/2026,
planilha já em set/2026), não erro do ETL.

Para criar o banco:

```powershell
cd etl
py -3.12 etl_excel_para_db.py --somente-validar          # confere sem gravar
py -3.12 etl_excel_para_db.py --somente-ddl --dialeto postgresql --ddl ddl\modelo_postgresql.sql
py -3.12 etl_excel_para_db.py --db-url "postgresql+psycopg2://usuario:senha@servidor/base" --schema ambiental
```

**Dois pontos de atenção antes de agendar a carga** (detalhados no
`etl/README.md`):

1. **Falta descobrir o caminho UNC da planilha.** Todos os scripts usam `U:`, e
   unidade mapeada não existe para tarefa agendada nem para serviço Windows. A
   origem tem que ser `\\servidor\share\...` e a conta da tarefa precisa de
   leitura nesse share.
2. **A instância do banco ainda não existe.** Enquanto não houver, a carga em
   SQLite (`etl/banco/ambiental.db`) serve para conferência.

---

## 6. Pendências conhecidas (não bloqueiam o deploy)

Herdadas de homologação, com o detalhe em cada documento de alterações:

1. **`kg/m²` somado acumula entre meses.** `res_por_obra` soma todas as linhas da
   obra, inclusive entre meses, enquanto a meta (36,6 · Senna 50) é mensal. Blue
   Coast tem 1 mês lançado e por isso aparece como 2ª melhor obra quando por mês
   é a 2ª pior. O valor cresce sozinho a cada mês novo.
2. **`.pbix` desatualizado** — é de 28/05/2026 e a planilha já está em set/2026.
   É a razão de o North não aparecer nos gráficos de resíduos.
3. **`m³/m²` ainda usa média** enquanto `kg/m²` virou soma, lado a lado na mesma
   tabela.
4. **Plano de Ação em 7,1%** — `Dim_Plano_acao` tem 42 linhas, 39 em branco, e o
   denominador conta todas.
5. **`score_ruido` não filtra `dB = 0`** — há 9 medições gravadas como 0 dB, que
   contam como conformes.
6. **Setor "Unknown" nas impressões** — 36.596 páginas, 2º do ranking; não é
   setor real, são impressões sem usuário mapeado.

---

## 7. Se algo sair errado

| Sintoma | Causa provável | O que fazer |
|---|---|---|
| IPA vem 79,1 | `pbixray` ainda em 0.9.0 | Reinstalar com `requirements.txt` e gerar de novo |
| Gráfico de ruído mostra abr/2026 | idem | idem |
| `.pbix` não encontrado | Cópia local ausente | Conferir os caminhos em `gerar_dashboard_pbix.py` linhas 44–49 |
| Dashboard sem as mudanças | Cache do navegador | Ctrl+F5 |
| Layout antigo (6 KPIs, sem rótulos) | `gerar_dashboard.py` errado | Conferir o tamanho: tem de ser ~203 KB, não ~166 KB |
| Precisa voltar atrás | — | Restaurar os `_backup_pre_deploy_*` do passo 3.2 e gerar de novo |

---

## 8. Contatos

- **Alterações e validação em homologação:** Natália Costa
- **Deploy em produção:** Eduardo
- **Decisões sobre dados e metas:** área de Qualidade / Sustentabilidade
