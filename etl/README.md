# ETL — Excel → modelo semântico → banco de dados

Este diretório contém o caminho que substitui o `.pbix` como fonte de dados do
dashboard ambiental: lê a planilha `Controle Ambiental 2025_oficial_db.xlsx`,
aplica **as mesmas transformações que o Power Query do `.pbix` aplica** e grava
as **19 tabelas do modelo semântico** em um banco relacional.

Hoje o `gerar_dashboard_pbix.py` lê essas mesmas 19 tabelas de dentro do
`.pbix`. Depois desta carga, elas passam a existir em banco, com o mesmo nome,
as mesmas colunas e os mesmos valores.

---

## Como funciona

```
Controle Ambiental 2025_oficial_db.xlsx
        │  (18 tabelas nomeadas / ListObjects)
        ▼
  etl_excel_para_db.py  ──  tabelas.py (o quê)  +  modelo_semantico.py (como)
        │
        ├──► PostgreSQL        (o destino)
        ├──► .parquet / .csv   (opcional)
        └──► ddl/*.sql         (opcional)
```

A receita não foi inventada: **o código Power Query (M) de cada tabela está
guardado dentro do próprio `.pbix`** e foi extraído com
`PBIXRay(...).power_query`. Cada função em `tabelas.py` reproduz os passos do M
da sua tabela, e o campo `m` de cada `Tabela` resume esses passos.

### Arquivos

| Arquivo | Papel |
|---|---|
| `etl_excel_para_db.py` | Programa principal (CLI): lê, valida, carrega. |
| `tabelas.py` | As 19 tabelas: origem, colunas, tipos, chave e transformações. |
| `modelo_semantico.py` | Operações genéricas equivalentes às do Power Query (substituição de texto, classificação de obra, coerção de tipos). |
| `validar_vs_pbix.py` | Compara o resultado do ETL com o `.pbix`, tabela a tabela. |
| `ddl/modelo_postgresql.sql` | DDL das tabelas. |
| `ddl/modelo_*_fk.sql` | As 26 FOREIGN KEYs do modelo (as que ainda não podem ser criadas saem comentadas). |
| `origem.py` | De onde vem a planilha, cópia local e detecção de mudança. |
| `Atualizar_Banco.bat` | Entrada da tarefa agendada. |

---

## Rodando

> **O destino do modelo é PostgreSQL.** O SQLite foi removido em 15/09/2026 —
> não há mais carga local de rascunho. `--db-url` é obrigatório para carregar, e
> sem ele o programa erra em vez de gravar em algum lugar por conta própria.

```powershell
# 1) conferir sem gravar nada (nao precisa de banco)
py -3.12 etl_excel_para_db.py --somente-validar

# 2) so o DDL do PostgreSQL, sem tocar em banco
py -3.12 etl_excel_para_db.py --somente-ddl --ddl ddl\modelo_postgresql.sql

# 3) carga em PostgreSQL, em schema proprio
py -3.12 etl_excel_para_db.py --db-url "postgresql+psycopg2://usuario:senha@servidor/base" --schema ambiental

# 4) carga agendada: so recarrega se a planilha mudou, com log
py -3.12 etl_excel_para_db.py --excel "\\servidor\share\...\Controle Ambiental 2025_oficial_db.xlsx" ^
    --db-url "postgresql+psycopg2://usuario@servidor/base" --schema ambiental ^
    --se-mudou --log logs\etl_banco.log

# 6) exportar tambem em parquet/csv
py -3.12 etl_excel_para_db.py --parquet saida\parquet --csv saida\csv

# 7) carga agendada: so recarrega se a planilha mudou, com log
py -3.12 etl_excel_para_db.py --excel "\\servidor\share\...\Controle Ambiental 2025_oficial_db.xlsx" ^
    --db-url "postgresql+psycopg2://usuario@servidor/base" --schema ambiental ^
    --se-mudou --log logs\etl_banco.log
```

A carga é **full refresh**: apaga e regrava todas as tabelas dentro de **uma
única transação**. Se qualquer tabela falhar, nada é gravado.

Duas tabelas de controle acompanham a carga:

- **`_dicionario_modelo`** — o mapa `nome no pbix ↔ nome no banco`, com o tipo
  lógico e a tabela nomeada de origem no Excel (157 linhas). É por aqui que se
  descobre que `Fat_soma_residuos."Kg/m²"` virou `fat_soma_residuos.kg_m2`.
- **`_relacionamentos_modelo`** — os 26 relacionamentos do modelo do pbix, com
  cardinalidade, filtro cruzado, nome da FK e quantas linhas órfãs existem hoje.
- **`_carga_log`** — quando cada carga rodou, de qual arquivo e quantas linhas
  por tabela.

### Nomes no banco

Os nomes do modelo semântico têm acento, `%`, `³`, `/` e espaço no fim
(`'Total de pág '`, `'Kg/m²'`, `'N° FMC'`) — não dá para usar em SQL. O ETL
converte para `snake_case` ASCII (`total_de_pag`, `kg_m2`, `n_fmc`) e registra
os dois nomes em `_dicionario_modelo`. **Os nomes lógicos das tabelas e colunas
continuam idênticos aos do pbix** — só a grafia física muda.

---

## Prova de que é o mesmo modelo

`validar_vs_pbix.py` compara, tabela a tabela: nomes e ordem das colunas, as
linhas como multiconjunto, e célula a célula nas chaves que existem dos dois
lados.

Resultado da última execução (14/09/2026):

```
TABELA                            ETL   PBIX  IGUAIS  SO_ETL SO_PBIX   VAZIO  MUDOU
Dim_Obras                          12     12      12       0       0       0      0
Dim_Calendario                   4017   4017    4017       0       0       0      0
Dim_Acoes_SGA                      18     17      16       2       1       1      0
Dim_Auditoria_Ambientais           45     36      28      17       8      19      0
Dim_NC_LAI_EIV                      8      8       8       0       0       0      0
Dim_Plano_acao                     11      3       2       9       1       0      0
Dim_emergencias_ambientais         20     10      10      10       0       0      0
Dim_reclamacoes_vizinhos           21     11      11      10       0       0      0
Fat_Conformidade_LAI_EIV          293    266     263      30       3       0      3
Fat_Eco_urbana                     14     14      14       0       0       0      0
Fat_Impressoes                     31     28      28       3       0       0      0
Fat_Impressoes_por_pessoa        4895   3854    3850    1045       4       0      0
Fat_Impressoes_por_setor          986    773     773     213       0       0      0
Fat_Metas_LAI_EIV                  12     12      12       0       0       0      0
Fat_Residuos_m³_m²                170    144     136      34       8       0      8
Fat_Ruido                         293    266     258      35       8      16      0
Fat_avanco_fisico                 172    145     145      27       0       0      0
Fat_consumo_agua_luz              666    612     557     109      55      70    132
Fat_soma_residuos                1860   1608    1294     566     314     287    544
TOTAL                           13544  11836   11434    2110     402     393    687
```

**As 19 tabelas têm exatamente as mesmas colunas, na mesma ordem, com os mesmos
tipos.** Nenhum erro de estrutura.

As diferenças de linha e de valor são **defasagem do `.pbix`, não erro do ETL**:
o `.pbix` foi atualizado pela última vez com dados até abr/2026 e a planilha já
está em set/2026.

- `SO_ETL` (2.110 linhas) — lançamentos novos na planilha.
- `SO_PBIX` (402 linhas) — quase tudo linha em branco que o Power Query
  importava como linha nula (o `.pbix` tem 39 linhas nulas em `Dim_Plano_acao`,
  32 em `Dim_emergencias_ambientais`, 24 em `Dim_Auditoria_Ambientais`); o ETL
  descarta linha totalmente vazia.
- `VAZIO` (393 células) — célula vazia no `.pbix` e preenchida na planilha
  (meta e resultado de auditoria de mai/2026, área `m²` de resíduos etc.).
- `MUDOU` (687 células) — valor diferente nos dois lados. Foi conferida uma
  amostra de cada tabela direto na planilha e **em todos os casos conferidos o
  valor do ETL é exatamente o que está na célula do Excel hoje**. Exemplos:
  - `Fat_soma_residuos`: `m²` era `1,0` (área não lançada) e hoje é a área real
    — 288 células; `Kg` era `'0'` e hoje tem valor — 126 células.
  - `Fat_consumo_agua_luz`: `Luz m²`/`Água m²` eram `1,0` e hoje são a área real.
  - `Fat_Conformidade_LAI_EIV`: 3 percentuais corrigidos na planilha.
  - `Fat_Residuos_m³_m²`: `Mês` foi de `'Março'` para `'março'` (só caixa).

### Duas diferenças que o validador achou e que foram corrigidas no ETL

1. **`Fat_Ruido.Horário`** — o `.pbix` guarda hora em ponto flutuante, então
   09:03 sai como `09:02:59,999999`. Truncar gerava 119 diferenças falsas de 1
   segundo; o validador passou a **arredondar** para o segundo.
2. **Espaço no fim de texto** — a planilha tem `'Emegência '` e o modelo tem
   `'Emegência'` (espaço no **início** é preservado, como em `' Ouvidoria'`).
   O ETL reproduz isso (`TRIM_TEXTO_FINAL` em `modelo_semantico.py`). Sem isso,
   `Dim_emergencias_ambientais`, `Dim_reclamacoes_vizinhos` e
   `Dim_Auditoria_Ambientais` acusavam divergência em toda linha de texto.

---

## Relacionamentos

Os **26 relacionamentos** do modelo do pbix foram copiados de
`PBIXRay(...).relationships` e estão em `RELACIONAMENTOS` (`tabelas.py`), com a
cardinalidade e o filtro cruzado de cada um. Eles existem no banco de três
formas:

1. **Índice único** nas duas dimensões (`uq_dim_obras`, `uq_dim_calendario`) —
   é o que uma FOREIGN KEY exige do lado "1". Criado junto com as tabelas.
2. **Índice** na coluna de junção de cada fato (`obra`, `data`, `mes_ano`).
3. **FOREIGN KEY** em `ddl/modelo_<dialeto>_fk.sql`, script separado.

No modelo puro do pbix, **13 das 26 não podiam ser criadas** — o dado tinha
órfão (valor de obra que não existe em `Dim_Obras`). Isso foi **corrigido por
padrão** (16/09/2026): o ETL agora saneia a coluna Obra antes de carregar, e as
**26 FKs valem hoje**, aplicadas de verdade sobre a carga real, zero violação.

O saneamento (`sanear_obras()` em `etl_excel_para_db.py`) resolve três causas
diferentes, cada uma com o tratamento certo — **nenhuma delas é defeito do
ETL**, é o estado do modelo no pbix, que nunca precisou de integridade
referencial:

| Causa | Tratamento | Onde doía |
|---|---|---|
| `Outros` — valor que o próprio Power Query atribui a setor que não é obra (Financeiro, Marketing, Sede...) | **Cadastrada** como obra válida em `Dim_Obras` (não é obra de verdade, mas precisa existir para o vínculo fechar) | Impressões por pessoa (4.279 linhas) e por setor (787), consumo água/luz (397) |
| `North Tower` — obra real, ainda sem linha em `Metas_LAI_EIV` (de onde sai `Dim_Obras`), grafada de 3 jeitos (`NORTH`, `North`, `North Tower`) | **Cadastrada** + grafias **normalizadas** para `North Tower` | Resíduos (60), reclamações, ruído, avanço físico, conformidade, auditorias |
| `Garden`, `Iconic Tower`, `Blue View Tower`, `ICONIC`, `IMPERIUM`, `GARDEN`, `Gran Place` — a **mesma obra já cadastrada**, só escrita diferente (abreviada, com sufixo "Tower" sobrando, maiúscula, ou erro de digitação) | Só **normalizadas** para a grafia oficial (`Garden Park`, `Iconic`, `Blue View`, `Imperium`, `Grand Place`) — nada é cadastrado, porque a obra já existe | Plano de ação (3), reclamações (8), emergências (5), eco urbana (1) |

Confirmado nos dados antes de normalizar: em `Dim_Plano_acao`, o valor
`'Garden'` aparece na mesma coluna ao lado de `'Imperium'`, `'Sapphire'` — é a
mesma obra abreviada, não uma obra nova. O mesmo padrão se repete nas outras
três tabelas.

**Este saneamento é só de vínculo — não muda nenhum fato.** Não cria, apaga
nem recalcula nenhuma linha; só corrige a etiqueta de obra para que a linha
passe a contar para a obra certa. As 3 linhas de "Garden" em Plano de Ação, por
exemplo, passam a aparecer nas ações de Garden Park (antes, ficavam soltas,
sem bater com nenhuma obra).

Use `--fiel-ao-pbix` para desligar o saneamento e obter a saída idêntica ao
pbix, órfãos incluídos — é o modo que `validar_vs_pbix.py` usa, para a prova
de fidelidade continuar valendo sobre o modelo *original*, sem a correção.

Duas observações sobre o modelo original:

- No pbix, o relacionamento de `Fat_Impressoes` está gravado **ao contrário**
  (`Dim_Calendario → Fat_Impressoes`, M:1). Parece engano de quem montou o
  modelo; aqui ficou na direção correta, fato → dimensão, e a nota está
  registrada em `_relacionamentos_modelo`.
- `cardinalidade` e `filtro_cruzado` (Single/Both) são conceitos do Power BI
  que não existem em SQL. Ficam registrados para quem for remontar o modelo em
  outra ferramenta.

---

## Defeitos do modelo reproduzidos de propósito

A regra deste ETL é **fidelidade ao pbix, inclusive nos defeitos que mudam
número de indicador** — corrigir um desses é decisão da área, não ajuste
técnico (mesmo critério já usado para a agregação de kg/m² de resíduos). Os
dois que **continuam em aberto**, aguardando decisão, estão em
`ALERTAS_MODELO` (`tabelas.py`) e são impressos a cada execução:

1. **`Fat_soma_residuos.Kg` é TEXTO**, com vírgula decimal pt-BR
   (`'1200,8399999999999'`). Em SQL exige
   `CAST(REPLACE(kg, ',', '.') AS FLOAT)`.
2. **`Fat_Metas_LAI_EIV` é tabela larga** — uma coluna por mês
   (`ago/25` … `dez/26`). Cada competência nova exige `ALTER TABLE`.

Quatro outros **já foram corrigidos por padrão** (16/09/2026) — três são
defeito de vínculo (ver seção "Relacionamentos" acima; não mudam nenhum valor
exibido, só ligam a linha à obra certa) e um muda um indicador de verdade, com
aprovação explícita da área. Continuam documentados em
`ALERTAS_MODELO_CORRIGIDOS` (`tabelas.py`) e só reaparecem com
`--fiel-ao-pbix`:

- `Fat_consumo_agua_luz.Obra` não reconhecia `SENNA` nem `NORTH TOWER`.
- `Dim_reclamacoes_vizinhos.Obra` não reconhecia `Iconic Tower`, `North Tower`,
  `Blue View Tower`.
- `Dim_Obras` tinha 12 obras; faltava `North Tower`.
- **`Fat_consumo_agua_luz."Água m³/m²"`** — ver a seção seguinte.

### Indicador de água corrigido (aprovado pela área em 16/09/2026)

O Power Query convertia essa coluna para inteiro: a planilha tem `0,206` e o
modelo guardava `0`. A tabela já tinha os dois ingredientes certos, decimais,
sem truncamento — `Água (m³)` e `Água m²` — então a correção é
`corrigir_intensidade_agua()` em `etl_excel_para_db.py`: recalcula
`Água m³/m²` como `Água (m³) ÷ Água m²` (nulo, não zero, quando a área é 0 ou
não lançada) e substitui o valor truncado.

Medido antes de aplicar (666 linhas, 521 com intensidade calculável): **72
linhas (13,8%) viravam zero só por causa do arredondamento**, não porque o
consumo foi zero. Confirmado depois de carregar de verdade: zero linhas
restantes com água > 0 e área > 0 mostrando intensidade zero.

Efeito colateral aceito: a coluna passa a ser **decimal** no banco (era
`BIGINT` no modelo fiel ao pbix). O ETL detecta isso sozinho pelo tipo real
do dado — tanto o schema (`_tipo_sql`) quanto o dicionário
(`_dicionario_modelo`) se ajustam automaticamente, sem precisar declarar nada
a mais em `tabelas.py`. Passe `--fiel-ao-pbix` para ver a coluna como no pbix
(truncada, `BIGINT`).

---

## Problema na planilha que o ETL detecta sozinho

O ETL lê o **intervalo declarado de cada tabela nomeada** — é exatamente isso
que o Power Query lê. Se alguém digitar linhas abaixo desse intervalo sem que a
tabela do Excel expanda, os dados **não entram no modelo e ninguém percebe**.
O ETL avisa:

```
ATENCAO Conformidade_LAI_EIV: ha 9 linha(s) com dados ABAIXO do intervalo
Conformidade LAI.EIV da tabela nomeada. Essas linhas nao entram no modelo.
```

Hoje essas 9 linhas (dez/2025, linhas 391–399) são **cópia** de linhas que já
estão dentro do intervalo — não há perda de dado. Mas vale limpar, porque o
aviso perde o valor se virar rotina.

O ETL também avisa quando:

- uma coluna esperada some da planilha (seria preenchida com NULL em silêncio);
- um valor de obra não existe em `Dim_Obras`;
- a chave de negócio se repete.

---

## Mantendo o banco atualizado

A planilha vive em
`\\servidor\Qualidade\AMBIENTAL\GESTÃO AMBIENTAL\INDICADORES AMBIENTAIS\POWER BI`
(a pasta que os usuários acessam como `U:\...`). A carga é **full refresh**:
apaga e regrava tudo numa transação só, então rodar de novo sem necessidade não
causa dano — mas também não é de graça.

### O ponto que quebra esse tipo de automação

**`U:` não existe para tarefa agendada nem para serviço Windows.** Unidade
mapeada é por sessão de usuário. Uma tarefa rodando como conta de serviço vê
`U:\...` como "caminho não encontrado".

Por isso a origem **tem que ser o caminho UNC** (`\\servidor\share\...`), e a
conta que roda a tarefa precisa de leitura no share. O ETL detecta esse caso e
falha com a mensagem explicando, em vez de deixar passar.

> **Falta descobrir qual é o UNC.** Nenhum script do projeto tem o caminho de
> rede — todos usam `U:`. Rode `net use U:` numa máquina onde o drive está
> mapeado, ou pergunte à TI qual é o servidor de arquivos.

### Agendando

1. Edite [`Atualizar_Banco.bat`](Atualizar_Banco.bat): o caminho UNC da
   planilha e os dados do PostgreSQL.
2. Crie a tarefa agendada (PowerShell como administrador):

```powershell
$acao    = New-ScheduledTaskAction -Execute 'C:\Apps\Dash_ambiental_qualidade\etl\Atualizar_Banco.bat'
$gatilho = New-ScheduledTaskTrigger -Daily -At 02:30
$conta   = New-ScheduledTaskPrincipal -UserId 'DOMINIO\svc_etl_ambiental' `
           -LogonType Password -RunLevel Limited
Register-ScheduledTask -TaskName 'ETL Ambiental - carga do banco' `
    -Action $acao -Trigger $gatilho -Principal $conta `
    -Description 'Recarrega o banco ambiental a partir da planilha da rede.'
```

Use uma **conta de serviço de domínio** com leitura no share — não
`LocalSystem`, que não tem identidade na rede. (O serviço atual do dashboard
roda como `LocalSystem`; isso é um problema separado, já anotado na análise do
projeto.)

Sugestão de horário: **02:30**, meia hora antes do refresh do dashboard (03:00),
para o dashboard já ler dado novo quando ele passar a vir do banco.

### O que o ETL faz sozinho a cada execução

| Passo | Por quê |
|---|---|
| Copia a planilha para disco local antes de ler | O arquivo fica aberto no Excel o dia inteiro; se a leitura cair durante um "Salvar", o `.xlsx` vem truncado. A cópia é validada (é um zip) e refeita até 3× com 20s de espera. |
| Calcula o `sha256` do arquivo | É o que responde "a planilha mudou?". |
| Com `--se-mudou`, pula a carga se o hash for igual ao da última carga OK | Evita regravar 13.544 linhas à toa. |
| Lê e **valida antes** de decidir pular | Assim um problema na origem (coluna renomeada, linha fora do intervalo da tabela) aparece mesmo em dia sem alteração. |
| Grava uma linha em `_carga_execucao` | Sucesso **e** falha, com o hash e a data de modificação da planilha. |
| Devolve código de saída | `0` ok/sem mudança · `1` falha na carga · `2` erro de estrutura · `3` planilha inacessível. |

### Monitorando

O log de texto fica em `logs/etl_banco.log`, mas ninguém lê log de texto. O que
vale monitorar é a tabela `_carga_execucao`:

```sql
-- a carga de hoje rodou e deu certo?
SELECT status, iniciado_em, terminado_em, linhas_total, mensagem
FROM ambiental._carga_execucao
ORDER BY terminado_em DESC LIMIT 5;

-- alerta: nenhuma carga OK nas ultimas 26 horas
SELECT COUNT(*) AS cargas_ok_recentes
FROM ambiental._carga_execucao
WHERE status = 'ok'
  AND terminado_em::timestamp > now() - interval '26 hours';
```

Se esse segundo `SELECT` voltar `0`, o banco está parado — seja porque a tarefa
não rodou, seja porque a planilha ficou inacessível.

### Quando a planilha for substituída

O ETL lê **tabelas nomeadas** (ListObjects), não abas. Se alguém recriar a
planilha e as tabelas nomeadas perderem o nome (`SOMA_RESIDUOS`,
`Conformidade_LAI_EIV`, ...), a carga falha com mensagem clara em vez de
carregar vazio. O mesmo vale para coluna renomeada: o ETL avisa em vez de
preencher `NULL` em silêncio.

---

## Próximos passos sugeridos

Nada disso foi feito, porque muda o modelo e precisa de decisão da área:

1. **Apontar o `gerar_dashboard_pbix.py` para o banco.** A troca é pequena: a
   função `T(nome)` passa a fazer `pd.read_sql` em vez de `m.get_table`. As
   colunas chegam com o nome físico, então precisa de um `rename` a partir de
   `_dicionario_modelo` — ou ajustar o gerador para os nomes novos.
2. **Despivotar `Fat_Metas_LAI_EIV`** para `(obra, competência, meta)`.
3. **Converter `Fat_soma_residuos.Kg` para numérico** no banco.
4. **Corrigir `"Água m³/m²"`** para decimal (hoje o indicador de água está
   zerado).
5. **Completar `Dim_Obras`** com North Tower e normalizar os nomes de obra de
   `Dim_reclamacoes_vizinhos`.
6. **Chave primária de verdade.** Hoje as chaves de negócio viram índice, não
   PK, porque a planilha repete chave em 4 tabelas
   (`Dim_Plano_acao`, `Fat_Impressoes_por_setor`, `Fat_Residuos_m³_m²`,
   `Fat_consumo_agua_luz`).

---

## Dependências

Todas estão no `requirements.txt` da raiz do projeto:

```powershell
py -3.12 -m pip install -r ..\requirements.txt
```

- `pandas`, `numpy`, `openpyxl` — leitura da planilha e transformações.
- `sqlalchemy` + **`psycopg2-binary`** — carga em PostgreSQL, o destino do
  modelo. Sem o `psycopg2` o ETL só valida e gera DDL; não carrega.
- `pbixray` — só para rodar `validar_vs_pbix.py`.

No servidor de homologação (10.1.1.155) o Python usado
(`C:\Users\julio.comin\AppData\Local\Programs\Python\Python312`) já tem
`sqlalchemy 2.0.49` e `psycopg2 2.9.9`. **Em produção (10.1.1.151) isso ainda
precisa ser conferido** — os dois entraram no `requirements.txt` só em 15/09/2026.

O ETL carrega **só em PostgreSQL** — SQLite e SQL Server foram removidos em
15/09/2026. Para voltar a gerar DDL de outro dialeto seria preciso reintroduzir
o parâmetro em `gerar_ddl`/`gerar_ddl_fk`.
