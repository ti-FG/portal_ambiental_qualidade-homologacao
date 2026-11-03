# Lançamentos Ambientais (FG · Qualidade)

Sistema web de lançamento que substitui a planilha
`Controle Ambiental 2025_oficial_db.xlsx`. A Qualidade lança os dados aqui, eles
são gravados no PostgreSQL (schema `portal_ambiental_qualidade`) e os painéis
passam a ler do banco.

**Homologação:** http://10.1.1.155:8095/ (servidor vSRV-HML-DEV)

## Estado em 21/10/2026: primeira versão para demonstração

**Regra de interface (pedido da área, 21/10/2026): fidelidade à planilha.** Cada aba
da planilha é uma tela com o MESMO nome, e as colunas têm os MESMOS nomes e a
mesma ordem (Ano, Mês, Mês/Ano no formato `set-26`...). Coluna que era fórmula
leva o marcador ƒ; coluna nova leva "novo" e vem depois das da planilha. A faixa
"Abas da planilha" mostra todas as abas na ordem do Excel — as que ainda não foram
migradas aparecem apagadas. Não renomear coluna sem combinar com a Qualidade.

Menu **Lançamentos** (suspenso) com as 17 abas da planilha, agrupadas: Resíduos (SOMA
RESÍDUOS, Resíduos m³m²) · Conformidade LAI/EIV (Conformidade LAI.EIV, Metas LAI EIV, NC LAI
e EIV, Ruído, Eco urbana) · Obra (Avanço físico) · Consumo (Água e luz, Impressões, Impressões
por setor, Impressões por pessoa) · Gestão ambiental (Auditoria Ambientais, Plano de Ação,
Emergências Ambientais, Reclamações vizinhos, Ações SGA). **Todas no sistema desde 21/10/2026**;
o painel não lê mais nada da planilha.

As 9 abas que eram listas usam um motor genérico (`LISTAS` em www/dados.js e `LISTAS_UI`
em public/index.html): cada aba é uma configuração de colunas. Impressões por setor/pessoa têm
o botão "Colar do relatório". Tabela da aba Auditoria Ambientais = `auditoria_ambiental`
(`auditoria` é o registro de acessos do sistema — não confundir).

**Reclamações vizinhos ← Salesforce (23/10/2026).** Botão "Buscar no Salesforce" na aba: lê do
DW corporativo (`dw_fg`, só leitura) os chamados com `Tipo_de_Cliente__c = 'Comunidade'` em
`silver.salesforce_case` e mostra uma prévia. Quem lança escolhe o que entra (vêm marcadas as do
tipo Reclamação). A obra é a conta do chamado (`AccountId`), porque `Empreendimento__c` vem
vazio. A obra escolhida para uma conta fica em `salesforce_conta`. Chamado já na lista pelo Cod.
Chamado só ganha o vínculo (`reclamacao.sf_case_id`), sem duplicar. Os novos entram com
origem `salesforce`, Fonte "Salesforce" e Categoria = tipo do chamado. A Ouvidoria continua manual.
Credencial em `DWC_PGHOST/PGPORT/PGDATABASE/PGUSER/PGPASSWORD` no `www/.env`. Hoje é
pessoal: pedir um usuário de serviço só leitura antes da produção. O Salesforce no DW só
tem chamados a partir de 18/12/2025.

| Tela | Substitui na planilha |
|---|---|
| Início | — (o que falta lançar no mês, por obra; fechar/reabrir mês; conferência e recarga da planilha) |
| Painel | o dashboard de sempre (8090), gerado DESTE banco e regerado ~5 s depois de cada gravação. Também em http://10.1.1.155:8095/painel |
| Avanço físico | aba Avanço físico |
| Conformidade e ruído | abas Conformidade LAI.EIV, Metas LAI EIV, NC LAI e EIV, Ruído |
| Resíduos | abas SOMA RESÍDUOS e Resíduos m³m² |
| Água e luz | aba Água e luz |
| Cadastros | obras, tipos de resíduo (densidade), unidades consumidoras, metas com vigência, assuntos de NC, parâmetros |
| Histórico | — (log de alterações campo a campo) |
| Usuários | — (aprovação de acesso: consulta / lançamento / administrador) |

**Ainda não feito (próximas etapas):**
- Telas de auditorias, plano de ação, emergências, reclamações, ações SGA e impressões.
- Hoje `painel/gerar_painel.py` roda o gerador oficial sem alterá-lo: troca o
  leitor do .pbix por tabelas montadas do banco (resíduos, conformidade, NCs,
  ruído, avanço, água/luz). O que ainda não tem tela (auditorias, plano de ação,
  emergências, reclamações, SGA, impressões) continua vindo da planilha pelo ETL.
  Se o gerador oficial mudar os trechos que o adaptador substitui, a geração
  falha com mensagem clara (não gera painel errado).
- Promoção para o 152 (pedido à TI: schema e role `app_ambiental`).

## Como está montado

Mesmo desenho do FG Metas Engenharia e do Portfólio de Projetos:
- Node sem framework, com `pg` e `ldapts`.
- Um único HTML em `www/public/`.
- Login no AD (`www/ldap.js`, copiado do Metas).
- Sessão no banco, guardando só o hash do token.
- API REST por recurso (ver o cabeçalho de `www/server.js`).

```
www/server.js        rotas, login, sessões, permissões
www/dados.js         todo o SQL (lançamentos com log campo a campo, trava de mês)
www/public/          index.html (telas) + logo
db/schema.sql        tabelas, cadastros e views de cálculo (idempotente)
db/carga_inicial.sql GERADO por carga/carga_planilha.py (histórico da planilha)
carga/               carga_planilha.py + relatorio_carga.txt (decisões de saneamento)
hml/memdb.mjs        banco da homologação (PGlite em 127.0.0.1:15439, persistido em hml/dados)
servico_hml.ps1      instala/atualiza os serviços NSSM
```

## Operação na homologação

Dois serviços NSSM, com início automático e logs em `logs/`:

| Serviço | O que roda | Porta |
|---|---|---|
| `FGAmbientalQualidadeMemdb` | `hml/memdb.mjs` | 127.0.0.1:15439 |
| `FGAmbientalQualidadeHML` | `www/server.js` (depende do memdb) | 0.0.0.0:8095 |

- **Mudou o servidor:** `nssm restart FGAmbientalQualidadeHML`. Ninguém é deslogado.
- **Mudou o schema:** reiniciar o memdb, que reaplica `db/schema.sql`. Antes
  disso, parar o HML (ele depende do memdb).
- **Voltar ao estado da planilha:** parar os dois serviços, apagar `hml/dados`,
  rodar `carga/carga_planilha.py` (se a planilha mudou) e subir os dois de novo.
- **Instalar ou atualizar os serviços:** `pwsh -File servico_hml.ps1`, como administrador.

**Pendente para o login funcionar:** preencher `LDAP_*` em `www/.env` com os
mesmos valores do Painel de Metas / Projetos. Sem isso, o login responde "login
pela rede ainda não configurado".

## Regras de dados (o que corrige da planilha)

- **Obra:** é um cadastro. Na planilha a mesma obra aparecia com cerca de 30
  grafias diferentes.
- **Competência:** é uma data (1º dia do mês). Antes era o mês por extenso
  convertido por PROCV.
- **Kg de resíduo:** é sempre `m³ × densidade do tipo`. Tipo sem densidade fica
  com kg em branco, e não zero.
- **Área do mês:** é `área total da obra × avanço do mês`. É o único denominador
  de kg/m², kWh/m² e m³/m².
- **Meta de conformidade:** tem vigência. Vale a última vigência até a
  competência, em vez do PROCV que lia sempre ago/25.
- **Sede:** é a soma das salas e não se digita. Na planilha, somar todos os
  setores contava a Sede em dobro.
- **Mês fechado:** só o administrador altera, com motivo obrigatório.
- **Histórico:** toda alteração fica em `log_alteracao`, com o valor de antes e o
  de depois.

**Sincronização automática (28/10/2026).** Salesforce e Prevision rodam sozinhos às 09:00, 13:00 e
17:00 (Brasília; se o servidor estava fora, roda ao voltar) e sob demanda (admin, tela Integrações).
Cada rodada fica em `sincronizacao`; `SYNC_AUTO=0` no `.env` desliga o agendador.
- *Salesforce:* importa só Reclamação com obra já conhecida (conta lembrada ou linha da lista) e mês
  aberto. Conta nova, obra só "sugerida" ou mês fechado = pendência (tela Integrações), resolvida em
  Reclamações vizinhos → Buscar no Salesforce; a obra escolhida fica em `salesforce_conta`.
- *Prevision:* `prevision_projeto` liga ProjectId → obra (ou "ignorar"). Projeto novo em construção sem
  ligação = pendência. Avanço = curva monetária, foto mais recente, realizado fim do mês − mês anterior,
  últimos 12 meses completos. Mês já lançado à mão que difere NÃO é sobrescrito: vira divergência.
  Mês fechado não é alterado. Origem gravada: `prevision`. Painel é regerado quando algo muda.
