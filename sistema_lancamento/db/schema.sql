-- =============================================================================
-- schema.sql — Lançamentos Ambientais (Qualidade FG), 21/10/2026
--
-- Substitui a planilha "Controle Ambiental 2025_oficial_db.xlsx". O que era
-- digitado na planilha vira tabela de LANÇAMENTO; o que era fórmula vira conta
-- do sistema (views no fim deste arquivo); o que era texto livre repetido em
-- cada aba (obra, tipo de resíduo, setor) vira CADASTRO.
--
-- Idempotente: pode rodar de novo sobre um banco que já tem tudo.
-- Guia de nomenclatura de schemas (TI DEV): tipo portal (sistema web de
-- lançamento) + assunto ambiental + departamento qualidade.
-- Nada em public. No 152 a TI cria o schema e o role (ver PEDIDO_TI_INFRA.md).
-- =============================================================================

CREATE SCHEMA IF NOT EXISTS portal_ambiental_qualidade;
SET search_path TO portal_ambiental_qualidade;

-- ── Acesso ──────────────────────────────────────────────────────────────────
-- Mesmo fluxo do FG Engenharia / Projetos: a conta do AD pede acesso na
-- primeira vez e um administrador aprova com um nível.
--   leitor  = só consulta
--   editor  = lança e corrige meses abertos
--   admin   = cadastros, usuários, fecha/reabre mês, corrige mês fechado
CREATE TABLE IF NOT EXISTS usuario (
  login          text PRIMARY KEY,
  nome           text,
  email          text,
  situacao       text NOT NULL DEFAULT 'pendente' CHECK (situacao IN ('pendente','aprovado','rejeitado')),
  nivel          text NOT NULL DEFAULT 'leitor'   CHECK (nivel IN ('leitor','editor','admin')),
  pedido_em      timestamptz NOT NULL DEFAULT now(),
  decidido_por   text,
  decidido_em    timestamptz,
  ultimo_acesso  timestamptz
);

-- só o sha256 do token, nunca o token (padrão do Projetos)
CREATE TABLE IF NOT EXISTS sessao (
  token_sha256  text PRIMARY KEY,
  login         text NOT NULL,
  nome          text,
  email         text,
  criada_em     timestamptz NOT NULL,
  expira_em     timestamptz NOT NULL
);
CREATE INDEX IF NOT EXISTS sessao_login_idx ON sessao (login);

-- eventos de acesso (login, aprovação, fechamento de mês)
CREATE TABLE IF NOT EXISTS auditoria (
  id       bigserial PRIMARY KEY,
  quando   timestamptz NOT NULL DEFAULT now(),
  login    text,
  acao     text NOT NULL,
  detalhe  jsonb
);

-- toda alteração de dado, campo a campo (padrão do Jurídico). Nada se apaga
-- sem deixar aqui o valor de antes.
CREATE TABLE IF NOT EXISTS log_alteracao (
  id       bigserial PRIMARY KEY,
  quando   timestamptz NOT NULL DEFAULT now(),
  login    text NOT NULL,
  tabela   text NOT NULL,
  chave    jsonb NOT NULL,
  acao     text NOT NULL CHECK (acao IN ('incluir','alterar','excluir')),
  antes    jsonb,
  depois   jsonb,
  motivo   text
);
CREATE INDEX IF NOT EXISTS log_alteracao_quando_idx ON log_alteracao (quando DESC);
CREATE INDEX IF NOT EXISTS log_alteracao_tabela_idx ON log_alteracao (tabela);

-- ── Cadastros ───────────────────────────────────────────────────────────────
-- Na planilha a mesma obra aparecia de ~30 jeitos (NORTH, North, North Tower,
-- "Senna ", GARDEN...). Aqui ela existe uma vez e todo lançamento aponta para ela.
CREATE TABLE IF NOT EXISTS obra (
  id                  serial PRIMARY KEY,
  nome                text NOT NULL UNIQUE,
  ativa               boolean NOT NULL DEFAULT true,
  area_total_m2       numeric(12,2),           -- ficha técnica; era "Total m²" repetido em cada linha do Avanço físico
  meta_residuo_kg_m2  numeric(10,3),           -- era "Meta Kg" digitado linha a linha (36,6 / Senna 50)
  observacao          text,
  atualizado_por      text,
  atualizado_em       timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS tipo_residuo (
  id               serial PRIMARY KEY,
  nome             text NOT NULL UNIQUE,
  classe           char(1) NOT NULL CHECK (classe IN ('A','B','C','D')),   -- CONAMA 307
  densidade_kg_m3  numeric(10,3),             -- era fixa num SE() aninhado da planilha; nula = sem densidade, kg não calculado
  perigoso         boolean NOT NULL DEFAULT false,
  ativo            boolean NOT NULL DEFAULT true,
  ordem            int NOT NULL DEFAULT 100,
  legado           boolean NOT NULL DEFAULT false,  -- só para o histórico até 2025, lançado por classe sem detalhe
  atualizado_por   text,
  atualizado_em    timestamptz NOT NULL DEFAULT now()
);

-- unidade consumidora de água/luz (obra, sala da sede, plantão, depósito...)
CREATE TABLE IF NOT EXISTS unidade_consumidora (
  id              serial PRIMARY KEY,
  nome            text NOT NULL UNIQUE,
  tipo            text NOT NULL CHECK (tipo IN ('obra','administrativo')),
  obra_id         int REFERENCES obra(id),
  agrupa_em_id    int REFERENCES unidade_consumidora(id),  -- Salas 02..07 -> Sede: a Sede é a SOMA, não se digita
  ativa           boolean NOT NULL DEFAULT true,
  ordem           int NOT NULL DEFAULT 100,
  atualizado_por  text,
  atualizado_em   timestamptz NOT NULL DEFAULT now(),
  CHECK (tipo <> 'obra' OR obra_id IS NOT NULL)
);

-- Meta de conformidade com vigência: vale a partir da competência e até a
-- próxima. Na planilha o PROCV lia sempre a coluna de ago/25.
CREATE TABLE IF NOT EXISTS meta_conformidade (
  obra_id         int NOT NULL REFERENCES obra(id),
  vigente_desde   date NOT NULL CHECK (extract(day FROM vigente_desde) = 1),
  meta            numeric(5,4) NOT NULL CHECK (meta BETWEEN 0 AND 1),
  atualizado_por  text,
  atualizado_em   timestamptz NOT NULL DEFAULT now(),
  PRIMARY KEY (obra_id, vigente_desde)
);

CREATE TABLE IF NOT EXISTS assunto_nc (
  id      serial PRIMARY KEY,
  tipo    text NOT NULL CHECK (tipo IN ('LAI','EIV')),
  nome    text NOT NULL,
  ativo   boolean NOT NULL DEFAULT true,
  UNIQUE (tipo, nome)
);

-- Desde 21/10/2026 a planilha não separa mais LAI de EIV: uma lista única de
-- tipos de NC. 'LAI/EIV' é esse tipo; LAI e EIV ficam para o histórico antigo.
ALTER TABLE assunto_nc DROP CONSTRAINT IF EXISTS assunto_nc_tipo_check;
ALTER TABLE assunto_nc ADD CONSTRAINT assunto_nc_tipo_check CHECK (tipo IN ('LAI','EIV','LAI/EIV'));
ALTER TABLE assunto_nc ADD COLUMN IF NOT EXISTS ordem int NOT NULL DEFAULT 100;

CREATE TABLE IF NOT EXISTS parametro (
  chave           text PRIMARY KEY,
  valor           numeric NOT NULL,
  descricao       text,
  atualizado_por  text,
  atualizado_em   timestamptz NOT NULL DEFAULT now()
);
INSERT INTO parametro (chave, valor, descricao) VALUES
  ('limite_ruido_db', 70, 'Limite de ruído (dB) para a medição ser considerada acima do permitido'),
  ('meta_residuo_kg_m2_padrao', 36.6, 'Meta de resíduos (kg/m²) para obra sem meta própria')
ON CONFLICT (chave) DO NOTHING;

-- ── Fechamento de mês ───────────────────────────────────────────────────────
-- Competência fechada: só admin altera, e com motivo obrigatório.
CREATE TABLE IF NOT EXISTS fechamento (
  competencia  date PRIMARY KEY CHECK (extract(day FROM competencia) = 1),
  fechado_por  text NOT NULL,
  fechado_em   timestamptz NOT NULL DEFAULT now()
);

-- ── Lançamentos ─────────────────────────────────────────────────────────────
-- competencia = 1º dia do mês. Na planilha: Ano (número) + Mês (texto) + PROCV
-- na aba Auxiliar para montar a data.
-- Toda tabela tem quem lançou e quem alterou por último; o detalhe fica em
-- log_alteracao.

CREATE TABLE IF NOT EXISTS avanco_fisico (
  obra_id        int  NOT NULL REFERENCES obra(id),
  competencia    date NOT NULL CHECK (extract(day FROM competencia) = 1),
  avanco_pct     numeric(8,6) NOT NULL CHECK (avanco_pct BETWEEN 0 AND 1),   -- avanço REALIZADO no mês
  origem         text NOT NULL DEFAULT 'sistema',                            -- 'planilha' = carga inicial
  criado_por     text, criado_em timestamptz NOT NULL DEFAULT now(),
  atualizado_por text, atualizado_em timestamptz NOT NULL DEFAULT now(),
  PRIMARY KEY (obra_id, competencia)
);

CREATE TABLE IF NOT EXISTS conformidade (
  obra_id          int  NOT NULL REFERENCES obra(id),
  competencia      date NOT NULL CHECK (extract(day FROM competencia) = 1),
  itens_conformes  int CHECK (itens_conformes >= 0),
  itens_avaliados  int CHECK (itens_avaliados > 0),
  percentual       numeric(6,4) NOT NULL CHECK (percentual BETWEEN 0 AND 1),
  observacao       text,
  origem           text NOT NULL DEFAULT 'sistema',
  criado_por     text, criado_em timestamptz NOT NULL DEFAULT now(),
  atualizado_por text, atualizado_em timestamptz NOT NULL DEFAULT now(),
  PRIMARY KEY (obra_id, competencia),
  CHECK (itens_conformes IS NULL OR itens_avaliados IS NULL OR itens_conformes <= itens_avaliados)
);

-- assuntos de não conformidade apontados no checklist do mês (seleção múltipla)
CREATE TABLE IF NOT EXISTS conformidade_nc (
  obra_id        int  NOT NULL REFERENCES obra(id),
  competencia    date NOT NULL,
  assunto_id     int  NOT NULL REFERENCES assunto_nc(id),
  criado_por     text, criado_em timestamptz NOT NULL DEFAULT now(),
  PRIMARY KEY (obra_id, competencia, assunto_id)
);

CREATE TABLE IF NOT EXISTS medicao_ruido (
  id             bigserial PRIMARY KEY,
  obra_id        int  NOT NULL REFERENCES obra(id),
  competencia    date NOT NULL CHECK (extract(day FROM competencia) = 1),
  ponto          text NOT NULL DEFAULT 'P1',
  horario        time,
  db             numeric(6,2) NOT NULL CHECK (db > 0 AND db < 200),
  empresa        text NOT NULL DEFAULT 'Alameda',
  origem         text NOT NULL DEFAULT 'sistema',
  criado_por     text, criado_em timestamptz NOT NULL DEFAULT now(),
  atualizado_por text, atualizado_em timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS medicao_ruido_comp_idx ON medicao_ruido (competencia, obra_id);

-- Volume por obra × mês × tipo. Classe e densidade vêm do cadastro do tipo;
-- kg e kg/m² são calculados (view), nunca digitados.
CREATE TABLE IF NOT EXISTS residuo (
  obra_id          int  NOT NULL REFERENCES obra(id),
  competencia      date NOT NULL CHECK (extract(day FROM competencia) = 1),
  tipo_residuo_id  int  NOT NULL REFERENCES tipo_residuo(id),
  volume_m3        numeric(12,3) NOT NULL CHECK (volume_m3 >= 0),
  origem           text NOT NULL DEFAULT 'sistema',
  criado_por     text, criado_em timestamptz NOT NULL DEFAULT now(),
  atualizado_por text, atualizado_em timestamptz NOT NULL DEFAULT now(),
  PRIMARY KEY (obra_id, competencia, tipo_residuo_id)
);

CREATE TABLE IF NOT EXISTS consumo (
  uc_id          int  NOT NULL REFERENCES unidade_consumidora(id),
  competencia    date NOT NULL CHECK (extract(day FROM competencia) = 1),
  luz_kwh        numeric(12,2) CHECK (luz_kwh >= 0),
  agua_m3        numeric(12,2) CHECK (agua_m3 >= 0),
  origem         text NOT NULL DEFAULT 'sistema',
  criado_por     text, criado_em timestamptz NOT NULL DEFAULT now(),
  atualizado_por text, atualizado_em timestamptz NOT NULL DEFAULT now(),
  linha          smallint NOT NULL DEFAULT 1,
  PRIMARY KEY (uc_id, competencia, linha)
);

-- 30/10/2026, pedido da área: (1) a linha pode ficar sem luz e sem água (a planilha tem linhas
-- em branco; fica NULL); (2) uma exceção de histórico: a planilha tem 2 linhas da Assistência
-- técnica em mar/25. De 2ª linha só mar/25 (linha = 2); em qualquer outro mês vale 1 linha por
-- unidade. O sistema só grava/edita a linha 1. Idempotente: roda a cada subida do banco.
ALTER TABLE consumo ADD COLUMN IF NOT EXISTS linha smallint NOT NULL DEFAULT 1;
ALTER TABLE consumo DROP CONSTRAINT IF EXISTS consumo_check;
ALTER TABLE consumo DROP CONSTRAINT IF EXISTS consumo_pkey;
ALTER TABLE consumo ADD CONSTRAINT consumo_pkey PRIMARY KEY (uc_id, competencia, linha);
ALTER TABLE consumo DROP CONSTRAINT IF EXISTS consumo_linha_check;
ALTER TABLE consumo ADD CONSTRAINT consumo_linha_check CHECK (linha = 1 OR (linha = 2 AND competencia = DATE '2025-03-01'));

-- ── Contas do sistema (o que eram fórmulas na planilha) ─────────────────────

-- meta de conformidade VIGENTE em cada competência (substitui o PROCV)
CREATE OR REPLACE VIEW v_conformidade AS
SELECT c.obra_id, o.nome AS obra, c.competencia, c.itens_conformes, c.itens_avaliados,
       c.percentual,
       (SELECT m.meta FROM meta_conformidade m
         WHERE m.obra_id = c.obra_id AND m.vigente_desde <= c.competencia
         ORDER BY m.vigente_desde DESC LIMIT 1) AS meta
  FROM conformidade c JOIN obra o ON o.id = c.obra_id;

-- área construída no mês = área total da obra × avanço do mês. É o ÚNICO
-- denominador de todas as intensidades (na planilha cada aba tinha o seu).
CREATE OR REPLACE VIEW v_area_mes AS
SELECT a.obra_id, o.nome AS obra, a.competencia, a.avanco_pct, o.area_total_m2,
       round(o.area_total_m2 * a.avanco_pct, 2) AS area_mes_m2
  FROM avanco_fisico a JOIN obra o ON o.id = a.obra_id;

CREATE OR REPLACE VIEW v_residuo AS
SELECT r.obra_id, o.nome AS obra, r.competencia, t.classe, t.nome AS tipo, t.legado,
       r.volume_m3, t.densidade_kg_m3,
       round(r.volume_m3 * t.densidade_kg_m3, 2) AS kg,
       am.area_mes_m2,
       CASE WHEN am.area_mes_m2 > 0 THEN round(r.volume_m3 * t.densidade_kg_m3 / am.area_mes_m2, 4) END AS kg_m2,
       coalesce(o.meta_residuo_kg_m2, (SELECT valor FROM parametro WHERE chave = 'meta_residuo_kg_m2_padrao')) AS meta_kg_m2
  FROM residuo r
  JOIN obra o ON o.id = r.obra_id
  JOIN tipo_residuo t ON t.id = r.tipo_residuo_id
  LEFT JOIN v_area_mes am ON am.obra_id = r.obra_id AND am.competencia = r.competencia;

-- consumo com a Sede (e qualquer grupo) somando as unidades filhas
CREATE OR REPLACE VIEW v_consumo AS
SELECT u.id AS uc_id, u.nome AS unidade, u.tipo, u.obra_id, o.nome AS obra, c.competencia,
       c.luz_kwh, c.agua_m3,
       am.area_mes_m2,
       CASE WHEN am.area_mes_m2 > 0 THEN round(c.luz_kwh / am.area_mes_m2, 4) END AS luz_kwh_m2,
       CASE WHEN am.area_mes_m2 > 0 THEN round(c.agua_m3 / am.area_mes_m2, 4) END AS agua_m3_m2
  FROM consumo c
  JOIN unidade_consumidora u ON u.id = c.uc_id
  LEFT JOIN obra o ON o.id = u.obra_id
  LEFT JOIN v_area_mes am ON am.obra_id = u.obra_id AND am.competencia = c.competencia;

-- =============================================================================
-- Abas que eram listas na planilha (21/10/2026). Uma tabela por aba, com as
-- MESMAS colunas da aba (nomes em snake_case). O painel remonta a aba com os
-- cabeçalhos originais e passa pela mesma transformação do Power Query.
-- =============================================================================

-- Auditoria Ambientais: Obra, Ano, Mês, Mês/Ano, Meta, Realizado, Principais ocorrencias da última auditória
CREATE TABLE IF NOT EXISTS auditoria_ambiental (
  id             bigserial PRIMARY KEY,
  obra_id        int  NOT NULL REFERENCES obra(id),
  competencia    date NOT NULL CHECK (extract(day FROM competencia) = 1),
  meta           numeric(6,4) CHECK (meta BETWEEN 0 AND 1),
  realizado      boolean,
  ocorrencias    text,
  origem         text NOT NULL DEFAULT 'sistema',
  criado_por     text, criado_em timestamptz NOT NULL DEFAULT now(),
  atualizado_por text, atualizado_em timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS auditoria_ambiental_comp_idx ON auditoria_ambiental (competencia);

-- Plano de Ação: Obra, Indicador, Análise, Ação, Responsável, Prazo, Status (sem mês: é uma lista viva)
CREATE TABLE IF NOT EXISTS plano_acao (
  id             bigserial PRIMARY KEY,
  obra_id        int  NOT NULL REFERENCES obra(id),
  indicador      text,
  analise        text,
  acao           text,
  responsavel    text,
  prazo          date,
  status         text,
  origem         text NOT NULL DEFAULT 'sistema',
  criado_por     text, criado_em timestamptz NOT NULL DEFAULT now(),
  atualizado_por text, atualizado_em timestamptz NOT NULL DEFAULT now()
);

-- Emergências Ambientais: Obra, Mês/Ano (data do evento), Fonte, N° FMC, Ocorrência, Status/Ações, Categoria
CREATE TABLE IF NOT EXISTS emergencia (
  id             bigserial PRIMARY KEY,
  obra_id        int  NOT NULL REFERENCES obra(id),
  data           date NOT NULL,
  competencia    date GENERATED ALWAYS AS (date_trunc('month', data::timestamp)::date) STORED,
  fonte          text,
  n_fmc          bigint,
  ocorrencia     text,
  status_acoes   text,
  categoria      text,
  origem         text NOT NULL DEFAULT 'sistema',
  criado_por     text, criado_em timestamptz NOT NULL DEFAULT now(),
  atualizado_por text, atualizado_em timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS emergencia_comp_idx ON emergencia (competencia);

-- Reclamações vizinhos: Mês/Ano (data), Obra (+ canal, que na planilha vinha no texto da obra), Fonte, Cod. Chamado, Ocorrência, Status/Ações, Categoria
CREATE TABLE IF NOT EXISTS reclamacao (
  id             bigserial PRIMARY KEY,
  obra_id        int  NOT NULL REFERENCES obra(id),
  canal          text,
  data           date NOT NULL,
  competencia    date GENERATED ALWAYS AS (date_trunc('month', data::timestamp)::date) STORED,
  fonte          text,
  cod_chamado    text,
  ocorrencia     text,
  status_acoes   text,
  categoria      text,
  origem         text NOT NULL DEFAULT 'sistema',
  criado_por     text, criado_em timestamptz NOT NULL DEFAULT now(),
  atualizado_por text, atualizado_em timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS reclamacao_comp_idx ON reclamacao (competencia);
-- Chamados importados do Salesforce (via DW corporativo, silver.salesforce_case).
-- sf_case_id evita importar o mesmo chamado duas vezes; linha da planilha com o
-- mesmo Cod. Chamado só ganha o vínculo, não é duplicada.
ALTER TABLE reclamacao ADD COLUMN IF NOT EXISTS sf_case_id text;
CREATE UNIQUE INDEX IF NOT EXISTS reclamacao_sf_case_uk ON reclamacao (sf_case_id) WHERE sf_case_id IS NOT NULL;
-- Conta do Salesforce -> obra. No Salesforce o Empreendimento__c vem vazio nos
-- chamados de vizinhos: a obra é a conta (AccountId). Escolhida uma vez na importação.
CREATE TABLE IF NOT EXISTS salesforce_conta (
  account_id     text PRIMARY KEY,
  nome           text,
  obra_id        int NOT NULL REFERENCES obra(id),
  atualizado_por text,
  atualizado_em  timestamptz NOT NULL DEFAULT now()
);

-- Eco urbana: Ano, Mês, Mês/Ano, Obra, Ponto_1, db, Ponto2, db3 (duas medições na mesma linha, como na planilha)
CREATE TABLE IF NOT EXISTS eco_urbana (
  id             bigserial PRIMARY KEY,
  obra_id        int  NOT NULL REFERENCES obra(id),
  competencia    date NOT NULL CHECK (extract(day FROM competencia) = 1),
  ponto_1        text,
  db             numeric(6,2),
  ponto2         text,
  db3            numeric(6,2),
  origem         text NOT NULL DEFAULT 'sistema',
  criado_por     text, criado_em timestamptz NOT NULL DEFAULT now(),
  atualizado_por text, atualizado_em timestamptz NOT NULL DEFAULT now()
);

-- Ações SGA: Ano, Mês, Pessoas. mes = 0 é o "Todos" dos anos fechados (2021-2024)
CREATE TABLE IF NOT EXISTS acoes_sga (
  id             bigserial PRIMARY KEY,
  ano            int  NOT NULL,
  mes            int  NOT NULL CHECK (mes BETWEEN 0 AND 12),
  pessoas        int  CHECK (pessoas >= 0),
  origem         text NOT NULL DEFAULT 'sistema',
  criado_por     text, criado_em timestamptz NOT NULL DEFAULT now(),
  atualizado_por text, atualizado_em timestamptz NOT NULL DEFAULT now(),
  UNIQUE (ano, mes)
);

-- Impressões: uma linha por mês. Só o que era digitado; Custos, % P&B, Pag P&B,
-- Pag Colorido, Pág Simples, Economia Total e Economia mínima são calculados.
CREATE TABLE IF NOT EXISTS impressoes (
  id                  bigserial PRIMARY KEY,
  competencia         date NOT NULL UNIQUE CHECK (extract(day FROM competencia) = 1),
  total_pag           int CHECK (total_pag >= 0),
  pct_colorido        numeric(8,6) CHECK (pct_colorido BETWEEN 0 AND 1),
  pag_dupla           numeric(8,6) CHECK (pag_dupla BETWEEN 0 AND 1),
  economia_pb         int,
  economia_colorido   int,
  simples_duplas      numeric(8,6),
  economia_folhas     int,
  origem              text NOT NULL DEFAULT 'sistema',
  criado_por     text, criado_em timestamptz NOT NULL DEFAULT now(),
  atualizado_por text, atualizado_em timestamptz NOT NULL DEFAULT now()
);

-- Impressões por setor: Mês/Ano, Nome, P&B, Cor, Páginas, % (export do NDD Print)
CREATE TABLE IF NOT EXISTS impressoes_setor (
  id             bigserial PRIMARY KEY,
  competencia    date NOT NULL CHECK (extract(day FROM competencia) = 1),
  nome           text NOT NULL,
  pb             int,
  cor            int,
  paginas        int,
  pct            numeric(10,6),
  origem         text NOT NULL DEFAULT 'sistema',
  criado_por     text, criado_em timestamptz NOT NULL DEFAULT now(),
  atualizado_por text, atualizado_em timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS impressoes_setor_comp_idx ON impressoes_setor (competencia);

-- Impressões por pessoa: Mês/Ano, Usuário, Conta atual, P&B, Cor, Páginas, %
CREATE TABLE IF NOT EXISTS impressoes_pessoa (
  id             bigserial PRIMARY KEY,
  competencia    date NOT NULL CHECK (extract(day FROM competencia) = 1),
  usuario        text NOT NULL,
  conta_atual    text,
  pb             int,
  cor            int,
  paginas        int,
  pct            numeric(12,6),
  origem         text NOT NULL DEFAULT 'sistema',
  criado_por     text, criado_em timestamptz NOT NULL DEFAULT now(),
  atualizado_por text, atualizado_em timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS impressoes_pessoa_comp_idx ON impressoes_pessoa (competencia);

INSERT INTO parametro (chave, valor, descricao) VALUES
  ('preco_pagina_pb', 0.05, 'Custo por página P&B (R$) — era a célula T7 da aba Impressões'),
  ('preco_pagina_cor', 0.30, 'Custo por página colorida (R$) — era a célula T8 da aba Impressões')
ON CONFLICT (chave) DO NOTHING;

-- 22/10/2026: Empresa do ruído não existe na planilha; sem valor padrão inventado
ALTER TABLE medicao_ruido ALTER COLUMN empresa DROP NOT NULL;
ALTER TABLE medicao_ruido ALTER COLUMN empresa DROP DEFAULT;

-- Abas de CONSULTA trazidas como estão na planilha (relatórios antigos, sem
-- formato de tabela — ex.: "Histórico de resíduos" 2015-2024). Uma linha por
-- linha da aba, com os valores das células na ordem das colunas.
CREATE TABLE IF NOT EXISTS aba_congelada (
  aba        text NOT NULL,
  linha      int  NOT NULL,
  valores    jsonb NOT NULL,
  importado_em timestamptz NOT NULL DEFAULT now(),
  PRIMARY KEY (aba, linha)
);

-- Sincronização automática com o DW (Salesforce e Prevision): 3x ao dia e sob demanda.
-- sincronizacao = uma linha por fonte por rodada (resumo e pendências da última aparecem em Integrações).
CREATE TABLE IF NOT EXISTS sincronizacao (
  id           bigserial PRIMARY KEY,
  fonte        text NOT NULL,                 -- 'salesforce' | 'prevision'
  origem       text NOT NULL,                 -- 'agendada' | 'manual'
  slot         text,                          -- 'AAAA-MM-DD HH:MM' da rodada agendada (evita repetir)
  iniciado_em  timestamptz NOT NULL DEFAULT now(),
  concluido_em timestamptz,
  status       text NOT NULL DEFAULT 'rodando',
  resumo       jsonb,
  erro         text
);
CREATE INDEX IF NOT EXISTS sincronizacao_fonte_idx ON sincronizacao (fonte, iniciado_em DESC);
-- Projeto do Prevision -> obra. Sem linha = pendência (ProjectId novo em construção); ignorar = não é obra do ambiental.
CREATE TABLE IF NOT EXISTS prevision_projeto (
  project_id     text PRIMARY KEY,
  nome           text,
  obra_id        int REFERENCES obra(id),
  ignorar        boolean NOT NULL DEFAULT false,
  atualizado_por text,
  atualizado_em  timestamptz NOT NULL DEFAULT now()
);

-- Diferenças entre a planilha e os sistemas (Salesforce e Prevision): a Qualidade decide em Integrações se mantém o valor
-- da planilha ou usa o do sistema. A decisão fica guardada; só volta a ser pedida se o valor do sistema mudar depois.
--   prevision: chave 'obra_id|AAAA-MM'  (planilha/sistema = {avanco_pct})
--   salesforce: chave = Id do chamado; registro_id = linha de reclamacao que veio da planilha
CREATE TABLE IF NOT EXISTS integracao_diferenca (
  id           bigserial PRIMARY KEY,
  fonte        text NOT NULL CHECK (fonte IN ('prevision', 'salesforce')),
  chave        text NOT NULL,
  titulo       text,
  planilha     jsonb,
  sistema      jsonb,
  registro_id  bigint,
  situacao     text NOT NULL DEFAULT 'pendente' CHECK (situacao IN ('pendente', 'planilha', 'sistema')),
  detectado_em timestamptz NOT NULL DEFAULT now(),
  decidido_por text,
  decidido_em  timestamptz,
  UNIQUE (fonte, chave)
);
-- O que uma pessoa excluiu de propósito (linha do Prevision ou chamado do Salesforce): a sincronização não traz de volta.
CREATE TABLE IF NOT EXISTS integracao_excluida (
  fonte       text NOT NULL,
  chave       text NOT NULL,
  excluido_por text,
  excluido_em timestamptz NOT NULL DEFAULT now(),
  PRIMARY KEY (fonte, chave)
);
