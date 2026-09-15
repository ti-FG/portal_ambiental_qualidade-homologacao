-- DDL do modelo semantico Ambiental/Qualidade (mssql)
-- Gerado por etl/etl_excel_para_db.py em 2026-09-15

CREATE TABLE _carga_log (
	carregado_em VARCHAR(32) NULL, 
	arquivo_origem VARCHAR(512) NULL, 
	tabela_db VARCHAR(128) NULL, 
	linhas VARCHAR(16) NULL
);

CREATE TABLE _dicionario_modelo (
	tabela_modelo VARCHAR(128) NULL, 
	coluna_modelo VARCHAR(256) NULL, 
	tabela_db VARCHAR(128) NULL, 
	coluna_db VARCHAR(128) NULL, 
	tipo_logico VARCHAR(16) NULL, 
	ordem VARCHAR(8) NULL, 
	fonte_excel VARCHAR(128) NULL, 
	chave VARCHAR(8) NULL, 
	descricao_tabela VARCHAR(512) NULL
);

CREATE TABLE _relacionamentos_modelo (
	tabela_origem VARCHAR(128) NULL, 
	coluna_origem VARCHAR(128) NULL, 
	tabela_destino VARCHAR(128) NULL, 
	coluna_destino VARCHAR(128) NULL, 
	tabela_origem_db VARCHAR(128) NULL, 
	coluna_origem_db VARCHAR(128) NULL, 
	tabela_destino_db VARCHAR(128) NULL, 
	coluna_destino_db VARCHAR(128) NULL, 
	cardinalidade VARCHAR(8) NULL, 
	filtro_cruzado VARCHAR(8) NULL, 
	nome_fk VARCHAR(128) NULL, 
	linhas_orfas VARCHAR(16) NULL, 
	fk_aplicavel VARCHAR(8) NULL, 
	nota VARCHAR(512) NULL
);

CREATE TABLE dim_acoes_sga (
	ano BIGINT NULL, 
	mes NVARCHAR(64) NULL, 
	pessoas BIGINT NULL
);
CREATE INDEX ix_dim_acoes_sga_chave ON dim_acoes_sga (ano, mes);

CREATE TABLE dim_auditoria_ambientais (
	obra NVARCHAR(64) NULL, 
	ano BIGINT NULL, 
	mes NVARCHAR(64) NULL, 
	mes_ano DATETIME NULL, 
	meta FLOAT(53) NULL, 
	realizado NVARCHAR(64) NULL, 
	principais_ocorrencias_da_ultima_auditoria NVARCHAR(704) NULL
);
CREATE INDEX ix_dim_auditoria_ambientais_chave ON dim_auditoria_ambientais (obra, mes_ano);
CREATE INDEX ix_dim_auditoria_ambientais_mes_ano ON dim_auditoria_ambientais (mes_ano);
CREATE INDEX ix_dim_auditoria_ambientais_obra ON dim_auditoria_ambientais (obra);

CREATE TABLE dim_calendario (
	data DATETIME NULL, 
	ano BIGINT NULL, 
	mesnum BIGINT NULL, 
	mes NVARCHAR(64) NULL, 
	mes_ano_key DATETIME NULL, 
	trimestre NVARCHAR(64) NULL, 
	dia BIGINT NULL, 
	dia_da_semana NVARCHAR(64) NULL, 
	chavedata BIGINT NULL, 
	mes_abrev NVARCHAR(64) NULL, 
	mes_ano_abrev NVARCHAR(64) NULL
);
CREATE INDEX ix_dim_calendario_data ON dim_calendario (data);
CREATE UNIQUE INDEX uq_dim_calendario ON dim_calendario (data);

CREATE TABLE dim_emergencias_ambientais (
	obra NVARCHAR(64) NULL, 
	mes_ano DATETIME NULL, 
	fonte NVARCHAR(64) NULL, 
	n_fmc BIGINT NULL, 
	ocorerencia NVARCHAR(128) NULL, 
	status_acoes NVARCHAR(128) NULL, 
	categoria NVARCHAR(64) NULL
);
CREATE INDEX ix_dim_emergencias_ambientais_chave ON dim_emergencias_ambientais (obra, mes_ano, n_fmc);
CREATE INDEX ix_dim_emergencias_ambientais_mes_ano ON dim_emergencias_ambientais (mes_ano);
CREATE INDEX ix_dim_emergencias_ambientais_obra ON dim_emergencias_ambientais (obra);

CREATE TABLE dim_nc_lai_eiv (
	obra NVARCHAR(64) NULL, 
	nc_lai NVARCHAR(128) NULL, 
	nc_eiv NVARCHAR(192) NULL, 
	data DATETIME NULL
);
CREATE INDEX ix_dim_nc_lai_eiv_chave ON dim_nc_lai_eiv (obra, data);
CREATE INDEX ix_dim_nc_lai_eiv_data ON dim_nc_lai_eiv (data);
CREATE INDEX ix_dim_nc_lai_eiv_obra ON dim_nc_lai_eiv (obra);

CREATE TABLE dim_obras (
	obra NVARCHAR(64) NULL
);
CREATE INDEX ix_dim_obras_obra ON dim_obras (obra);
CREATE UNIQUE INDEX uq_dim_obras ON dim_obras (obra);

CREATE TABLE dim_plano_acao (
	obra NVARCHAR(64) NULL, 
	indicador NVARCHAR(64) NULL, 
	analise NVARCHAR(192) NULL, 
	acao NVARCHAR(64) NULL, 
	responsavel NVARCHAR(64) NULL, 
	prazo DATETIME NULL, 
	status NVARCHAR(64) NULL
);
CREATE INDEX ix_dim_plano_acao_chave ON dim_plano_acao (obra, acao);
CREATE INDEX ix_dim_plano_acao_obra ON dim_plano_acao (obra);

CREATE TABLE dim_reclamacoes_vizinhos (
	mes_ano DATETIME NULL, 
	fonte NVARCHAR(64) NULL, 
	cod_chamado BIGINT NULL, 
	ocorrencia NVARCHAR(320) NULL, 
	status_acoes NVARCHAR(1536) NULL, 
	categoria NVARCHAR(64) NULL, 
	obra NVARCHAR(64) NULL, 
	canal NVARCHAR(64) NULL
);
CREATE INDEX ix_dim_reclamacoes_vizinhos_chave ON dim_reclamacoes_vizinhos (mes_ano, cod_chamado);
CREATE INDEX ix_dim_reclamacoes_vizinhos_mes_ano ON dim_reclamacoes_vizinhos (mes_ano);
CREATE INDEX ix_dim_reclamacoes_vizinhos_obra ON dim_reclamacoes_vizinhos (obra);

CREATE TABLE fat_avanco_fisico (
	obra NVARCHAR(64) NULL, 
	ano BIGINT NULL, 
	mes NVARCHAR(64) NULL, 
	data DATETIME NULL, 
	avanco_pct FLOAT(53) NULL, 
	avanco_m2 FLOAT(53) NULL, 
	total_m2 FLOAT(53) NULL
);
CREATE INDEX ix_fat_avanco_fisico_chave ON fat_avanco_fisico (obra, data);
CREATE INDEX ix_fat_avanco_fisico_data ON fat_avanco_fisico (data);
CREATE INDEX ix_fat_avanco_fisico_obra ON fat_avanco_fisico (obra);

CREATE TABLE fat_conformidade_lai_eiv (
	obra NVARCHAR(64) NULL, 
	ano BIGINT NULL, 
	mes NVARCHAR(64) NULL, 
	data DATETIME NULL, 
	percentual FLOAT(53) NULL, 
	meta FLOAT(53) NULL, 
	meta2 FLOAT(53) NULL
);
CREATE INDEX ix_fat_conformidade_lai_eiv_chave ON fat_conformidade_lai_eiv (obra, data);
CREATE INDEX ix_fat_conformidade_lai_eiv_data ON fat_conformidade_lai_eiv (data);
CREATE INDEX ix_fat_conformidade_lai_eiv_obra ON fat_conformidade_lai_eiv (obra);

CREATE TABLE fat_consumo_agua_luz (
	ano BIGINT NULL, 
	mes NVARCHAR(64) NULL, 
	data DATETIME NULL, 
	nome_setor NVARCHAR(64) NULL, 
	luz_kwh FLOAT(53) NULL, 
	agua_m3 FLOAT(53) NULL, 
	obra NVARCHAR(64) NULL, 
	luz_m2 FLOAT(53) NULL, 
	luz_kwh_m3_m2 FLOAT(53) NULL, 
	agua_m2 FLOAT(53) NULL, 
	agua_m3_m2 BIGINT NULL
);
CREATE INDEX ix_fat_consumo_agua_luz_chave ON fat_consumo_agua_luz (nome_setor, data);
CREATE INDEX ix_fat_consumo_agua_luz_data ON fat_consumo_agua_luz (data);
CREATE INDEX ix_fat_consumo_agua_luz_obra ON fat_consumo_agua_luz (obra);

CREATE TABLE fat_eco_urbana (
	ano BIGINT NULL, 
	mes NVARCHAR(64) NULL, 
	mes_ano DATETIME NULL, 
	obra NVARCHAR(64) NULL, 
	db BIGINT NULL, 
	ponto2 NVARCHAR(64) NULL, 
	db3 FLOAT(53) NULL, 
	ponto_1 NVARCHAR(64) NULL
);
CREATE INDEX ix_fat_eco_urbana_chave ON fat_eco_urbana (obra, mes_ano);
CREATE INDEX ix_fat_eco_urbana_mes_ano ON fat_eco_urbana (mes_ano);
CREATE INDEX ix_fat_eco_urbana_obra ON fat_eco_urbana (obra);

CREATE TABLE fat_impressoes (
	ano BIGINT NULL, 
	mes NVARCHAR(64) NULL, 
	data DATETIME NULL, 
	total_de_pag BIGINT NULL, 
	custos NUMERIC(19, 4) NULL, 
	pct_pb FLOAT(53) NULL, 
	pag_pb FLOAT(53) NULL, 
	pct_colorido FLOAT(53) NULL, 
	pag_colorido FLOAT(53) NULL, 
	pag_simples FLOAT(53) NULL, 
	pag_dupla FLOAT(53) NULL, 
	economia_preto_e_branco BIGINT NULL, 
	economia_colorido BIGINT NULL, 
	economia_total BIGINT NULL, 
	impressoes_simples_poderiam_ter_sido_feitas_duplas FLOAT(53) NULL, 
	economia_de_folhas BIGINT NULL, 
	economia_minima NUMERIC(19, 4) NULL
);
CREATE INDEX ix_fat_impressoes_chave ON fat_impressoes (data);
CREATE INDEX ix_fat_impressoes_data ON fat_impressoes (data);

CREATE TABLE fat_impressoes_por_pessoa (
	data DATETIME NULL, 
	usuario NVARCHAR(192) NULL, 
	conta_atual NVARCHAR(128) NULL, 
	pb BIGINT NULL, 
	cor BIGINT NULL, 
	paginas BIGINT NULL, 
	pct FLOAT(53) NULL, 
	obra NVARCHAR(64) NULL
);
CREATE INDEX ix_fat_impressoes_por_pessoa_chave ON fat_impressoes_por_pessoa (data, usuario);
CREATE INDEX ix_fat_impressoes_por_pessoa_data ON fat_impressoes_por_pessoa (data);
CREATE INDEX ix_fat_impressoes_por_pessoa_obra ON fat_impressoes_por_pessoa (obra);

CREATE TABLE fat_impressoes_por_setor (
	data DATETIME NULL, 
	nome_setor NVARCHAR(128) NULL, 
	pb BIGINT NULL, 
	cor BIGINT NULL, 
	paginas BIGINT NULL, 
	pct FLOAT(53) NULL, 
	obra NVARCHAR(64) NULL
);
CREATE INDEX ix_fat_impressoes_por_setor_chave ON fat_impressoes_por_setor (data, nome_setor);
CREATE INDEX ix_fat_impressoes_por_setor_data ON fat_impressoes_por_setor (data);
CREATE INDEX ix_fat_impressoes_por_setor_obra ON fat_impressoes_por_setor (obra);

CREATE TABLE fat_metas_lai_eiv (
	obra NVARCHAR(64) NULL, 
	ago_25 FLOAT(53) NULL, 
	set_25 FLOAT(53) NULL, 
	out_25 FLOAT(53) NULL, 
	nov_25 FLOAT(53) NULL, 
	dez_25 FLOAT(53) NULL, 
	jan_26 FLOAT(53) NULL, 
	fev_26 FLOAT(53) NULL, 
	mar_26 FLOAT(53) NULL, 
	abr_26 FLOAT(53) NULL, 
	mai_26 FLOAT(53) NULL, 
	jun_26 FLOAT(53) NULL, 
	jul_26 FLOAT(53) NULL, 
	ago_26 FLOAT(53) NULL, 
	set_26 FLOAT(53) NULL, 
	out_26 FLOAT(53) NULL, 
	nov_26 FLOAT(53) NULL, 
	dez_26 FLOAT(53) NULL
);
CREATE INDEX ix_fat_metas_lai_eiv_chave ON fat_metas_lai_eiv (obra);
CREATE INDEX ix_fat_metas_lai_eiv_obra ON fat_metas_lai_eiv (obra);

CREATE TABLE fat_residuos_m3_m2 (
	ano BIGINT NULL, 
	mes NVARCHAR(64) NULL, 
	data DATETIME NULL, 
	obra NVARCHAR(64) NULL, 
	residuos_m3 FLOAT(53) NULL, 
	avanco_fisico_m2 FLOAT(53) NULL, 
	residuos_m3_m2 FLOAT(53) NULL
);
CREATE INDEX ix_fat_residuos_m3_m2_chave ON fat_residuos_m3_m2 (obra, data);
CREATE INDEX ix_fat_residuos_m3_m2_data ON fat_residuos_m3_m2 (data);
CREATE INDEX ix_fat_residuos_m3_m2_obra ON fat_residuos_m3_m2 (obra);

CREATE TABLE fat_ruido (
	ano BIGINT NULL, 
	mes NVARCHAR(64) NULL, 
	data DATETIME NULL, 
	obra NVARCHAR(64) NULL, 
	ponto NVARCHAR(64) NULL, 
	horario DATETIME NULL, 
	db FLOAT(53) NULL
);
CREATE INDEX ix_fat_ruido_chave ON fat_ruido (obra, data, ponto);
CREATE INDEX ix_fat_ruido_data ON fat_ruido (data);
CREATE INDEX ix_fat_ruido_obra ON fat_ruido (obra);

CREATE TABLE fat_soma_residuos (
	ano BIGINT NULL, 
	mes NVARCHAR(64) NULL, 
	data DATETIME NULL, 
	obra NVARCHAR(64) NULL, 
	classe NVARCHAR(64) NULL, 
	m3 BIGINT NULL, 
	residuos_detalhados NVARCHAR(128) NULL, 
	kg NVARCHAR(64) NULL, 
	densidade FLOAT(53) NULL, 
	m2 FLOAT(53) NULL, 
	kg_m2 FLOAT(53) NULL, 
	meta_kg FLOAT(53) NULL
);
CREATE INDEX ix_fat_soma_residuos_chave ON fat_soma_residuos (obra, data, classe, residuos_detalhados);
CREATE INDEX ix_fat_soma_residuos_data ON fat_soma_residuos (data);
CREATE INDEX ix_fat_soma_residuos_obra ON fat_soma_residuos (obra);
