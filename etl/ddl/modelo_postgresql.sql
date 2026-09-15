-- DDL do modelo semantico Ambiental/Qualidade (postgresql)
-- Gerado por etl/etl_excel_para_db.py em 2026-09-15

CREATE TABLE _carga_log (
	carregado_em VARCHAR(32), 
	arquivo_origem VARCHAR(512), 
	tabela_db VARCHAR(128), 
	linhas VARCHAR(16)
);

CREATE TABLE _dicionario_modelo (
	tabela_modelo VARCHAR(128), 
	coluna_modelo VARCHAR(256), 
	tabela_db VARCHAR(128), 
	coluna_db VARCHAR(128), 
	tipo_logico VARCHAR(16), 
	ordem VARCHAR(8), 
	fonte_excel VARCHAR(128), 
	chave VARCHAR(8), 
	descricao_tabela VARCHAR(512)
);

CREATE TABLE _relacionamentos_modelo (
	tabela_origem VARCHAR(128), 
	coluna_origem VARCHAR(128), 
	tabela_destino VARCHAR(128), 
	coluna_destino VARCHAR(128), 
	tabela_origem_db VARCHAR(128), 
	coluna_origem_db VARCHAR(128), 
	tabela_destino_db VARCHAR(128), 
	coluna_destino_db VARCHAR(128), 
	cardinalidade VARCHAR(8), 
	filtro_cruzado VARCHAR(8), 
	nome_fk VARCHAR(128), 
	linhas_orfas VARCHAR(16), 
	fk_aplicavel VARCHAR(8), 
	nota VARCHAR(512)
);

CREATE TABLE dim_acoes_sga (
	ano BIGINT, 
	mes VARCHAR(64), 
	pessoas BIGINT
);
CREATE INDEX ix_dim_acoes_sga_chave ON dim_acoes_sga (ano, mes);

CREATE TABLE dim_auditoria_ambientais (
	obra VARCHAR(64), 
	ano BIGINT, 
	mes VARCHAR(64), 
	mes_ano DATE, 
	meta FLOAT(53), 
	realizado VARCHAR(64), 
	principais_ocorrencias_da_ultima_auditoria VARCHAR(704)
);
CREATE INDEX ix_dim_auditoria_ambientais_chave ON dim_auditoria_ambientais (obra, mes_ano);
CREATE INDEX ix_dim_auditoria_ambientais_mes_ano ON dim_auditoria_ambientais (mes_ano);
CREATE INDEX ix_dim_auditoria_ambientais_obra ON dim_auditoria_ambientais (obra);

CREATE TABLE dim_calendario (
	data DATE, 
	ano BIGINT, 
	mesnum BIGINT, 
	mes VARCHAR(64), 
	mes_ano_key DATE, 
	trimestre VARCHAR(64), 
	dia BIGINT, 
	dia_da_semana VARCHAR(64), 
	chavedata BIGINT, 
	mes_abrev VARCHAR(64), 
	mes_ano_abrev VARCHAR(64)
);
CREATE INDEX ix_dim_calendario_data ON dim_calendario (data);
CREATE UNIQUE INDEX uq_dim_calendario ON dim_calendario (data);

CREATE TABLE dim_emergencias_ambientais (
	obra VARCHAR(64), 
	mes_ano DATE, 
	fonte VARCHAR(64), 
	n_fmc BIGINT, 
	ocorerencia VARCHAR(128), 
	status_acoes VARCHAR(128), 
	categoria VARCHAR(64)
);
CREATE INDEX ix_dim_emergencias_ambientais_chave ON dim_emergencias_ambientais (obra, mes_ano, n_fmc);
CREATE INDEX ix_dim_emergencias_ambientais_mes_ano ON dim_emergencias_ambientais (mes_ano);
CREATE INDEX ix_dim_emergencias_ambientais_obra ON dim_emergencias_ambientais (obra);

CREATE TABLE dim_nc_lai_eiv (
	obra VARCHAR(64), 
	nc_lai VARCHAR(128), 
	nc_eiv VARCHAR(192), 
	data DATE
);
CREATE INDEX ix_dim_nc_lai_eiv_chave ON dim_nc_lai_eiv (obra, data);
CREATE INDEX ix_dim_nc_lai_eiv_data ON dim_nc_lai_eiv (data);
CREATE INDEX ix_dim_nc_lai_eiv_obra ON dim_nc_lai_eiv (obra);

CREATE TABLE dim_obras (
	obra VARCHAR(64)
);
CREATE INDEX ix_dim_obras_obra ON dim_obras (obra);
CREATE UNIQUE INDEX uq_dim_obras ON dim_obras (obra);

CREATE TABLE dim_plano_acao (
	obra VARCHAR(64), 
	indicador VARCHAR(64), 
	analise VARCHAR(192), 
	acao VARCHAR(64), 
	responsavel VARCHAR(64), 
	prazo DATE, 
	status VARCHAR(64)
);
CREATE INDEX ix_dim_plano_acao_chave ON dim_plano_acao (obra, acao);
CREATE INDEX ix_dim_plano_acao_obra ON dim_plano_acao (obra);

CREATE TABLE dim_reclamacoes_vizinhos (
	mes_ano DATE, 
	fonte VARCHAR(64), 
	cod_chamado BIGINT, 
	ocorrencia VARCHAR(320), 
	status_acoes VARCHAR(1536), 
	categoria VARCHAR(64), 
	obra VARCHAR(64), 
	canal VARCHAR(64)
);
CREATE INDEX ix_dim_reclamacoes_vizinhos_chave ON dim_reclamacoes_vizinhos (mes_ano, cod_chamado);
CREATE INDEX ix_dim_reclamacoes_vizinhos_mes_ano ON dim_reclamacoes_vizinhos (mes_ano);
CREATE INDEX ix_dim_reclamacoes_vizinhos_obra ON dim_reclamacoes_vizinhos (obra);

CREATE TABLE fat_avanco_fisico (
	obra VARCHAR(64), 
	ano BIGINT, 
	mes VARCHAR(64), 
	data DATE, 
	avanco_pct FLOAT(53), 
	avanco_m2 FLOAT(53), 
	total_m2 FLOAT(53)
);
CREATE INDEX ix_fat_avanco_fisico_chave ON fat_avanco_fisico (obra, data);
CREATE INDEX ix_fat_avanco_fisico_data ON fat_avanco_fisico (data);
CREATE INDEX ix_fat_avanco_fisico_obra ON fat_avanco_fisico (obra);

CREATE TABLE fat_conformidade_lai_eiv (
	obra VARCHAR(64), 
	ano BIGINT, 
	mes VARCHAR(64), 
	data DATE, 
	percentual FLOAT(53), 
	meta FLOAT(53), 
	meta2 FLOAT(53)
);
CREATE INDEX ix_fat_conformidade_lai_eiv_chave ON fat_conformidade_lai_eiv (obra, data);
CREATE INDEX ix_fat_conformidade_lai_eiv_data ON fat_conformidade_lai_eiv (data);
CREATE INDEX ix_fat_conformidade_lai_eiv_obra ON fat_conformidade_lai_eiv (obra);

CREATE TABLE fat_consumo_agua_luz (
	ano BIGINT, 
	mes VARCHAR(64), 
	data DATE, 
	nome_setor VARCHAR(64), 
	luz_kwh FLOAT(53), 
	agua_m3 FLOAT(53), 
	obra VARCHAR(64), 
	luz_m2 FLOAT(53), 
	luz_kwh_m3_m2 FLOAT(53), 
	agua_m2 FLOAT(53), 
	agua_m3_m2 BIGINT
);
CREATE INDEX ix_fat_consumo_agua_luz_chave ON fat_consumo_agua_luz (nome_setor, data);
CREATE INDEX ix_fat_consumo_agua_luz_data ON fat_consumo_agua_luz (data);
CREATE INDEX ix_fat_consumo_agua_luz_obra ON fat_consumo_agua_luz (obra);

CREATE TABLE fat_eco_urbana (
	ano BIGINT, 
	mes VARCHAR(64), 
	mes_ano DATE, 
	obra VARCHAR(64), 
	db BIGINT, 
	ponto2 VARCHAR(64), 
	db3 FLOAT(53), 
	ponto_1 VARCHAR(64)
);
CREATE INDEX ix_fat_eco_urbana_chave ON fat_eco_urbana (obra, mes_ano);
CREATE INDEX ix_fat_eco_urbana_mes_ano ON fat_eco_urbana (mes_ano);
CREATE INDEX ix_fat_eco_urbana_obra ON fat_eco_urbana (obra);

CREATE TABLE fat_impressoes (
	ano BIGINT, 
	mes VARCHAR(64), 
	data DATE, 
	total_de_pag BIGINT, 
	custos NUMERIC(19, 4), 
	pct_pb FLOAT(53), 
	pag_pb FLOAT(53), 
	pct_colorido FLOAT(53), 
	pag_colorido FLOAT(53), 
	pag_simples FLOAT(53), 
	pag_dupla FLOAT(53), 
	economia_preto_e_branco BIGINT, 
	economia_colorido BIGINT, 
	economia_total BIGINT, 
	impressoes_simples_poderiam_ter_sido_feitas_duplas FLOAT(53), 
	economia_de_folhas BIGINT, 
	economia_minima NUMERIC(19, 4)
);
CREATE INDEX ix_fat_impressoes_chave ON fat_impressoes (data);
CREATE INDEX ix_fat_impressoes_data ON fat_impressoes (data);

CREATE TABLE fat_impressoes_por_pessoa (
	data DATE, 
	usuario VARCHAR(192), 
	conta_atual VARCHAR(128), 
	pb BIGINT, 
	cor BIGINT, 
	paginas BIGINT, 
	pct FLOAT(53), 
	obra VARCHAR(64)
);
CREATE INDEX ix_fat_impressoes_por_pessoa_chave ON fat_impressoes_por_pessoa (data, usuario);
CREATE INDEX ix_fat_impressoes_por_pessoa_data ON fat_impressoes_por_pessoa (data);
CREATE INDEX ix_fat_impressoes_por_pessoa_obra ON fat_impressoes_por_pessoa (obra);

CREATE TABLE fat_impressoes_por_setor (
	data DATE, 
	nome_setor VARCHAR(128), 
	pb BIGINT, 
	cor BIGINT, 
	paginas BIGINT, 
	pct FLOAT(53), 
	obra VARCHAR(64)
);
CREATE INDEX ix_fat_impressoes_por_setor_chave ON fat_impressoes_por_setor (data, nome_setor);
CREATE INDEX ix_fat_impressoes_por_setor_data ON fat_impressoes_por_setor (data);
CREATE INDEX ix_fat_impressoes_por_setor_obra ON fat_impressoes_por_setor (obra);

CREATE TABLE fat_metas_lai_eiv (
	obra VARCHAR(64), 
	ago_25 FLOAT(53), 
	set_25 FLOAT(53), 
	out_25 FLOAT(53), 
	nov_25 FLOAT(53), 
	dez_25 FLOAT(53), 
	jan_26 FLOAT(53), 
	fev_26 FLOAT(53), 
	mar_26 FLOAT(53), 
	abr_26 FLOAT(53), 
	mai_26 FLOAT(53), 
	jun_26 FLOAT(53), 
	jul_26 FLOAT(53), 
	ago_26 FLOAT(53), 
	set_26 FLOAT(53), 
	out_26 FLOAT(53), 
	nov_26 FLOAT(53), 
	dez_26 FLOAT(53)
);
CREATE INDEX ix_fat_metas_lai_eiv_chave ON fat_metas_lai_eiv (obra);
CREATE INDEX ix_fat_metas_lai_eiv_obra ON fat_metas_lai_eiv (obra);

CREATE TABLE fat_residuos_m3_m2 (
	ano BIGINT, 
	mes VARCHAR(64), 
	data DATE, 
	obra VARCHAR(64), 
	residuos_m3 FLOAT(53), 
	avanco_fisico_m2 FLOAT(53), 
	residuos_m3_m2 FLOAT(53)
);
CREATE INDEX ix_fat_residuos_m3_m2_chave ON fat_residuos_m3_m2 (obra, data);
CREATE INDEX ix_fat_residuos_m3_m2_data ON fat_residuos_m3_m2 (data);
CREATE INDEX ix_fat_residuos_m3_m2_obra ON fat_residuos_m3_m2 (obra);

CREATE TABLE fat_ruido (
	ano BIGINT, 
	mes VARCHAR(64), 
	data DATE, 
	obra VARCHAR(64), 
	ponto VARCHAR(64), 
	horario TIME WITHOUT TIME ZONE, 
	db FLOAT(53)
);
CREATE INDEX ix_fat_ruido_chave ON fat_ruido (obra, data, ponto);
CREATE INDEX ix_fat_ruido_data ON fat_ruido (data);
CREATE INDEX ix_fat_ruido_obra ON fat_ruido (obra);

CREATE TABLE fat_soma_residuos (
	ano BIGINT, 
	mes VARCHAR(64), 
	data DATE, 
	obra VARCHAR(64), 
	classe VARCHAR(64), 
	m3 BIGINT, 
	residuos_detalhados VARCHAR(128), 
	kg VARCHAR(64), 
	densidade FLOAT(53), 
	m2 FLOAT(53), 
	kg_m2 FLOAT(53), 
	meta_kg FLOAT(53)
);
CREATE INDEX ix_fat_soma_residuos_chave ON fat_soma_residuos (obra, data, classe, residuos_detalhados);
CREATE INDEX ix_fat_soma_residuos_data ON fat_soma_residuos (data);
CREATE INDEX ix_fat_soma_residuos_obra ON fat_soma_residuos (obra);
