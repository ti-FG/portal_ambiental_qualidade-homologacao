-- FOREIGN KEYs do modelo semantico Ambiental/Qualidade (postgresql)
-- Gerado por etl/etl_excel_para_db.py em 2026-09-15
--
-- As FKs COMENTADAS so podem ser criadas depois de limpar as linhas
-- orfas indicadas. Ate la o relacionamento existe como documentacao
-- na tabela _relacionamentos_modelo e como indice.
-- Depende do indice unico uq_dim_calendario, uq_dim_obras criado junto com as tabelas.

ALTER TABLE "fat_conformidade_lai_eiv" ADD CONSTRAINT fk_fat_conformidade_lai_eiv_data FOREIGN KEY ("data") REFERENCES "dim_calendario" ("data");

ALTER TABLE "fat_residuos_m3_m2" ADD CONSTRAINT fk_fat_residuos_m3_m2_data FOREIGN KEY ("data") REFERENCES "dim_calendario" ("data");

ALTER TABLE "fat_consumo_agua_luz" ADD CONSTRAINT fk_fat_consumo_agua_luz_data FOREIGN KEY ("data") REFERENCES "dim_calendario" ("data");

ALTER TABLE "fat_eco_urbana" ADD CONSTRAINT fk_fat_eco_urbana_mes_ano FOREIGN KEY ("mes_ano") REFERENCES "dim_calendario" ("data");

ALTER TABLE "fat_impressoes" ADD CONSTRAINT fk_fat_impressoes_data FOREIGN KEY ("data") REFERENCES "dim_calendario" ("data");

ALTER TABLE "fat_impressoes_por_setor" ADD CONSTRAINT fk_fat_impressoes_por_setor_data FOREIGN KEY ("data") REFERENCES "dim_calendario" ("data");

ALTER TABLE "fat_impressoes_por_pessoa" ADD CONSTRAINT fk_fat_impressoes_por_pessoa_data FOREIGN KEY ("data") REFERENCES "dim_calendario" ("data");

ALTER TABLE "fat_soma_residuos" ADD CONSTRAINT fk_fat_soma_residuos_data FOREIGN KEY ("data") REFERENCES "dim_calendario" ("data");

ALTER TABLE "fat_ruido" ADD CONSTRAINT fk_fat_ruido_data FOREIGN KEY ("data") REFERENCES "dim_calendario" ("data");

ALTER TABLE "fat_avanco_fisico" ADD CONSTRAINT fk_fat_avanco_fisico_data FOREIGN KEY ("data") REFERENCES "dim_calendario" ("data");

ALTER TABLE "dim_nc_lai_eiv" ADD CONSTRAINT fk_dim_nc_lai_eiv_data FOREIGN KEY ("data") REFERENCES "dim_calendario" ("data");

-- 3 linha(s) orfa(s) (ex.: North, North Tower) - limpe antes de criar:
-- ALTER TABLE "fat_avanco_fisico" ADD CONSTRAINT fk_fat_avanco_fisico_obra FOREIGN KEY ("obra") REFERENCES "dim_obras" ("obra");

-- 4 linha(s) orfa(s) (ex.: NORTH) - limpe antes de criar:
-- ALTER TABLE "fat_conformidade_lai_eiv" ADD CONSTRAINT fk_fat_conformidade_lai_eiv_obra FOREIGN KEY ("obra") REFERENCES "dim_obras" ("obra");

-- 4279 linha(s) orfa(s) (ex.: Outros) - limpe antes de criar:
-- ALTER TABLE "fat_impressoes_por_pessoa" ADD CONSTRAINT fk_fat_impressoes_por_pessoa_obra FOREIGN KEY ("obra") REFERENCES "dim_obras" ("obra");

-- 787 linha(s) orfa(s) (ex.: Outros) - limpe antes de criar:
-- ALTER TABLE "fat_impressoes_por_setor" ADD CONSTRAINT fk_fat_impressoes_por_setor_obra FOREIGN KEY ("obra") REFERENCES "dim_obras" ("obra");

ALTER TABLE "fat_metas_lai_eiv" ADD CONSTRAINT fk_fat_metas_lai_eiv_obra FOREIGN KEY ("obra") REFERENCES "dim_obras" ("obra");

-- 2 linha(s) orfa(s) (ex.: North Tower) - limpe antes de criar:
-- ALTER TABLE "fat_residuos_m3_m2" ADD CONSTRAINT fk_fat_residuos_m3_m2_obra FOREIGN KEY ("obra") REFERENCES "dim_obras" ("obra");

-- 4 linha(s) orfa(s) (ex.: North) - limpe antes de criar:
-- ALTER TABLE "fat_ruido" ADD CONSTRAINT fk_fat_ruido_obra FOREIGN KEY ("obra") REFERENCES "dim_obras" ("obra");

-- 60 linha(s) orfa(s) (ex.: North Tower) - limpe antes de criar:
-- ALTER TABLE "fat_soma_residuos" ADD CONSTRAINT fk_fat_soma_residuos_obra FOREIGN KEY ("obra") REFERENCES "dim_obras" ("obra");

-- 1 linha(s) orfa(s) (ex.: Gran Place) - limpe antes de criar:
-- ALTER TABLE "fat_eco_urbana" ADD CONSTRAINT fk_fat_eco_urbana_obra FOREIGN KEY ("obra") REFERENCES "dim_obras" ("obra");

ALTER TABLE "dim_nc_lai_eiv" ADD CONSTRAINT fk_dim_nc_lai_eiv_obra FOREIGN KEY ("obra") REFERENCES "dim_obras" ("obra");

-- 397 linha(s) orfa(s) (ex.: Outros) - limpe antes de criar:
-- ALTER TABLE "fat_consumo_agua_luz" ADD CONSTRAINT fk_fat_consumo_agua_luz_obra FOREIGN KEY ("obra") REFERENCES "dim_obras" ("obra");

-- 1 linha(s) orfa(s) (ex.: NORTH TOWER) - limpe antes de criar:
-- ALTER TABLE "dim_auditoria_ambientais" ADD CONSTRAINT fk_dim_auditoria_ambientais_obra FOREIGN KEY ("obra") REFERENCES "dim_obras" ("obra");

-- 3 linha(s) orfa(s) (ex.: Garden) - limpe antes de criar:
-- ALTER TABLE "dim_plano_acao" ADD CONSTRAINT fk_dim_plano_acao_obra FOREIGN KEY ("obra") REFERENCES "dim_obras" ("obra");

-- 8 linha(s) orfa(s) (ex.: GARDEN, ICONIC, IMPERIUM, NORTH) - limpe antes de criar:
-- ALTER TABLE "dim_emergencias_ambientais" ADD CONSTRAINT fk_dim_emergencias_ambientais_obra FOREIGN KEY ("obra") REFERENCES "dim_obras" ("obra");

-- 6 linha(s) orfa(s) (ex.: Blue View Tower, Garden, Iconic Tower, North Tower) - limpe antes de criar:
-- ALTER TABLE "dim_reclamacoes_vizinhos" ADD CONSTRAINT fk_dim_reclamacoes_vizinhos_obra FOREIGN KEY ("obra") REFERENCES "dim_obras" ("obra");
