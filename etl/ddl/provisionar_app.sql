-- ===========================================================================
--  Provisionamento do app Dash Ambiental/Qualidade na base fg_apps
--  Servidor: 10.1.1.152:5432   Base: fg_apps
--
--  PROPOSTA para quem administra o 152 (Edu) revisar e rodar.
--  NAO roda com o role do app: exige dono da base / superusuario.
--
--  Nomenclatura conforme o "Guia de Nomenclatura de Schemas" (TI DEV,
--  03/08/2026), padrao tipo_assunto_departamento:
--
--    tipo         = bi          -> "ferramentas de indicadores, dashboards e
--                                 relatorios analiticos" (secao 3). E o que
--                                 este sistema e: o dashboard de indicadores
--                                 ambientais. O ETL e uma etapa dele, nao um
--                                 sistema separado.
--    assunto      = ambiental   -> dominio de negocio (controle ambiental).
--    departamento = qualidade   -> DONO do sistema, nao quem apenas usa
--                                 (secao 6). Ver observacoes em meta.schemas.
--
--    => schema: bi_ambiental_qualidade
--    => role  : app_ambiental        (padrao app_<assunto>)
--
--  Rodar assim:
--      psql -h 10.1.1.152 -U <dono_da_base> -d fg_apps -f provisionar_app.sql
--
--  O script PERGUNTA a senha do role (\prompt logo abaixo). Assim ela nao
--  fica gravada neste arquivo nem na linha de comando, onde apareceria na
--  lista de processos e no historico do shell. Atencao: o psql ecoa o que
--  voce digita -- rode de um terminal sem plateia.
--
--  Tudo em UMA transacao: ou o app inteiro fica provisionado e registrado
--  no catalogo, ou nada acontece. O registro em meta.schemas e parte da
--  criacao (secao 9 do guia), nao um passo posterior que alguem esquece.
-- ===========================================================================

\set ON_ERROR_STOP on

\prompt 'Senha do role app_ambiental: ' senha_do_app

BEGIN;

-- 1) Role dedicado do app ---------------------------------------------------
--    Um role por app (app_<assunto>), nunca "postgres" nem usuario do SO.
--
--    A senha entra por um parametro de sessao em vez de ser interpolada no
--    texto do CREATE ROLE: o SET LOCAL morre no COMMIT e o valor nao fica
--    junto do comando em pg_stat_activity. Se o role JA existir, o script
--    NAO mexe na senha dele -- trocar senha e decisao consciente, nao efeito
--    colateral de rodar o provisionamento outra vez.
SET LOCAL provisionar.senha_do_app = :'senha_do_app';

DO $$
BEGIN
    IF NOT EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'app_ambiental') THEN
        EXECUTE format('CREATE ROLE app_ambiental LOGIN PASSWORD %L',
                       current_setting('provisionar.senha_do_app'));
        RAISE NOTICE 'role app_ambiental criado';
    ELSE
        RAISE NOTICE 'role app_ambiental ja existe -- senha preservada';
    END IF;
END
$$;

GRANT CONNECT ON DATABASE fg_apps TO app_ambiental;

-- 2) Schema do app, pertencente ao proprio role -----------------------------
--    Com AUTHORIZATION o role vira DONO do schema. O GRANT logo abaixo e
--    redundante para o dono, mas deixa explicito o que o app pode fazer --
--    CREATE e requisito real aqui, porque o ETL cria as tabelas sozinho
--    (md.create_all em etl/etl_excel_para_db.py).
CREATE SCHEMA IF NOT EXISTS bi_ambiental_qualidade AUTHORIZATION app_ambiental;

GRANT USAGE, CREATE ON SCHEMA bi_ambiental_qualidade TO app_ambiental;

-- 3) Confinar o role ao schema dele -----------------------------------------
--    public nao recebe tabela de negocio de nenhum app (secao 8 do guia).
REVOKE CREATE ON SCHEMA public FROM PUBLIC;
REVOKE ALL    ON SCHEMA public FROM app_ambiental;

-- 4) search_path fixo, so no schema do app ----------------------------------
--    Sem "public" na lista: se algum SQL esquecer de qualificar e apontar
--    para fora do schema, o erro aparece na hora em vez de criar objeto no
--    lugar errado.
ALTER ROLE app_ambiental IN DATABASE fg_apps SET search_path = bi_ambiental_qualidade;

-- 5) Catalogo de metadados (secao 9 do guia) --------------------------------
--    Criado aqui de forma idempotente caso este seja o primeiro app a rodar;
--    se o catalogo ja existir, nada muda.
CREATE SCHEMA IF NOT EXISTS meta;

CREATE TABLE IF NOT EXISTS meta.schemas (
    schema_name      text primary key,
    tipo             text not null,
    assunto          text not null,
    departamento     text not null,
    sistema_real     text,
    responsavel      text,
    observacoes      text,
    criado_em        date default current_date
);

INSERT INTO meta.schemas
    (schema_name, tipo, assunto, departamento, sistema_real, responsavel, observacoes)
VALUES (
    'bi_ambiental_qualidade',
    'bi',
    'ambiental',
    'qualidade',
    'Dash Ambiental/Qualidade (C:\Apps\Dash_ambiental_qualidade; producao 10.1.1.151, homologacao 10.1.1.155)',
    'Natalia Costa (TI DEV) - alteracoes e homologacao; Eduardo - deploy e administracao do 152',
    'Departamento = qualidade por ser o dono/responsavel na acepcao da secao 6 do guia: '
    'a area de Qualidade/Sustentabilidade decide dados e metas (DEPLOY_PRODUCAO.md, secao 8) '
    'e e dona da planilha de origem (\\SERVIDOR\Qualidade\AMBIENTAL\...). TI DEV constroi e '
    'opera, mas nao e a dona do sistema. Role do app: app_ambiental, confinado a este schema. '
    'Carga: ETL full refresh a partir do Excel (etl/etl_excel_para_db.py), 23 tabelas do '
    'modelo semantico, em transacao unica.'
)
ON CONFLICT (schema_name) DO UPDATE SET
    tipo         = EXCLUDED.tipo,
    assunto      = EXCLUDED.assunto,
    departamento = EXCLUDED.departamento,
    sistema_real = EXCLUDED.sistema_real,
    responsavel  = EXCLUDED.responsavel,
    observacoes  = EXCLUDED.observacoes;

GRANT USAGE ON SCHEMA meta TO app_ambiental;
GRANT SELECT ON meta.schemas TO app_ambiental;

COMMIT;

-- ===========================================================================
--  Conferencia depois de rodar, conectando COMO app_ambiental:
--
--      select current_user, current_database(), current_schema();
--        -- esperado: app_ambiental | fg_apps | bi_ambiental_qualidade
--
--      show search_path;
--        -- esperado: bi_ambiental_qualidade
--
--      select has_schema_privilege('bi_ambiental_qualidade','CREATE');  -- t
--      select has_schema_privilege('public','CREATE');                  -- f
--
--  Ou, pronto para rodar: py etl\validar_conexao.py
-- ===========================================================================
