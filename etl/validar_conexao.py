#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Confere a conexao com o PostgreSQL consolidado ANTES do cutover.

Roda sem gravar nada: so conecta com as credenciais do .env e verifica se o
provisionamento (etl/ddl/provisionar_app.sql) foi aplicado como o Guia de
Nomenclatura de Schemas manda.

    py -3.12 etl\\validar_conexao.py

Codigos de saida:
    0 = tudo conforme, pode fazer o cutover
    1 = conectou, mas alguma verificacao falhou
    2 = nao conseguiu conectar (senha, rede, role inexistente)
    3 = configuracao ausente no .env

O que ele verifica, um a um:
    - conecta em 10.1.1.152/fg_apps com o role do app
    - o usuario conectado e o role dedicado, nao "postgres" nem outro
    - o search_path e exatamente o schema do app
    - o role PODE criar no schema dele (o ETL cria as tabelas sozinho)
    - o role NAO pode criar em public (secao 8 do guia)
    - o schema esta registrado em meta.schemas (secao 9 do guia)
"""

import os
import sys

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CAMINHO_ENV = os.path.join(RAIZ, '.env')

ESPERADO_USUARIO = 'app_ambiental'
ESPERADO_SCHEMA = 'bi_ambiental_qualidade'
ESPERADO_BASE = 'fg_apps'


def _cor(texto, ok):
    marca = 'OK  ' if ok else 'FALHA'
    return '  [%s] %s' % (marca, texto)


def carregar_env():
    """Le o .env para o ambiente. Sem ele nao ha o que validar."""
    try:
        from dotenv import load_dotenv
    except ImportError:
        print('ERRO: pacote python-dotenv ausente. Instale com:')
        print('      py -3.12 -m pip install -r requirements.txt')
        return None

    if not os.path.exists(CAMINHO_ENV):
        print('ERRO: %s nao encontrado.' % CAMINHO_ENV)
        return None

    load_dotenv(CAMINHO_ENV, override=False)

    url = os.environ.get('DATABASE_URL')
    schema = os.environ.get('DASHFG_DB_SCHEMA')
    if not url:
        print('ERRO: DATABASE_URL ausente no .env')
        return None
    if not schema:
        print('ERRO: DASHFG_DB_SCHEMA ausente no .env')
        return None
    if not os.environ.get('PGPASSWORD') and '@' in url.split('//', 1)[-1].split('/')[0]:
        # senha pode vir embutida na URL; so avisa se nao houver nenhuma fonte
        if ':' not in url.split('//', 1)[-1].split('@')[0]:
            print('AVISO: sem PGPASSWORD no .env e sem senha na DATABASE_URL.')
    return url, schema


def main():
    cfg = carregar_env()
    if cfg is None:
        return 3
    url, schema = cfg

    print('Validando conexao do app Dash Ambiental/Qualidade')
    print('  .env   : %s' % CAMINHO_ENV)
    print('  destino: %s' % url)
    print('  schema : %s' % schema)
    print('')

    try:
        from sqlalchemy import create_engine, text
    except ImportError:
        print('ERRO: SQLAlchemy ausente. py -3.12 -m pip install -r requirements.txt')
        return 3

    try:
        eng = create_engine(url, future=True,
                            connect_args={'connect_timeout': 10})
        con = eng.connect()
    except Exception as e:                                      # noqa: BLE001
        print('NAO CONECTOU: %s: %s' % (type(e).__name__, e))
        print('')
        print('Causas comuns, em ordem de probabilidade:')
        print('  1. senha do role errada (PGPASSWORD no .env)')
        print('  2. role %s ainda nao existe -- rodar etl/ddl/provisionar_app.sql'
              % ESPERADO_USUARIO)
        print('  3. pg_hba.conf do 10.1.1.152 nao libera este host')
        return 2

    falhas = []
    with con:
        def uma(rotulo, sql, esperado=None, condicao=None):
            try:
                valor = con.execute(text(sql)).scalar()
            except Exception as e:                              # noqa: BLE001
                print(_cor('%s -> erro: %s' % (rotulo, e), False))
                falhas.append(rotulo)
                return None
            ok = condicao(valor) if condicao else (valor == esperado)
            detalhe = '%s -> %s' % (rotulo, valor)
            if not ok and esperado is not None:
                detalhe += '   (esperado: %s)' % esperado
            print(_cor(detalhe, ok))
            if not ok:
                falhas.append(rotulo)
            return valor

        uma('usuario conectado', 'select current_user', ESPERADO_USUARIO)
        uma('base', 'select current_database()', ESPERADO_BASE)
        uma('search_path', 'show search_path', schema)
        uma('schema do app existe',
            "select exists(select 1 from information_schema.schemata "
            "where schema_name = '%s')" % schema, True)
        uma('CREATE no schema do app',
            "select has_schema_privilege('%s','CREATE')" % schema, True)
        uma('CREATE em public bloqueado',
            "select has_schema_privilege('public','CREATE')", False)
        uma('registrado em meta.schemas',
            "select exists(select 1 from meta.schemas "
            "where schema_name = '%s')" % schema, True)

        tabelas = con.execute(text(
            "select count(*) from information_schema.tables "
            "where table_schema = '%s'" % schema)).scalar()
        print('  [info ] tabelas ja existentes no schema: %s' % tabelas)

    print('')
    if falhas:
        print('RESULTADO: %d verificacao(oes) falharam -- NAO faca o cutover:'
              % len(falhas))
        for f in falhas:
            print('  - %s' % f)
        print('')
        print('Ajustar rodando etl/ddl/provisionar_app.sql no 10.1.1.152.')
        return 1

    print('RESULTADO: conexao conforme o guia. Pode carregar o banco:')
    print('  py -3.12 etl\\etl_excel_para_db.py --se-mudou --log etl\\logs\\etl_banco.log')
    return 0


if __name__ == '__main__':
    sys.exit(main())
