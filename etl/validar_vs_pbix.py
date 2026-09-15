# -*- coding: utf-8 -*-
"""
Prova de que o ETL produz o MESMO modelo semantico do .pbix.

Para cada uma das 19 tabelas compara, entre o resultado do ETL (Excel de hoje)
e a tabela lida do .pbix:

  1. os nomes das colunas e a ordem delas;
  2. as linhas, como multiconjunto (o .pbix e um column store: a ordem das
     linhas que o PBIXRay devolve nao e a ordem original, entao comparar
     posicionalmente daria diferenca falsa);
  3. quando a tabela tem chave, celula a celula nas chaves que existem dos
     dois lados - e isso que pega erro de transformacao.

A planilha esta MAIS NOVA que o .pbix (o .pbix foi atualizado pela ultima vez
com dados ate abr/2026). Entao linhas "so no Excel" sao esperadas: sao os
lancamentos novos. O que NAO pode aparecer e divergencia de valor numa chave
que existe dos dois lados.

Uso:
    py -3.12 validar_vs_pbix.py
    py -3.12 validar_vs_pbix.py --pbix ../AMBIENTAL_QUALIDADE_202602_v1_DEV.pbix
    py -3.12 validar_vs_pbix.py --detalhe Fat_soma_residuos
"""
from __future__ import annotations

import argparse
import datetime as _dt
import os
import sys

# Sob NSSM/servico o stdout do Windows fica em cp1252 e qualquer m3/m2
# impresso derruba o processo com UnicodeEncodeError.
try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass
import tempfile
from decimal import Decimal

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402

from etl_excel_para_db import EXCEL_CANDIDATOS, extrair  # noqa: E402
from tabelas import TABELAS, TABELAS_POR_NOME  # noqa: E402

DIR_SCRIPT = os.path.dirname(os.path.abspath(__file__))
DIR_PROJETO = os.path.dirname(DIR_SCRIPT)
PBIX_PADRAO = os.path.join(DIR_PROJETO, 'AMBIENTAL_QUALIDADE_202602_v1_DEV.pbix')
CACHE_PADRAO = os.path.join(tempfile.gettempdir(), 'pbix_ref_ambiental')


# ------------------------------------------------------------- leitura pbix --
def carregar_pbix(caminho, cache):
    """Le as tabelas do .pbix, com cache em disco (a leitura leva minutos)."""
    os.makedirs(cache, exist_ok=True)
    carimbo = os.path.join(cache, '_origem.txt')
    assinatura = '%s|%s' % (os.path.abspath(caminho),
                            os.path.getmtime(caminho))
    valido = (os.path.exists(carimbo)
              and open(carimbo, encoding='utf-8').read() == assinatura)

    nomes = [t.nome for t in TABELAS]
    if valido and all(os.path.exists(os.path.join(cache, _arq(n))) for n in nomes):
        print('Usando cache do .pbix em %s' % cache)
        return {n: pd.read_pickle(os.path.join(cache, _arq(n))) for n in nomes}

    print('Lendo .pbix (demora alguns minutos): %s' % caminho)
    from pbixray import PBIXRay
    m = PBIXRay(caminho)
    saida = {}
    for n in nomes:
        try:
            df = m.get_table(n)
        except Exception as e:                       # noqa: BLE001
            print('  AVISO: tabela %s nao lida do pbix (%s)' % (n, e))
            df = pd.DataFrame()
        saida[n] = df
        df.to_pickle(os.path.join(cache, _arq(n)))
    with open(carimbo, 'w', encoding='utf-8') as fh:
        fh.write(assinatura)
    return saida


def _arq(nome):
    return nome.replace('³', '3').replace('²', '2').replace('/', '_') + '.pkl'


# ----------------------------------------------------------- canonicalizacao --
def _canon(v):
    """Converte um valor para uma forma comparavel entre pandas/pbixray."""
    if v is None or v is pd.NA or v is pd.NaT:
        return None
    if isinstance(v, float) and np.isnan(v):
        return None
    if isinstance(v, (bool, np.bool_)):
        return bool(v)
    if isinstance(v, Decimal):
        v = float(v)
    if isinstance(v, (np.integer,)):
        v = int(v)
    if isinstance(v, (np.floating,)):
        v = float(v)
    if isinstance(v, int):
        return float(v)
    if isinstance(v, float):
        if np.isnan(v):
            return None
        return float('%.9g' % v)
    if isinstance(v, pd.Timestamp):
        v = v.to_pydatetime()
    if isinstance(v, _dt.datetime):
        # Horario do pbix vem como 1899-12-31 + hora guardada em float: 09:03
        # vira 09:02:59,999999. Por isso ARREDONDA pro segundo, nao trunca -
        # truncar produziria 119 diferencas falsas de 1 segundo em Fat_Ruido.
        if v.year <= 1900:
            return (v + _dt.timedelta(microseconds=500000)).time()                     .replace(microsecond=0).isoformat()
        return v.date().isoformat() if v.time() == _dt.time(0, 0) else v.isoformat()
    if isinstance(v, _dt.date):
        return v.isoformat()
    if isinstance(v, _dt.time):
        base = _dt.datetime(2000, 1, 1, v.hour, v.minute, v.second, v.microsecond)
        return (base + _dt.timedelta(microseconds=500000)).time()                 .replace(microsecond=0).isoformat()
    s = str(v)
    return None if s.strip() == '' else s


def _normalizar(df, colunas):
    """DataFrame -> DataFrame de valores canonicos, sem linhas totalmente nulas."""
    if df is None or df.empty:
        return pd.DataFrame(columns=colunas)
    presentes = [c for c in colunas if c in df.columns]
    d = df[presentes].copy()
    for c in presentes:
        # dtype object: atribuir numa coluna Float64 converteria None em NaN
        # de volta, e ai vazio de um lado nao casaria com vazio do outro.
        d[c] = pd.Series([_canon(v) for v in d[c]], index=d.index, dtype=object)
    d = d[~d.apply(lambda r: all(v is None for v in r), axis=1)]
    return d.reset_index(drop=True)


def _chaveavel(v):
    if v is None or v is pd.NA or v is pd.NaT:
        return ''
    if isinstance(v, float) and np.isnan(v):
        return ''
    return str(v)


# --------------------------------------------------------------- comparacao --
def comparar(nome, df_etl, df_pbix, detalhe=False):
    t = TABELAS_POR_NOME[nome]
    colunas = [c for c, _ in t.colunas]
    res = {'tabela': nome, 'colunas_ok': True, 'erros': [], 'notas': []}

    cols_pbix = list(df_pbix.columns) if df_pbix is not None else []
    if cols_pbix != colunas:
        faltam = [c for c in colunas if c not in cols_pbix]
        sobram = [c for c in cols_pbix if c not in colunas]
        if faltam or sobram:
            res['colunas_ok'] = False
            res['erros'].append('colunas diferentes do pbix — faltando no ETL: %s'
                                ' | so no pbix: %s' % (faltam or '-', sobram or '-'))
        else:
            res['notas'].append('mesmas colunas, ordem diferente da do pbix')

    a = _normalizar(df_etl, colunas)
    b = _normalizar(df_pbix, colunas)
    res['linhas_etl'] = len(a)
    res['linhas_pbix'] = len(b)

    ta = [tuple(_chaveavel(v) for v in r) for r in a.itertuples(index=False, name=None)]
    tb = [tuple(_chaveavel(v) for v in r) for r in b.itertuples(index=False, name=None)]
    ca, cb = pd.Series(ta).value_counts(), pd.Series(tb).value_counts()
    idx = ca.index.union(cb.index)
    ca = ca.reindex(idx, fill_value=0)
    cb = cb.reindex(idx, fill_value=0)
    iguais = int(np.minimum(ca.values, cb.values).sum())
    res['linhas_iguais'] = iguais
    res['so_etl'] = int((ca - cb).clip(lower=0).sum())
    res['so_pbix'] = int((cb - ca).clip(lower=0).sum())

    chave = [c for c in t.chave if c in colunas]
    res['divergencias'] = 0
    res['div_vazio'] = res['div_ambos'] = 0
    res['exemplos'] = []
    if chave and len(a) and len(b):
        ka = a.assign(_k=[tuple(_chaveavel(v) for v in r)
                          for r in a[chave].itertuples(index=False, name=None)])
        kb = b.assign(_k=[tuple(_chaveavel(v) for v in r)
                          for r in b[chave].itertuples(index=False, name=None)])
        # so chaves que aparecem uma unica vez dos dois lados sao comparaveis
        ua = ka['_k'].value_counts()
        ub = kb['_k'].value_counts()
        comuns = [k for k in ua.index
                  if ua[k] == 1 and k in ub.index and ub[k] == 1]
        res['chaves_comparadas'] = len(comuns)
        if comuns:
            # filtra ANTES de indexar: reindex nao aceita eixo com rotulo
            # repetido, e as chaves nao-comuns podem repetir.
            alvo = set(comuns)
            ia = ka[ka['_k'].isin(alvo)].set_index('_k').reindex(comuns)
            ib = kb[kb['_k'].isin(alvo)].set_index('_k').reindex(comuns)
            for c in colunas:
                oa, ob = ia[c].tolist(), ib[c].tolist()
                va = [_chaveavel(v) for v in oa]
                vb = [_chaveavel(v) for v in ob]
                difs = [i for i, (x, y) in enumerate(zip(va, vb)) if x != y]
                if not difs:
                    continue
                # Separa defasagem de mudanca de valor: se um dos lados esta
                # VAZIO, a celula foi preenchida (ou esvaziada) na planilha
                # depois da ultima atualizacao do .pbix - nao e erro de ETL.
                ambos = [i for i in difs if va[i] != '' and vb[i] != '']
                res['divergencias'] += len(difs)
                res['div_vazio'] += len(difs) - len(ambos)
                res['div_ambos'] += len(ambos)
                if ambos:
                    i0 = ambos[0]
                    res['exemplos'].append(
                        '%s: %d celula(s) com valor diferente (de %d chaves '
                        'comparadas); ex. chave=%s  planilha=%r  pbix=%r'
                        % (c, len(ambos), len(comuns), comuns[i0],
                           oa[i0], ob[i0]))
    else:
        res['chaves_comparadas'] = 0

    if detalhe and res['so_pbix']:
        faltantes = (cb - ca).clip(lower=0)
        faltantes = faltantes[faltantes > 0]
        res['notas'].append('exemplos de linhas so no pbix: %s'
                            % list(faltantes.index[:3]))
    return res


def main(argv=None):
    p = argparse.ArgumentParser()
    p.add_argument('--pbix', default=PBIX_PADRAO)
    p.add_argument('--excel')
    p.add_argument('--cache', default=CACHE_PADRAO)
    p.add_argument('--detalhe', help='mostra exemplos de linhas so no pbix')
    a = p.parse_args(argv)

    excel = a.excel or next((c for c in EXCEL_CANDIDATOS if os.path.exists(c)), None)
    avisos = []
    etl = extrair(excel, avisos)
    pbix = carregar_pbix(a.pbix, a.cache)

    print('\n%-30s %6s %6s %7s %7s %7s %7s %6s' % (
        'TABELA', 'ETL', 'PBIX', 'IGUAIS', 'SO_ETL', 'SO_PBIX', 'VAZIO', 'MUDOU'))
    print('-' * 94)
    com_erro = []
    resultados = []
    for t in TABELAS:
        r = comparar(t.nome, etl[t.nome], pbix.get(t.nome),
                     detalhe=(a.detalhe == t.nome))
        resultados.append(r)
        if r['erros'] or r['div_ambos']:
            com_erro.append(r)
        print('%-30s %6d %6d %7d %7d %7d %7d %6d' % (
            t.nome, r['linhas_etl'], r['linhas_pbix'], r['linhas_iguais'],
            r['so_etl'], r['so_pbix'], r['div_vazio'], r['div_ambos']))

    print('\nLegenda (VAZIO e MUDOU olham so as chaves que existem dos 2 lados):')
    print('  SO_ETL   linhas novas na planilha - esperado, o .pbix e mais antigo')
    print('  SO_PBIX  linhas que sairam da planilha desde a ultima atualizacao')
    print('  VAZIO    celula vazia de um lado, preenchida do outro - tambem e '
          'defasagem')
    print('  MUDOU    celula com valor diferente nos DOIS lados - so isso pode '
          'ser erro de')
    print('           transformacao; cada caso vem detalhado abaixo para '
          'conferencia.')

    for r in resultados:
        for n in r['notas']:
            print('  nota  %s: %s' % (r['tabela'], n))
    if com_erro:
        print('\n--- A CONFERIR (valor diferente nos dois lados) ---')
        for r in com_erro:
            print('  %s:' % r['tabela'])
            for e in r['erros']:
                print('    ! %s' % e)
            for e in r['exemplos'][:6]:
                print('    - %s' % e)
    print('\nCelulas com valor diferente nos dois lados: %d'
          % sum(r['div_ambos'] for r in resultados))
    print('Celulas vazias de um dos lados (defasagem): %d'
          % sum(r['div_vazio'] for r in resultados))
    # So erro de ESTRUTURA reprova. Diferenca de valor pede leitura humana: a
    # planilha e mais nova que o .pbix, entao diferenca ali costuma ser edicao.
    return 1 if any(r['erros'] for r in resultados) else 0


if __name__ == '__main__':
    sys.exit(main())
