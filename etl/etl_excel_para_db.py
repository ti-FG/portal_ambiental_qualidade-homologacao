# -*- coding: utf-8 -*-
"""
ETL: Controle Ambiental (Excel) -> modelo semantico -> banco de dados.

Le as tabelas nomeadas (ListObjects) da planilha
`Controle Ambiental 2025_oficial_db.xlsx`, aplica as mesmas transformacoes que
o Power Query do .pbix aplica, e grava as 19 tabelas do modelo semantico em um
banco relacional (qualquer um suportado pelo SQLAlchemy).

Uso tipico
----------
    # carga local em SQLite (padrao), so pra conferir
    py -3.12 etl_excel_para_db.py

    # carga em SQL Server
    py -3.12 etl_excel_para_db.py --db-url "mssql+pyodbc://usr:senha@servidor/BaseAmbiental?driver=ODBC+Driver+17+for+SQL+Server"

    # carga em PostgreSQL, em um schema proprio
    py -3.12 etl_excel_para_db.py --db-url "postgresql+psycopg2://usr:senha@servidor/base" --schema ambiental

    # so gerar o DDL, sem tocar em banco nenhum
    py -3.12 etl_excel_para_db.py --somente-ddl --dialeto mssql

A carga e SEMPRE full refresh: apaga e regrava todas as linhas dentro de uma
unica transacao. Se qualquer tabela falhar, nada e commitado.
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

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd  # noqa: E402
from openpyxl import load_workbook  # noqa: E402
from openpyxl.utils.cell import range_boundaries  # noqa: E402

from modelo_semantico import (  # noqa: E402
    DATE, DATETIME, DECIMAL, FLOAT, INT, PERCENT, TEXT, TIME,
    aplicar_tipos, nome_fisico,
)
import origem as org  # noqa: E402
from tabelas import (  # noqa: E402
    ALERTAS_MODELO, CHAVES_DIMENSAO, RELACIONAMENTOS, TABELAS,
)

DIR_SCRIPT = os.path.dirname(os.path.abspath(__file__))
DIR_PROJETO = os.path.dirname(DIR_SCRIPT)

EXCEL_CANDIDATOS = [
    os.path.join(DIR_PROJETO, 'Controle Ambiental 2025_oficial_db.xlsx'),
    r'U:\Qualidade\AMBIENTAL\GESTÃO AMBIENTAL\INDICADORES AMBIENTAIS\POWER BI'
    r'\Controle Ambiental 2025_oficial_db.xlsx',
]
TABELA_DICIONARIO = '_dicionario_modelo'
TABELA_LOG = '_carga_log'
TABELA_RELACIONAMENTOS = '_relacionamentos_modelo'
TABELA_EXECUCAO = '_carga_execucao'


# ---------------------------------------------------------------- leitura ---
def _achar_tabela_nomeada(wb, nome):
    for ws in wb.worksheets:
        if nome in ws.tables:
            return ws, ws.tables[nome]
    return None, None


def ler_tabela_nomeada(wb, nome, avisos):
    """Le um ListObject do Excel como DataFrame de valores crus.

    Usa o intervalo declarado da tabela nomeada - e exatamente isso que o
    Power Query le. Linhas totalmente vazias sao descartadas (o pbix as
    importa como linhas nulas; ver README).
    """
    ws, t = _achar_tabela_nomeada(wb, nome)
    if t is None:
        raise KeyError('Tabela nomeada "%s" nao existe na planilha. '
                       'Alguem renomeou ou apagou o ListObject.' % nome)

    min_col, min_row, max_col, max_row = range_boundaries(t.ref)
    cabecalho = min_row + (t.headerRowCount if t.headerRowCount is not None else 1)
    colunas = [c.name for c in t.tableColumns]

    linhas = []
    for row in ws.iter_rows(min_row=cabecalho, max_row=max_row,
                            min_col=min_col, max_col=max_col, values_only=True):
        linhas.append(list(row))

    brutas = len(linhas)
    linhas = [r for r in linhas
              if not all(v is None or (isinstance(v, str) and not v.strip())
                         for v in r)]
    if brutas != len(linhas):
        avisos.append('%s: %d linha(s) em branco descartadas (o pbix as importa '
                      'como linhas nulas).' % (nome, brutas - len(linhas)))

    _conferir_dados_fora_do_intervalo(ws, nome, min_col, max_col, max_row, avisos)
    return pd.DataFrame(linhas, columns=colunas)


def _conferir_dados_fora_do_intervalo(ws, nome, min_col, max_col, max_row, avisos):
    """Avisa se ha dados abaixo do intervalo declarado da tabela nomeada.

    Se a tabela do Excel parar de expandir sozinha, linhas novas ficam de fora
    do ListObject - e somem do modelo sem ninguem perceber.
    """
    extras = 0
    for row in ws.iter_rows(min_row=max_row + 1, max_row=min(ws.max_row, max_row + 500),
                            min_col=min_col, max_col=max_col, values_only=True):
        if any(v is not None and str(v).strip() for v in row):
            extras += 1
    if extras:
        avisos.append('ATENCAO %s: ha %d linha(s) com dados ABAIXO do intervalo '
                      '%s da tabela nomeada. Essas linhas nao entram no modelo. '
                      'Corrija o intervalo da tabela no Excel.'
                      % (nome, extras, ws.title))


def extrair(caminho_excel, avisos):
    """Le a planilha e devolve {nome_da_tabela_do_modelo: DataFrame tipado}."""
    print('Lendo planilha: %s' % caminho_excel)
    wb = load_workbook(caminho_excel, data_only=True)
    try:
        resultado = {}
        for t in TABELAS:
            if t.fonte is None:
                bruto = pd.DataFrame()
            else:
                bruto = ler_tabela_nomeada(wb, t.fonte, avisos)
            df = t.construir(bruto)
            # aplicar_tipos preenche com NULL toda coluna do modelo que nao
            # tiver sido produzida - isso seria uma falha silenciosa.
            ausentes = [c for c, _ in t.colunas if c not in df.columns]
            if ausentes:
                avisos.append('ATENCAO %s: coluna(s) do modelo ausentes na origem '
                              'e preenchidas com NULL: %s. Provavel renomeacao de '
                              'coluna no Excel.' % (t.nome, ', '.join(ausentes)))
            df = aplicar_tipos(df, t.colunas)
            resultado[t.nome] = df.reset_index(drop=True)
            print('  %-30s %6d linhas  (fonte: %s)'
                  % (t.nome, len(df), t.fonte or 'gerada'))
        return resultado
    finally:
        wb.close()


# ------------------------------------------------------------ verificacoes ---
def verificar(dados, avisos):
    """Checagens de qualidade que NAO alteram o dado, so reportam."""
    problemas = []

    fisicos = {}
    for t in TABELAS:
        vistos = {}
        for c, _ in t.colunas:
            f = nome_fisico(c)
            if f in vistos:
                problemas.append('%s: as colunas "%s" e "%s" viram o mesmo nome '
                                 'fisico "%s".' % (t.nome, vistos[f], c, f))
            vistos[f] = c
        if t.tabela in fisicos:
            problemas.append('Nome fisico "%s" repetido entre %s e %s.'
                             % (t.tabela, fisicos[t.tabela], t.nome))
        fisicos[t.tabela] = t.nome

    for r in integridade(dados):
        if r['orfas']:
            avisos.append('%s.%s: %d linha(s) com valor que nao existe em %s.%s '
                          '-> %s' % (r['origem'], r['coluna_origem'], r['orfas'],
                                     r['destino'], r['coluna_destino'],
                                     ', '.join(r['exemplos'][:8])
                                     + (' ...' if len(r['exemplos']) > 8 else '')))

    for t in TABELAS:
        if not t.chave:
            continue
        df = dados[t.nome]
        cols = [c for c in t.chave if c in df.columns]
        if not cols:
            continue
        dup = df.duplicated(subset=cols, keep=False).sum()
        if dup:
            avisos.append('%s: %d linha(s) repetem a chave %s (a fonte Excel nao '
                          'garante unicidade; a chave vira indice, nao PK).'
                          % (t.nome, dup, '+'.join(cols)))
    return problemas


# ----------------------------------------------------------------- destino ---
def _tipo_sql(tipo, serie):
    from sqlalchemy import BigInteger, Date, DateTime, Float, Numeric, Time
    from sqlalchemy import Unicode, UnicodeText
    if tipo == INT:
        return BigInteger()
    if tipo in (FLOAT, PERCENT):
        return Float(53)
    if tipo == DECIMAL:
        return Numeric(19, 4)
    if tipo == DATE:
        return Date()
    if tipo == TIME:
        return Time()
    if tipo == DATETIME:
        return DateTime()
    maior = 0
    for v in serie.dropna():
        maior = max(maior, len(str(v)))
    if maior > 2000:
        return UnicodeText()
    tamanho = max(64, ((maior * 2) // 64 + 1) * 64)
    return Unicode(min(tamanho, 4000))


def integridade(dados):
    """Confere cada relacionamento do modelo e conta as linhas orfas.

    Orfa = valor na tabela do lado "muitos" que nao existe na dimensao. Enquanto
    houver orfa, a FOREIGN KEY correspondente nao pode ser criada - por isso as
    FKs saem num script a parte (ver `gerar_ddl_fk`).
    """
    saida = []
    for r in RELACIONAMENTOS:
        df = dados.get(r.origem)
        dim = dados.get(r.destino)
        if df is None or dim is None or r.coluna_origem not in df.columns:
            continue
        validos = {_canonico(v) for v in dim[r.coluna_destino].dropna()}
        orfas, exemplos = 0, set()
        for v in df[r.coluna_origem]:
            c = _canonico(v)
            if c is None or c in validos:
                continue
            orfas += 1
            if len(exemplos) < 12:
                exemplos.add(c)
        saida.append({
            'origem': r.origem, 'coluna_origem': r.coluna_origem,
            'destino': r.destino, 'coluna_destino': r.coluna_destino,
            'cardinalidade': r.cardinalidade, 'filtro': r.filtro,
            'linhas': len(df), 'orfas': orfas,
            'exemplos': sorted(exemplos), 'nome_fk': r.nome_fk, 'nota': r.nota,
        })
    return saida


def _canonico(v):
    """Chave comparavel entre fato e dimensao (nulo -> None)."""
    if v is None or v is pd.NA or v is pd.NaT:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    s = str(v).strip()
    return s or None


def montar_metadata(dados, schema=None):
    from sqlalchemy import Column, Index, MetaData, String, Table

    md = MetaData(schema=schema)
    for t in TABELAS:
        df = dados[t.nome]
        cols = [Column(nome_fisico(c), _tipo_sql(tp, df[c]), nullable=True)
                for c, tp in t.colunas]
        tab = Table(t.tabela, md, *cols, comment=t.descricao or None)
        # Lado "1" dos relacionamentos: indice UNICO, que e o que uma FOREIGN
        # KEY exige do outro lado. So cabe nas dimensoes (Dim_Obras,
        # Dim_Calendario); os fatos repetem chave de proposito.
        if t.nome in CHAVES_DIMENSAO:
            Index('uq_%s' % t.tabela,
                  *[tab.c[nome_fisico(c)] for c in CHAVES_DIMENSAO[t.nome]],
                  unique=True)
        elif t.chave:
            Index('ix_%s_chave' % t.tabela,
                  *[tab.c[nome_fisico(c)] for c in t.chave])
        for col in ('Obra', 'Obra ', 'Data', 'Mês/Ano'):
            fis = nome_fisico(col)
            if any(nome_fisico(c) == fis for c, _ in t.colunas):
                nome_ix = 'ix_%s_%s' % (t.tabela, fis)
                if nome_ix not in {i.name for i in tab.indexes}:
                    Index(nome_ix, tab.c[fis])

    Table(TABELA_DICIONARIO, md,
          Column('tabela_modelo', String(128)), Column('coluna_modelo', String(256)),
          Column('tabela_db', String(128)), Column('coluna_db', String(128)),
          Column('tipo_logico', String(16)), Column('ordem', String(8)),
          Column('fonte_excel', String(128)), Column('chave', String(8)),
          Column('descricao_tabela', String(512)),
          comment='Mapa modelo semantico (pbix) <-> nomes fisicos no banco.')
    Table(TABELA_LOG, md,
          Column('carregado_em', String(32)), Column('arquivo_origem', String(512)),
          Column('tabela_db', String(128)), Column('linhas', String(16)),
          comment='Auditoria de cada carga do ETL.')
    Table(TABELA_EXECUCAO, md,
          Column('iniciado_em', String(32)), Column('terminado_em', String(32)),
          Column('status', String(16)), Column('arquivo_origem', String(512)),
          Column('origem_modificado_em', String(32)),
          Column('origem_tamanho', String(24)), Column('origem_sha256', String(64)),
          Column('linhas_total', String(16)), Column('mensagem', String(1024)),
          comment='Uma linha por execucao do ETL, com a impressao digital da '
                  'planilha lida. E daqui que sai a resposta "a planilha mudou?".')
    Table(TABELA_RELACIONAMENTOS, md,
          Column('tabela_origem', String(128)), Column('coluna_origem', String(128)),
          Column('tabela_destino', String(128)), Column('coluna_destino', String(128)),
          Column('tabela_origem_db', String(128)), Column('coluna_origem_db', String(128)),
          Column('tabela_destino_db', String(128)), Column('coluna_destino_db', String(128)),
          Column('cardinalidade', String(8)), Column('filtro_cruzado', String(8)),
          Column('nome_fk', String(128)), Column('linhas_orfas', String(16)),
          Column('fk_aplicavel', String(8)), Column('nota', String(512)),
          comment='Os 26 relacionamentos do modelo do pbix, com a contagem de '
                  'linhas orfas que hoje impede a FOREIGN KEY.')
    return md


def dicionario(dados):
    linhas = []
    for t in TABELAS:
        for i, (c, tp) in enumerate(t.colunas, start=1):
            linhas.append({
                'tabela_modelo': t.nome, 'coluna_modelo': c,
                'tabela_db': t.tabela, 'coluna_db': nome_fisico(c),
                'tipo_logico': tp, 'ordem': str(i),
                'fonte_excel': t.fonte or '(gerada)',
                'chave': 'S' if c in t.chave else 'N',
                'descricao_tabela': t.descricao,
            })
    return pd.DataFrame(linhas)


def _registros(df):
    """DataFrame -> lista de dicts com None no lugar de NA/NaN/NaT."""
    saida = []
    colunas = list(df.columns)
    fisicas = [nome_fisico(c) for c in colunas]
    for tupla in df.itertuples(index=False, name=None):
        reg = {}
        for fis, v in zip(fisicas, tupla):
            if v is pd.NA or v is None:
                reg[fis] = None
            elif isinstance(v, float) and pd.isna(v):
                reg[fis] = None
            elif v is pd.NaT:
                reg[fis] = None
            else:
                reg[fis] = v
        saida.append(reg)
    return saida


def ultima_carga(dados, db_url, schema=None):
    """Impressao digital da planilha na ultima carga OK, ou None.

    Devolve None tambem quando o banco ainda nao existe - primeira carga.
    """
    from sqlalchemy import create_engine, select

    md = montar_metadata(dados, schema)
    tab = md.tables[(schema + '.' if schema else '') + TABELA_EXECUCAO]
    eng = create_engine(db_url, future=True)
    try:
        with eng.connect() as con:
            if not eng.dialect.has_table(con, TABELA_EXECUCAO, schema=schema):
                return None
            linha = con.execute(
                select(tab.c.origem_sha256, tab.c.terminado_em,
                       tab.c.origem_modificado_em, tab.c.linhas_total)
                .where(tab.c.status == 'ok')
                .order_by(tab.c.terminado_em.desc()).limit(1)).first()
    except Exception as e:                                    # noqa: BLE001
        print('  aviso: nao deu para ler %s (%s); seguindo como primeira carga.'
              % (TABELA_EXECUCAO, e))
        return None
    if not linha:
        return None
    return {'sha256': linha[0], 'terminado_em': linha[1],
            'modificado_em': linha[2], 'linhas': linha[3]}


def carregar(dados, db_url, schema=None, lote=1000, arquivo_origem='',
             digital=None, iniciado_em=None):
    from sqlalchemy import create_engine

    md = montar_metadata(dados, schema)
    eng = create_engine(db_url, future=True)
    print('Destino: %s' % eng.url.render_as_string(hide_password=True))
    sha, tamanho, modificado = digital or ('', '', '')
    with eng.begin() as con:
        md.create_all(con, checkfirst=True)
        agora = _dt.datetime.now().isoformat(timespec='seconds')
        log = []
        for t in TABELAS:
            tab = md.tables[(schema + '.' if schema else '') + t.tabela]
            con.execute(tab.delete())
            regs = _registros(dados[t.nome])
            for i in range(0, len(regs), lote):
                con.execute(tab.insert(), regs[i:i + lote])
            print('  gravado %-30s %6d linhas' % (t.tabela, len(regs)))
            log.append({'carregado_em': agora, 'arquivo_origem': arquivo_origem,
                        'tabela_db': t.tabela, 'linhas': str(len(regs))})

        dic = md.tables[(schema + '.' if schema else '') + TABELA_DICIONARIO]
        con.execute(dic.delete())
        con.execute(dic.insert(), dicionario(dados).to_dict('records'))

        trel = md.tables[(schema + '.' if schema else '') + TABELA_RELACIONAMENTOS]
        con.execute(trel.delete())
        con.execute(trel.insert(), [
            {'tabela_origem': r['origem'], 'coluna_origem': r['coluna_origem'],
             'tabela_destino': r['destino'], 'coluna_destino': r['coluna_destino'],
             'tabela_origem_db': nome_fisico(r['origem']),
             'coluna_origem_db': nome_fisico(r['coluna_origem']),
             'tabela_destino_db': nome_fisico(r['destino']),
             'coluna_destino_db': nome_fisico(r['coluna_destino']),
             'cardinalidade': r['cardinalidade'], 'filtro_cruzado': r['filtro'],
             'nome_fk': r['nome_fk'], 'linhas_orfas': str(r['orfas']),
             'fk_aplicavel': 'S' if r['orfas'] == 0 else 'N',
             'nota': r['nota']}
            for r in integridade(dados)])

        tlog = md.tables[(schema + '.' if schema else '') + TABELA_LOG]
        con.execute(tlog.insert(), log)

        texec = md.tables[(schema + '.' if schema else '') + TABELA_EXECUCAO]
        con.execute(texec.insert(), [{
            'iniciado_em': iniciado_em or agora, 'terminado_em': agora,
            'status': 'ok', 'arquivo_origem': arquivo_origem,
            'origem_modificado_em': modificado, 'origem_tamanho': str(tamanho),
            'origem_sha256': sha,
            'linhas_total': str(sum(int(x['linhas']) for x in log)),
            'mensagem': '',
        }])
    print('Carga concluida (full refresh, uma transacao).')


def gerar_ddl(dados, dialeto, schema=None):
    from sqlalchemy.schema import CreateIndex, CreateTable
    dialetos = {
        'postgresql': 'postgresql.psycopg2.dialect',
        'mssql': 'mssql.pyodbc.dialect',
    }
    if dialeto not in dialetos:
        raise SystemExit('dialeto deve ser um de: %s' % ', '.join(dialetos))
    mod = __import__('sqlalchemy.dialects.' + dialetos[dialeto].rsplit('.', 1)[0],
                     fromlist=['dialect'])
    d = mod.dialect()
    md = montar_metadata(dados, schema)
    partes = ['-- DDL do modelo semantico Ambiental/Qualidade (%s)' % dialeto,
              '-- Gerado por etl/etl_excel_para_db.py em %s'
              % _dt.date.today().isoformat(), '']
    for tab in md.sorted_tables:
        partes.append(str(CreateTable(tab).compile(dialect=d)).strip() + ';')
        for ix in sorted(tab.indexes, key=lambda i: i.name):
            partes.append(str(CreateIndex(ix).compile(dialect=d)).strip() + ';')
        partes.append('')
    return '\n'.join(partes)


def gerar_ddl_fk(dados, dialeto, schema=None):
    """Script das 26 FOREIGN KEYs do modelo, separado do DDL das tabelas.

    Fica separado de proposito: hoje varias delas NAO podem ser aplicadas
    porque o dado tem orfa (North Tower fora de Dim_Obras, "Outros" em
    impressoes, "Iconic Tower " em reclamacoes). As que nao dao sao geradas
    comentadas, com a contagem de orfas ao lado.
    """
    pre = (schema + '.') if schema else ''
    aspas = '"' if dialeto != 'mssql' else ''
    linhas = [
        '-- FOREIGN KEYs do modelo semantico Ambiental/Qualidade (%s)' % dialeto,
        '-- Gerado por etl/etl_excel_para_db.py em %s' % _dt.date.today().isoformat(),
        '--',
        '-- As FKs COMENTADAS so podem ser criadas depois de limpar as linhas',
        '-- orfas indicadas. Ate la o relacionamento existe como documentacao',
        '-- na tabela %s e como indice.' % TABELA_RELACIONAMENTOS,
        '',
    ]
    for r in integridade(dados):
        alvo_uq = 'uq_%s' % nome_fisico(r['destino'])
        cmd = ('ALTER TABLE %s%s%s%s ADD CONSTRAINT %s FOREIGN KEY (%s%s%s) '
               'REFERENCES %s%s%s%s (%s%s%s);'
               % (pre, aspas, nome_fisico(r['origem']), aspas, r['nome_fk'],
                  aspas, nome_fisico(r['coluna_origem']), aspas,
                  pre, aspas, nome_fisico(r['destino']), aspas,
                  aspas, nome_fisico(r['coluna_destino']), aspas))
        if r['orfas'] == 0:
            linhas.append(cmd)
        else:
            linhas.append('-- %d linha(s) orfa(s) (ex.: %s) - limpe antes de criar:'
                          % (r['orfas'], ', '.join(r['exemplos'][:4])))
            linhas.append('-- ' + cmd)
        linhas.append('')
    linhas.insert(6, '-- Depende do indice unico %s criado junto com as tabelas.'
                  % ', '.join(sorted({'uq_' + nome_fisico(d)
                                      for d in CHAVES_DIMENSAO})))
    return '\n'.join(linhas)


def exportar_arquivos(dados, dir_parquet=None, dir_csv=None):
    for destino, ext in ((dir_parquet, 'parquet'), (dir_csv, 'csv')):
        if not destino:
            continue
        os.makedirs(destino, exist_ok=True)
        for t in TABELAS:
            df = dados[t.nome].copy()
            df.columns = [nome_fisico(c) for c in df.columns]
            caminho = os.path.join(destino, '%s.%s' % (t.tabela, ext))
            if ext == 'parquet':
                df.astype({c: 'string' for c in df.columns
                           if df[c].dtype == object}).to_parquet(caminho, index=False)
            else:
                df.to_csv(caminho, index=False, sep=';', decimal=',',
                          encoding='utf-8-sig')
        print('Exportado %s em %s' % (ext, destino))


# --------------------------------------------------------------------- CLI ---
class _Tee:
    """Escreve na saida padrao e tambem no arquivo de log."""

    def __init__(self, fluxo, arquivo):
        self.fluxo, self.arquivo = fluxo, arquivo

    def write(self, texto):
        self.fluxo.write(texto)
        self.arquivo.write(texto)

    def flush(self):
        self.fluxo.flush()
        self.arquivo.flush()


def _ligar_log(caminho):
    pasta = os.path.dirname(os.path.abspath(caminho))
    if pasta:
        os.makedirs(pasta, exist_ok=True)
    fh = open(caminho, 'a', encoding='utf-8')
    sys.stdout = _Tee(sys.stdout, fh)
    sys.stderr = _Tee(sys.stderr, fh)


def registrar_falha(db_url, schema, arquivo, mensagem, iniciado_em):
    """Grava a falha em _carga_execucao, se a tabela ja existir.

    Sem isso, uma tarefa agendada que quebra toda noite so aparece no log de
    texto - e ninguem le log de texto. Com isso, da para monitorar por SQL.
    """
    from sqlalchemy import create_engine, text
    alvo = ('%s.%s' % (schema, TABELA_EXECUCAO)) if schema else TABELA_EXECUCAO
    agora = _dt.datetime.now().isoformat(timespec='seconds')
    try:
        eng = create_engine(db_url, future=True)
        with eng.begin() as con:
            con.execute(text(
                'INSERT INTO %s (iniciado_em, terminado_em, status, '
                'arquivo_origem, mensagem) VALUES (:i, :t, :s, :a, :m)' % alvo),
                {'i': iniciado_em, 't': agora, 's': 'erro', 'a': arquivo or '',
                 'm': mensagem[:1000]})
        print('Falha registrada em %s.' % alvo)
    except Exception as e:                                    # noqa: BLE001
        print('Nao deu para registrar a falha no banco (%s).' % e)


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('--excel', help='caminho do .xlsx de origem')
    p.add_argument('--db-url',
                   help='URL SQLAlchemy do destino PostgreSQL, ex.: '
                        '"postgresql+psycopg2://usuario@servidor/base". '
                        'Obrigatorio para carregar; --somente-validar e '
                        '--somente-ddl nao precisam')
    p.add_argument('--schema', default=None, help='schema do banco (ex: ambiental)')
    p.add_argument('--parquet', help='diretorio para exportar .parquet tambem')
    p.add_argument('--csv', help='diretorio para exportar .csv tambem')
    p.add_argument('--ddl', help='arquivo onde gravar o DDL gerado')
    p.add_argument('--dialeto', default='postgresql',
                   choices=['postgresql', 'mssql'],
                   help='dialeto do DDL (padrao: postgresql, o destino do modelo)')
    p.add_argument('--somente-validar', action='store_true',
                   help='le e valida, mas nao grava em lugar nenhum')
    p.add_argument('--somente-ddl', action='store_true',
                   help='so imprime/grava o DDL, sem carregar o banco')
    p.add_argument('--se-mudou', action='store_true',
                   help='so carrega se a planilha mudou desde a ultima carga OK '
                        '(use na tarefa agendada)')
    p.add_argument('--copiar-local', choices=['auto', 'sim', 'nao'], default='auto',
                   help='copiar a planilha para disco local antes de ler; "auto" '
                        '(padrao) copia quando a origem esta em rede')
    p.add_argument('--tentativas', type=int, default=3,
                   help='tentativas de copia, para o caso de a planilha estar '
                        'sendo salva no Excel (padrao: 3)')
    p.add_argument('--log', help='arquivo de log (append) alem da saida padrao')
    a = p.parse_args(argv)

    if a.log:
        _ligar_log(a.log)

    inicio = _dt.datetime.now()
    print('=== ETL Ambiental/Qualidade - %s ===' % inicio.isoformat(timespec='seconds'))

    try:
        excel = org.resolver(a.excel, EXCEL_CANDIDATOS)
    except org.OrigemIndisponivel as e:
        print('ERRO: %s' % e)
        return 3

    if a.copiar_local == 'sim' or (a.copiar_local == 'auto' and org.eh_rede(excel)):
        print('Origem em rede - copiando para disco local antes de ler.')
        try:
            local = org.copiar_local(excel, tentativas=a.tentativas)
        except org.OrigemIndisponivel as e:
            print('ERRO: %s' % e)
            return 3
    else:
        local = excel

    digital = org.impressao_digital(local)
    print('Planilha: %s\n  modificada em %s, %s bytes, sha256 %s...'
          % (excel, digital[2], digital[1], digital[0][:12]))

    avisos = []
    dados = extrair(local, avisos)
    problemas = verificar(dados, avisos)

    if avisos:
        print('\n--- Avisos (dado carregado, mas confira) ---')
        for x in avisos:
            print('  * %s' % x)
    if ALERTAS_MODELO:
        print('\n--- Defeitos conhecidos do modelo, reproduzidos de proposito ---')
        for tab, col, txt in ALERTAS_MODELO:
            print('  * %s.%s: %s' % (tab, col, txt))
    if problemas:
        print('\n--- ERROS ---')
        for x in problemas:
            print('  ! %s' % x)
        raise SystemExit(2)

    print('\n--- Relacionamentos do modelo (integridade) ---')
    rels = integridade(dados)
    ok = [r for r in rels if r['orfas'] == 0]
    print('  %d relacionamentos; %d viram FOREIGN KEY hoje, %d ainda nao:'
          % (len(rels), len(ok), len(rels) - len(ok)))
    for r in rels:
        if r['orfas']:
            print('    %-27s %-8s -> %-15s %5d/%-5d orfas (%s)'
                  % (r['origem'], r['coluna_origem'].strip(), r['destino'],
                     r['orfas'], r['linhas'], ', '.join(r['exemplos'][:3])))

    if a.ddl or a.somente_ddl:
        ddl = gerar_ddl(dados, a.dialeto, a.schema)
        fks = gerar_ddl_fk(dados, a.dialeto, a.schema)
        if a.ddl:
            os.makedirs(os.path.dirname(os.path.abspath(a.ddl)), exist_ok=True)
            with open(a.ddl, 'w', encoding='utf-8') as fh:
                fh.write(ddl)
            alvo_fk = a.ddl.replace('.sql', '_fk.sql')
            with open(alvo_fk, 'w', encoding='utf-8') as fh:
                fh.write(fks)
            print('\nDDL (%s) gravado em %s' % (a.dialeto, a.ddl))
            print('FKs (%s) gravadas em %s' % (a.dialeto, alvo_fk))
        else:
            print(ddl)
            print(fks)
    if a.somente_ddl:
        return 0

    exportar_arquivos(dados, a.parquet, a.csv)

    if a.somente_validar:
        print('\n--somente-validar: nada foi gravado no banco.')
        return 0

    # Sem destino explicito nao ha carga. O padrao antigo era um SQLite local,
    # que gravava em silencio e deixava quem preparava o banco achando que
    # tinha carregado o PostgreSQL.
    if not a.db_url:
        print('\nERRO: --db-url e obrigatorio para carregar o banco.')
        print('      ex.: --db-url "postgresql+psycopg2://usuario@servidor/base"'
              ' --schema ambiental')
        print('      Para so conferir a planilha, use --somente-validar.')
        return 2

    # --se-mudou so pula DEPOIS de ler e validar a planilha: assim a tarefa
    # agendada continua avisando de problema na origem mesmo em dia parado.
    if a.se_mudou:
        anterior = ultima_carga(dados, a.db_url, a.schema)
        if anterior and anterior['sha256'] == digital[0]:
            print('\nPlanilha inalterada desde a carga de %s (mesmo sha256). '
                  'Nada a fazer.' % anterior['terminado_em'])
            print('Use sem --se-mudou para forcar a recarga.')
            return 0
        if anterior:
            print('\nPlanilha mudou desde %s - recarregando.'
                  % anterior['terminado_em'])
        else:
            print('\nPrimeira carga neste banco.')

    try:
        carregar(dados, a.db_url, a.schema, arquivo_origem=excel,
                 digital=digital, iniciado_em=inicio.isoformat(timespec='seconds'))
    except Exception as e:                                    # noqa: BLE001
        import traceback
        traceback.print_exc()
        registrar_falha(a.db_url, a.schema, excel, '%s: %s' % (type(e).__name__, e),
                        inicio.isoformat(timespec='seconds'))
        return 1
    print('Duracao: %.1fs' % (_dt.datetime.now() - inicio).total_seconds())
    return 0


if __name__ == '__main__':
    sys.exit(main())
