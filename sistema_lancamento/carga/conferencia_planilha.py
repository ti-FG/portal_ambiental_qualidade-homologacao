"""
conferencia_planilha.py — confere a planilha contra o sistema, linha a linha, 21/10/2026

Para a Qualidade validar a migração com a planilha aberta ao lado. Gera um
.xlsx com:
  - Resumo        quantas linhas batem, quantas foram ajustadas e por quê
  - De-Para       cada grafia de obra/unidade na planilha -> nome no sistema
  - uma aba por assunto (Resíduos, Conformidade, Metas, NCs, Avanço físico,
    Ruído, Água e luz): ABA e LINHA do Excel, o texto como estava escrito, o
    valor da planilha, o valor que está HOJE no sistema e a situação.

Compara com o banco atual (não com a carga): se alguém alterou um valor no
sistema depois da carga, aparece como "alterado no sistema", com quem e quando.

Uso:  python conferencia_planilha.py [saida.xlsx]
"""
import datetime as dt
import os
import re
import sys
import warnings
from collections import defaultdict

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)
sys.path.insert(0, AQUI)

import pandas as pd  # noqa: E402
import psycopg2  # noqa: E402
from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402
from openpyxl.utils import get_column_letter, range_boundaries  # noqa: E402

import carga_planilha as cp  # noqa: E402  (mesmas regras de nome e data da carga)

warnings.filterwarnings('ignore', message='pandas only supports SQLAlchemy')
MESES = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez']

IGUAL = 'igual'
DIFERENTE = 'DIFERENTE'
ALTERADO = 'alterado no sistema'
AJUSTE = 'ajustado na migração'
NAO_MIGRADO = 'não migrado'
CORES = {IGUAL: 'E4F4F3', DIFERENTE: 'F9E4E2', ALTERADO: 'E3ECF8', AJUSTE: 'F8F5DE', NAO_MIGRADO: 'EEEEEE'}


# ── leitura da planilha guardando aba e linha ──────────────────────────────
def ler(wb, nome):
    for ws in wb.worksheets:
        if nome in ws.tables:
            t = ws.tables[nome]
            break
    else:
        raise KeyError(nome)
    c0, r0, c1, r1 = range_boundaries(t.ref)
    cab = r0 + (t.headerRowCount if t.headerRowCount is not None else 1)
    cols = [c.name for c in t.tableColumns]
    for i, row in enumerate(ws.iter_rows(min_row=cab, max_row=r1, min_col=c0, max_col=c1, values_only=True)):
        if all(v is None or (isinstance(v, str) and not v.strip()) for v in row):
            continue
        d = dict(zip(cols, row))
        d['_aba'], d['_linha'] = ws.title, cab + i
        yield d


def mes_txt(c):
    return '%s/%s' % (MESES[c.month - 1], str(c.year)[2:]) if c else ''


def igual(a, b, tol=0.005):
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    return abs(float(a) - float(b)) <= tol * max(1.0, abs(float(a)))


def quem(r):
    if r is None or r.get('atualizado_por') in (None, 'carga_planilha'):
        return ''
    return '%s em %s' % (r['atualizado_por'], r['atualizado_em'].strftime('%d/%m/%Y %H:%M') if r.get('atualizado_em') is not None else '')


def situacao(v_plan, v_sis, reg):
    if igual(v_plan, v_sis):
        return IGUAL
    if reg is not None and reg.get('atualizado_por') not in (None, 'carga_planilha'):
        return ALTERADO
    return DIFERENTE


# ── banco ──────────────────────────────────────────────────────────────────
def banco():
    env = {}
    try:
        with open(os.path.join(RAIZ, 'www', '.env'), encoding='utf-8') as f:
            for linha in f:
                linha = linha.strip()
                if linha and not linha.startswith('#') and '=' in linha:
                    k, v = linha.split('=', 1)
                    env[k.strip()] = v.strip().strip('"\'')
    except OSError:
        pass
    env.update({k: v for k, v in os.environ.items() if k.startswith('PG') or k == 'APP_SCHEMA'})
    con = psycopg2.connect(host=env.get('PGHOST'), port=int(env.get('PGPORT', 5432)), dbname=env.get('PGDATABASE'),
                           user=env.get('PGUSER'), password=env.get('PGPASSWORD') or None, connect_timeout=10)
    con.cursor().execute('SET search_path TO %s' % env.get('APP_SCHEMA', 'portal_ambiental_qualidade'))

    def q(sql):
        df = pd.read_sql_query(sql, con)
        # o pandas troca NULL por NaN; aqui vazio tem de continuar vazio (None)
        return df.astype(object).where(pd.notna(df), None).to_dict('records')
    B = {
        'residuo': q('''SELECT o.nome AS obra, r.competencia, t.nome AS tipo, r.volume_m3, t.densidade_kg_m3,
                               round(r.volume_m3 * t.densidade_kg_m3, 2) AS kg, am.area_mes_m2, r.atualizado_por,
                               r.atualizado_em AT TIME ZONE 'America/Sao_Paulo' AS atualizado_em
                          FROM residuo r JOIN obra o ON o.id = r.obra_id JOIN tipo_residuo t ON t.id = r.tipo_residuo_id
                          LEFT JOIN v_area_mes am ON am.obra_id = r.obra_id AND am.competencia = r.competencia'''),
        'conf': q('''SELECT v.obra, v.competencia, v.percentual, v.meta, c.atualizado_por,
                            c.atualizado_em AT TIME ZONE 'America/Sao_Paulo' AS atualizado_em
                       FROM v_conformidade v JOIN conformidade c ON c.obra_id = v.obra_id AND c.competencia = v.competencia'''),
        'metas': q('SELECT o.nome AS obra, m.vigente_desde, m.meta FROM meta_conformidade m JOIN obra o ON o.id = m.obra_id'),
        'ncs': q('''SELECT o.nome AS obra, c.competencia, a.tipo, a.nome AS assunto FROM conformidade_nc c
                      JOIN obra o ON o.id = c.obra_id JOIN assunto_nc a ON a.id = c.assunto_id'''),
        'avanco': q('''SELECT v.obra, v.competencia, v.avanco_pct, v.area_total_m2, v.area_mes_m2, a.atualizado_por,
                              a.atualizado_em AT TIME ZONE 'America/Sao_Paulo' AS atualizado_em
                         FROM v_area_mes v JOIN avanco_fisico a ON a.obra_id = v.obra_id AND a.competencia = v.competencia'''),
        'ruido': q('''SELECT o.nome AS obra, r.competencia, r.ponto, to_char(r.horario, 'HH24:MI') AS horario, r.db,
                             r.atualizado_por, r.atualizado_em AT TIME ZONE 'America/Sao_Paulo' AS atualizado_em
                        FROM medicao_ruido r JOIN obra o ON o.id = r.obra_id'''),
        'consumo': q('''SELECT u.nome AS unidade, c.competencia, c.linha, c.luz_kwh, c.agua_m3, c.atualizado_por,
                               c.atualizado_em AT TIME ZONE 'America/Sao_Paulo' AS atualizado_em
                          FROM consumo c JOIN unidade_consumidora u ON u.id = c.uc_id'''),
        'ucs': q('SELECT u.nome, p.nome AS pai FROM unidade_consumidora u LEFT JOIN unidade_consumidora p ON p.id = u.agrupa_em_id'),
        # abas que são listas
        'l_auditoria': q('SELECT o.nome AS obra, x.competencia, x.meta, x.realizado, x.ocorrencias FROM auditoria_ambiental x JOIN obra o ON o.id = x.obra_id'),
        'l_plano': q('SELECT o.nome AS obra, x.indicador, x.analise, x.acao, x.responsavel, x.prazo, x.status FROM plano_acao x JOIN obra o ON o.id = x.obra_id'),
        'l_emerg': q('SELECT o.nome AS obra, x.data, x.n_fmc, x.ocorrencia, x.status_acoes, x.categoria FROM emergencia x JOIN obra o ON o.id = x.obra_id'),
        'l_recl': q('SELECT o.nome AS obra, x.data, x.cod_chamado, x.canal FROM reclamacao x JOIN obra o ON o.id = x.obra_id'),
        'l_eco': q('SELECT o.nome AS obra, x.competencia, x.ponto_1, x.db, x.ponto2, x.db3 FROM eco_urbana x JOIN obra o ON o.id = x.obra_id'),
        'l_sga': q('SELECT ano, mes, pessoas FROM acoes_sga'),
        'l_imp': q('SELECT competencia, total_pag, pct_colorido FROM impressoes'),
        'l_setor': q('SELECT competencia, nome, pb, cor FROM impressoes_setor'),
        'l_pessoa': q('SELECT competencia, usuario, pb, cor FROM impressoes_pessoa'),
    }
    con.close()
    for k in ('residuo', 'conf', 'avanco', 'ruido', 'consumo', 'ncs'):
        for r in B[k]:
            r['competencia'] = pd.Timestamp(r['competencia']).date()
    for r in B['metas']:
        r['vigente_desde'] = pd.Timestamp(r['vigente_desde']).date()
    return B


# ── abas que são listas: cada linha da planilha tem de existir no sistema ──
def _t(v):
    return '' if cp.vazio(v) else ' '.join(str(v).split()).lower()


def _n(v, casas=4):
    x = cp.num(v)
    return None if x is None else round(float(x), casas)


def _d(v):
    v = cp.data(v)  # datetime (planilha), date ou Timestamp (banco) -> date: chave sem hora
    return v.isoformat() if v else ''


def conferir_listas(wb, B, obra):
    """Para cada aba: chave da linha na planilha × chaves do banco (multiconjunto).
    Linha que acha par = igual; sem par = DIFERENTE (ou não migrado, com motivo)."""
    def casar(nome, linhas_plan, chaves_banco, cab_extra, montar):
        resto = defaultdict(int)
        for k in chaves_banco:
            resto[k] += 1
        out = []
        for d, chave, mostrar, motivo in linhas_plan:
            if motivo:
                st, obs = (AJUSTE if motivo.startswith('ajuste:') else NAO_MIGRADO), motivo.replace('ajuste: ', '')
            elif resto[chave] > 0:
                resto[chave] -= 1
                st, obs = IGUAL, ''
            else:
                st, obs = DIFERENTE, 'não achei esta linha no sistema com os mesmos valores'
            out.append([d['_aba'], d['_linha']] + mostrar + [st, obs])
        return nome, (['Aba', 'Linha Excel'] + cab_extra + ['Situação', 'Observação'], out)

    abas = {}
    m_ = {'Todos': 0, **{k.capitalize(): v for k, v in cp._MES_N.items()}}

    L = []
    for d in ler(wb, 'Auditoria_Ambientais'):
        o, c = obra(d['_aba'], d['Obra ']), cp.data(d['Mês/Ano'])
        comp = dt.date(c.year, c.month, 1) if c else None
        real = _t(d['Realizado '])
        L.append((d, (o, comp, _n(cp.pct(d['Meta '])), None if not real else real.startswith('s'), _t(d['Principais ocorrencias da última auditória'])),
                  [d['Obra '], o or '', mes_txt(comp), d['Realizado ']], None if o and comp else 'obra ou Mês/Ano não reconhecido'))
    n, a = casar('Auditoria Ambientais', L, [(r['obra'], pd.Timestamp(r['competencia']).date(), _n(r['meta']), r['realizado'], _t(r['ocorrencias'])) for r in B['l_auditoria']],
                 ['Obra (como na planilha)', 'Obra (sistema)', 'Competência', 'Realizado'], None)
    abas[n] = a

    L = []
    for d in ler(wb, 'Plano_acao'):
        o = obra(d['_aba'], d['Obra'])
        L.append((d, (o, _t(d['Indicador ']), _t(d['Análise ']), _t(d['Ação']), _t(d['Responsável ']), _d(d['Prazo']), _t(d['Status'])),
                  [d['Obra'], o or '', d['Indicador '], d['Status']], None if o else 'obra não reconhecida'))
    n, a = casar('Plano de Ação', L, [(r['obra'], _t(r['indicador']), _t(r['analise']), _t(r['acao']), _t(r['responsavel']), _d(r['prazo']), _t(r['status'])) for r in B['l_plano']],
                 ['Obra (como na planilha)', 'Obra (sistema)', 'Indicador', 'Status'], None)
    abas[n] = a

    L = []
    for d in ler(wb, 'emergencias_ambientais'):
        o = obra(d['_aba'], d['Obra '])
        vazia = all(cp.vazio(d[c]) for c in ('N° FMC', 'Ocorerência', 'Status/Ações'))
        cat = _t(d['Categoria '])
        motivo = 'linha só com obra e data, sem ocorrência/FMC/status' if vazia else None if o else 'obra não reconhecida'
        L.append((d, (o, _d(d['Mês/Ano']), _n(d['N° FMC'], 0), _t(d['Ocorerência']), _t(d['Status/Ações']), 'emergência' if cat == 'emegência' else cat),
                  [d['Obra '], o or '', _d(d['Mês/Ano']), d['Categoria ']], motivo))
    n, a = casar('Emergências Ambientais', L, [(r['obra'], _d(r['data']), _n(r['n_fmc'], 0), _t(r['ocorrencia']), _t(r['status_acoes']), _t(r['categoria'])) for r in B['l_emerg']],
                 ['Obra (como na planilha)', 'Obra (sistema)', 'Mês/Ano', 'Categoria'], None)
    abas[n] = a
    for linha in a[1]:  # categoria corrigida conta como ajuste
        if linha[-2] == IGUAL and _t(linha[5]) == 'emegência':
            linha[-2], linha[-1] = AJUSTE, 'categoria "Emegência" (erro de digitação) gravada como "Emergência"'

    L = []
    for d in ler(wb, 'reclamacoes_vizinhos'):
        partes = [p.strip() for p in re.split(r'\s*[-–]\s*', _t(d['Obra']) and str(d['Obra'])) if p.strip()] if not cp.vazio(d['Obra']) else []
        o = obra(d['_aba'], partes[0] if partes else None)
        cod = d['Cod. Chamado']
        cod = None if cp.vazio(cod) else str(int(cod)) if isinstance(cod, float) else str(cod)
        L.append((d, (o, _d(d['Mês/Ano']), cod), [d['Obra'], o or '', partes[-1] if len(partes) > 1 else '', cod], None if o else 'obra não reconhecida'))
    n, a = casar('Reclamações vizinhos', L, [(r['obra'], _d(r['data']), r['cod_chamado']) for r in B['l_recl']],
                 ['Obra (como na planilha)', 'Obra (sistema)', 'Canal', 'Cod. Chamado'], None)
    abas[n] = a

    L = []
    for d in ler(wb, 'Eco_urbana'):
        o, c = obra(d['_aba'], d['Obra']), cp.competencia(d['Mês/Ano'])
        L.append((d, (o, c, _t(d['Ponto_1']), _n(d['db'], 2), _t(d['Ponto2']), _n(d['db3'], 2)), [d['Obra'], o or '', mes_txt(c)], None if o and c else 'obra ou mês não reconhecido'))
    n, a = casar('Eco urbana', L, [(r['obra'], pd.Timestamp(r['competencia']).date(), _t(r['ponto_1']), _n(r['db'], 2), _t(r['ponto2']), _n(r['db3'], 2)) for r in B['l_eco']],
                 ['Obra (como na planilha)', 'Obra (sistema)', 'Competência'], None)
    abas[n] = a

    L = []
    for d in ler(wb, 'Acoes_SGA'):
        m = m_.get(str(d['Mês']).strip().capitalize())
        L.append((d, (int(cp.num(d['Ano'])), m, _n(d['Pessoas'], 0)), [d['Ano'], d['Mês'], cp.num(d['Pessoas'])], None if m is not None else 'mês não reconhecido'))
    n, a = casar('Ações SGA', L, [(int(r['ano']), int(r['mes']), _n(r['pessoas'], 0)) for r in B['l_sga']], ['Ano', 'Mês', 'Pessoas'], None)
    abas[n] = a

    L = []
    for d in ler(wb, 'Impressoes'):
        c = cp.competencia(d['Mês/Ano'])
        if not c and cp.num(d['Ano']) and cp.mes_num(d['Mês']):
            c = dt.date(int(cp.num(d['Ano'])), cp.mes_num(d['Mês']), 1)
        L.append((d, (c, _n(d['Total de pág '], 0)), [mes_txt(c), cp.num(d['Total de pág ']), cp.num(d['% Colorido'])], None if c else 'sem mês'))
    n, a = casar('Impressões', L, [(pd.Timestamp(r['competencia']).date(), _n(r['total_pag'], 0)) for r in B['l_imp']], ['Competência', 'Total de pág', '% Colorido'], None)
    abas[n] = a

    for nome, tab, col, chave_b in (('Impressões por setor', 'Impressoes_por_setor', 'Nome', 'l_setor'), ('Impressões por pessoa', 'Impressoes_por_pessoa', 'Usuário', 'l_pessoa')):
        L = []
        for d in ler(wb, tab):
            c = cp.competencia(d['Mês/Ano'])
            L.append((d, (c, _t(d[col]) or '-', _n(d['P&B'], 0), _n(d['Cor'], 0)), [mes_txt(c), d[col], cp.num(d['P&B']), cp.num(d['Cor'])], None if c else 'sem Mês/Ano'))
        campo = 'nome' if chave_b == 'l_setor' else 'usuario'
        n, a = casar(nome, L, [(pd.Timestamp(r['competencia']).date(), _t(r[campo]) or '-', _n(r['pb'], 0), _n(r['cor'], 0)) for r in B[chave_b]],
                     ['Competência', col, 'P&B', 'Cor'], None)
        abas[n] = a
    return abas


# ── conferência ────────────────────────────────────────────────────────────
def main():
    saida = sys.argv[1] if len(sys.argv) > 1 else os.path.join(AQUI, 'conferencia_planilha.xlsx')
    wb = load_workbook(cp.PLANILHA, data_only=True)
    B = banco()
    # obra nova na planilha (ex.: Central Tower): mesma regra da carga
    cp.descobrir_obras({n: pd.DataFrame(list(ler(wb, n))) for n in ('Avanco_fisico', 'Metas_LAI_EIV')})
    depara = defaultdict(int)       # (aba, como escrito, nome no sistema) -> linhas
    abas = {}

    def obra(aba, texto):
        o = cp.obra_oficial(texto)
        depara[(aba, '' if texto is None else str(texto), o or '(não reconhecida)')] += 1
        return o

    # Resíduos ---------------------------------------------------------------
    sis = {(r['obra'], r['competencia'], r['tipo']): r for r in B['residuo']}
    linhas_plan = list(ler(wb, 'SOMA_RESIDUOS'))
    grupos = defaultdict(list)
    for d in linhas_plan:
        o = obra(d['_aba'], d['Obra'])
        c = cp.competencia(d['Mês/Ano'])
        tipo = str(d['Resíduos Detalhados']).strip() if not cp.vazio(d['Resíduos Detalhados']) else 'Classe %s (sem detalhe)' % str(d['Classe']).strip()
        d['_k'] = (o, c, tipo)
        if cp.num(d['m³']) is not None and o and c:
            grupos[d['_k']].append(d['_linha'])
    out = []
    for d in linhas_plan:
        o, c, tipo = d['_k']
        m3 = cp.num(d['m³'])
        r = sis.get(d['_k'])
        obs = []
        if m3 is None:
            st, obs = NAO_MIGRADO, ['linha sem m³ (grade pré-criada, sem lançamento)']
        elif not o:
            st, obs = NAO_MIGRADO, ['obra não reconhecida']
        else:
            outras = [x for x in grupos[d['_k']] if x != d['_linha']]
            soma = sum(cp.num(x['m³']) for x in linhas_plan if x['_linha'] in grupos[d['_k']])
            st = situacao(soma, r['volume_m3'] if r else None, r)
            if outras:
                obs.append('mesma obra/mês/tipo nas linhas %s: o sistema guarda a soma (%s m³)' % (', '.join(map(str, outras)), cp.sql(soma)))
            if tipo.endswith('(sem detalhe)'):
                obs.append('lançamento antigo só por classe: tipo "%s", sem densidade' % tipo)
            elif r and r['densidade_kg_m3'] is None:
                obs.append('tipo sem densidade no cadastro: kg fica em branco (na planilha saía 0)')
            m2 = cp.num(d['m²'])
            if r and r['area_mes_m2'] and (m2 or 0) <= 1:
                obs.append('m² da planilha = %s (não lançado); sistema usa a área do mês do avanço físico: %s m²' % (cp.sql(m2), cp.sql(float(r['area_mes_m2']))))
            elif r and r['area_mes_m2'] and not igual(m2, r['area_mes_m2'], 0.01):
                obs.append('m² da planilha (%s) ≠ área do mês do avanço físico (%s): o sistema usa a do avanço' % (cp.sql(m2), cp.sql(float(r['area_mes_m2']))))
            elif r and not r['area_mes_m2'] and m2 and m2 > 1:
                obs.append('m² da planilha = %s, mas não há avanço físico lançado no mês: kg/m² fica em branco no sistema' % cp.sql(m2))
            if st == IGUAL and obs and not outras:
                st = AJUSTE if any('sem densidade' in x or 'm² da planilha' in x for x in obs) else IGUAL
        out.append([d['_aba'], d['_linha'], d['Obra'], o or '', mes_txt(c), d['Classe'], d['Resíduos Detalhados'] or '', tipo,
                    m3, r['volume_m3'] if r else None, cp.num(d['Kg']), r['kg'] if r else None,
                    cp.num(d['m²']), r['area_mes_m2'] if r else None, st, quem(r), '; '.join(obs)])
    abas['Resíduos'] = (['Aba', 'Linha Excel', 'Obra (como na planilha)', 'Obra (sistema)', 'Competência', 'Classe', 'Tipo (planilha)',
                         'Tipo (sistema)', 'm³ planilha', 'm³ sistema', 'Kg planilha', 'Kg sistema', 'm² planilha', 'Área do mês sistema',
                         'Situação', 'Alterado por', 'Observação'], out)

    # Conformidade -----------------------------------------------------------
    sis = {(r['obra'], r['competencia']): r for r in B['conf']}
    out = []
    for d in ler(wb, 'Conformidade_LAI_EIV'):
        o = obra(d['_aba'], d['OBRA'])
        c = cp.competencia(d['Mês/Ano'])
        p = cp.num(d['PERCENTUAL'])
        if p is not None and p > 1:
            p = p / 100.0
        r = sis.get((o, c))
        obs = []
        if p is None:
            st = NAO_MIGRADO; obs.append('sem percentual')
        else:
            st = situacao(p, r['percentual'] if r else None, r)
            mp = cp.num(d['META'])
            if r and not igual(mp, r['meta']):
                obs.append('meta: planilha %s (o PROCV lia sempre ago/25), sistema %s (meta vigente no mês)'
                           % ('vazia' if mp is None else '%.0f%%' % (mp * 100), 'sem meta' if r['meta'] is None else '%.0f%%' % (r['meta'] * 100)))
        out.append([d['_aba'], d['_linha'], d['OBRA'], o or '', mes_txt(c), p, r['percentual'] if r else None,
                    cp.num(d['META']), r['meta'] if r else None, st, quem(r), '; '.join(obs)])
    abas['Conformidade'] = (['Aba', 'Linha Excel', 'Obra (como na planilha)', 'Obra (sistema)', 'Competência', '% planilha', '% sistema',
                             'Meta planilha', 'Meta sistema (vigente)', 'Situação', 'Alterado por', 'Observação'], out)

    # Metas (formato largo) --------------------------------------------------
    out = []
    metas = defaultdict(list)
    for r in B['metas']:
        metas[r['obra']].append(r)
    mapa_m = {m: i + 1 for i, m in enumerate(MESES)}
    for d in ler(wb, 'Metas_LAI_EIV'):
        o = obra(d['_aba'], d['Obra'])
        for col in [k for k in d if '/' in str(k) and not str(k).startswith('_')]:
            m, a = col.split('/')
            comp = dt.date(2000 + int(a), mapa_m[m], 1)
            v = cp.num(d[col])
            vig = [x for x in sorted(metas.get(o, []), key=lambda x: x['vigente_desde']) if x['vigente_desde'] <= comp]
            vs = vig[-1]['meta'] if vig else None
            if v is None:
                st, obs = (IGUAL if vs is None else AJUSTE), ('' if vs is None else 'vazia na planilha: o sistema mantém a meta anterior')
            else:
                st, obs = (IGUAL if igual(v, vs) else DIFERENTE), ''
            out.append([d['_aba'], d['_linha'], col, d['Obra'], o or '', v, vs, st, obs])
    abas['Metas conformidade'] = (['Aba', 'Linha Excel', 'Coluna (mês)', 'Obra (como na planilha)', 'Obra (sistema)', 'Meta planilha',
                                   'Meta sistema (vigente)', 'Situação', 'Observação'], out)

    # NCs --------------------------------------------------------------------
    sis = {(r['obra'], r['competencia'], r['tipo'], r['assunto']) for r in B['ncs']}
    out = []
    for d in ler(wb, 'NC_LAI_EIV'):
        o = obra(d['_aba'], d['Obra'])
        c = cp.competencia(d['Mês/Ano'])
        cols = (('LAI/EIV', 'NC LAI/EIV'),) if 'NC LAI/EIV' in d else (('LAI', 'NC LAI'), ('EIV', 'NC EIV'))
        for tipo, col in cols:
            if cp.vazio(d[col]):
                continue
            for a in str(d[col]).split('|'):
                a = a.strip()
                if a and a != '-':
                    ok = (o, c, tipo, a) in sis
                    out.append([d['_aba'], d['_linha'], d['Obra'], o or '', mes_txt(c), tipo, a, 'sim' if ok else 'não', IGUAL if ok else DIFERENTE])
    abas['NC LAI-EIV'] = (['Aba', 'Linha Excel', 'Obra (como na planilha)', 'Obra (sistema)', 'Competência', 'Tipo', 'Assunto',
                           'No sistema?', 'Situação'], out)

    # Avanço físico ----------------------------------------------------------
    sis = {(r['obra'], r['competencia']): r for r in B['avanco']}
    out = []
    for d in ler(wb, 'Avanco_fisico'):
        o = obra(d['_aba'], d['Obra'])
        c = cp.competencia(d['Mês/Ano'])
        r = sis.get((o, c))
        p = cp.num(d['Avanço %'])
        st = NAO_MIGRADO if p is None else situacao(p, r['avanco_pct'] if r else None, r)
        out.append([d['_aba'], d['_linha'], d['Obra'], o or '', mes_txt(c), p, r['avanco_pct'] if r else None,
                    cp.num(d['Total m²']), r['area_total_m2'] if r else None, cp.num(d['Avanço m²']), r['area_mes_m2'] if r else None,
                    st, quem(r), 'Total m² agora é o cadastro da obra' if r and not igual(cp.num(d['Total m²']), r['area_total_m2']) else ''])
    abas['Avanço físico'] = (['Aba', 'Linha Excel', 'Obra (como na planilha)', 'Obra (sistema)', 'Competência', 'Avanço % planilha',
                              'Avanço % sistema', 'Total m² planilha', 'Área total sistema', 'Avanço m² planilha', 'Área do mês sistema',
                              'Situação', 'Alterado por', 'Observação'], out)

    # Ruído ------------------------------------------------------------------
    sis = defaultdict(list)
    for r in B['ruido']:
        sis[(r['obra'], r['competencia'])].append(r)
    out = []
    for d in ler(wb, 'Ruido'):
        o = obra(d['_aba'], d['Obra'])
        c = cp.competencia(d['Mês/Ano'])
        db = cp.num(d['dB'])
        cand = sis.get((o, c), [])
        r = next((x for x in cand if igual(x['db'], db, 0.0001)), None)
        h = d['Horário'].strftime('%H:%M') if isinstance(d['Horário'], dt.time) else ('' if d['Horário'] is None else str(d['Horário']))
        if not db:
            st, obs = NAO_MIGRADO, 'dB 0 ou vazio = obra sem medição no mês (na planilha entrava na média como 0)'
        elif r:
            st, obs = IGUAL, ''
        else:
            r = cand[0] if cand else None
            st = ALTERADO if r and r.get('atualizado_por') not in (None, 'carga_planilha') else DIFERENTE
            obs = ''
        if isinstance(d['Horário'], dt.datetime):
            obs = (obs + '; ' if obs else '') + 'horário gravado como data na planilha: ficou sem horário'
        out.append([d['_aba'], d['_linha'], d['Obra'], o or '', mes_txt(c), d['Ponto'], h, db, r['db'] if r else None,
                    r['horario'] if r else None, st, quem(r), obs])
    abas['Ruído'] = (['Aba', 'Linha Excel', 'Obra (como na planilha)', 'Obra (sistema)', 'Competência', 'Ponto', 'Horário planilha',
                      'dB planilha', 'dB sistema', 'Horário sistema', 'Situação', 'Alterado por', 'Observação'], out)

    # Água e luz -------------------------------------------------------------
    sis = {(r['unidade'], r['competencia']): r for r in B['consumo'] if r['linha'] == 1}
    sis2 = {(r['unidade'], r['competencia']): r for r in B['consumo'] if r['linha'] == 2}   # 2ª linha de mar/25 (exceção)
    filhos = defaultdict(list)
    for u in B['ucs']:
        if u['pai']:
            filhos[u['pai']].append(u['nome'])
    plan = list(ler(wb, 'agua_luz'))
    vistos = defaultdict(list)
    for d in plan:
        vistos[(str(d['Obra']).strip(), cp.competencia(d['Mês/Ano']))].append(d['_linha'])
    out = []
    for d in plan:
        texto = '' if cp.vazio(d['Obra']) else str(d['Obra']).strip()
        o = cp.obra_oficial(texto)
        uc = o or texto
        depara[(d['_aba'], texto, '(linha sem obra/setor)' if not texto else uc + ('' if o else ' (unidade administrativa)'))] += 1
        c = cp.competencia(d['Mês/Ano'])
        luz, agua = cp.num(d['Luz (kwh)']), cp.num(d['Água (m³)'])
        pos = vistos[(str(d['Obra']).strip(), c)].index(d['_linha'])   # mesma chave de `vistos`
        r = (sis if pos == 0 else sis2).get((uc, c))
        obs = ''
        if luz is None and agua is None:
            if texto and r is not None and r['luz_kwh'] is None and r['agua_m3'] is None:
                st, obs = IGUAL, 'linha em branco na planilha, guardada em branco (NULL)'
            else:
                st, obs = NAO_MIGRADO, 'sem luz e sem água'
        elif not texto:
            st, obs = NAO_MIGRADO, 'linha com consumo e SEM obra/setor na coluna Obra: conferir na planilha'
        elif uc in filhos and not r:
            soma_l = sum((sis.get((f, c)) or {}).get('luz_kwh') or 0 for f in filhos[uc])
            soma_a = sum((sis.get((f, c)) or {}).get('agua_m3') or 0 for f in filhos[uc])
            st = AJUSTE if igual(luz or 0, soma_l, 0.02) and igual(agua or 0, soma_a, 0.02) else DIFERENTE
            obs = '%s é calculada = soma de %s (luz %s, água %s). Na planilha ela entrava junto e contava em dobro' % (
                uc, ', '.join(filhos[uc]), cp.sql(soma_l), cp.sql(soma_a))
            if st == DIFERENTE:
                if igual(luz or 0, soma_l, 0.02) and agua is None:
                    obs += '. A luz bate; a água da %s está VAZIA na planilha, mas as unidades têm água' % uc
                elif luz:
                    # o valor digitado bate com a soma das unidades de OUTRO mês? (mês trocado na digitação)
                    outros = sorted({k[1] for k in sis if k[0] in filhos[uc]} - {c})
                    for c2 in outros:
                        s2 = sum((sis.get((f, c2)) or {}).get('luz_kwh') or 0 for f in filhos[uc])
                        if igual(luz, s2, 0.0001):
                            obs += '. ATENÇÃO: o valor da planilha (%s) é a soma das unidades de %s — mês trocado na digitação' % (cp.sql(luz), mes_txt(c2))
                            break
            r = {'luz_kwh': soma_l, 'agua_m3': soma_a}
        else:
            linhas = vistos[(texto, c)]
            if pos > 0 and r is None:
                st, obs = NAO_MIGRADO, 'linha duplicada fora da exceção de mar/25: o sistema ficou com a da linha %d' % linhas[-1]
            else:
                if len(linhas) > 1:
                    obs = 'unidade com %d linhas na planilha neste mês: guardada como %dª linha' % (len(linhas), pos + 1)
                st_l, st_a = situacao(luz, r['luz_kwh'] if r else None, r), situacao(agua, r['agua_m3'] if r else None, r)
                st = IGUAL if st_l == st_a == IGUAL else (ALTERADO if ALTERADO in (st_l, st_a) else DIFERENTE)
        out.append([d['_aba'], d['_linha'], texto, uc, mes_txt(c), luz, r['luz_kwh'] if r else None, agua, r['agua_m3'] if r else None,
                    st, quem(r) if r and 'atualizado_por' in r else '', obs])
    abas['Água e luz'] = (['Aba', 'Linha Excel', 'Obra/setor (como na planilha)', 'Unidade (sistema)', 'Competência', 'Luz kWh planilha',
                           'Luz kWh sistema', 'Água m³ planilha', 'Água m³ sistema', 'Situação', 'Alterado por', 'Observação'], out)

    abas.update(conferir_listas(wb, B, obra))

    # ── grava o xlsx ──────────────────────────────────────────────────────
    w = Workbook()
    res = w.active
    res.title = 'Resumo'
    res.append(['Conferência planilha × sistema de Lançamentos Ambientais'])
    res.append(['Planilha: %s' % cp.PLANILHA])
    res.append(['Gerado em %s. Sistema = valores de hoje no banco (inclui o que foi alterado depois da carga).' % dt.datetime.now().strftime('%d/%m/%Y %H:%M')])
    res.append([])
    res.append(['Assunto', 'Linhas na planilha', IGUAL, AJUSTE, ALTERADO, NAO_MIGRADO, DIFERENTE])
    for nome, (cab, linhas) in abas.items():
        i = cab.index('Situação')
        cont = defaultdict(int)
        for l in linhas:
            cont[l[i]] += 1
        res.append([nome, len(linhas), cont[IGUAL], cont[AJUSTE], cont[ALTERADO], cont[NAO_MIGRADO], cont[DIFERENTE]])
    res.append([])
    res.append(['Como ler a coluna Situação'])
    for st, txt in [(IGUAL, 'o valor da planilha está no sistema exatamente igual'),
                    (AJUSTE, 'migrado com uma correção de regra (ver Observação): Sede calculada, meta vigente, área real, densidade'),
                    (ALTERADO, 'era igual na carga e alguém alterou no sistema depois (ver "Alterado por")'),
                    (NAO_MIGRADO, 'a linha não virou lançamento (grade vazia, dB 0, duplicada...) — motivo na Observação'),
                    (DIFERENTE, 'NÃO bate e não há explicação: conferir')]:
        res.append([st, txt])
        res.cell(res.max_row, 1).fill = PatternFill('solid', fgColor=CORES[st])
    res['A1'].font = Font(bold=True, size=14)
    for c in res[5]:
        c.font = Font(bold=True)
    res.column_dimensions['A'].width = 24
    for col in 'BCDEFG':
        res.column_dimensions[col].width = 20

    dp = w.create_sheet('De-Para nomes')
    dp.append(['Aba da planilha', 'Como está escrito na planilha', 'Nome no sistema', 'Linhas'])
    for (aba, texto, nome), n in sorted(depara.items()):
        dp.append([aba, texto, nome, n])
        if texto.strip() != nome and '(unidade administrativa)' not in nome:
            dp.cell(dp.max_row, 2).fill = PatternFill('solid', fgColor='F8F5DE')
    for nome, (cab, linhas) in abas.items():
        ws = w.create_sheet(nome[:31])
        ws.append(cab)
        i = cab.index('Situação')
        for l in linhas:
            ws.append(l)
            ws.cell(ws.max_row, i + 1).fill = PatternFill('solid', fgColor=CORES.get(l[i], 'FFFFFF'))
    for ws in w.worksheets[1:]:
        for c in ws[1]:
            c.font = Font(bold=True, color='FFFFFF')
            c.fill = PatternFill('solid', fgColor='006071')
            c.alignment = Alignment(wrap_text=True, vertical='center')
        ws.freeze_panes = 'C2'
        ws.auto_filter.ref = ws.dimensions
        for j, c in enumerate(ws[1], 1):
            larg = max(len(str(c.value or '')), *(len(str(x.value or '')) for x in list(ws.iter_cols(min_col=j, max_col=j, min_row=2, max_row=200))[0])) if ws.max_row > 1 else 12
            ws.column_dimensions[get_column_letter(j)].width = min(max(10, larg + 2), 60)
    w.save(saida)
    print('conferência -> %s' % saida)
    for row in res.iter_rows(min_row=5, max_row=5 + len(abas), values_only=True):
        print('  ', row)


if __name__ == '__main__':
    main()
