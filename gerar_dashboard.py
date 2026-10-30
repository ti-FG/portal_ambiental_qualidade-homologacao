# -*- coding: utf-8 -*-
"""
Dashboard FG Viva+ / Sustentabilidade — Gerador HTML
Gera dashboard_fg.html a partir do Excel de controle ambiental.
"""
import pandas as pd
import numpy as np
import json, os, re, warnings
from datetime import datetime, date
warnings.filterwarnings('ignore')

EXCEL_PATH = "U:\\Qualidade\\AMBIENTAL\\GESTÃO AMBIENTAL\\INDICADORES AMBIENTAIS\\TESTE POWER BI\\Controle Ambiental 2025_oficial_db.xlsx"
OUTPUT_PATH = "U:\\Qualidade\\AMBIENTAL\\GESTÃO AMBIENTAL\\INDICADORES AMBIENTAIS\\TESTE POWER BI\\dashboard_fg.html"
TODAY = datetime.today()

# ─── Utilitários ──────────────────────────────────────────────────────────────

OBRA_MAP = {
    'BLUE COAST': 'Blue Coast', 'BLUE VIEW': 'Blue View',
    'BOREAL': 'Boreal', 'SAPPHIRE': 'Sapphire', 'TITANIUM': 'Titanium',
    'IMPERIUM': 'Imperium', 'ICONIC': 'Iconic', 'HORIZON': 'Horizon',
    'SENNA': 'Senna', 'SCENARIUM': 'Scenarium', 'GRAND PLACE': 'Grand Place',
    'GARDEN PARK': 'Garden Park', 'GRANDPLACE': 'Grand Place',
}

def norm_obra(s):
    if pd.isna(s): return ''
    s = str(s).strip()
    key = re.sub(r'\s+', ' ', s.upper().split('–')[0].split('-')[0]).strip()
    return OBRA_MAP.get(key, s.strip())

def sf(v, d=0.0):
    try:
        f = float(v)
        return d if np.isnan(f) else f
    except: return d

def read(xl, name, **kw):
    if name is None:
        print("  Aviso: aba nao encontrada")
        return pd.DataFrame()
    try:
        return pd.read_excel(xl, sheet_name=name, **kw)
    except Exception as e:
        print(f"  Aviso: nao foi possivel ler '{name}': {e}")
        return pd.DataFrame()

class DateEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, (datetime, date)): return obj.strftime('%Y-%m-%d')
        if isinstance(obj, (np.integer,)): return int(obj)
        if isinstance(obj, (np.floating,)): return None if np.isnan(obj) else float(obj)
        if isinstance(obj, np.ndarray): return obj.tolist()
        return super().default(obj)

# ─── Leitura dos dados ────────────────────────────────────────────────────────
print("Lendo Excel...")
xl = pd.ExcelFile(EXCEL_PATH)
sheets = xl.sheet_names

import unicodedata as _ud

def _asc(s):
    """Remove accents for fuzzy matching."""
    return _ud.normalize('NFKD', str(s)).encode('ascii', 'ignore').decode().lower()

def sheet(name):
    """Find sheet by substring match (accent-insensitive)."""
    name_n = _asc(name)
    for s in sheets:
        try:
            if name_n in _asc(s): return s
        except: pass
    return None

# Conformidade
df_conf = read(xl, sheet('Conformidade'))
df_conf.columns = [c.strip() for c in df_conf.columns]
df_conf['obra'] = df_conf[df_conf.columns[0]].apply(norm_obra)
for c in df_conf.columns:
    if 'percentual' in c.lower(): df_conf['perc'] = pd.to_numeric(df_conf[c], errors='coerce')
    if c.upper() == 'META': df_conf['meta'] = pd.to_numeric(df_conf[c], errors='coerce')
    if 'mes' in _asc(c) and 'ano' in _asc(c): df_conf['mes_ano'] = pd.to_datetime(df_conf[c], errors='coerce')
df_conf = df_conf[df_conf['obra'].str.len() > 0].copy()

# Resíduos
df_res = read(xl, sheet('SOMA'))
df_res.columns = [c.strip() for c in df_res.columns]
for c in df_res.columns:
    if 'obra' in c.lower(): df_res['obra'] = df_res[c].apply(norm_obra)
    if c.lower() in ('kg',): df_res['kg'] = pd.to_numeric(df_res[c], errors='coerce')
    if 'kg/m' in c.lower(): df_res['kg_m2'] = pd.to_numeric(df_res[c], errors='coerce')
    if 'meta' in c.lower(): df_res['meta_kg'] = pd.to_numeric(df_res[c], errors='coerce')
    if 'classe' in c.lower(): df_res['classe'] = df_res[c].astype(str).str.strip()
    if 'res' in c.lower() and 'det' in c.lower(): df_res['tipo'] = df_res[c].astype(str).str.strip()
    if 'mes' in _asc(c) and 'ano' in _asc(c): df_res['mes_ano'] = pd.to_datetime(df_res[c], errors='coerce')
    if c.lower() == 'm³' or (c.lower().startswith('m') and '³' in c and 'kg' not in c.lower() and 'm²' not in c.lower()):
        if 'm3_waste' not in df_res.columns: df_res['m3_waste'] = pd.to_numeric(df_res[c], errors='coerce')
if 'obra' not in df_res.columns: df_res['obra'] = ''
df_res = df_res[df_res['obra'].str.len() > 0].copy()

# Ruído
df_ruido = read(xl, sheet('ido'))  # Ruído
df_ruido.columns = [c.strip() for c in df_ruido.columns]
for c in df_ruido.columns:
    if c.lower() == 'db' or c.lower() == 'db': df_ruido['db'] = pd.to_numeric(df_ruido[c], errors='coerce')
    if 'obra' in c.lower(): df_ruido['obra'] = df_ruido[c].apply(norm_obra)
    if 'mes' in _asc(c) and 'ano' in _asc(c): df_ruido['mes_ano'] = pd.to_datetime(df_ruido[c], errors='coerce')
    if 'ponto' in c.lower(): df_ruido['ponto'] = df_ruido[c].astype(str)
if 'db' not in df_ruido.columns:
    for c in df_ruido.columns:
        vals = pd.to_numeric(df_ruido[c], errors='coerce')
        if vals.between(30, 120).sum() > 3:
            df_ruido['db'] = vals; break
if 'obra' not in df_ruido.columns: df_ruido['obra'] = ''

# Água e Luz
df_al = read(xl, sheet('gua'))
df_al.columns = [c.strip() for c in df_al.columns]
for c in df_al.columns:
    if 'obra' in c.lower(): df_al['obra'] = df_al[c].apply(norm_obra)
    if 'kwh' in c.lower() and 'm' in c.lower(): df_al['kwh_m2'] = pd.to_numeric(df_al[c], errors='coerce')
    if 'agua' in c.lower() and '/m' in c.lower(): df_al['m3_m2'] = pd.to_numeric(df_al[c], errors='coerce')
    if 'kwh' in c.lower() and 'm²' not in c.lower() and 'kwh' == c.lower().replace('(', '').replace(')', '').replace('luz', '').strip():
        if 'kwh_total' not in df_al.columns: df_al['kwh_total'] = pd.to_numeric(df_al[c], errors='coerce')
    if 'mes' in _asc(c) and 'ano' in _asc(c): df_al['mes_ano'] = pd.to_datetime(df_al[c], errors='coerce')
if 'obra' not in df_al.columns: df_al['obra'] = ''

# Auditoria
df_aud = read(xl, sheet('Audit'))
df_aud.columns = [c.strip() for c in df_aud.columns]
for c in df_aud.columns:
    if 'obra' in c.lower(): df_aud['obra'] = df_aud[c].apply(norm_obra)
    if 'realiz' in c.lower(): df_aud['realizado'] = df_aud[c].astype(str).str.strip()
    if 'mes' in _asc(c) and 'ano' in _asc(c): df_aud['mes_ano'] = pd.to_datetime(df_aud[c], errors='coerce')
if 'obra' not in df_aud.columns: df_aud['obra'] = ''

# Plano de Ação
df_pa = read(xl, sheet('Plano'))
df_pa.columns = [c.strip() for c in df_pa.columns]
for c in df_pa.columns:
    if 'obra' in c.lower(): df_pa['obra'] = df_pa[c].apply(norm_obra)
    if 'status' in c.lower(): df_pa['status'] = df_pa[c].astype(str).str.strip()
    if 'prazo' in c.lower(): df_pa['prazo'] = pd.to_datetime(df_pa[c], errors='coerce')
    if 'indicador' in c.lower(): df_pa['indicador'] = df_pa[c].astype(str).str.strip()
    if 'responsavel' in c.lower() or 'respons' in c.lower(): df_pa['responsavel'] = df_pa[c].astype(str).str.strip()
if 'obra' not in df_pa.columns: df_pa['obra'] = ''

# NC LAI e EIV
df_nc = read(xl, sheet('NC LAI'))
df_nc.columns = [c.strip() for c in df_nc.columns]
for c in df_nc.columns:
    if 'obra' in c.lower(): df_nc['obra'] = df_nc[c].apply(norm_obra)
    if 'nc lai' in c.lower(): df_nc['nc_lai'] = df_nc[c].astype(str)
    if 'nc eiv' in c.lower(): df_nc['nc_eiv'] = df_nc[c].astype(str)
    if 'mes' in _asc(c) and 'ano' in _asc(c): df_nc['mes_ano'] = pd.to_datetime(df_nc[c], errors='coerce')
if 'obra' not in df_nc.columns: df_nc['obra'] = ''

# Reclamações
df_rec = read(xl, sheet('clama'))
df_rec.columns = [c.strip() for c in df_rec.columns]
for c in df_rec.columns:
    if 'obra' in c.lower(): df_rec['obra'] = df_rec[c].apply(norm_obra)
    if 'categoria' in c.lower(): df_rec['categoria'] = df_rec[c].astype(str).str.strip()
    if 'status' in c.lower(): df_rec['status'] = df_rec[c].astype(str).str.strip()
    if 'mes' in _asc(c) and 'ano' in _asc(c): df_rec['mes_ano'] = pd.to_datetime(df_rec[c], errors='coerce')
if 'obra' not in df_rec.columns: df_rec['obra'] = ''

# Emergências
df_emerg = read(xl, sheet('merg'))
df_emerg.columns = [c.strip() for c in df_emerg.columns]
for c in df_emerg.columns:
    if 'obra' in c.lower(): df_emerg['obra'] = df_emerg[c].apply(norm_obra)
    if 'categoria' in c.lower(): df_emerg['categoria'] = df_emerg[c].astype(str).str.strip()
    if 'mes' in _asc(c) and 'ano' in _asc(c): df_emerg['mes_ano'] = pd.to_datetime(df_emerg[c], errors='coerce')
if 'obra' not in df_emerg.columns: df_emerg['obra'] = ''

# SGA
df_sga = read(xl, sheet('SGA'))
df_sga.columns = [c.strip() for c in df_sga.columns]
for c in df_sga.columns:
    if 'pessoa' in c.lower(): df_sga['pessoas'] = pd.to_numeric(df_sga[c], errors='coerce')

# Impressões
df_imp = read(xl, sheet('press'))
df_imp.columns = [c.strip() for c in df_imp.columns]
for c in df_imp.columns:
    if 'total' in c.lower() and 'p' in c.lower(): df_imp['total_pag'] = pd.to_numeric(df_imp[c], errors='coerce')
    if '%' in c and 'p&b' in c.lower(): df_imp['pct_pb'] = pd.to_numeric(df_imp[c], errors='coerce')
    if '%' in c and 'color' in c.lower(): df_imp['pct_color'] = pd.to_numeric(df_imp[c], errors='coerce')
    if 'custo' in c.lower(): df_imp['custo'] = pd.to_numeric(df_imp[c], errors='coerce')
    if 'mes' in _asc(c) and 'ano' in _asc(c): df_imp['mes_ano'] = pd.to_datetime(df_imp[c], errors='coerce')

# Avanço físico
df_av = read(xl, sheet('avan'))
df_av.columns = [c.strip() for c in df_av.columns]
for c in df_av.columns:
    if 'obra' in c.lower(): df_av['obra'] = df_av[c].apply(norm_obra)
    if 'total' in c.lower() and 'm' in c.lower(): df_av['total_m2'] = pd.to_numeric(df_av[c], errors='coerce')
    if 'avan' in c.lower() and '%' in c.lower(): df_av['avanco_pct'] = pd.to_numeric(df_av[c], errors='coerce')
    if 'avan' in c.lower() and 'm' in c.lower() and '%' not in c: df_av['avanco_m2'] = pd.to_numeric(df_av[c], errors='coerce')
    if 'mes' in _asc(c) and 'ano' in _asc(c): df_av['mes_ano'] = pd.to_datetime(df_av[c], errors='coerce')
if 'obra' not in df_av.columns: df_av['obra'] = ''

print("Processando KPIs...")

# ─── KPIs por obra ────────────────────────────────────────────────────────────

OBRAS = sorted(set(
    list(df_conf['obra'].unique()) +
    list(df_res['obra'].unique()) +
    list(df_aud['obra'].unique())
))
OBRAS = [o for o in OBRAS if o and o not in ('', 'nan', 'NaN')]

def score_conformidade_obra(obra):
    d = df_conf[df_conf['obra'] == obra]
    if d.empty or 'perc' not in d.columns: return 0.5
    perc = d['perc'].mean()
    meta = d['meta'].mean() if 'meta' in d.columns else 0.9
    if pd.isna(meta) or meta == 0: meta = 0.9
    return min(float(perc / meta), 1.0) if not pd.isna(perc) else 0.5

def score_residuos_obra(obra):
    d = df_res[df_res['obra'] == obra]
    if d.empty or 'kg_m2' not in d.columns: return 0.5
    val = d['kg_m2'].replace(0, np.nan).mean()
    meta_val = d['meta_kg'].mean() if 'meta_kg' in d.columns else 50.0
    if pd.isna(meta_val) or meta_val == 0: meta_val = 50.0
    if pd.isna(val) or val == 0: return 0.5
    return min(float(meta_val / val), 1.0)

def score_ruido_obra(obra):
    d = df_ruido[df_ruido['obra'] == obra] if 'obra' in df_ruido.columns else pd.DataFrame()
    if 'db' not in d.columns or d.empty: return 0.5
    d2 = d.dropna(subset=['db'])
    if d2.empty: return 0.5
    return float((d2['db'] <= 70).sum() / len(d2))

def score_auditoria_obra(obra):
    d = df_aud[df_aud['obra'] == obra] if 'obra' in df_aud.columns else pd.DataFrame()
    if 'realizado' not in d.columns or d.empty: return 0.5
    total = len(d)
    sim = (d['realizado'].str.lower().isin(['sim', 'yes', 's', '1', 'true'])).sum()
    return float(sim / total) if total > 0 else 0.5

def score_plano_obra(obra):
    d = df_pa[df_pa['obra'] == obra] if 'obra' in df_pa.columns else pd.DataFrame()
    if 'status' not in d.columns or d.empty: return 0.5
    total = len(d)
    no_prazo = d['status'].str.lower().isin(['realizado', 'em andamento']).sum()
    return float(no_prazo / total) if total > 0 else 0.5

# Pesos do IPA definidos pela área (out/2026): conformidade 45%, auditoria 35%,
# resíduos 20% — fecha 100%. Ruído e plano de ação saíram da composição (os
# scores continuam calculados e publicados em D.ipa, só não pesam mais).
IPA_PESOS = {'conf': 0.45, 'aud': 0.35, 'res': 0.20}

def calc_ipa(obra):
    sc = score_conformidade_obra(obra)
    sr = score_residuos_obra(obra)
    sru = score_ruido_obra(obra)
    sa = score_auditoria_obra(obra)
    spa = score_plano_obra(obra)
    raw = sc * IPA_PESOS['conf'] + sa * IPA_PESOS['aud'] + sr * IPA_PESOS['res']
    return round(raw * 100, 1)

def ipa_class(v):
    if v >= 90: return 'excelente'
    if v >= 80: return 'bom'
    if v >= 70: return 'atencao'
    return 'critico'

# Compute all IPA values
ipa_data = {}
for o in OBRAS:
    val = calc_ipa(o)
    ipa_data[o] = {
        'ipa': val,
        'classe': ipa_class(val),
        'score_conf': round(score_conformidade_obra(o) * 100, 1),
        'score_res': round(score_residuos_obra(o) * 100, 1),
        'score_ruido': round(score_ruido_obra(o) * 100, 1),
        'score_aud': round(score_auditoria_obra(o) * 100, 1),
        'score_pa': round(score_plano_obra(o) * 100, 1),
    }

global_ipa = round(np.mean([v['ipa'] for v in ipa_data.values()]), 1) if ipa_data else 0

# ─── Conformidade por obra × mês (heatmap) ────────────────────────────────────
conf_heatmap = {}
if not df_conf.empty and 'mes_ano' in df_conf.columns:
    df_conf_valid = df_conf.dropna(subset=['mes_ano', 'perc'])
    meses = sorted(df_conf_valid['mes_ano'].dt.to_period('M').unique())[-12:]
    meses_str = [str(m) for m in meses]
    for o in OBRAS:
        row = {}
        d = df_conf_valid[df_conf_valid['obra'] == o]
        for m in meses:
            m_data = d[d['mes_ano'].dt.to_period('M') == m]
            perc_val = m_data['perc'].mean() if not m_data.empty else None
            meta_val = m_data['meta'].mean() if not m_data.empty and 'meta' in m_data.columns else None
            row[str(m)] = {
                'perc': round(float(perc_val) * 100, 1) if perc_val is not None and not pd.isna(perc_val) else None,
                'meta': round(float(meta_val) * 100, 1) if meta_val is not None and not pd.isna(meta_val) else None,
            }
        conf_heatmap[o] = row
    conf_meses = meses_str
else:
    conf_meses = []

# Conformidade atual (última média por obra)
conf_atual = {}
for o in OBRAS:
    d = df_conf[df_conf['obra'] == o]
    if not d.empty and 'perc' in d.columns:
        perc = d['perc'].mean()
        meta = d['meta'].mean() if 'meta' in d.columns else 0.9
        conf_atual[o] = {
            'perc': round(float(perc) * 100, 1) if not pd.isna(perc) else None,
            'meta': round(float(meta) * 100, 1) if not pd.isna(meta) else 90.0,
        }
    else:
        conf_atual[o] = {'perc': None, 'meta': 90.0}

# Obras em risco crítico (< 85%)
obras_risco = [o for o, v in conf_atual.items() if v['perc'] is not None and v['perc'] < 85]

# ─── IPA mensal por obra (histórico últimos 12 meses) ─────────────────────────
# Reaproveita as mesmas fórmulas/pesos do IPA agregado (calc_ipa acima), mas
# recalcula cada componente usando só os dados daquele mês. Quando a obra não
# tem nenhum registro de um componente naquele mês específico (ex.: nenhuma
# medição de ruído feita, nenhuma auditoria no mês), cai pro score agregado
# (histórico completo) daquele componente — evita que o IPA do mês "zere" por
# ausência pontual de dado, em vez de refletir uma piora real.
# Plano de ação não tem uma coluna de data histórica confiável na fonte (o
# campo 'status' é um retrato do estado atual, não uma linha do tempo), então
# esse componente fica constante (o mesmo score agregado) em todos os meses.
def _score_conf_mes(o, mp):
    if 'obra' not in df_conf.columns or 'mes_ano' not in df_conf.columns: return None
    d = df_conf[(df_conf['obra'] == o) & (df_conf['mes_ano'].dt.to_period('M') == mp)]
    if d.empty or 'perc' not in d.columns: return None
    p = d['perc'].mean(); mt = d['meta'].mean() if 'meta' in d.columns else 0.9
    if pd.isna(mt) or mt == 0: mt = 0.9
    return float(min(p/mt, 1.0)) if not pd.isna(p) else None

def _score_res_mes(o, mp):
    if 'obra' not in df_res.columns or 'mes_ano' not in df_res.columns: return None
    d = df_res[(df_res['obra'] == o) & (df_res['mes_ano'].dt.to_period('M') == mp)]
    if d.empty or 'kg_m2' not in d.columns: return None
    v = d['kg_m2'].replace(0, np.nan).mean()
    if pd.isna(v) or v == 0: return None
    mt = d['meta_kg'].mean() if 'meta_kg' in d.columns else 50.0
    if pd.isna(mt) or mt == 0: mt = 50.0
    return float(min(mt/v, 1.0))

def _score_ruido_mes(o, mp):
    if 'obra' not in df_ruido.columns or 'mes_ano' not in df_ruido.columns or 'db' not in df_ruido.columns:
        return None
    d = df_ruido[(df_ruido['obra'] == o) & (df_ruido['mes_ano'].dt.to_period('M') == mp)].dropna(subset=['db'])
    if d.empty: return None
    return float((d['db'] <= 70).sum() / len(d))

def _score_aud_mes(o, mp):
    if 'obra' not in df_aud.columns or 'mes_ano' not in df_aud.columns or 'realizado' not in df_aud.columns:
        return None
    d = df_aud[(df_aud['obra'] == o) & (df_aud['mes_ano'].dt.to_period('M') == mp)]
    if d.empty: return None
    t = len(d); s = d['realizado'].str.lower().isin(['sim', 'yes', 's', '1', 'true']).sum()
    return float(s/t) if t > 0 else None

_MESES_PT_IPA = ['','Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez']
def _mes_label_ipa(mp_str):
    try:
        y, mn = mp_str.split('-')
        return f"{_MESES_PT_IPA[int(mn)]}/{y[2:]}"
    except Exception:
        return mp_str

ipa_mensal = {}
for o in OBRAS:
    sc_agg, sr_agg, sru_agg = score_conformidade_obra(o), score_residuos_obra(o), score_ruido_obra(o)
    sa_agg, spa_agg = score_auditoria_obra(o), score_plano_obra(o)
    serie = []
    for mp_str in conf_meses:
        mp = pd.Period(mp_str, freq='M')
        sc  = _score_conf_mes(o, mp)
        sr  = _score_res_mes(o, mp)
        sru = _score_ruido_mes(o, mp)
        sa  = _score_aud_mes(o, mp)
        val = round((
            (sc if sc is not None else sc_agg) * IPA_PESOS['conf'] +
            (sa if sa is not None else sa_agg) * IPA_PESOS['aud'] +
            (sr if sr is not None else sr_agg) * IPA_PESOS['res']
        ) * 100, 1)
        serie.append({'mes': mp_str, 'mes_label': _mes_label_ipa(mp_str), 'ipa': val})
    ipa_mensal[o] = serie

# ─── Ruído mensal por obra (histórico últimos 12 meses) ───────────────────────
# Mesma janela de meses do ipa_mensal (conf_meses). Diferente do IPA, aqui não
# há fallback pro agregado no mês sem medição — é uma medição real (dB), não
# um score computável a partir de outro dado; o mês sem leitura simplesmente
# não entra na série (o histórico fica "compactado" nos meses com dado).
def _db_mes(o, mp):
    if 'obra' not in df_ruido.columns or 'mes_ano' not in df_ruido.columns or 'db' not in df_ruido.columns:
        return None
    d = df_ruido[(df_ruido['obra'] == o) & (df_ruido['mes_ano'].dt.to_period('M') == mp)].dropna(subset=['db'])
    if d.empty: return None
    return float(d['db'].mean())

ruido_mensal = {}
for o in OBRAS:
    serie = []
    for mp_str in conf_meses:
        mp = pd.Period(mp_str, freq='M')
        db = _db_mes(o, mp)
        if db is not None:
            serie.append({'mes': mp_str, 'mes_label': _mes_label_ipa(mp_str), 'db': round(db, 1)})
    ruido_mensal[o] = serie

# ─── Resíduos ─────────────────────────────────────────────────────────────────
res_por_obra = {}
for o in OBRAS:
    d = df_res[df_res['obra'] == o]
    if d.empty:
        res_por_obra[o] = {'kg_m2': None, 'meta': 50.0, 'total_kg': 0}
        continue
    kg_m2 = d['kg_m2'].replace(0, np.nan).mean() if 'kg_m2' in d.columns else None
    total_kg = d['kg'].sum() if 'kg' in d.columns else 0
    meta = d['meta_kg'].mean() if 'meta_kg' in d.columns else 50.0
    if pd.isna(meta) or meta == 0: meta = 50.0
    res_por_obra[o] = {
        'kg_m2': round(float(kg_m2), 2) if kg_m2 is not None and not pd.isna(kg_m2) else None,
        'meta': round(float(meta), 1),
        'total_kg': round(float(total_kg), 0),
    }

# Mix por classe CONAMA
res_classe = {}
if 'classe' in df_res.columns and 'kg' in df_res.columns:
    gc = df_res.groupby('classe')['kg'].sum()
    for k, v in gc.items():
        if str(k) not in ('nan', 'NaN', ''):
            res_classe[str(k)] = round(float(v), 0)

# Mix por tipo
res_tipo = {}
if 'tipo' in df_res.columns and 'kg' in df_res.columns:
    gt = df_res.groupby('tipo')['kg'].sum().nlargest(8)
    for k, v in gt.items():
        if str(k) not in ('nan', 'NaN', ''):
            res_tipo[str(k)] = round(float(v), 0)

total_res_kg = df_res['kg'].sum() if 'kg' in df_res.columns else 0
media_kg_m2 = df_res['kg_m2'].replace(0, np.nan).mean() if 'kg_m2' in df_res.columns else 0

# ─── Ruído ────────────────────────────────────────────────────────────────────
ruido_por_obra = {}
if 'db' in df_ruido.columns and 'obra' in df_ruido.columns:
    d_valid = df_ruido.dropna(subset=['db', 'obra'])
    for o in OBRAS:
        d = d_valid[d_valid['obra'] == o]
        if d.empty:
            ruido_por_obra[o] = {'media_db': None, 'pct_conforme': None, 'n': 0}
            continue
        media = d['db'].mean()
        pct = (d['db'] <= 70).mean() * 100
        ruido_por_obra[o] = {
            'media_db': round(float(media), 1),
            'pct_conforme': round(float(pct), 1),
            'n': len(d),
        }

# ─── Água e Luz ───────────────────────────────────────────────────────────────
al_por_obra = {}
if 'obra' in df_al.columns:
    for o in OBRAS:
        d = df_al[df_al['obra'] == o]
        kwh = d['kwh_m2'].replace(0, np.nan).mean() if 'kwh_m2' in d.columns else None
        m3 = d['m3_m2'].replace(0, np.nan).mean() if 'm3_m2' in d.columns else None
        al_por_obra[o] = {
            'kwh_m2': round(float(kwh), 3) if kwh is not None and not pd.isna(kwh) else None,
            'm3_m2': round(float(m3), 3) if m3 is not None and not pd.isna(m3) else None,
        }

# ─── Auditoria ────────────────────────────────────────────────────────────────
aud_por_obra = {}
if 'obra' in df_aud.columns and 'realizado' in df_aud.columns:
    for o in OBRAS:
        d = df_aud[df_aud['obra'] == o]
        total = len(d)
        sim = (d['realizado'].str.lower().isin(['sim', 'yes', 's'])).sum() if total > 0 else 0
        aud_por_obra[o] = {'total': total, 'realizadas': int(sim), 'pct': round(sim/total*100, 1) if total > 0 else 0}
total_aud = len(df_aud)
aud_sim = (df_aud['realizado'].str.lower().isin(['sim', 'yes', 's'])).sum() if 'realizado' in df_aud.columns else 0

# ─── Plano de Ação ────────────────────────────────────────────────────────────
pa_status = {}
if 'status' in df_pa.columns:
    for s in df_pa['status'].unique():
        if str(s) not in ('nan', 'NaN', ''):
            cnt = (df_pa['status'] == s).sum()
            pa_status[str(s).strip()] = int(cnt)

pa_atrasadas = 0
if 'prazo' in df_pa.columns and 'status' in df_pa.columns:
    atrasadas = df_pa[
        (df_pa['prazo'] < pd.Timestamp(TODAY)) &
        (~df_pa['status'].str.lower().isin(['realizado', 'nan']))
    ]
    pa_atrasadas = len(atrasadas)

pa_no_prazo_pct = 0
total_pa = len(df_pa)
if total_pa > 0 and 'status' in df_pa.columns:
    no_prazo = df_pa['status'].str.strip().str.lower().isin(['realizado', 'em andamento']).sum()
    pa_no_prazo_pct = round(float(no_prazo / total_pa) * 100, 1)

pa_por_obra = {}
if 'obra' in df_pa.columns and 'status' in df_pa.columns:
    for o in OBRAS:
        d = df_pa[df_pa['obra'] == o]
        pa_por_obra[o] = {}
        for s in ['Realizado', 'Em Andamento', 'Pendente', 'Atrasado']:
            cnt = d['status'].str.strip().str.lower().eq(s.lower()).sum()
            pa_por_obra[o][s] = int(cnt)

# ─── NC de resíduos: 1 ocorrência por obra-mês em que o kg/m² realizado
# ultrapassa a meta (Senna Tower 50 kg/m² · demais obras 36,6 kg/m²). Isso soma
# ao total de NCs registradas a pedido da área — não é uma NC LAI/EIV, mas
# passa a contar junto.
# Importante: aqui usa SOMA dos Kg/m² das linhas do mês (uma linha por tipo de
# resíduo), não média — cada linha já é kg daquele tipo ÷ área, então a
# intensidade real da obra no mês é a soma dessas frações (= kg total ÷ área).
# O resto do dashboard (KPI Resíduos, tabela, IPA) ainda usa média — problema
# em aberto, aguardando decisão — ver memória "kgm2-agregacao-media-vs-soma".
# Corrigido só aqui a pedido da usuária, pra não subestimar as NCs de resíduos.
def meta_obra(obra):
    return 50.0 if obra and 'senna' in str(obra).lower() else 36.6

def _residuos_kgm2_mes(o, mp):
    if 'obra' not in df_res.columns or 'mes_ano' not in df_res.columns or 'kg_m2' not in df_res.columns:
        return None
    d = df_res[(df_res['obra'] == o) & (df_res['mes_ano'].dt.to_period('M') == mp)]
    if d.empty: return None
    return float(d['kg_m2'].sum(skipna=True))

_meses_res_todos = sorted(df_res['mes_ano'].dropna().dt.to_period('M').astype(str).unique()) if 'mes_ano' in df_res.columns else []
res_nc_mensal = {}  # obra -> {mes_str: True/False (excedeu meta naquele mês)}
for o in OBRAS:
    res_nc_mensal[o] = {}
    for mp_str in _meses_res_todos:
        v = _residuos_kgm2_mes(o, pd.Period(mp_str, freq='M'))
        res_nc_mensal[o][mp_str] = (v is not None and v > meta_obra(o))
res_nc_por_obra = {o: sum(1 for excedeu in res_nc_mensal[o].values() if excedeu) for o in OBRAS}
res_nc_total = sum(res_nc_por_obra.values())

# ─── NC LAI e EIV ─────────────────────────────────────────────────────────────
nc_por_obra = {}
if 'obra' in df_nc.columns:
    for o in OBRAS:
        d = df_nc[df_nc['obra'] == o]
        nc_por_obra[o] = len(d) + res_nc_por_obra.get(o, 0)
total_ncs = len(df_nc) + res_nc_total

# Top NCs recorrentes (por tipo de NC)
top_ncs = []
if 'nc_lai' in df_nc.columns:
    ncs = df_nc['nc_lai'].value_counts().head(5)
    for nc, cnt in ncs.items():
        if str(nc) not in ('nan', '-', ''):
            top_ncs.append({'nc': str(nc), 'count': int(cnt)})

# ─── Reclamações ──────────────────────────────────────────────────────────────
rec_por_obra = {}
if 'obra' in df_rec.columns:
    for o in OBRAS:
        rec_por_obra[o] = int((df_rec['obra'] == o).sum())
total_rec = len(df_rec)

rec_categoria = {}
if 'categoria' in df_rec.columns:
    for k, v in df_rec['categoria'].value_counts().head(5).items():
        if str(k) not in ('nan', 'NaN', ''):
            rec_categoria[str(k)] = int(v)

# ─── NC e Reclamações mensal por obra (pra filtro de período reagir na aba Obra Individual) ─
# Diferente do ruído, aqui 0 é uma contagem real (nenhuma NC/reclamação naquele mês),
# não "sem dado" — por isso todo mês do período entra na série, mesmo com count=0.
def _count_mes(df, o, mp):
    if 'obra' not in df.columns or 'mes_ano' not in df.columns: return None
    return int(len(df[(df['obra'] == o) & (df['mes_ano'].dt.to_period('M') == mp)]))

nc_mensal = {}
rec_mensal = {}
for o in OBRAS:
    serie_nc, serie_rec = [], []
    for mp_str in conf_meses:
        mp = pd.Period(mp_str, freq='M')
        cnt_nc = _count_mes(df_nc, o, mp)
        cnt_rec = _count_mes(df_rec, o, mp)
        if cnt_nc is not None:
            if res_nc_mensal.get(o, {}).get(mp_str): cnt_nc += 1
            serie_nc.append({'mes': mp_str, 'mes_label': _mes_label_ipa(mp_str), 'count': cnt_nc})
        if cnt_rec is not None:
            serie_rec.append({'mes': mp_str, 'mes_label': _mes_label_ipa(mp_str), 'count': cnt_rec})
    nc_mensal[o] = serie_nc
    rec_mensal[o] = serie_rec

# ─── Emergências ──────────────────────────────────────────────────────────────
emerg_por_obra = {}
if 'obra' in df_emerg.columns:
    for o in OBRAS:
        emerg_por_obra[o] = int((df_emerg['obra'] == o).sum())
total_emerg = len(df_emerg)

# ─── SGA ──────────────────────────────────────────────────────────────────────
sga_total = 0
if 'pessoas' in df_sga.columns:
    sga_total = int(df_sga['pessoas'].sum())

# ─── Impressões ───────────────────────────────────────────────────────────────
imp_total = 0
imp_custo = 0.0
imp_pct_pb = 0.0
imp_historico = []
if 'total_pag' in df_imp.columns:
    imp_total = int(df_imp['total_pag'].sum())
    imp_custo = round(float(df_imp['custo'].sum() if 'custo' in df_imp.columns else 0), 2)
    if 'pct_pb' in df_imp.columns:
        imp_pct_pb = round(float(df_imp['pct_pb'].mean() * 100), 1)
    if 'mes_ano' in df_imp.columns:
        df_imp_valid = df_imp.dropna(subset=['mes_ano', 'total_pag'])
        for _, r in df_imp_valid.tail(12).iterrows():
            imp_historico.append({
                'mes': r['mes_ano'].strftime('%m/%Y'),
                'total': int(r['total_pag']),
                'custo': round(float(r['custo']), 2) if 'custo' in df_imp.columns else 0,
            })

# ─── Conformidade evolução (linha 12 meses global) ────────────────────────────
conf_evolucao = []
if not df_conf.empty and 'mes_ano' in df_conf.columns and 'perc' in df_conf.columns:
    df_cg = df_conf.dropna(subset=['mes_ano', 'perc'])
    gm = df_cg.groupby(df_cg['mes_ano'].dt.to_period('M'))['perc'].mean()
    for p, v in gm.tail(12).items():
        conf_evolucao.append({'mes': str(p), 'perc': round(float(v) * 100, 1)})

# Conformidade por obra para barras
conf_barras = []
for o in OBRAS:
    d = df_conf[df_conf['obra'] == o]
    if not d.empty and 'perc' in d.columns:
        perc = d['perc'].mean()
        meta = d['meta'].mean() if 'meta' in d.columns else 0.9
        if not pd.isna(perc):
            conf_barras.append({
                'obra': o,
                'perc': round(float(perc) * 100, 1),
                'meta': round(float(meta if not pd.isna(meta) else 0.9) * 100, 1),
            })
conf_barras.sort(key=lambda x: x['perc'], reverse=True)

# Ruído por obra para barras
ruido_barras = [
    {'obra': o, 'db': v['media_db'], 'pct': v['pct_conforme']}
    for o, v in ruido_por_obra.items()
    if v['media_db'] is not None
]
ruido_barras.sort(key=lambda x: x['db'] or 0, reverse=True)

# Resíduos por obra para barras
res_barras = [
    {'obra': o, 'kg_m2': v['kg_m2'], 'meta': v['meta']}
    for o, v in res_por_obra.items()
    if v['kg_m2'] is not None
]
res_barras.sort(key=lambda x: x['kg_m2'] or 0, reverse=True)

# ─── Compilar JSON de dados ────────────────────────────────────────────────────
DATA = {
    'obras': OBRAS,
    'global': {
        'ipa': global_ipa,
        'ipa_classe': ipa_class(global_ipa),
        'conf_media': round(np.mean([v['perc'] for v in conf_atual.values() if v['perc'] is not None]), 1) if conf_atual else 0,
        'total_ncs': total_ncs,
        'media_kg_m2': round(float(media_kg_m2), 2) if not pd.isna(media_kg_m2) else 0,
        'total_rec': total_rec,
        'pa_no_prazo_pct': pa_no_prazo_pct,
        'total_emerg': total_emerg,
        'sga_pessoas': sga_total,
        'total_auds': total_aud,
        'aud_realizadas': int(aud_sim),
    },
    'ipa': ipa_data,
    'ipa_mensal': ipa_mensal,
    'ruido_mensal': ruido_mensal,
    'nc_mensal': nc_mensal,
    'rec_mensal': rec_mensal,
    'conf_heatmap': conf_heatmap,
    'conf_meses': conf_meses,
    'conf_atual': conf_atual,
    'conf_barras': conf_barras,
    'conf_evolucao': conf_evolucao,
    'obras_risco': obras_risco,
    'top_ncs': top_ncs,
    'nc_por_obra': nc_por_obra,
    'res_por_obra': res_por_obra,
    'res_classe': res_classe,
    'res_tipo': res_tipo,
    'res_barras': res_barras,
    'total_res_kg': round(float(total_res_kg), 0),
    'ruido_por_obra': ruido_por_obra,
    'ruido_barras': ruido_barras,
    'al_por_obra': al_por_obra,
    'aud_por_obra': aud_por_obra,
    'pa_status': pa_status,
    'pa_atrasadas': pa_atrasadas,
    'pa_por_obra': pa_por_obra,
    'rec_por_obra': rec_por_obra,
    'rec_categoria': rec_categoria,
    'emerg_por_obra': emerg_por_obra,
    'imp_total': imp_total,
    'imp_custo': imp_custo,
    'imp_pct_pb': imp_pct_pb,
    'imp_historico': imp_historico,
}

DATA_JSON = json.dumps(DATA, cls=DateEncoder, ensure_ascii=False, indent=None)

print(f"  Obras encontradas: {OBRAS}")
print(f"  IPA global: {global_ipa}")

# ─── HTML Template ────────────────────────────────────────────────────────────
print("Gerando HTML...")

MESES_LABEL = {
    '01': 'Jan', '02': 'Fev', '03': 'Mar', '04': 'Abr',
    '05': 'Mai', '06': 'Jun', '07': 'Jul', '08': 'Ago',
    '09': 'Set', '10': 'Out', '11': 'Nov', '12': 'Dez',
}

HTML = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard Indicadores Ambientais — FG Viva+ / Sustentabilidade</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@3.8.0/dist/tabler-icons.min.css">
<style>
  :root {
    --fg1: #173404; --fg2: #27500A; --fg3: #3B6D11; --fg4: #639922;
    --fg5: #97C459; --fg6: #C0DD97; --fg7: #EAF3DE;
    /* Header customizado */
    --header-bg: #00484C;          /* Teal escuro para o cabeçalho */
    --header-accent: #006B70;      /* Variação clara para borda inferior */
    --tab-active-bg: #D9D9D9;      /* Cinza claro p/ aba selecionada */
    --tab-active-fg: #00484C;      /* Texto da aba selecionada */
    --az1: #042C53; --az2: #0C447C; --az3: #185FA5; --az4: #378ADD;
    --az5: #85B7EB; --az6: #B5D4F4; --az7: #E6F1FB;
    --am1: #412402; --am2: #633806; --am3: #BA7517; --am4: #EF9F27;
    --am5: #FAC775; --am6: #FAEEDA;
    --vr1: #6B3A3A; --vr2: #8B5252; --vr3: #A86D6D; --vr4: #CE8686;
    --vr5: #E0A8A8; --vr6: #EDC8C8; --vr7: #FAEAEA;
    --gz1: #2C2C2A; --gz2: #444441; --gz3: #5F5E5A; --gz4: #888780;
    --gz5: #B4B2A9; --gz6: #D3D1C7; --gz7: #F1EFE8;
    --bg: #F8F7F3; --white: #FFFFFF;
    --radius: 12px; --gap: 14px;
  }
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Inter, system-ui, sans-serif; background: var(--bg); color: var(--gz1); font-size: 13px; line-height: 1.4; }
  h1 { font-size: 22px; font-weight: 500; }
  h2 { font-size: 18px; font-weight: 500; }
  h3 { font-size: 13px; font-weight: 500; text-transform: none; }

  /* Top nav */
  .top-nav { background: var(--header-bg); color: #fff; display: flex; align-items: center; gap: 32px; padding: 0 24px; height: 52px; position: sticky; top: 0; z-index: 100; border-bottom: 2px solid var(--header-accent); }
  .brand { display: flex; align-items: center; gap: 10px; }
  .brand-logo { height: 42px; display: flex; align-items: center; justify-content: center; }
  .brand-logo img { height: 35px; width: auto; display: block; }
  .brand-logo i { font-size: 20px; padding: 4px 6px; background: var(--header-accent); border-radius: 8px; }
  .brand-name { font-size: 20px; font-weight: 500; letter-spacing: 0.4px; color: #D9D9D9; }
  .brand-sub { font-size: 12px; color: rgba(217,217,217,0.6); }
  .nav-tabs { display: flex; gap: 2px; margin-left: auto; }
  .nav-tab { background: transparent; border: none; color: rgba(217,217,217,0.6); padding: 8px 18px; border-radius: 6px; cursor: pointer; font-size: 15px; font-weight: 400; display: flex; align-items: center; gap: 6px; transition: all 0.15s; }
  .nav-tab:hover { background: rgba(255,255,255,0.12); color: #D9D9D9; }
  .nav-tab.active { background: var(--tab-active-bg); color: var(--tab-active-fg); font-weight: 500; }
  .nav-right { margin-left: 24px; font-size: 11px; color: rgba(255,255,255,0.7); }

  /* Filter bar — colada abaixo do header (52px), acompanha o scroll junto com ele */
  .filter-bar { background: var(--white); border-bottom: 1px solid var(--gz6); padding: 8px 24px; display: flex; align-items: center; gap: 16px; flex-wrap: wrap; position: sticky; top: 52px; z-index: 90; transition: box-shadow 0.15s; }
  .filter-bar.is-stuck { box-shadow: 0 4px 10px -4px rgba(0,0,0,0.18); }
  .filter-group { display: flex; align-items: center; gap: 6px; font-size: 11px; color: var(--gz3); }
  .filter-group label { font-weight: 500; }
  .filter-group select { border: 1px solid var(--gz6); border-radius: 6px; padding: 4px 8px; font-size: 12px; color: var(--gz1); background: var(--bg); cursor: pointer; }
  .filter-group select:focus { outline: 2px solid var(--fg4); }

  /* Page layout */
  .page { display: none; padding: 20px 24px; max-width: 1600px; margin: 0 auto; }
  .page.active { display: block; }

  /* KPI row */
  .kpi-row { display: grid; gap: var(--gap); margin-bottom: var(--gap); }
  .kpi-row.cols-6 { grid-template-columns: repeat(6, 1fr); }
  .kpi-row.cols-5 { grid-template-columns: repeat(5, 1fr); }
  .kpi-row.cols-4 { grid-template-columns: repeat(4, 1fr); }
  .kpi-row.cols-3 { grid-template-columns: repeat(3, 1fr); }
  .kpi-row.cols-2 { grid-template-columns: 1fr 1fr; }

  /* KPI card */
  .kpi { background: var(--white); border: 0.5px solid var(--gz6); border-radius: var(--radius); padding: 14px; border-left: 3px solid var(--fg4); position: relative; overflow: hidden; }
  .kpi.green { border-left-color: var(--fg4); }
  .kpi.blue { border-left-color: var(--az3); }
  .kpi.amber { border-left-color: var(--am4); }
  .kpi.red { border-left-color: var(--vr4); }
  .kpi.gray { border-left-color: var(--gz4); }
  .kpi-label { font-size: 10px; color: var(--gz3); font-weight: 500; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px; display: flex; align-items: center; gap: 5px; }
  /* Título do KPI (caixa alta acima do valor) +10%, mesmo escopo por página
     do .kpi-sub — Dicionário fica de fora (não tem KPI). */
  #page-consolidado .kpi-label, #page-obra .kpi-label, #page-residuos .kpi-label { font-size: 11px; }
  .kpi-value { font-size: 26px; font-weight: 500; color: var(--gz1); line-height: 1; }
  .kpi-value.large { font-size: 32px; }
  .kpi-sub { font-size: 10px; color: var(--gz4); margin-top: 4px; }
  /* Subtítulo dos KPIs (texto abaixo do valor) +10%, só em Consolidado, Obra
     individual e Resíduos — mesmo escopo por página das demais mudanças de
     fonte. Os KPIs são gerados por JS (innerHTML) dentro do container de
     cada aba, então o seletor por #page-id alcança tanto o HTML estático
     quanto o que é montado em runtime. */
  #page-consolidado .kpi-sub, #page-obra .kpi-sub, #page-residuos .kpi-sub { font-size: 11px; }
  .kpi-delta { display: inline-flex; align-items: center; gap: 3px; font-size: 11px; padding: 2px 6px; border-radius: 4px; margin-top: 4px; }
  .kpi-delta.up { background: var(--fg7); color: var(--fg2); }
  .kpi-delta.down { background: var(--vr7); color: var(--vr2); }
  .kpi-delta.neutral { background: var(--gz7); color: var(--gz3); }
  .kpi-icon { position: absolute; right: 12px; top: 12px; font-size: 22px; opacity: 0.15; }

  /* Cards grid */
  .cards-grid { display: grid; gap: var(--gap); margin-bottom: var(--gap); }
  .cards-grid.cols-2-1 { grid-template-columns: 3fr 2fr; }
  .cards-grid.cols-1-2 { grid-template-columns: 2fr 3fr; }
  /* Plano de ação (rosca compacta) + Reclamações por obra lado a lado */
  .cards-grid.cols-pa-rec { grid-template-columns: 1fr 1.4fr; }
  /* Resíduos: demais obras (mais barras) + Senna Tower ao lado */
  .cards-grid.cols-res { grid-template-columns: 1.6fr 1fr; }
  @media (max-width: 1100px) { .cards-grid.cols-res { grid-template-columns: 1fr; } }
  @media (max-width: 1100px) { .cards-grid.cols-pa-rec { grid-template-columns: 1fr; } }
  .cards-grid.cols-3 { grid-template-columns: 1fr 1fr 1fr; }
  /* Classe A e Classe B: um painel por tipo de resíduo, escala própria em cada um */
  .ct-classe { font-size: 12px; font-weight: 600; color: var(--gz2); margin: 6px 0 8px; display: flex; align-items: center; gap: 8px; }
  .ct-classe .ct-dot { width: 10px; height: 10px; border-radius: 2px; display: inline-block; }
  .ct-classe .ct-tot { font-weight: 400; color: var(--gz4); }
  .ct-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 12px; margin-bottom: 14px; }
  .ct-painel { border: 1px solid var(--gz6); border-radius: 8px; padding: 8px 10px 4px; min-width: 0; }
  .ct-painel-titulo { font-size: 12px; font-weight: 600; color: var(--gz1); display: flex; justify-content: space-between; gap: 6px; margin-bottom: 2px; }
  .ct-painel-titulo span:last-child { font-weight: 400; color: var(--gz4); white-space: nowrap; }
  @media (max-width: 1100px) { .ct-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); } }
  @media (max-width: 700px)  { .ct-grid { grid-template-columns: 1fr; } }
  .cards-grid.cols-2 { grid-template-columns: 1fr 1fr; }

  /* Card */
  .card { background: var(--white); border: 0.5px solid var(--gz6); border-radius: var(--radius); padding: 16px; }
  /* Variante mais compacta — ~10% menor que o padrão */
  .card.compact { padding: 12px; }
  .card.compact .card-header { margin-bottom: 8px; padding-bottom: 7px; }
  .card.compact .nc-item { padding: 5px 0; }
  .card-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; padding-bottom: 10px; border-bottom: 1px solid var(--gz7); }
  .card-title { font-size: 14px; font-weight: 500; color: var(--gz1); display: flex; align-items: center; gap: 6px; }
  .card-badge { font-size: 12px; background: rgba(0, 72, 76, 0.08); color: #00484C; padding: 2px 8px; border-radius: 10px; }
  /* Títulos/subtítulos dos cards +10% só em Consolidado, Obra individual e
     Resíduos — escopado por página para não alterar o cabeçalho do card
     "Dicionário de dados e termos", que reaproveita .card-title. */
  #page-consolidado .card-title, #page-obra .card-title, #page-residuos .card-title { font-size: 15.4px; }
  #page-consolidado .card-badge, #page-obra .card-badge, #page-residuos .card-badge { font-size: 13.2px; }
  .card-body { }
  canvas { max-height: 280px; }

  /* Heatmap table */
  .heatmap-wrap { overflow-x: auto; }
  table.heatmap { width: 100%; border-collapse: collapse; font-size: 11px; }
  table.heatmap th { background: var(--gz7); padding: 6px 8px; text-align: center; font-weight: 500; white-space: nowrap; border: 1px solid var(--gz6); }
  table.heatmap th.obra-col { text-align: left; min-width: 110px; }
  table.heatmap td { padding: 5px 8px; text-align: center; border: 1px solid var(--gz6); font-size: 11px; white-space: nowrap; }
  table.heatmap td.obra-name { text-align: left; font-weight: 500; background: var(--gz7) !important; color: var(--gz1); }
  .cell-ex { background: var(--fg7); color: var(--fg1); }
  .cell-ok { background: #dff0d8; color: #2a6a1a; }
  .cell-am { background: var(--am6); color: var(--am2); }
  .cell-vr { background: var(--vr7); color: var(--vr2); }
  .cell-nd { background: var(--gz7); color: var(--gz4); }

  .badge { display: inline-flex; align-items: center; padding: 2px 8px; border-radius: 10px; font-size: 12px; font-weight: 500; gap: 4px; }
  .badge-ex { background: var(--fg7); color: var(--fg2); }
  .badge-bom { background: var(--az7); color: var(--az2); }
  .badge-am { background: var(--am6); color: var(--am2); }
  .badge-cr { background: var(--vr7); color: var(--vr2); }

  /* Semáforo */
  .semaforo { display: inline-block; width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; }
  .sem-green { background: var(--fg4); }
  .sem-amber { background: var(--am4); }
  .sem-red { background: var(--vr4); }
  .sem-gray { background: var(--gz5); }

  /* Progress bar */
  .prog-wrap { background: var(--gz7); border-radius: 4px; height: 6px; overflow: hidden; margin-top: 4px; }
  .prog-bar { height: 100%; border-radius: 4px; transition: width 0.3s; }
  .prog-green { background: var(--fg4); }
  .prog-amber { background: var(--am4); }
  .prog-red { background: var(--vr4); }
  .prog-blue { background: var(--az3); }

  /* NC list */
  .nc-list { list-style: none; }
  .nc-item { display: flex; align-items: center; gap: 8px; padding: 6px 0; border-bottom: 1px solid var(--gz7); }
  .nc-item:last-child { border-bottom: none; }
  .nc-count { font-weight: 500; font-size: 16px; color: var(--vr3); width: 32px; text-align: center; flex-shrink: 0; }
  .nc-name { font-size: 11px; color: var(--gz1); flex: 1; }

  /* IPA gauge */
  .ipa-gauge { display: flex; align-items: center; justify-content: center; flex-direction: column; padding: 16px 0; }
  .ipa-number { font-size: 52px; font-weight: 500; line-height: 1; }
  .ipa-label { font-size: 12px; color: var(--gz3); margin-top: 4px; }
  .ipa-ex { color: var(--fg3); }
  .ipa-bom { color: var(--az3); }
  .ipa-am { color: var(--am3); }
  .ipa-cr { color: var(--vr3); }

  /* Tabela conformidade detalhe */
  .conf-table { width: 100%; border-collapse: collapse; font-size: 11px; }
  .conf-table th { background: var(--gz7); padding: 6px 10px; text-align: left; font-weight: 500; border-bottom: 2px solid var(--gz6); }
  .conf-table td { padding: 6px 10px; border-bottom: 1px solid var(--gz7); }
  .conf-table tr:hover td { background: var(--gz7); }

  /* Divider */
  .section-title { font-size: 11px; font-weight: 500; color: var(--gz3); text-transform: uppercase; letter-spacing: 0.5px; margin: 18px 0 10px; display: flex; align-items: center; gap: 8px; }
  .section-title::after { content: ''; flex: 1; height: 1px; background: var(--gz6); }

  /* Obra selector */
  .obra-selector { background: var(--white); border: 0.5px solid var(--gz6); border-radius: var(--radius); padding: 14px; margin-bottom: var(--gap); display: flex; align-items: center; gap: 16px; }
  .obra-selector label { font-weight: 500; font-size: 12px; color: var(--gz3); }
  .obra-selector select { border: 1.5px solid var(--fg4); border-radius: 8px; padding: 6px 12px; font-size: 14px; font-weight: 500; color: var(--gz1); background: var(--fg7); cursor: pointer; }

  /* Dicionário de dados e termos */
  .dic-intro { font-size: 13.2px; color: var(--gz3); line-height: 1.5; margin-bottom: 12px; max-width: 900px; }
  .dic-search { width: 100%; max-width: 420px; border: 1px solid var(--gz6); border-radius: 8px; padding: 7px 12px; font-size: 13.2px; color: var(--gz1); background: var(--bg); font-family: inherit; }
  .dic-search:focus { outline: 2px solid var(--fg4); }
  .dic-chips { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 12px; }
  .dic-chip { border: 1px solid var(--gz6); background: var(--white); color: var(--gz2); border-radius: 14px; padding: 4px 11px; font-size: 12.1px; cursor: pointer; font-family: inherit; transition: all 0.15s; }
  .dic-chip:hover { background: var(--fg7); border-color: var(--fg4); color: var(--fg1); }
  .dic-table { width: 100%; border-collapse: collapse; font-size: 12.65px; }
  .dic-table th { background: var(--gz7); padding: 7px 10px; text-align: left; font-weight: 500; color: var(--gz2); border-bottom: 2px solid var(--gz6); white-space: nowrap; }
  .dic-table td { padding: 8px 10px; border-bottom: 1px solid var(--gz7); vertical-align: top; line-height: 1.45; }
  .dic-table tr:last-child td { border-bottom: none; }
  .dic-table tr:hover td { background: var(--gz7); }
  .dic-table td.dic-nome { font-weight: 500; color: var(--gz1); white-space: nowrap; }
  .dic-table td.dic-onde { color: var(--gz3); white-space: nowrap; }
  .dic-code { font-family: Consolas, 'Courier New', monospace; font-size: 12.1px; background: var(--gz7); color: var(--gz1); padding: 1px 5px; border-radius: 4px; }
  .dic-nota { display: block; margin-top: 4px; font-size: 11.55px; color: var(--gz4); }
  .dic-alerta { color: var(--am3); }
  .dic-peso-wrap { display: flex; align-items: center; gap: 8px; }
  .dic-peso-bar { flex: 1; max-width: 160px; height: 7px; background: var(--gz7); border-radius: 4px; overflow: hidden; }
  .dic-peso-fill { height: 100%; background: var(--fg4); border-radius: 4px; }
  .dic-peso-num { font-weight: 500; font-size: 12.1px; color: var(--gz2); width: 34px; text-align: right; }
  .dic-vazio { padding: 14px; font-size: 13.2px; color: var(--gz4); text-align: center; }
  .dic-hidden { display: none !important; }

  /* Footer */
  .dash-footer { margin-top: 24px; padding: 12px 24px; background: var(--gz7); border-top: 1px solid var(--gz6); text-align: center; font-size: 10px; color: var(--gz4); }

  /* Responsivo simplificado */
  @media (max-width: 1100px) {
    .kpi-row.cols-6 { grid-template-columns: repeat(3, 1fr); }
    .cards-grid.cols-2-1 { grid-template-columns: 1fr; }
  }
  @media (max-width: 700px) {
    .kpi-row.cols-6, .kpi-row.cols-5, .kpi-row.cols-4, .kpi-row.cols-3 { grid-template-columns: repeat(2, 1fr); }
    .page { padding: 12px; }
    .top-nav { padding: 0 12px; gap: 16px; }
  }

  .empty-msg { text-align: center; color: var(--gz4); padding: 24px; font-size: 12px; }
  .mini-table { width: 100%; border-collapse: collapse; font-size: 11px; }
  .mini-table th { background: var(--gz7); padding: 5px 8px; font-weight: 500; text-align: left; }
  .mini-table td { padding: 5px 8px; border-bottom: 1px solid var(--gz7); }

  /* Tabela de Emergências Ambientais */
  /* Painel "Reclamações — assuntos e maiores incidências" */
  .rec-detalhe-wrap { max-height: 460px; overflow-y: auto; }
  .rec-table { width: 100%; border-collapse: collapse; font-size: 10.5px; }
  .rec-table thead th {
    background: var(--gz7); padding: 6px 8px; text-align: left; font-weight: 500;
    color: var(--gz2); position: sticky; top: 0; border-bottom: 2px solid var(--gz6); white-space: nowrap;
  }
  .rec-table td { padding: 5px 8px; border-bottom: 1px solid var(--gz7); vertical-align: top; color: var(--gz1); }
  .rec-table .rec-obra { white-space: nowrap; font-weight: 500; }
  /* Linha que inicia um novo bloco de obra ganha separador mais forte, para as
     reclamações da mesma obra lerem como um grupo só */
  .rec-table tr.rec-grupo td { border-top: 2px solid var(--gz6); }
  .rec-table tr.rec-grupo:first-child td { border-top: none; }
  .rec-tag {
    display: inline-block; padding: 1px 6px; border-radius: 10px;
    background: rgba(0, 107, 107, 0.10); color: #00484C; font-size: 10px; white-space: nowrap;
  }
  .emerg-wrap { max-height: 340px; overflow-y: auto; border: 1px solid var(--gz7); border-radius: 8px; }
  .emerg-table { width: 100%; border-collapse: collapse; font-size: 11px; }
  .emerg-table thead th {
    background: var(--gz7); padding: 8px 10px; text-align: left;
    font-weight: 500; color: var(--gz2); position: sticky; top: 0;
    border-bottom: 2px solid var(--gz6); white-space: nowrap;
  }
  .emerg-table td {
    padding: 7px 10px; border-bottom: 1px solid var(--gz7);
    vertical-align: top; color: var(--gz1);
  }
  .emerg-table tr:last-child td { border-bottom: none; }
  .emerg-table tr:hover td { background: var(--gz7); }
  .emerg-table .col-mes { white-space: nowrap; font-weight: 500; color: var(--gz2); }
  .emerg-table .col-fmc { white-space: nowrap; font-family: 'Consolas', monospace; color: var(--gz3); font-size: 10px; }
  .emerg-table .col-ocorr { max-width: 380px; }
  .emerg-table .col-status { max-width: 260px; }
  .emerg-table .obra-tag {
    display: inline-block; padding: 2px 7px; border-radius: 10px;
    background: var(--fg7); color: var(--fg2); font-size: 10px;
    font-weight: 500; white-space: nowrap;
  }
  .emerg-table .empty-row td { text-align: center; padding: 24px; color: var(--gz4); }

  /* Section header com filtros inline (Impressões) */
  .imp-section-header { display: flex; align-items: center; gap: 14px; margin: 18px 0 10px; flex-wrap: wrap; }
  .imp-section-title { font-size: 11px; font-weight: 500; color: var(--gz3); text-transform: uppercase; letter-spacing: 0.5px; display: inline-flex; align-items: center; gap: 6px; white-space: nowrap; }
  /* Só existe hoje dentro do Consolidado (blocos de energia/água/impressões);
     escopado do mesmo jeito para acompanhar as outras duas abas se um dia
     um bloco assim aparecer lá também. */
  #page-consolidado .imp-section-title, #page-obra .imp-section-title, #page-residuos .imp-section-title { font-size: 12.1px; }
  .imp-filters { display: inline-flex; align-items: center; gap: 8px; font-size: 11px; color: var(--gz3); background: var(--white); border: 1px solid var(--gz6); border-radius: 8px; padding: 4px 10px; }
  .imp-filters label { font-weight: 500; }
  .imp-filters select { border: 1px solid var(--gz6); border-radius: 5px; padding: 3px 8px; font-size: 11px; color: var(--gz1); background: var(--bg); cursor: pointer; }
  .imp-filters select:focus { outline: 2px solid var(--fg4); }
  .imp-filters .imp-clear { background: transparent; border: none; color: var(--gz4); cursor: pointer; padding: 2px 4px; border-radius: 4px; display: flex; align-items: center; font-size: 13px; }
  .imp-filters .imp-clear:hover { background: var(--gz7); color: var(--gz1); }
  .imp-section-line { flex: 1; height: 1px; background: var(--gz6); min-width: 30px; }
</style>
</head>
<body>

<!-- Top Navigation -->
<nav class="top-nav">
  <div class="brand">
    <div class="brand-logo"><i class="ti ti-leaf" style="color:#fff"></i></div>
    <div>
      <div class="brand-name">INDICADORES AMBIENTAIS</div>
      <div class="brand-sub">FG Viva+ / Sustentabilidade</div>
    </div>
  </div>
  <div class="nav-tabs">
    <button class="nav-tab active" onclick="showPage('consolidado')" id="tab-consolidado">
      <i class="ti ti-layout-dashboard"></i> Consolidado
    </button>
    <button class="nav-tab" onclick="showPage('obra')" id="tab-obra">
      <i class="ti ti-building-skyscraper"></i> Obra individual
    </button>
    <button class="nav-tab" onclick="showPage('residuos')" id="tab-residuos">
      <i class="ti ti-recycle"></i> Resíduos
    </button>
    <button class="nav-tab" onclick="showPage('dicionario')" id="tab-dicionario">
      <i class="ti ti-book"></i> Dicionário
    </button>
  </div>
  <div class="nav-right" id="last-update">Atualizado em __DATA_DATE__</div>
</nav>

<!-- Sentinela: 1px acima da filter-bar. Enquanto ela intersecta a viewport,
     a filter-bar ainda não colou no header. Ver initFilterBarSombra(). -->
<div id="filter-bar-sentinela" style="height:1px"></div>

<!-- Filter bar (shared) -->
<div class="filter-bar" id="filter-bar-consolidado">
  <i class="ti ti-filter" style="color:var(--gz4)"></i>
  <div class="filter-group">
    <label>Obras:</label>
    <select id="filter-obra-global" onchange="applyFilters()">
      <option value="">Todas as obras</option>
    </select>
  </div>
  <div class="filter-group">
    <label>Período:</label>
    <select id="filter-periodo" onchange="applyFilters()">
      <option value="all">Todo o período</option>
      <option value="6m">Últimos 6 meses</option>
      <option value="12m">Últimos 12 meses</option>
      <option value="2025">2025</option>
      <option value="2026">2026</option>
    </select>
  </div>
  <span style="margin-left:auto; font-size:10px; color:var(--gz4)">
    <i class="ti ti-info-circle"></i> Dados lidos diretamente do Excel
  </span>
</div>

<!-- ═══════════════════════════════════════════════════════════
     PAGE 1 — CONSOLIDADO
════════════════════════════════════════════════════════════ -->
<div id="page-consolidado" class="page active">

  <!-- 6 KPI cards — mesma ordem na aba Obra Individual (ver kpiCardsHtml) -->
  <div class="kpi-row cols-6" id="kpis-global"></div>

  <!-- Bloco principal -->
  <div class="cards-grid cols-2-1" style="margin-bottom:var(--gap)">
    <!-- Heatmap conformidade (esquerda) -->
    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="ti ti-file-certificate" style="color:var(--fg4)"></i> Conformidade LAI/EIV por obra × mês</div>
        <div class="card-badge" id="heatmap-badge">últimos 12 meses</div>
      </div>
      <div class="card-body heatmap-wrap" id="heatmap-container"></div>
    </div>
    <!-- Top NCs (direita) -->
    <div style="display:flex; flex-direction:column; gap:var(--gap); height:100%">
      <div class="card" style="flex:1 1 auto; display:flex; flex-direction:column">
        <div class="card-header">
          <div class="card-title"><i class="ti ti-list-check" style="color:var(--am4)"></i> Top não conformidades recorrentes</div>
          <span class="card-badge" id="top-ncs-badge">—</span>
        </div>
        <ul class="nc-list" id="top-ncs-list" style="flex:1; display:flex; flex-direction:column; justify-content:space-around"></ul>
      </div>
    </div>
  </div>

  <!-- Monitoramento de ruído por obra × mês -->
  <div class="card" style="margin-bottom:var(--gap)">
    <div class="card-header">
      <div class="card-title"><i class="ti ti-volume" style="color:var(--az3)"></i> Monitoramento de ruído — dB por obra × mês</div>
      <span class="card-badge" id="monitor-ruido-badge">Itajaí: 85 dB · BC: 60 dB</span>
    </div>
    <canvas id="chart-monitor-ruido" style="max-height:340px"></canvas>
  </div>

  <!-- Plano de ação Auditorias (rosca compacta, abertura por obra no tooltip)
       lado a lado com Reclamações de Vizinhos por Obra -->
  <div class="cards-grid cols-pa-rec" style="margin-bottom:var(--gap)">
    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="ti ti-clipboard-check" style="color:var(--fg4)"></i> Plano de ação Auditorias distribuição de status</div>
        <span class="card-badge" id="pa-status-badge">—</span>
      </div>
      <canvas id="chart-pa-status" style="max-height:210px"></canvas>
    </div>
    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="ti ti-message-report" style="color:var(--am4)"></i> Reclamações de Vizinhos por Obra</div>
        <span class="card-badge" id="rec-obra-badge">—</span>
      </div>
      <canvas id="chart-rec-obra" style="max-height:200px"></canvas>
    </div>
  </div>

  <!-- Detalhe das reclamações — "Assunto e ação por obra" primeiro, depois
       "Maiores incidências por assunto"; cada um na largura toda e mais alto
       para mostrar mais linhas sem rolagem -->
  <div class="card" style="margin-bottom:var(--gap)">
    <div class="card-header">
      <div class="card-title"><i class="ti ti-list-details" style="color:var(--fg4)"></i> Assunto e ação por obra</div>
      <span class="card-badge" id="rec-detalhe-badge">—</span>
    </div>
    <div id="rec-detalhe" class="rec-detalhe-wrap"></div>
  </div>
  <div class="card" style="margin-bottom:var(--gap)">
    <div class="card-header">
      <div class="card-title"><i class="ti ti-chart-bar" style="color:var(--fg4)"></i> Maiores incidências por assunto</div>
      <span class="card-badge" id="rec-rank-badge">—</span>
    </div>
    <!-- Altura acompanha o número de assuntos (ver renderRecDetalhe) -->
    <div id="rec-rank-wrap" style="position:relative; height:320px">
      <canvas id="chart-rec-rank"></canvas>
    </div>
  </div>

  <!-- Emergências Ambientais (tabela) -->
  <div class="card" style="margin-bottom:var(--gap)">
    <div class="card-header">
      <div class="card-title"><i class="ti ti-alert-octagon" style="color:var(--vr4)"></i> Emergências ambientais</div>
      <span class="card-badge" id="emerg-table-badge">—</span>
    </div>
    <div class="emerg-wrap">
      <table class="emerg-table" id="emerg-table-content"></table>
    </div>
  </div>

  <!-- Resíduos gerados: demais obras (meta 36,6) e Senna Tower (meta 50) lado a
       lado — réguas diferentes, por isso em gráficos separados -->
  <div class="cards-grid cols-res" style="margin-bottom:var(--gap)">
    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="ti ti-trash" style="color:var(--fg4)"></i> <span id="res-barras-titulo">Resíduos Gerados kg/m² por obra</span></div>
        <span class="card-badge" id="res-barras-badge">—</span>
      </div>
      <canvas id="chart-res-barras" style="max-height:320px"></canvas>
    </div>
    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="ti ti-building-skyscraper" style="color:var(--fg4)"></i> Resíduos Gerados kg/m² — Senna Tower</div>
        <span style="display:inline-flex; gap:8px; align-items:center">
          <span class="card-badge" id="res-senna-acum" hidden></span>
          <span class="card-badge" id="res-senna-badge">—</span>
        </span>
      </div>
      <canvas id="chart-res-senna" style="max-height:320px"></canvas>
    </div>
  </div>

  <!-- Energia + Água + Impressões. O bloco aparece também na aba Obra Individual:
       é o MESMO nó, movido entre os slots por showPage() — mantém os IDs, os filtros
       de setor/período e o estado dos gráficos, sem duplicar código. -->
  <div id="slot-consumo-consolidado">
  <div id="bloco-consumo">
  <!-- Consumo de Energia (kWh + intensidade kWh/m²) -->
  <div class="imp-section-header">
    <span class="imp-section-title"><i class="ti ti-bolt"></i> Consumo de energia</span>
    <span class="imp-filters">
      <i class="ti ti-filter" style="color:var(--gz4); font-size:12px"></i>
      <label>Setor:</label>
      <select id="filter-energia-setor" onchange="renderConsumoEnergiaAll()">
        <option value="">Todos os setores</option>
      </select>
      <label>Período:</label>
      <select id="filter-energia-periodo" onchange="renderConsumoEnergiaAll()">
        <option value="6m">Últimos 6 meses</option>
        <option value="12m" selected>Últimos 12 meses</option>
        <option value="all">Todo o período</option>
      </select>
    </span>
    <span class="imp-section-line"></span>
  </div>
  <div class="cards-grid cols-2" style="margin-bottom:var(--gap)">
    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="ti ti-bolt" style="color:var(--am4)"></i> Consumo de Energia Obras (kWh)</div>
        <span class="card-badge" id="energia-badge">—</span>
      </div>
      <canvas id="chart-consumo-energia" style="max-height:260px"></canvas>
    </div>
    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="ti ti-ruler-measure" style="color:var(--am4)"></i> Intensidade Energética Obras (kWh/m²)</div>
        <span class="card-badge" id="energia-m2-badge">—</span>
      </div>
      <canvas id="chart-consumo-energia-m2" style="max-height:260px"></canvas>
    </div>
  </div>

  <!-- Consumo de Água (m³ + intensidade m³/m²) -->
  <div class="imp-section-header">
    <span class="imp-section-title"><i class="ti ti-droplet"></i> Consumo de água</span>
    <span class="imp-filters">
      <i class="ti ti-filter" style="color:var(--gz4); font-size:12px"></i>
      <label>Setor:</label>
      <select id="filter-agua-setor" onchange="renderConsumoAguaAll()">
        <option value="">Todos os setores</option>
      </select>
      <label>Período:</label>
      <select id="filter-agua-periodo" onchange="renderConsumoAguaAll()">
        <option value="6m">Últimos 6 meses</option>
        <option value="12m" selected>Últimos 12 meses</option>
        <option value="all">Todo o período</option>
      </select>
    </span>
    <span class="imp-section-line"></span>
  </div>
  <div class="cards-grid cols-2" style="margin-bottom:var(--gap)">
    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="ti ti-droplet" style="color:var(--az4)"></i> Consumo de Água Obras (m³)</div>
        <span class="card-badge" id="agua-badge">—</span>
      </div>
      <canvas id="chart-consumo-agua" style="max-height:260px"></canvas>
    </div>
    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="ti ti-ruler-measure" style="color:var(--az4)"></i> Intensidade Hídrica Obras (m³/m²)</div>
        <span class="card-badge" id="agua-m2-badge">—</span>
      </div>
      <canvas id="chart-consumo-agua-m2" style="max-height:260px"></canvas>
    </div>
  </div>

  <!-- Impressões — histórico + setores (com filtros) -->
  <div class="imp-section-header">
    <span class="imp-section-title"><i class="ti ti-printer"></i> Impressões — histórico mensal e maiores setores</span>
    <span class="imp-filters">
      <i class="ti ti-filter" style="color:var(--gz4); font-size:12px"></i>
      <label>Setor:</label>
      <select id="filter-imp-setor" onchange="renderImpressoes()">
        <option value="">Todos os setores</option>
      </select>
      <label>Período:</label>
      <select id="filter-imp-periodo" onchange="renderImpressoes()">
        <option value="all">Todo o período</option>
        <option value="6m">Últimos 6 meses</option>
        <option value="12m">Últimos 12 meses</option>
      </select>
      <button class="imp-clear" onclick="clearImpFilters()" title="Limpar filtros"><i class="ti ti-x"></i></button>
    </span>
    <span class="imp-section-line"></span>
  </div>
  <div class="cards-grid cols-2-1" style="margin-bottom:var(--gap)">
    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="ti ti-chart-line" style="color:var(--az3)"></i> Histórico mensal de páginas impressas</div>
        <span class="card-badge" id="imp-media-badge">—</span>
      </div>
      <canvas id="chart-imp-historico" style="max-height:280px"></canvas>
    </div>
    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="ti ti-building" style="color:var(--az3)"></i> Maiores setores impressores</div>
        <span class="card-badge" id="imp-setores-badge">top 10</span>
      </div>
      <div id="imp-setores-table"></div>
    </div>
  </div>
  </div>
  </div>

  <!-- Eventos Ambientais (Dim_Acoes_SGA) — pessoas por ano. Último gráfico da aba -->
  <div class="card" style="margin-bottom:var(--gap)">
    <div class="card-header">
      <div class="card-title"><i class="ti ti-users" style="color:var(--fg4)"></i> Eventos Ambientais</div>
      <span class="card-badge" id="sga-badge">acumulado desde 2021</span>
    </div>
    <canvas id="chart-sga-anual" style="max-height:240px"></canvas>
  </div>

</div>

<!-- ═══════════════════════════════════════════════════════════
     PAGE 2 — OBRA INDIVIDUAL
════════════════════════════════════════════════════════════ -->
<div id="page-obra" class="page">
  <!-- Mesmos 6 KPIs, na mesma ordem, da aba Consolidado (kpiCardsHtml) -->
  <div class="kpi-row cols-6" id="kpis-obra" style="margin-bottom:var(--gap)"></div>

  <!-- Conformidade + IPA -->
  <div class="cards-grid cols-2" style="margin-bottom:var(--gap)">
    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="ti ti-file-certificate" style="color:var(--fg4)"></i> Conformidade LAI/EIV — histórico</div>
      </div>
      <div id="obra-conf-table"></div>
    </div>
    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="ti ti-chart-line" style="color:var(--fg4)"></i> <span id="ipa-obras-title">IPA por obra — Índice de Performance Ambiental</span></div>
      </div>
      <canvas id="chart-ipa-obras" style="max-height:280px"></canvas>
    </div>
  </div>

  <!-- Ruído + Resíduos + Plano de ação (3 colunas) -->
  <div class="cards-grid cols-3" style="margin-bottom:var(--gap)">
    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="ti ti-volume" style="color:var(--az3)"></i> Ruído — medições</div>
      </div>
      <canvas id="chart-obra-ruido" style="max-height:300px"></canvas>
    </div>
    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="ti ti-chart-pie" style="color:var(--fg4)"></i> Geração Resíduos (ton)</div>
      </div>
      <canvas id="chart-obra-res-donut" style="max-height:170px"></canvas>
    </div>
    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="ti ti-clipboard-check" style="color:var(--fg4)"></i> Plano Ação Auditorias</div>
        <span class="card-badge" id="obra-pa-badge">—</span>
      </div>
      <canvas id="chart-obra-pa" style="max-height:170px"></canvas>
    </div>
  </div>

  <!-- Emergências Ambientais — analítica da obra -->
  <div class="card" style="margin-top:var(--gap)">
    <div class="card-header">
      <div class="card-title"><i class="ti ti-alert-octagon" style="color:var(--vr4)"></i> Emergências ambientais — analítica</div>
      <span class="card-badge" id="obra-emerg-badge">—</span>
    </div>
    <div class="emerg-wrap">
      <table class="emerg-table" id="obra-emerg-table"></table>
    </div>
  </div>

  <!-- Auditorias Ambientais (Dim_Auditoria_Ambientais) — movida da aba Resíduos -->
  <div class="card" style="margin-top:var(--gap)">
    <div class="card-header">
      <div class="card-title"><i class="ti ti-clipboard-check" style="color:var(--fg4)"></i> Auditorias ambientais</div>
      <span style="display:inline-flex; gap:8px; align-items:center; font-size:11px; color:var(--gz3)">
        <label>Período:</label>
        <select id="filter-aud-mes" onchange="renderAuditorias()"
          style="border:1px solid var(--gz6); border-radius:5px; padding:3px 8px; font-size:11px; color:var(--gz1); background:var(--bg); cursor:pointer">
          <option value="all">Todo o período</option>
          <option value="6m">Últimos 6 meses</option>
          <option value="12m">Últimos 12 meses</option>
          <option value="2025">2025</option>
          <option value="2026">2026</option>
        </select>
        <span class="card-badge" id="aud-badge" style="margin-left:6px">—</span>
      </span>
    </div>
    <div id="aud-tabela"></div>
    <div style="margin-top: 12px; padding: 10px 12px; background: var(--gz7); border-radius: 8px; border-left: 3px solid var(--fg4);">
      <div style="font-size: 11px; font-weight: 500; color: var(--gz2); margin-bottom: 6px;">
        <i class="ti ti-notes" style="color:var(--gz3)"></i> Principais ocorrências da última auditoria
      </div>
      <div id="aud-ocorrencias" style="font-size: 11px; color: var(--gz1);"></div>
    </div>
  </div>

  <!-- Energia + Água + Impressões: o bloco da aba Consolidado é movido para cá
       quando esta aba é aberta (showPage) -->
  <div id="slot-consumo-obra" style="margin-top:var(--gap)"></div>

  <!-- Volume por tipo de resíduo — mesmo gráfico da aba Resíduos, no recorte da obra -->
  <div class="card">
    <div class="card-header">
      <div class="card-title"><i class="ti ti-trash" style="color:var(--fg4)"></i> Volume por tipo de resíduo</div>
      <span class="card-badge" id="obra-res-tipo-badge">—</span>
    </div>
    <canvas id="chart-obra-res-tipo" style="max-height:340px"></canvas>
  </div>

</div>

<!-- ═══════════════════════════════════════════════════════════
     PAGE 3 — RESÍDUOS
════════════════════════════════════════════════════════════ -->
<div id="page-residuos" class="page">

  <!-- 4 KPI cards (Volume · Geração Resíduo · Senna Tower consolidado · Classes CONAMA) -->
  <div class="kpi-row cols-4" id="kpis-residuos" style="margin-bottom:var(--gap)"></div>

  <!-- Geração Resíduo: demais obras (meta 36,6) e Senna Tower (meta 50) lado a
       lado — mesmo padrão do card "Resíduos Gerados kg/m² por obra" da aba
       Consolidado (réguas diferentes, por isso em gráficos separados) -->
  <div class="cards-grid cols-res" style="margin-bottom:var(--gap)">
    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="ti ti-package" style="color:var(--fg4)"></i> <span id="res-ger-titulo">Geração Resíduo</span></div>
        <span class="card-badge" id="res-ger-badge">—</span>
      </div>
      <canvas id="chart-res-intensidade" style="max-height:320px"></canvas>
    </div>
    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="ti ti-building-skyscraper" style="color:var(--fg4)"></i> Geração Resíduo — Senna Tower</div>
        <span style="display:inline-flex; gap:8px; align-items:center">
          <span class="card-badge" id="res-ger-senna-acum" hidden></span>
          <span class="card-badge" id="res-ger-senna-badge">—</span>
        </span>
      </div>
      <canvas id="chart-res-intensidade-senna" style="max-height:320px"></canvas>
    </div>
  </div>

  <!-- Geração Resíduos (ton) -->
  <div class="card" style="margin-bottom:var(--gap)">
    <div class="card-header">
      <div class="card-title"><i class="ti ti-chart-bar" style="color:var(--fg4)"></i> Geração Resíduos (ton)</div>
    </div>
    <canvas id="chart-res-conama" style="max-height:280px"></canvas>
  </div>

  <!-- Classe A e Classe B: o que compõe cada classe do "Geração Resíduos (ton)" -->
  <div class="card" style="margin-bottom:var(--gap)">
    <div class="card-header">
      <div class="card-title"><i class="ti ti-list-tree" style="color:var(--fg4)"></i> Classe A e Classe B — tipos de resíduo por empreendimento</div>
      <span class="card-badge" id="res-classe-tipo-badge">—</span>
    </div>
    <!-- Painéis montados em renderResiduos() — um por tipo, agrupados pela classe -->
    <div id="res-classe-tipo-paineis"></div>
  </div>

  <!-- Tipo x mes: 6 colunas (meses) por tipo, para os 8 tipos com maior volume -->
  <div class="card" style="margin-bottom:var(--gap)">
    <div class="card-header">
      <div class="card-title"><i class="ti ti-trash" style="color:var(--fg4)"></i> Volume por tipo de resíduo (top 8)</div>
      <span class="card-badge" id="res-tipo-badge">—</span>
    </div>
    <canvas id="chart-res-tipo" style="max-height:340px"></canvas>
  </div>

  <!-- Tabela detalhada -->
  <!-- Série mensal consolidada (m³, kg, m³/m², kg/m²) -->
  <div class="cards-grid cols-2" style="margin-bottom:var(--gap)">
    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="ti ti-calendar-stats" style="color:var(--fg4)"></i> Série Mensal de Resíduos: Consolidado</div>
        <span class="card-badge">Senna: 50 · demais: 36,6 kg/m²</span>
      </div>
      <div id="res-mensal-tabela"></div>
    </div>
    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="ti ti-biohazard" style="color:var(--vr4)"></i> % Resíduos perigosos com destinação adequada</div>
        <span class="card-badge">Meta: 100%</span>
      </div>
      <div id="res-perigosos-tabela"></div>
    </div>
  </div>

  <div class="card">
    <div class="card-header">
      <div class="card-title"><i class="ti ti-table" style="color:var(--fg4)"></i> Resíduos por Obra</div>
    </div>
    <div id="res-tabela-detalhe"></div>
  </div>
</div>

<!-- ═══════════════════════════════════════════════════════════
     PAGE 4 — DICIONÁRIO DE DADOS E TERMOS
════════════════════════════════════════════════════════════ -->
<div id="page-dicionario" class="page">

  <div class="card" style="margin-bottom:var(--gap)">
    <div class="card-header">
      <div class="card-title"><i class="ti ti-book" style="color:var(--fg3)"></i> Dicionário de dados e termos</div>
      <span class="card-badge" id="dic-contador">—</span>
    </div>
    <p class="dic-intro">
      Referência das regras de negócio, das fórmulas dos indicadores e da origem de cada dado
      apresentado neste dashboard. O objetivo é que qualquer número visto nas outras abas possa ser
      rastreado até a regra que o produziu e até a tabela do modelo semântico de onde ele saiu.
      Os valores marcados em <span class="dic-code">destaque</span> são parâmetros fixos definidos
      no gerador — mudá-los exige alteração de código e nova validação.
    </p>
    <input id="dic-busca" class="dic-search" type="search" autocomplete="off"
           placeholder="Buscar termo, indicador, tabela ou regra..." oninput="filtrarDicionario()">
    <div class="dic-chips">
      <button class="dic-chip" onclick="irParaDic('dic-sec-kpis')">Indicadores e KPIs</button>
      <button class="dic-chip" onclick="irParaDic('dic-sec-ipa')">Composição do IPA</button>
      <button class="dic-chip" onclick="irParaDic('dic-sec-regras')">Regras de negócio</button>
      <button class="dic-chip" onclick="irParaDic('dic-sec-modelo')">Modelo de dados</button>
      <button class="dic-chip" onclick="irParaDic('dic-sec-glossario')">Glossário</button>
      <button class="dic-chip" onclick="irParaDic('dic-sec-limites')">Limitações conhecidas</button>
    </div>
  </div>

  <!-- ─── Indicadores e KPIs ─────────────────────────────────────────────── -->
  <div class="dic-sec" id="dic-sec-kpis">
    <div class="section-title"><i class="ti ti-chart-bar"></i> Indicadores e KPIs</div>
    <div class="card">
      <table class="dic-table">
        <thead>
          <tr><th style="width:190px">Indicador</th><th style="width:130px">Onde aparece</th><th>Como é calculado</th><th style="width:210px">Fonte</th></tr>
        </thead>
        <tbody>
          <tr class="dic-row">
            <td class="dic-nome">IPA — Índice de Performance Ambiental</td>
            <td class="dic-onde">Consolidado · Obra individual</td>
            <td>Média ponderada de três scores de 0 a 1, multiplicada por 100 e arredondada em 1 casa:
              conformidade 45%, auditoria 35% e resíduos 20% — fecha 100%.
              Cada score é calculado por obra — ver a seção <em>Composição do IPA</em>.
              <span class="dic-nota">Faixas: ≥ 90 excelente · ≥ 80 bom · ≥ 70 atenção · &lt; 70 crítico.</span></td>
            <td>Todas as tabelas dos cinco componentes</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">IPA global</td>
            <td class="dic-onde">Consolidado</td>
            <td>Média <strong>simples</strong> do IPA das obras — cada obra pesa igual, independentemente
              de porte, área construída ou volume de lançamentos.</td>
            <td>Derivado do IPA por obra</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">IPA mensal</td>
            <td class="dic-onde">Obra individual</td>
            <td>Mesma fórmula recalculada mês a mês nos últimos 12 meses. Quando falta o dado de um
              componente naquele mês, o score agregado da obra é usado como retaguarda.
</td>
            <td>Derivado dos scores mensais</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Conformidade LAI/EIV</td>
            <td class="dic-onde">Consolidado · Obra individual</td>
            <td>Percentual de conformidade lançado por obra e mês, comparado à meta da obra
              (padrão <span class="dic-code">90%</span> quando a meta não está preenchida).
              O card consolidado é a média das obras.</td>
            <td>Fat_Conformidade_LAI_EIV · Fat_Metas_LAI_EIV</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">NC registradas LAI/EIV</td>
            <td class="dic-onde">Consolidado</td>
            <td>Contagem por <strong>assunto</strong>, não por linha. Cada linha é uma auditoria e pode
              apontar vários assuntos separados por <span class="dic-code">|</span> nas colunas
              <em>NC LAI</em> e <em>NC EIV</em>; as duas colunas são somadas.
              <span class="dic-nota">Exemplo: “Armazenamento de resíduos | Ruído | Placa da LAI” conta 3 NCs. Campos “-” ou vazios não contam.</span></td>
            <td>Dim_NC_LAI_EIV</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Top NC recorrentes</td>
            <td class="dic-onde">Consolidado</td>
            <td>Os 10 assuntos de NC mais frequentes, pelo mesmo critério do card — por isso os dois
              números sempre fecham entre si.</td>
            <td>Dim_NC_LAI_EIV</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Reclamações de vizinhos</td>
            <td class="dic-onde">Consolidado · Obra individual</td>
            <td>Número de reclamações <strong>reais</strong>: linhas com ao menos um campo preenchido
              (obra, ocorrência, status ou canal).
              <span class="dic-nota">As linhas em branco da grade da planilha são descartadas — é o que faz o card bater com o gráfico por obra.</span></td>
            <td>Dim_reclamacoes_vizinhos</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Resíduos gerados (kg/m²)</td>
            <td class="dic-onde">Consolidado · Resíduos</td>
            <td>Soma dos <span class="dic-code">Kg/m²</span> das linhas do recorte. Cada linha já é a
              intensidade de um tipo de resíduo naquele mês (Kg do tipo ÷ m² construído da obra), então
              somá-las dá a intensidade total. Zeros são tratados como ausência de lançamento.
              <span class="dic-nota dic-alerta">Atenção: na visão por obra a soma acumula entre meses, enquanto a meta é mensal — ver Limitações conhecidas.</span></td>
            <td>Fat_soma_residuos</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Volume total de resíduos</td>
            <td class="dic-onde">Resíduos</td>
            <td>Soma da coluna de peso (Kg) do recorte, exibida em toneladas nos gráficos de geração.</td>
            <td>Fat_soma_residuos</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Classe CONAMA de maior geração</td>
            <td class="dic-onde">Resíduos</td>
            <td>Classe com maior peso gerado dentro do recorte, com o total em toneladas e o percentual
              que representa do total.</td>
            <td>Fat_soma_residuos (coluna Classe)</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Plano Ação Auditorias</td>
            <td class="dic-onde">Consolidado</td>
            <td>Percentual de ações com status <em>Realizado</em> ou <em>Em andamento</em> sobre o total
              de linhas do plano de ação.
              <span class="dic-nota dic-alerta">O denominador conta todas as linhas, inclusive as em branco da planilha — por isso o percentual sai baixo.</span></td>
            <td>Dim_Plano_acao</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Ações atrasadas</td>
            <td class="dic-onde">Consolidado</td>
            <td>Ações cujo prazo já passou da data de geração do dashboard e cujo status ainda não é
              <em>Realizado</em>.</td>
            <td>Dim_Plano_acao</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Auditorias realizadas</td>
            <td class="dic-onde">Obra individual</td>
            <td>Auditorias com <em>Realizado = Sim</em> dividido pelas auditorias com o campo
              <strong>preenchido</strong> (Sim ou Não). Linhas sem preenchimento ficam fora do
              numerador e do denominador.</td>
            <td>Dim_Auditoria_Ambientais</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Ruído — média e conformidade</td>
            <td class="dic-onde">Consolidado · Obra individual</td>
            <td>Média das medições em dB e percentual de medições dentro do limite de
              <span class="dic-code">70 dB</span>. Medições sem valor são descartadas.
              <span class="dic-nota dic-alerta">Medições gravadas como 0 dB hoje contam como conformes — ver Limitações conhecidas.</span></td>
            <td>Fat_Ruido</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Emergências ambientais</td>
            <td class="dic-onde">Consolidado</td>
            <td>Contagem de ocorrências registradas (vazamentos, derrames e afins).</td>
            <td>Dim_emergencias_ambientais</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Pessoas alcançadas — SGA</td>
            <td class="dic-onde">Consolidado</td>
            <td>Soma da coluna de pessoas atingidas pelas ações do Sistema de Gestão Ambiental,
              agrupada por ano.</td>
            <td>Dim_Acoes_SGA</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Consumo de água e energia</td>
            <td class="dic-onde">Consolidado</td>
            <td>Consumo absoluto (m³ e kWh) e intensidade por m² construído. A intensidade usa
              <strong>média</strong> das linhas, diferente dos resíduos, que usam soma.</td>
            <td>Fat_consumo_agua_luz</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Impressões</td>
            <td class="dic-onde">Consolidado</td>
            <td>Total de páginas, custo acumulado e percentual médio de impressão em preto e branco,
              com quebra por setor e histórico mensal.</td>
            <td>Fat_Impressoes · Fat_Impressoes_por_setor</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">NC de resíduos (obra-mês acima da meta)</td>
            <td class="dic-onde">Não exibido</td>
            <td>Conta 1 ocorrência por obra e mês em que o kg/m² realizado ultrapassa a meta da obra.
              <span class="dic-nota">Está calculado e disponível nos dados, mas <strong>não</strong> é somado ao card de NC LAI/EIV.</span></td>
            <td>Fat_soma_residuos + meta por obra</td>
          </tr>
        </tbody>
      </table>
    </div>
    <div class="dic-vazio dic-hidden">Nenhum item desta seção corresponde à busca.</div>
  </div>

  <!-- ─── Composição do IPA ──────────────────────────────────────────────── -->
  <div class="dic-sec" id="dic-sec-ipa">
    <div class="section-title"><i class="ti ti-math-function"></i> Composição do IPA</div>
    <div class="card">
      <table class="dic-table">
        <thead>
          <tr><th style="width:170px">Componente</th><th style="width:190px">Peso</th><th>Score (0 a 1)</th><th style="width:210px">Fonte</th></tr>
        </thead>
        <tbody>
          <tr class="dic-row">
            <td class="dic-nome">Conformidade LAI/EIV</td>
            <td><div class="dic-peso-wrap"><div class="dic-peso-bar"><div class="dic-peso-fill" style="width:100%"></div></div><span class="dic-peso-num">45%</span></div></td>
            <td>Conformidade média da obra ÷ meta da obra, com teto em 1,0.</td>
            <td>Fat_Conformidade_LAI_EIV</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Resíduos</td>
            <td><div class="dic-peso-wrap"><div class="dic-peso-bar"><div class="dic-peso-fill" style="width:44.4%"></div></div><span class="dic-peso-num">20%</span></div></td>
            <td>Meta da obra ÷ kg/m² médio da obra, com teto em 1,0 — quanto menor a geração, maior o
              score. Aqui a agregação é <strong>média</strong> das linhas, e não soma.
              <span class="dic-nota">Usa a meta fixa por obra, não a coluna “Meta Kg” do modelo.</span></td>
            <td>Fat_soma_residuos</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Auditoria</td>
            <td><div class="dic-peso-wrap"><div class="dic-peso-bar"><div class="dic-peso-fill" style="width:77.8%"></div></div><span class="dic-peso-num">35%</span></div></td>
            <td>Auditorias com <em>Realizado = Sim</em> ÷ total de auditorias da obra.</td>
            <td>Dim_Auditoria_Ambientais</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Sem dado — retaguarda</td>
            <td class="dic-onde">—</td>
            <td>Quando a obra não tem dado de um componente, aquele score entra como
              <span class="dic-code">0,5</span> (neutro), e não como zero. É uma decisão deliberada:
              obra sem lançamento não é penalizada como se tivesse desempenho ruim, mas também não é
              premiada.
              <span class="dic-nota dic-alerta">Efeito colateral: uma obra com pouquíssimo dado tende ao IPA 50, o que pode mascarar tanto um bom quanto um mau desempenho real.</span></td>
            <td>—</td>
          </tr>
        </tbody>
      </table>
    </div>
    <div class="dic-vazio dic-hidden">Nenhum item desta seção corresponde à busca.</div>
  </div>

  <!-- ─── Regras de negócio ──────────────────────────────────────────────── -->
  <div class="dic-sec" id="dic-sec-regras">
    <div class="section-title"><i class="ti ti-gavel"></i> Regras de negócio</div>
    <div class="card">
      <table class="dic-table">
        <thead>
          <tr><th style="width:230px">Regra</th><th>Definição</th></tr>
        </thead>
        <tbody>
          <tr class="dic-row">
            <td class="dic-nome">Meta de resíduos</td>
            <td><span class="dic-code" id="dic-meta-default">36,6</span> kg/m² para todas as obras e
              <span class="dic-code" id="dic-meta-senna">50</span> kg/m² para a Senna Tower.
              A meta é <strong>mensal</strong> e está fixa no gerador — não vem da planilha nem da
              coluna “Meta Kg” do modelo.</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Senna Tower fora da visão consolidada</td>
            <td>A Senna tem meta diferente e intensidade muito acima das demais obras. Se entrasse na
              mesma escala, achataria todas as outras barras. Por isso o gráfico consolidado mostra as
              demais obras com a meta padrão e a Senna aparece num recorte próprio, pelo botão
              “Ver Senna Tower”.</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Limite de ruído</td>
            <td>Referência de <span class="dic-code">70 dB</span> para classificar uma medição como
              conforme.</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Meta de conformidade</td>
            <td>Quando a obra não tem meta preenchida, o cálculo assume <span class="dic-code">90%</span>.</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Obra é o mesmo que empreendimento</td>
            <td>Não existe dimensão de empreendimento separada de obra. As obras de Dim_Obras
              (Blue Coast, Sapphire, Titanium e demais) <strong>são</strong> os empreendimentos.</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Normalização do nome da obra</td>
            <td>O mesmo empreendimento aparece grafado de formas diferentes entre as tabelas de origem.
              O gerador normaliza antes de cruzar: remove sufixos após hífen, colapsa espaços e aplica
              um mapa de equivalências.
              <span class="dic-nota">Exemplo: “Garden” em reclamações e “Garden Park” nas demais tabelas são a mesma obra — sem essa regra, as reclamações dessa obra somem do gráfico.</span></td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Precedência dos filtros</td>
            <td>O filtro global de obra tem precedência sobre os botões de recorte dos gráficos
              (como o de Senna Tower). Com uma obra selecionada, o gráfico mostra aquela obra
              independentemente do estado do botão.</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Sem dado não é zero</td>
            <td>A distinção vale por indicador: em ruído, mês sem medição é descartado da série;
              em NC e reclamações, mês sem registro conta como zero real, porque ali a ausência de
              ocorrência é a informação.</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Contagem por assunto</td>
            <td>Não-conformidades são contadas por assunto apontado, não por linha de auditoria.
              Uma auditoria que aponta três problemas conta três NCs.</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Linha em branco não é ocorrência</td>
            <td>Reclamações e emergências vêm de planilhas com grade pré-formatada, cheia de linhas
              vazias. Só entram no indicador as linhas com conteúdo real.</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Origem dos números publicados</td>
            <td>O dashboard é gerado a partir do modelo semântico do Power BI (arquivo .pbix), não
              diretamente da planilha. Um lançamento feito na planilha só aparece aqui depois que o
              modelo é atualizado e o dashboard é regerado.</td>
          </tr>
        </tbody>
      </table>
    </div>
    <div class="dic-vazio dic-hidden">Nenhum item desta seção corresponde à busca.</div>
  </div>

  <!-- ─── Modelo de dados ────────────────────────────────────────────────── -->
  <div class="dic-sec" id="dic-sec-modelo">
    <div class="section-title"><i class="ti ti-database"></i> Modelo de dados — tabelas de origem</div>
    <div class="card">
      <p class="dic-intro" style="margin-bottom:10px">
        As 19 tabelas do modelo semântico. Cada uma vem de uma tabela nomeada da planilha
        <em>Controle Ambiental — oficial</em>; a coluna <em>Grão</em> indica o que identifica uma linha.
      </p>
      <table class="dic-table">
        <thead>
          <tr><th style="width:215px">Tabela</th><th style="width:150px">Origem na planilha</th><th style="width:235px">Grão (chave)</th><th>Conteúdo</th></tr>
        </thead>
        <tbody>
          <tr class="dic-row"><td class="dic-nome">Dim_Obras</td><td class="dic-onde">Metas_LAI_EIV</td><td class="dic-onde">Obra</td><td>Lista canônica de obras, com a área construída usada como denominador dos indicadores por m².</td></tr>
          <tr class="dic-row"><td class="dic-nome">Dim_Calendario</td><td class="dic-onde">gerada</td><td class="dic-onde">Data</td><td>Calendário gerado pelo modelo; não depende da planilha. Não é usada pelo dashboard.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Dim_Acoes_SGA</td><td class="dic-onde">Acoes_SGA</td><td class="dic-onde">Ano + Mês</td><td>Pessoas alcançadas pelas ações do Sistema de Gestão Ambiental.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Dim_Auditoria_Ambientais</td><td class="dic-onde">Auditoria_Ambientais</td><td class="dic-onde">Obra + Mês/Ano</td><td>Auditorias ambientais por obra e mês, com o campo Realizado (Sim/Não).<span class="dic-nota">O mês correto vem das colunas de texto Mês e Ano — a coluna de data tem dias variando por obra e não reflete o mês real.</span></td></tr>
          <tr class="dic-row"><td class="dic-nome">Dim_NC_LAI_EIV</td><td class="dic-onde">NC_LAI_EIV</td><td class="dic-onde">Obra + Data</td><td>Não-conformidades de LAI e EIV em texto livre, com vários assuntos por linha separados por barra vertical.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Dim_Plano_acao</td><td class="dic-onde">Plano_acao</td><td class="dic-onde">Obra + Ação</td><td>Plano de ação das auditorias, com status e prazo.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Dim_emergencias_ambientais</td><td class="dic-onde">emergencias_ambientais</td><td class="dic-onde">Obra + Mês/Ano + N° FMC</td><td>Emergências ambientais registradas (vazamentos, derrames e afins).</td></tr>
          <tr class="dic-row"><td class="dic-nome">Dim_reclamacoes_vizinhos</td><td class="dic-onde">reclamacoes_vizinhos</td><td class="dic-onde">Mês/Ano + Cód. Chamado</td><td>Reclamações de vizinhos, com canal de entrada (Ouvidoria ou Relacionamento), categoria e status.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Fat_Conformidade_LAI_EIV</td><td class="dic-onde">Conformidade_LAI_EIV</td><td class="dic-onde">Obra + Data</td><td>Percentual de conformidade LAI/EIV por obra e mês, com a meta do período.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Fat_Eco_urbana</td><td class="dic-onde">Eco_urbana</td><td class="dic-onde">Obra + Mês/Ano</td><td>Medições de ruído feitas por terceiros. Não é usada pelo dashboard.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Fat_Impressoes</td><td class="dic-onde">Impressoes</td><td class="dic-onde">Data</td><td>Consumo mensal de impressão da empresa: páginas, custo e percentual em preto e branco.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Fat_Impressoes_por_pessoa</td><td class="dic-onde">Impressoes_por_pessoa</td><td class="dic-onde">Data + Usuário</td><td>Impressões por usuário, com a obra deduzida da conta. Não é usada pelo dashboard.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Fat_Impressoes_por_setor</td><td class="dic-onde">Impressoes_por_setor</td><td class="dic-onde">Data + Nome do setor</td><td>Impressões por setor e obra; alimenta o ranking de setores.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Fat_Metas_LAI_EIV</td><td class="dic-onde">Metas_LAI_EIV</td><td class="dic-onde">Obra</td><td>Meta de conformidade por obra, com uma coluna por mês (formato largo herdado da planilha).</td></tr>
          <tr class="dic-row"><td class="dic-nome">Fat_Residuos_m³_m²</td><td class="dic-onde">Residuos_m³_m²</td><td class="dic-onde">Obra + Data</td><td>Volume de resíduos em m³ e a intensidade por m² construído.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Fat_Ruido</td><td class="dic-onde">Ruido</td><td class="dic-onde">Obra + Data + Ponto</td><td>Medições de ruído por obra e ponto de medição, com horário e valor em dB.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Fat_avanco_fisico</td><td class="dic-onde">Avanco_fisico</td><td class="dic-onde">Obra + Data</td><td>Avanço físico da obra em percentual e em m². É a base do denominador dos indicadores por m².</td></tr>
          <tr class="dic-row"><td class="dic-nome">Fat_consumo_agua_luz</td><td class="dic-onde">agua_luz</td><td class="dic-onde">Setor + Data</td><td>Consumo de água e energia por setor e obra, em valor absoluto e por m².</td></tr>
          <tr class="dic-row"><td class="dic-nome">Fat_soma_residuos</td><td class="dic-onde">SOMA_RESIDUOS</td><td class="dic-onde">Obra + Data + Classe + Tipo</td><td>Resíduos por obra, mês, classe CONAMA e tipo. É a tabela mais usada do dashboard: Kg = m³ × densidade e Kg/m² = Kg ÷ m² construído.</td></tr>
        </tbody>
      </table>
    </div>
    <div class="dic-vazio dic-hidden">Nenhum item desta seção corresponde à busca.</div>
  </div>

  <!-- ─── Glossário ──────────────────────────────────────────────────────── -->
  <div class="dic-sec" id="dic-sec-glossario">
    <div class="section-title"><i class="ti ti-vocabulary"></i> Glossário de termos</div>
    <div class="card">
      <table class="dic-table">
        <thead>
          <tr><th style="width:200px">Termo</th><th>Significado</th></tr>
        </thead>
        <tbody>
          <tr class="dic-row"><td class="dic-nome">LAI</td><td>Licença Ambiental de Instalação — autorização para implantar o empreendimento, com condicionantes que a obra precisa cumprir e comprovar.</td></tr>
          <tr class="dic-row"><td class="dic-nome">EIV</td><td>Estudo de Impacto de Vizinhança — estudo urbanístico que gera medidas mitigadoras e compensatórias exigidas da obra.</td></tr>
          <tr class="dic-row"><td class="dic-nome">NC — não-conformidade</td><td>Item exigido pela LAI ou pelo EIV que a auditoria encontrou descumprido. Aqui a unidade de contagem é o assunto apontado, não a linha de auditoria.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Conformidade</td><td>Percentual de itens da LAI e do EIV atendidos pela obra no mês.</td></tr>
          <tr class="dic-row"><td class="dic-nome">IPA</td><td>Índice de Performance Ambiental — nota única de 0 a 100 que resume conformidade LAI/EIV, auditoria e resíduos de uma obra.</td></tr>
          <tr class="dic-row"><td class="dic-nome">SGA</td><td>Sistema de Gestão Ambiental — conjunto de práticas e ações de conscientização; o indicador aqui é o número de pessoas alcançadas.</td></tr>
          <tr class="dic-row"><td class="dic-nome">CONAMA 307</td><td>Resolução que classifica os resíduos da construção civil em quatro classes.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Classe A</td><td>Resíduos reutilizáveis ou recicláveis como agregado: concreto, argamassa, alvenaria, entulho em geral. É a maior parcela do volume gerado.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Classe B</td><td>Recicláveis para outras destinações: plástico, papel e papelão, metal, madeira, gesso.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Classe C</td><td>Resíduos sem tecnologia ou aplicação viável de reciclagem.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Classe D</td><td>Resíduos perigosos: tintas, solventes, óleos, materiais contaminados. Exigem destinação específica.</td></tr>
          <tr class="dic-row"><td class="dic-nome">kg/m²</td><td>Intensidade de geração: peso de resíduo dividido pelos m² construídos da obra. Permite comparar obras de portes diferentes.</td></tr>
          <tr class="dic-row"><td class="dic-nome">m³/m²</td><td>Mesma lógica em volume: m³ de resíduo por m² construído.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Densidade</td><td>Fator de conversão de volume para peso por tipo de resíduo. É o que transforma m³ lançado em Kg.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Avanço físico</td><td>Percentual e área executados da obra no período. Fornece os m² que servem de denominador das intensidades.</td></tr>
          <tr class="dic-row"><td class="dic-nome">dB — decibel</td><td>Unidade de intensidade sonora das medições de ruído. O limite de referência adotado é 70 dB.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Ponto de medição</td><td>Local fixo da obra onde o ruído é medido. A mesma obra tem vários pontos, medidos no mesmo dia.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Obra-mês</td><td>Unidade de análise que combina uma obra e um mês. Usada, por exemplo, para contar quantas vezes uma obra passou da meta de resíduos.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Score</td><td>Nota de 0 a 1 de um componente do IPA, antes da ponderação. 1 significa atingiu ou superou a referência.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Recorte</td><td>Combinação dos filtros ativos (obra e período) que define o conjunto de linhas usado para calcular o que está na tela.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Modelo semântico</td><td>Conjunto de tabelas tratadas e relacionadas que fica dentro do arquivo .pbix. É o que o dashboard lê — não a planilha bruta.</td></tr>
          <tr class="dic-row"><td class="dic-nome">Ouvidoria / Relacionamento</td><td>Canais por onde a reclamação de vizinho chega até a empresa.</td></tr>
          <tr class="dic-row"><td class="dic-nome">FMC</td><td>Número do formulário que identifica o registro de uma emergência ambiental.</td></tr>
        </tbody>
      </table>
    </div>
    <div class="dic-vazio dic-hidden">Nenhum item desta seção corresponde à busca.</div>
  </div>

  <!-- ─── Limitações conhecidas ──────────────────────────────────────────── -->
  <div class="dic-sec" id="dic-sec-limites">
    <div class="section-title"><i class="ti ti-alert-triangle"></i> Limitações conhecidas</div>
    <div class="card">
      <p class="dic-intro" style="margin-bottom:10px">
        Pontos já identificados e ainda em aberto. Estão listados aqui para que quem lê o dashboard
        interprete os números com o contexto certo — nenhum deles é erro de geração.
      </p>
      <table class="dic-table">
        <thead>
          <tr><th style="width:250px">Ponto</th><th>O que acontece</th></tr>
        </thead>
        <tbody>
          <tr class="dic-row">
            <td class="dic-nome">kg/m² acumula entre meses</td>
            <td>Na visão consolidada por obra, o kg/m² soma todas as linhas da obra, inclusive de meses
              diferentes, enquanto a meta é mensal. O valor cresce a cada mês lançado e o ranking fica
              distorcido: uma obra com poucos meses lançados aparece melhor do que realmente é.
              <span class="dic-nota">No eixo mensal o problema não existe — ali cada barra já é um mês.</span></td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Agregações diferentes lado a lado</td>
            <td>O kg/m² usa soma e o m³/m² usa média, e os dois aparecem na mesma tabela. Comparar as
              duas colunas diretamente induz a erro.</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Plano de ação com denominador inflado</td>
            <td>A tabela do plano de ação tem a maioria das linhas em branco e todas entram no
              denominador, o que derruba o percentual exibido.</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Medições de 0 dB contam como conformes</td>
            <td>Existem medições gravadas com 0 dB. Como 0 é menor que 70, elas entram como conformes
              e inflam o score de ruído da obra.</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Setor “Unknown” nas impressões</td>
            <td>Aparece entre os primeiros do ranking de setores, mas não é um setor real: são
              impressões sem usuário mapeado.</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Água por m² zerada na origem</td>
            <td>A coluna de água por m² é convertida para número inteiro ainda no modelo, então valores
              fracionários viram zero. O indicador nasce zerado antes de chegar ao dashboard.</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Defasagem do modelo</td>
            <td>O dashboard reflete a última atualização do modelo Power BI, que pode estar atrás da
              planilha. Obra nova ou lançamento recente só aparece depois que o modelo é atualizado.</td>
          </tr>
          <tr class="dic-row">
            <td class="dic-nome">Nomes de obra ainda divergentes na origem</td>
            <td>Algumas obras são grafadas de formas diferentes entre as planilhas. O gerador corrige
              as variações conhecidas, mas uma grafia nova que apareça na origem pode fazer a obra
              sumir de um gráfico até o mapa ser atualizado.</td>
          </tr>
        </tbody>
      </table>
    </div>
    <div class="dic-vazio dic-hidden">Nenhum item desta seção corresponde à busca.</div>
  </div>

  <div id="dic-sem-resultado" class="card dic-hidden" style="margin-top:var(--gap)">
    <div class="dic-vazio">Nenhum resultado para a busca. Tente outro termo.</div>
  </div>

</div>

<div class="dash-footer">
  FG Viva+ / Sustentabilidade · Dashboard Indicadores Ambientais · Gerado em __DATA_DATE__
</div>

<script>
// ═══════════════════════════════════════════════════════════
//  Dados
// ═══════════════════════════════════════════════════════════
const D = __DATA_JSON__;

// ═══════════════════════════════════════════════════════════
//  Navegação
// ═══════════════════════════════════════════════════════════
const PAGES = ['consolidado', 'obra', 'residuos', 'dicionario'];
function showPage(id) {
  PAGES.forEach(p => {
    document.getElementById('page-' + p).classList.toggle('active', p === id);
    document.getElementById('tab-' + p).classList.toggle('active', p === id);
  });
  // Energia + Água + Impressões aparecem no Consolidado e na Obra Individual:
  // o mesmo bloco é movido para o slot da aba aberta (os gráficos Chart.js
  // seguem vivos e se redimensionam pelo ResizeObserver).
  const blocoConsumo = document.getElementById('bloco-consumo');
  const slotConsumo  = document.getElementById(id === 'obra' ? 'slot-consumo-obra' : 'slot-consumo-consolidado');
  if (blocoConsumo && slotConsumo && (id === 'obra' || id === 'consolidado')
      && blocoConsumo.parentNode !== slotConsumo) slotConsumo.appendChild(blocoConsumo);
  if (id === 'obra') renderObra();
  if (id === 'dicionario') renderDicionario();
  // A barra de filtros nao tem efeito sobre o dicionario (conteudo de referencia).
  const fb = document.getElementById('filter-bar-consolidado');
  if (fb) fb.style.display = (id === 'dicionario') ? 'none' : '';
}

// ═══════════════════════════════════════════════════════════
//  Utilitários
// ═══════════════════════════════════════════════════════════
const COLORS = {
  verde: ['#639922','#97C459','#C0DD97','#3B6D11','#27500A','#EAF3DE','#173404'],
  azul:  ['#185FA5','#378ADD','#85B7EB','#0C447C','#042C53','#B5D4F4','#E6F1FB'],
  ambar: ['#EF9F27','#FAC775','#BA7517','#633806','#FAEEDA'],
  verm:  ['#CE8686','#E0A8A8','#A86D6D','#791F1F','#FCEBEB'],
  cinza: ['#888780','#B4B2A9','#D3D1C7','#5F5E5A','#F1EFE8'],
};
// ═══════════════════════════════════════════════════════════
//  Filtro de Período — helper único
// ═══════════════════════════════════════════════════════════
// "Últimos 6/12 meses" é uma JANELA DE CALENDÁRIO ÚNICA, igual para todos os
// gráficos, ancorada em D.periodo_ref (o mês mais recente com dado real,
// calculado no gerador).
//
// Antes cada gráfico pegava os N últimos meses da SUA PRÓPRIA fonte. Como cada
// tabela termina num mês diferente — conformidade em 2026-05, água em 2026-03,
// resíduos com grade zerada até 2026-07, auditorias com planejamento até
// 2026-10 —, a mesma seleção produzia sete janelas diferentes na mesma tela:
// "últimos 6 meses" mostrava Dez/25–Mai/26 na conformidade, Out/25–Mar/26 na
// água e Fev/26–Jul/26 nos resíduos.
//
// Agora a janela é a mesma em todo lugar; cada gráfico mostra o que tem dentro
// dela (pode ter menos barras que o teto, quando a fonte acaba antes).
function mesAnterior(mes, n) {
  const [a, m] = String(mes).split('-').map(Number);
  const total = a * 12 + (m - 1) - n;
  return String(Math.floor(total / 12)).padStart(4, '0') + '-' + String(total % 12 + 1).padStart(2, '0');
}
// Janela {de, ate} em 'YYYY-MM' para o valor do filtro; null = sem recorte.
function janelaPeriodo(anoF) {
  if (!anoF || anoF === 'all') return null;
  const ref = (typeof D !== 'undefined' && D.periodo_ref) || null;
  if (anoF === '6m' || anoF === '12m') {
    if (!ref) return null;                       // sem âncora, não recorta
    const n = anoF === '6m' ? 5 : 11;
    return { de: mesAnterior(ref, n), ate: ref };
  }
  return { de: anoF + '-01', ate: anoF + '-12' };  // ano cheio: '2025' / '2026'
}
// true se o mês 'YYYY-MM' (ou 'YYYY-MM-DD') cai dentro do recorte.
function noPeriodo(mes, anoF) {
  const j = janelaPeriodo(anoF);
  if (!j) return true;
  const m = String(mes || '').slice(0, 7);
  return !!m && m >= j.de && m <= j.ate;
}
// Set dos meses do universo que caem no recorte. O universo continua servindo
// para o caso "todo o período" e para as funções que precisam da lista pronta.
function mesesDoPeriodo(anoF, universo) {
  const ms = [...new Set((universo || []).map(m => String(m || '').slice(0, 7)))].filter(Boolean).sort();
  const j = janelaPeriodo(anoF);
  if (!j) return new Set(ms);
  return new Set(ms.filter(m => m >= j.de && m <= j.ate));
}
// Rótulo da janela, ex.: "Dez/25 – Mai/26" — usado nos badges.
function faixaPeriodo(anoF) {
  const j = janelaPeriodo(anoF);
  if (!j) return 'todo o período';
  return labelMesRes(j.de) + ' – ' + labelMesRes(j.ate);
}
// Valor corrente do filtro global de Período.
function periodoGlobal() {
  return (document.getElementById('filter-periodo') || {}).value || 'all';
}
// Um filtro de card (água, energia, impressões, auditorias) tem precedência
// sobre o global quando o usuário escolheu algo nele; em 'all' ele obedece ao
// filtro global do topo.
function periodoEfetivo(idSelectCard, padraoCard) {
  const v = (document.getElementById(idSelectCard) || {}).value;
  if (v && v !== (padraoCard || 'all')) return v;
  const g = periodoGlobal();
  return g !== 'all' ? g : (v || 'all');
}
// Rótulo curto do valor do filtro de Período, para badges e mensagens.
function labelPeriodo(anoF) {
  return { all: 'todo o período', '6m': 'últimos 6 meses', '12m': 'últimos 12 meses' }[anoF] || anoF;
}
function pct(v) { return v != null ? v.toFixed(1) + '%' : '—'; }
function num(v, dec=0) { return v != null ? Number(v).toFixed(dec) : '—'; }
// Número no padrão pt-BR: milhar "." e decimal "," (ex.: 1.234,56)
function numBR(v, dec=0) { return v != null ? Number(v).toLocaleString('pt-BR', { minimumFractionDigits: dec, maximumFractionDigits: dec }) : '—'; }
function semClass(v, meta) {
  if (v == null) return 'cell-nd';
  const r = v / meta;
  if (r >= 1.0) return 'cell-ex';
  if (r >= 0.90) return 'cell-ok';
  if (r >= 0.80) return 'cell-am';
  return 'cell-vr';
}
function ipaClass(v) {
  if (v >= 90) return 'ipa-ex';
  if (v >= 80) return 'ipa-bom';
  if (v >= 70) return 'ipa-am';
  return 'ipa-cr';
}
function badgeIpa(v) {
  if (v >= 90) return '<span class="badge badge-ex">🟢 Excelente</span>';
  if (v >= 80) return '<span class="badge badge-bom">🟡 Bom</span>';
  if (v >= 70) return '<span class="badge badge-am">🟠 Atenção</span>';
  return '<span class="badge badge-cr">🔴 Crítico</span>';
}
// Cor do card "Resíduos gerados": verde ABAIXO da meta da obra, vermelho a
// partir dela. A meta é 36,6 kg/m² em todas as obras — Senna Tower é a única
// exceção (50 kg/m²), vinda de meta_residuos.por_obra / res_por_obra[o].meta.
function resClass(v, meta) {
  if (v == null) return 'gray';
  return v < (meta || 36.6) ? 'green' : 'red';
}
function progColor(v, meta) {
  const r = v / meta;
  if (r >= 1) return 'prog-green';
  if (r >= 0.9) return 'prog-green';
  if (r >= 0.8) return 'prog-amber';
  return 'prog-red';
}
const chartReg = {}; // track chart instances
function makeChart(id, config) {
  if (chartReg[id]) { chartReg[id].destroy(); }
  const ctx = document.getElementById(id);
  if (!ctx) return null;
  chartReg[id] = new Chart(ctx, config);
  return chartReg[id];
}

// Plugin reutilizável: rótulo de valor no topo de cada barra (dataset índice 0,
// as barras "de dado" — datasets extras de linha/meta ficam de fora do rótulo).
function barValueLabels(formatFn, datasetIndex = 0) {
  return {
    id: 'barValueLabels',
    afterDatasetsDraw(chart) {
      const meta = chart.getDatasetMeta(datasetIndex);
      const data = (chart.data.datasets[datasetIndex] || {}).data || [];
      if (!meta) return;
      const { ctx } = chart;
      ctx.save();
      ctx.font = '600 11px "Segoe UI", Inter, system-ui, sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'bottom';
      ctx.fillStyle = '#3A3833';
      meta.data.forEach((bar, i) => {
        const v = data[i];
        if (v == null) return;
        ctx.fillText(formatFn(v), bar.x, bar.y - 4);
      });
      ctx.restore();
    }
  };
}

// Plugin reutilizável: rótulo de valor no centro de cada fatia de doughnut/pie
// (dataset índice 0).
// Rótulo por FORA de cada fatia (com linha-guia até a borda do arco) — usado
// quando a rosca é pequena/fatias finas demais pro rótulo caber dentro.
function donutValueLabels(formatFn) {
  return {
    id: 'donutValueLabels',
    afterDatasetsDraw(chart) {
      const meta = chart.getDatasetMeta(0);
      const data = (chart.data.datasets[0] || {}).data || [];
      if (!meta) return;
      const total = data.reduce((a, v) => a + (v || 0), 0);
      const { ctx } = chart;
      ctx.save();
      ctx.font = '600 11px "Segoe UI", Inter, system-ui, sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillStyle = '#fff';
      meta.data.forEach((arc, i) => {
        const v = data[i];
        if (!v) return;
        const pos = arc.getCenterPoint();
        ctx.fillText(formatFn(v, total), pos.x, pos.y);
      });
      ctx.restore();
    }
  };
}
Chart.defaults.font.family = "'Segoe UI', Inter, system-ui, sans-serif";
Chart.defaults.font.size = 11;

// ═══════════════════════════════════════════════════════════
//  KPIs — os 6 cards, na mesma ordem, no Consolidado e na Obra Individual
// ═══════════════════════════════════════════════════════════
// Ordem definida pela área: IPA → Conformidade LAI/EIV → Plano Ação Auditorias →
// Emergências Ambientais → Resíduos gerados → Reclamações de Vizinhos.

// Emergências no recorte obra × período — mesmo filtro da tabela (renderEmergencias)
function emergFiltradas(obraF, anoF) {
  let data = D.emergencias_lista || [];
  if (obraF) data = data.filter(r => r.obra === obraF);
  if (anoF && anoF !== 'all') data = data.filter(r => noPeriodo(r.mes || r.mes_ano, anoF));
  return data;
}
const emergSubDe = anoF => (!anoF || anoF === 'all') ? 'acumulado' : 'acumulado no período';

// % de ações Realizado/Em andamento do plano de ação (uma obra ou todas)
function paPctDe(obraF) {
  if (!obraF) return D.global.pa_no_prazo_pct;
  const pa_o = (D.pa_por_obra || {})[obraF] || {};
  const total = Object.values(pa_o).reduce((a, b) => a + b, 0);
  const done  = (pa_o['Realizado'] || 0) + (pa_o['Em Andamento'] || 0);
  return total > 0 ? Math.round(done / total * 100 * 10) / 10 : 0;
}

// k = { ipaLabel, ipa, conf, confMeta, pa, emerg, emergSub, kgM2, resMeta, resSub, rec, recLimite }
function kpiCardsHtml(k) {
  const ipaColor = k.ipa >= 90 ? 'green' : k.ipa >= 80 ? 'blue' : k.ipa >= 70 ? 'amber' : 'red';
  const confCor  = k.confMeta == null ? 'green'
                 : (k.conf || 0) >= k.confMeta ? 'green'
                 : (k.conf || 0) >= k.confMeta * 0.9 ? 'amber' : 'red';
  return `
    <div class="kpi ${ipaColor}">
      <div class="kpi-label"><i class="ti ti-chart-line"></i> ${k.ipaLabel}</div>
      <div class="kpi-value large ${ipaClass(k.ipa)}">${k.ipa}</div>
      <div class="kpi-sub">Índice de Performance Ambiental</div>
      <div style="margin-top:6px">${badgeIpa(k.ipa)}</div>
    </div>
    <div class="kpi ${confCor}">
      <div class="kpi-label"><i class="ti ti-file-certificate"></i> Conformidade LAI/EIV</div>
      <div class="kpi-value">${pct(k.conf)}</div>
      <div class="kpi-sub">${k.confMeta == null ? 'média LAI/EIV' : 'meta: ' + pct(k.confMeta)}</div>
      <i class="ti ti-file-certificate kpi-icon"></i>
    </div>
    <div class="kpi ${k.pa >= 80 ? 'green' : 'amber'}">
      <div class="kpi-label"><i class="ti ti-clipboard-check"></i> Plano Ação Auditorias</div>
      <div class="kpi-value">${pct(k.pa)}</div>
      <div class="kpi-sub">ações em dia</div>
      <i class="ti ti-clipboard-check kpi-icon"></i>
    </div>
    <div class="kpi ${k.emerg > 0 ? 'amber' : 'green'}">
      <div class="kpi-label"><i class="ti ti-alert-octagon"></i> Emergências Ambientais</div>
      <div class="kpi-value">${k.emerg}</div>
      <div class="kpi-sub">${k.emergSub}</div>
      <i class="ti ti-alert-octagon kpi-icon"></i>
    </div>
    <div class="kpi ${resClass(k.kgM2, k.resMeta)}">
      <div class="kpi-label"><i class="ti ti-package"></i> Resíduos gerados</div>
      <div class="kpi-value">${num(k.kgM2, 2)}</div>
      <div class="kpi-sub">${k.resSub}</div>
      <i class="ti ti-trash kpi-icon"></i>
    </div>
    <div class="kpi ${k.rec > k.recLimite ? 'amber' : 'green'}">
      <div class="kpi-label"><i class="ti ti-message-report"></i> Reclamações de Vizinhos</div>
      <div class="kpi-value">${k.rec}</div>
      <div class="kpi-sub">vizinhos no período</div>
      <i class="ti ti-message-report kpi-icon"></i>
    </div>
  `;
}

// ═══════════════════════════════════════════════════════════
//  PAGE 1 — Consolidado
// ═══════════════════════════════════════════════════════════
function renderConsolidado() {
  // KPI cards — mesma função usada quando os filtros mudam
  updateKpisFiltered((document.getElementById('filter-obra-global') || {}).value || '', periodoGlobal());

  // Bloco SGA + Energia + Água + Impressões
  renderSGAAnual();
  renderConsumoEnergiaAll();
  renderConsumoAguaAll();
  renderImpressoes();

  // Heatmap
  renderHeatmap();

  renderTopNcs();

  // Charts
  renderResBarras();
  renderResSenna();
  renderPaStatus();  // sem filtro inicial (será atualizado via applyFilters se houver)
  renderRecPorObra();
  // IPA por obra agora está na aba Obra Individual — renderizado em renderObra()
  renderMonitoramentoRuido();
  renderEmergencias();
}

// ═══════════════════════════════════════════════════════════
//  Top não conformidades recorrentes — respeita obra e período
// ═══════════════════════════════════════════════════════════
// Conta por ASSUNTO (mesmo critério da medida de NC), a partir de
// D.nc_assuntos_lista, que traz obra × mês × assunto. Sem essa lista o gráfico
// caía no agregado estático D.top_ncs e não reagia a filtro nenhum.
function renderTopNcs(obraF, anoF) {
  const ncsEl = document.getElementById('top-ncs-list');
  const badge = document.getElementById('top-ncs-badge');
  if (!ncsEl) return;
  if (!obraF) obraF = (document.getElementById('filter-obra-global') || {}).value || '';
  if (!anoF)  anoF  = (document.getElementById('filter-periodo') || {}).value || 'all';

  const lista = D.nc_assuntos_lista || [];
  let dados;
  if (lista.length) {
    // Janela de meses: mesma régua do heatmap/IPA (conf_meses)
    const allMeses = D.conf_meses || [];
    const mesesSet = (anoF && anoF !== 'all') ? mesesDoPeriodo(anoF, allMeses) : null;

    let filt = lista;
    if (obraF) filt = filt.filter(r => r.obra === obraF);
    if (mesesSet) filt = filt.filter(r => r.mes && mesesSet.has(r.mes));

    const freq = {};
    filt.forEach(r => { freq[r.assunto] = (freq[r.assunto] || 0) + 1; });
    dados = Object.entries(freq)
      .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'pt-BR'))
      .slice(0, 10)
      .map(([nc, count]) => ({ nc, count }));
  } else {
    dados = (D.top_ncs || []).slice(0, 10);   // fallback: base antiga sem granularidade
  }

  const total = dados.reduce((a, d) => a + d.count, 0);
  if (badge) {
    const periodoLbl = { all: 'todo o período', '6m': 'últ. 6 meses', '12m': 'últ. 12 meses' }[anoF] || anoF;
    badge.textContent = total
      ? `${total} NC${total !== 1 ? 's' : ''} · ${obraF || periodoLbl}`
      : (obraF || periodoLbl);
  }

  if (!dados.length) {
    const ctx = obraF ? ` para ${obraF}` : '';
    ncsEl.innerHTML = `<li class="empty-msg">Nenhuma NC no recorte${ctx}</li>`;
    return;
  }

  const max = dados[0].count;
  ncsEl.innerHTML = dados.map(n => `
    <li class="nc-item">
      <span class="nc-count">${n.count}</span>
      <span class="nc-name">${n.nc}</span>
      <div style="width:60px">
        <div class="prog-wrap"><div class="prog-bar prog-red" style="width:${Math.round(n.count / max * 100)}%"></div></div>
      </div>
    </li>`).join('');
}

function renderHeatmap(obraF, anoF) {
  const cont = document.getElementById('heatmap-container');
  let meses = D.conf_meses || [];
  let obras = Object.keys(D.conf_heatmap || {});
  if (obraF) obras = obras.filter(o => o === obraF);
  if (anoF && anoF !== 'all') meses = meses.filter(m => noPeriodo(m, anoF));
  if (!meses.length || !obras.length) {
    cont.innerHTML = '<div class="empty-msg">Sem dados de conformidade</div>';
    return;
  }
  const labelMes = m => {
    const p = m.split('-');
    const mNames = ['','Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
    return (mNames[parseInt(p[1])] || p[1]) + '/' + p[0].slice(2);
  };
  let html = '<table class="heatmap"><thead><tr><th class="obra-col">Obra</th>';
  meses.forEach(m => { html += `<th>${labelMes(m)}</th>`; });
  html += '</tr></thead><tbody>';
  obras.forEach(o => {
    const row = D.conf_heatmap[o] || {};
    const meta = D.conf_atual[o] ? D.conf_atual[o].meta : 90;
    html += `<tr><td class="obra-name">${o}</td>`;
    meses.forEach(m => {
      const cell = row[m] || {};
      const v = cell.perc;
      const cls = v != null ? semClass(v, meta||90) : 'cell-nd';
      html += `<td class="${cls}">${v != null ? v.toFixed(0)+'%' : '—'}</td>`;
    });
    html += '</tr>';
  });
  html += '</tbody></table>';
  cont.innerHTML = html;
}

// ═══════════════════════════════════════════════════════════
//  Obras ocultas por gráfico (pedido da área)
//  Cada gráfico tem a sua lista — NÃO é um filtro global de obra. Para voltar a
//  exibir uma obra, basta removê-la da lista correspondente.
// ═══════════════════════════════════════════════════════════
const RUIDO_OBRAS_OCULTAS = ['Blue Coast'];
// Reclamações: nenhuma obra oculta. Blue Coast e Horizon têm reclamação real
// registrada e voltaram ao gráfico para que ele feche com o card (11).
const REC_OBRAS_OCULTAS   = [];

// Rótulo de mês pt-BR a partir de "YYYY-MM" (ex.: Jan/26). Global: usado tanto
// pelos gráficos da aba Resíduos quanto pelo "Resíduos Gerados kg/m² por obra".
const MESES_RES = ['', 'Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
const labelMesRes = m => {
  const pt = String(m).split('-');
  return (MESES_RES[parseInt(pt[1])] || pt[1]) + '/' + pt[0].slice(2);
};

// ═══════════════════════════════════════════════════════════
//  Rótulo de dados em barras (reaproveitado por vários gráficos)
// ═══════════════════════════════════════════════════════════
// Escreve o valor no topo (barra vertical) ou à direita (barra horizontal).
// Feito à mão em vez de chartjs-plugin-datalabels para não acrescentar
// dependência de CDN — o dashboard roda em rede interna.
function barLabels(fmt, opts) {
  opts = opts || {};
  return {
    id: 'barLabels_' + (opts.id || Math.random().toString(36).slice(2)),
    afterDatasetsDraw(chart) {
      const { ctx } = chart;
      const horizontal = chart.options.indexAxis === 'y';
      ctx.save();
      ctx.font = opts.font || '600 10px "Segoe UI", Inter, system-ui, sans-serif';
      ctx.fillStyle = opts.color || '#444441';
      ctx.textAlign = horizontal ? 'left' : 'center';
      ctx.textBaseline = horizontal ? 'middle' : 'bottom';
      chart.data.datasets.forEach((ds, di) => {
        if (opts.dataset != null && di !== opts.dataset) return;  // só a série pedida
        if (opts.dataset == null && ds.type === 'line') return;   // não rotula linha de meta
        const meta = chart.getDatasetMeta(di);
        if (meta.hidden) return;
        meta.data.forEach((bar, i) => {
          const v = ds.data[i];
          if (v == null || v === 0) return;
          const txt = fmt(v, i, ds);
          if (!txt) return;
          if (horizontal) ctx.fillText(txt, bar.x + 4, bar.y);
          else ctx.fillText(txt, bar.x, bar.y - 3);
        });
      });
      ctx.restore();
    }
  };
}

const SENNA_RE = /senna/i;

// ═══════════════════════════════════════════════════════════
//  Resíduos Gerados kg/m² — dois cards lado a lado
// ═══════════════════════════════════════════════════════════
// Senna Tower tem meta própria (50 kg/m², contra 36,6 das demais). Misturar as
// duas réguas no mesmo eixo distorce a leitura, então ela fica num card ao lado
// em vez de escondida atrás de um botão.
//
// AGREGAÇÃO (definida em 15/09/2026):
//   intensidade do obra-mês      = SOMA das linhas daquele mês
//   barra da obra no consolidado = MÉDIA dessas intensidades mensais
//
// Cada linha da fonte é um tipo de resíduo e já vem como "kg do tipo ÷ área do
// mês", com a MESMA área em todas as linhas do mês — somar devolve kg total ÷
// área, a intensidade real. A média das linhas, usada até 15/09, dividia esse
// valor pelo número de tipos lançados: Blue Coast jan/26 mostrava 16,74 em vez
// de 83,69 porque tinha 5 tipos, e a obra "melhorava" ao separar melhor o
// resíduo. No consolidado usamos a MÉDIA das intensidades mensais (não a soma
// delas): somar meses acumularia e deixaria de ser comparável com a meta, que é
// mensal. Ver memórias "kgm2-agregacao-media-vs-soma" e
// "kgm2-soma-acumula-entre-meses".
const RES_COR_BARRA = '#006B6B';
const RES_COR_META  = '#EF9F27';
const resFmtKg = v => v.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const resMetaDe = o => (D.meta_residuos && D.meta_residuos.por_obra && D.meta_residuos.por_obra[o]) || 36.6;
const resIntensidadeMes = arr => (arr && arr.length) ? arr.reduce((a, b) => a + b, 0) : null;

// Linhas válidas agrupadas por obra × mês (guarda os valores individuais)
function resLinhasPorObraMes(anoF) {
  const linhas = {};
  // Recorte de período sobre o universo de meses COM lançamento de kg/m²
  const mesesSet = mesesDoPeriodo(
    anoF !== undefined ? anoF : periodoGlobal(),
    (D.residuos_raw || []).filter(r => r.kg_m2 > 0).map(r => r.mes_ano)
  );
  (D.residuos_raw || []).forEach(r => {
    if (!(r.kg_m2 > 0)) return;
    if (!mesesSet.has(r.mes_ano)) return;
    const o = String(r.obra || '').trim();
    if (!o) return;
    if (!linhas[o]) linhas[o] = {};
    (linhas[o][r.mes_ano] = linhas[o][r.mes_ano] || []).push(r.kg_m2);
  });
  return linhas;
}
function resMeses6(anoF) {
  const todos = (D.residuos_raw || []).filter(r => r.kg_m2 > 0).map(r => r.mes_ano);
  // Aplica o período primeiro e só depois pega os 6 últimos: com "2025"
  // selecionado, mostra os 6 últimos meses DE 2025, não os 6 últimos globais.
  const ms = [...mesesDoPeriodo(anoF !== undefined ? anoF : periodoGlobal(), todos)].sort();
  return ms.slice(-6);
}

// Gráfico mês a mês de UMA obra.
// `comMeta` liga a linha tracejada da meta: o card da Senna Tower usa, e o card
// principal filtrado por empreendimento não (pedido da área — ali a meta polui
// a leitura da evolução mensal).
function resGraficoMensal(canvasId, obra, serie, meses, badgeEl, comMeta) {
  const meta = resMetaDe(obra);
  if (badgeEl) {
    badgeEl.textContent = `${meses.length} ${meses.length !== 1 ? 'meses' : 'mês'}`
      + (comMeta ? ` · meta ${String(meta).replace('.', ',')} kg/m²` : ' com lançamento');
  }
  makeChart(canvasId, {
    type: 'bar',
    data: {
      labels: meses.map(labelMesRes),
      datasets: [
        {
          label: 'kg/m²',
          data: meses.map(m => +resIntensidadeMes(serie[m]).toFixed(2)),
          backgroundColor: RES_COR_BARRA, borderRadius: 4, order: 2,
        }
      ].concat(comMeta ? [{
        label: `Meta (${String(meta).replace('.', ',')})`,
        data: meses.map(() => meta),
        type: 'line',
        borderColor: RES_COR_META, borderWidth: 2, borderDash: [6, 4],
        pointRadius: 0, fill: false, tension: 0, order: 1,
      }] : [])
    },
    options: {
      responsive: true,
      layout: { padding: { top: 18 } },
      plugins: {
        legend: { display: !!comMeta, position: 'top', labels: { boxWidth: 12, padding: 10 } },
        tooltip: {
          callbacks: {
            label: ctx => ctx.dataset.type === 'line'
              ? ` Meta: ${meta} kg/m²`
              : ` ${resFmtKg(ctx.raw)} kg/m²`,
            afterBody: items => {
              if (!items.length) return '';
              const n = (serie[meses[items[0].dataIndex]] || []).length;
              return ['', `soma de ${n} tipo${n !== 1 ? 's' : ''} de resíduo`];
            }
          }
        }
      },
      scales: { y: { beginAtZero: true, title: { display: true, text: 'kg/m²', font: { size: 12.1 } } } }
    },
    plugins: [barLabels(v => resFmtKg(v), { id: canvasId })]
  });
}

// ─── Card principal: demais obras (ou a obra filtrada, mês a mês) ───
function renderResBarras(obraF, anoF) {
  if (anoF === undefined) anoF = periodoGlobal();
  const badge = document.getElementById('res-barras-badge');
  const titulo = document.getElementById('res-barras-titulo');
  const limpar = msg => {
    if (chartReg['chart-res-barras']) { chartReg['chart-res-barras'].destroy(); delete chartReg['chart-res-barras']; }
    if (badge) badge.textContent = msg;
  };
  if (obraF === undefined) obraF = (document.getElementById('filter-obra-global') || {}).value || '';

  const linhas = resLinhasPorObraMes(anoF);
  const meses6 = resMeses6(anoF);
  if (!meses6.length) { limpar('sem lançamento no período'); return; }

  // Obra filtrada no topo → mês a mês daquela obra
  if (obraF) {
    const serie = linhas[obraF] || {};
    const meses = meses6.filter(m => (serie[m] || []).length);
    if (titulo) titulo.textContent = `Resíduos Gerados kg/m² — ${obraF}`;
    if (!meses.length) { limpar(`sem lançamento · ${obraF}`); return; }
    resGraficoMensal('chart-res-barras', obraF, serie, meses, badge, false);
    return;
  }

  // ─── Consolidado: uma barra por obra, SEM a Senna Tower ───
  if (titulo) titulo.textContent = 'Resíduos Gerados kg/m² por obra';

  const mediaDe = o => {
    const mensais = meses6
      .map(m => resIntensidadeMes((linhas[o] || {})[m]))
      .filter(v => v != null);
    return mensais.length ? mensais.reduce((a, b) => a + b, 0) / mensais.length : null;
  };

  const obras = Object.keys(linhas)
    .filter(o => !SENNA_RE.test(o) && mediaDe(o) != null)
    .sort((a, b) => mediaDe(b) - mediaDe(a));
  if (!obras.length) { limpar('sem lançamento no período'); return; }

  const presentes = new Set();
  obras.forEach(o => meses6.forEach(m => { if ((linhas[o] || {})[m]) presentes.add(m); }));
  const ord = [...presentes].sort();
  if (badge) {
    badge.textContent = ord.length > 1
      ? `média ${labelMesRes(ord[0])} – ${labelMesRes(ord[ord.length - 1])} · meta 36,6 kg/m²`
      : `${labelMesRes(ord[0])} · meta 36,6 kg/m²`;
  }

  makeChart('chart-res-barras', {
    type: 'bar',
    data: {
      labels: obras,
      datasets: [
        {
          label: 'kg/m²',
          data: obras.map(o => +mediaDe(o).toFixed(2)),
          backgroundColor: RES_COR_BARRA, borderRadius: 4, order: 2,
        },
        {
          label: 'Meta (36,6)', data: obras.map(o => resMetaDe(o)), type: 'line',
          stepped: 'middle', borderColor: RES_COR_META, borderWidth: 2, borderDash: [6, 4],
          pointRadius: 0, fill: false, order: 1,
        }
      ]
    },
    options: {
      responsive: true,
      layout: { padding: { top: 18 } },
      plugins: {
        legend: { display: true, position: 'top', labels: { boxWidth: 12, padding: 10 } },
        tooltip: {
          callbacks: {
            label: ctx => ctx.dataset.type === 'line'
              ? ` Meta da obra: ${resMetaDe(ctx.label)} kg/m²`
              : ` ${resFmtKg(ctx.raw)} kg/m² (média)`,
            afterBody: items => {
              if (!items.length) return '';
              const o = items[0].label;
              const det = meses6.filter(m => (linhas[o] || {})[m])
                .map(m => `${labelMesRes(m)}: ${resIntensidadeMes(linhas[o][m]).toFixed(1)}`);
              return det.length ? [''].concat(det) : '';
            }
          }
        }
      },
      scales: {
        x: { ticks: { autoSkip: false, maxRotation: 30 } },
        y: { beginAtZero: true, title: { display: true, text: 'kg/m²', font: { size: 12.1 } } }
      }
    },
    plugins: [barLabels(v => resFmtKg(v), { id: 'resobra' })]
  });
}

// ─── Card ao lado: Senna Tower, mês a mês, com a meta de 50 ───
function renderResSenna(obraF, anoF) {
  if (anoF === undefined) anoF = periodoGlobal();
  const badge = document.getElementById('res-senna-badge');
  const acum = document.getElementById('res-senna-acum');
  const canvas = document.getElementById('chart-res-senna');
  if (!canvas) return;
  if (obraF === undefined) obraF = (document.getElementById('filter-obra-global') || {}).value || '';

  const limpar = msg => {
    if (chartReg['chart-res-senna']) { chartReg['chart-res-senna'].destroy(); delete chartReg['chart-res-senna']; }
    if (badge) badge.textContent = msg;
    if (acum) acum.hidden = true;
  };

  const linhas = resLinhasPorObraMes(anoF);
  const obra = Object.keys(linhas).find(o => SENNA_RE.test(o))
            || (D.obras || []).find(o => SENNA_RE.test(o)) || '';
  if (!obra) { limpar('obra não encontrada'); return; }

  // Respeita o filtro global: com outra obra selecionada, a Senna sai de cena
  if (obraF && !SENNA_RE.test(obraF)) { limpar(`fora do filtro · ${obraF}`); return; }

  const serie = linhas[obra] || {};
  const meses = resMeses6(anoF).filter(m => (serie[m] || []).length);
  if (!meses.length) { limpar('sem lançamento no período'); return; }

  // Acumulado: média das intensidades de TODOS os meses lançados (não só os 6)
  if (acum) {
    const todos = Object.keys(serie).map(m => resIntensidadeMes(serie[m])).filter(v => v != null);
    if (todos.length) {
      acum.textContent = `acumulado: ${resFmtKg(todos.reduce((a, b) => a + b, 0) / todos.length)} kg/m²`;
      acum.hidden = false;
    } else {
      acum.hidden = true;
    }
  }

  resGraficoMensal('chart-res-senna', obra, serie, meses, badge, true);
}


// ═══════════════════════════════════════════════════════════
//  Emergências Ambientais — tabela detalhada
// ═══════════════════════════════════════════════════════════
function renderEmergencias(obraF, anoF) {
  if (anoF === undefined) anoF = periodoGlobal();
  const all = D.emergencias_lista || [];
  let data = all;
  if (obraF) data = data.filter(r => r.obra === obraF);
  // mes_ano das emergências tem DIA ('2026-04-14'); o campo `mes` é o mesmo
  // valor truncado no mês. Comparar direto o mes_ano fazia o filtro recortar
  // por dia, não por mês.
  if (anoF && anoF !== 'all') {
    data = data.filter(r => noPeriodo(r.mes || r.mes_ano, anoF));
  }

  const tblEl = document.getElementById('emerg-table-content');
  const badge = document.getElementById('emerg-table-badge');
  if (!tblEl) return;

  const escapeHtml = s => String(s == null ? '' : s).replace(/[&<>"']/g, c =>
    ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]));

  if (!data.length) {
    tblEl.innerHTML = `
      <thead><tr>
        <th class="col-mes">Mês/Ano</th><th>Obra</th><th>Fonte</th><th>FMC</th>
        <th class="col-ocorr">Ocorrência</th><th class="col-status">Status/Ações</th><th>Categoria</th>
      </tr></thead>
      <tbody><tr class="empty-row"><td colspan="7">Nenhum registro de emergência${obraF ? ' para ' + obraF : ''}</td></tr></tbody>`;
    if (badge) badge.textContent = obraF ? `0 registros · ${obraF}` : '0 registros';
    return;
  }

  let html = `
    <thead><tr>
      <th class="col-mes">Mês/Ano</th>
      <th>Obra</th>
      <th>Fonte</th>
      <th>FMC</th>
      <th class="col-ocorr">Ocorrência</th>
      <th class="col-status">Status/Ações</th>
      <th>Categoria</th>
    </tr></thead>
    <tbody>`;
  data.forEach(r => {
    html += `<tr>
      <td class="col-mes">${escapeHtml(r.mes_label) || '—'}</td>
      <td>${r.obra ? `<span class="obra-tag">${escapeHtml(r.obra)}</span>` : '—'}</td>
      <td>${escapeHtml(r.fonte) || '—'}</td>
      <td class="col-fmc">${escapeHtml(r.fmc) || '—'}</td>
      <td class="col-ocorr">${escapeHtml(r.ocorrencia) || '—'}</td>
      <td class="col-status">${escapeHtml(r.status) || '—'}</td>
      <td>${escapeHtml(r.categoria) || '—'}</td>
    </tr>`;
  });
  html += '</tbody>';
  tblEl.innerHTML = html;

  if (badge) {
    const obrasDistintas = new Set(data.map(r => r.obra).filter(Boolean));
    badge.textContent = obraF
      ? `${data.length} registro${data.length !== 1 ? 's' : ''} · ${obraF}`
      : `${data.length} registros · ${obrasDistintas.size} obra${obrasDistintas.size !== 1 ? 's' : ''}`;
  }
}

// ═══════════════════════════════════════════════════════════
//  Monitoramento de ruído — dB por obra × mês (grouped bar)
// ═══════════════════════════════════════════════════════════
function renderMonitoramentoRuido(obraF, anoF) {
  if (anoF === undefined) anoF = periodoGlobal();
  let all = (D.ruido_obra_mes || []).filter(r => !RUIDO_OBRAS_OCULTAS.includes(r.obra));
  if (obraF) all = all.filter(r => r.obra === obraF);
  // Período aplicado ANTES do corte dos 6 meses do eixo
  const mesesPer = mesesDoPeriodo(anoF, (D.ruido_obra_mes || []).map(r => r.mes_ano));
  all = all.filter(r => mesesPer.has(r.mes_ano));

  const badge = document.getElementById('monitor-ruido-badge');
  if (!all.length) {
    if (chartReg['chart-monitor-ruido']) { chartReg['chart-monitor-ruido'].destroy(); delete chartReg['chart-monitor-ruido']; }
    if (badge) badge.textContent = 'sem medições · Itajaí: 85 dB · BC: 60 dB';
    return;
  }

  // Últimos 6 meses no eixo, todas as obras com dados no eixo X
  const mesesAll = [...new Set(all.map(r => r.mes_ano))].sort();
  // `all` já vem recortado pelo período; aqui é só o teto de 6 colunas do eixo
  const meses = mesesAll.slice(-6);
  const mesesSet = new Set(meses);
  const filt = all.filter(r => mesesSet.has(r.mes_ano));

  // Eixo X em ordem ALFABÉTICA (a pedido da área — antes era por dB médio).
  // Só entram obras COM medição no recorte; obra sem dado não ocupa espaço no
  // eixo e passa a aparecer sozinha assim que tiver medição lançada.
  let obras = [...new Set(filt.map(r => r.obra))]
    .filter(o => !RUIDO_OBRAS_OCULTAS.includes(o))
    .sort((a, b) => a.localeCompare(b, 'pt-BR'));
  if (obraF) obras = obras.filter(o => o === obraF);

  // Map de label por mes_ano
  const lblMap = {};
  filt.forEach(r => { lblMap[r.mes_ano] = r.mes_label; });

  // Paleta dos 6 meses (ordem cronológica do mais antigo ao mais recente):
  //   25-out  →  Amarelo/Lima       #C5C639
  //   25-nov  →  Verde teal         #006B6B
  //   25-dez  →  Azul petróleo      #005E70
  //   26-jan  →  Azul acinzentado   #336B87
  //   26-fev  →  Rosa muted         #CE8686
  //   26-mar  →  Verde escuro       #00484C
  const palette = ['#C5C639', '#006B6B', '#005E70', '#336B87', '#CE8686', '#00484C'];

  // Um dataset por mês
  const datasets = meses.map((m, idx) => {
    const dataByObra = {};
    filt.filter(r => r.mes_ano === m).forEach(r => { dataByObra[r.obra] = r.db; });
    return {
      label: lblMap[m] || m,
      data: obras.map(o => dataByObra[o] != null ? dataByObra[o] : null),
      backgroundColor: palette[idx % palette.length],
      borderRadius: 3,
      maxBarThickness: 28,
      categoryPercentage: 0.85,
      barPercentage: 0.95,
      order: idx + 10,
    };
  });

  // Linhas de limite — não entram na agrupamento das barras
  datasets.push({
    label: 'Limite Itajaí (85 dB)',
    data: obras.map(() => 85),
    type: 'line',
    borderColor: '#EBB626',
    backgroundColor: '#EBB626',
    borderWidth: 2,
    borderDash: [8, 4],
    pointRadius: 0,
    fill: false,
    order: 0,
  });
  datasets.push({
    label: 'Limite BC (60 dB)',
    data: obras.map(() => 60),
    type: 'line',
    borderColor: '#C17979',
    backgroundColor: '#C17979',
    borderWidth: 2,
    borderDash: [3, 3],
    pointRadius: 0,
    fill: false,
    order: 1,
  });

  // Plugin: escreve o valor em dB no topo de cada barra (estilo PBI)
  const valuePlugin = {
    id: 'monitorRuidoLabels',
    afterDatasetsDraw(chart) {
      const { ctx } = chart;
      ctx.save();
      ctx.font = '500 9px "Segoe UI", Inter, system-ui, sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'bottom';
      chart.data.datasets.forEach((ds, dsIdx) => {
        if (ds.type === 'line') return;
        const meta = chart.getDatasetMeta(dsIdx);
        if (meta.hidden) return;
        meta.data.forEach((bar, i) => {
          const v = ds.data[i];
          if (v == null) return;
          ctx.fillStyle = '#444441';
          ctx.fillText(v, bar.x, bar.y - 2);
        });
      });
      ctx.restore();
    }
  };

  makeChart('chart-monitor-ruido', {
    type: 'bar',
    data: { labels: obras, datasets },
    options: {
      responsive: true,
      layout: { padding: { top: 18, right: 60 } },
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: {
          position: 'top',
          labels: { boxWidth: 14, padding: 10, font: { size: 11 } }
        },
        tooltip: {
          backgroundColor: '#141414',
          titleColor: '#FFFFFF',
          bodyColor: '#FFFFFF',
          borderColor: '#1A3668',
          borderWidth: 1,
          padding: 10,
          callbacks: {
            label: ctx => {
              if (ctx.dataset.type === 'line') return ` ${ctx.dataset.label}`;
              const v = ctx.raw;
              if (v == null) return ` ${ctx.dataset.label}: sem dado`;
              let s = '🟢 OK';
              if (v > 85) s = '🔴 Acima Itajaí (>85)';
              else if (v > 60) s = '🟡 Acima BC (>60)';
              return [` ${ctx.dataset.label}: ${v} dB`, ` ${s}`];
            }
          }
        }
      },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 11 }, color: '#1A3668' } },
        y: { beginAtZero: true, suggestedMax: 95, title: { display: true, text: 'dB', font: { size: 12.1 }, color: '#1A3668' }, grid: { color: '#EBEBEB' }, ticks: { color: '#1A3668' } }
      }
    },
    plugins: [valuePlugin]
  });

  if (badge) {
    const periodoLbl = `${lblMap[meses[0]] || meses[0]} → ${lblMap[meses[meses.length - 1]] || meses[meses.length - 1]}`;
    badge.textContent = `${periodoLbl}${obraF ? ' · ' + obraF : ''} · Itajaí 85 / BC 60`;
  }
}

// Paleta de status compartilhada pelos dois gráficos de Plano de Ação
const PA_CORES = {
  'Realizado': '#005E70', 'Em Andamento': '#89B1C5',
  'Pendente': '#888780', 'Atrasado': '#CE8686',
};
const PA_STATUS_INVALIDO = k => {
  const s = String(k || '').trim().toLowerCase();
  return !s || s === 'nan' || s === '<na>' || s === 'none' || s === 'null';
};

// Plugin: escreve "qtde (pct%)" dentro de cada fatia da rosca. Feito à mão em
// vez de usar chartjs-plugin-datalabels pra não acrescentar dependência de CDN
// ao dashboard, que roda em rede interna.
const paDonutLabels = {
  id: 'paDonutLabels',
  afterDatasetsDraw(chart) {
    const { ctx } = chart;
    const ds = chart.data.datasets[0];
    if (!ds) return;
    const total = ds.data.reduce((a, b) => a + b, 0);
    if (!total) return;
    const meta = chart.getDatasetMeta(0);
    ctx.save();
    ctx.textBaseline = 'middle';
    // Rótulo SEMPRE fora do arco, com um traço ligando à fatia. Dentro do anel
    // ele era cortado quando a rosca fica pequena ou a fatia é estreita.
    ctx.font = '600 11px "Segoe UI", Inter, system-ui, sans-serif';
    meta.data.forEach((arc, i) => {
      const v = ds.data[i];
      if (!v) return;
      const pct = Math.round(v / total * 100);
      const txt = v + ' (' + pct + '%)';
      const ang = (arc.startAngle + arc.endAngle) / 2;
      const cos = Math.cos(ang), sin = Math.sin(ang);
      const x0 = arc.x + cos * arc.outerRadius;
      const y0 = arc.y + sin * arc.outerRadius;
      const x1 = arc.x + cos * (arc.outerRadius + 9);
      const y1 = arc.y + sin * (arc.outerRadius + 9);
      ctx.strokeStyle = '#B4B2A9';
      ctx.lineWidth = 1;
      ctx.beginPath(); ctx.moveTo(x0, y0); ctx.lineTo(x1, y1); ctx.stroke();
      const dir = cos >= 0 ? 1 : -1;
      ctx.textAlign = dir > 0 ? 'left' : 'right';
      ctx.fillStyle = '#2C2C2A';
      ctx.fillText(txt, x1 + dir * 3, y1);
    });
    ctx.restore();
  }
};

function renderPaStatus(obraF, anoF) {
  if (anoF === undefined) anoF = periodoGlobal();
  // ATENÇÃO — este é o ÚNICO gráfico que não reage ao filtro de Período.
  // Dim_Plano_acao não tem data de registro da ação, só o campo Prazo (que é
  // data futura de conclusão, não do fato). Filtrar por prazo responderia
  // "ações que vencem no período", que é outra pergunta. Em vez de filtrar
  // errado em silêncio, o card avisa no badge que está fora do recorte.
  const st = obraF ? (D.pa_por_obra[obraF] || {}) : (D.pa_status || {});
  const entries = Object.entries(st)
    .filter(([k, v]) => !PA_STATUS_INVALIDO(k) && v > 0)
    .sort(([, a], [, b]) => b - a);

  const badge = document.getElementById('pa-status-badge');
  if (!entries.length) {
    if (chartReg['chart-pa-status']) { chartReg['chart-pa-status'].destroy(); delete chartReg['chart-pa-status']; }
    if (badge) badge.textContent = obraF ? 'sem ações · ' + obraF : 'sem ações';
    return;
  }

  const totalPa = entries.reduce((a, [, v]) => a + v, 0);
  // Avisa quando há período selecionado, porque este card o ignora (ver acima)
  const avisoPeriodo = (anoF && anoF !== 'all') ? ' · histórico completo (sem data na fonte)' : '';
  if (badge) badge.textContent = totalPa + (totalPa !== 1 ? ' ações' : ' ação')
    + (obraF ? ' · ' + obraF : '') + avisoPeriodo + ' · passe o mouse para abrir por obra';

  // Ações do recorte, para abrir cada fatia por obra dentro do tooltip
  let acoes = (D.plano_acao_lista || []).filter(r => r.status && !PA_STATUS_INVALIDO(r.status));
  if (obraF) acoes = acoes.filter(r => r.obra === obraF);
  const mesmoStatus = (a, b) => String(a).trim().toLowerCase() === String(b).trim().toLowerCase();

  makeChart('chart-pa-status', {
    type: 'doughnut',
    data: {
      labels: entries.map(([k]) => k),
      datasets: [{
        data: entries.map(([, v]) => v),
        backgroundColor: entries.map(([k]) => PA_CORES[k] || '#B4B2A9'),
        borderWidth: 2, borderColor: '#fff'
      }]
    },
    options: {
      responsive: true,
      layout: { padding: { top: 14, bottom: 6, left: 58, right: 58 } },
      plugins: {
        legend: { position: 'bottom', labels: { boxWidth: 12, padding: 12 } },
        tooltip: {
          backgroundColor: '#141414',
          titleColor: '#FFFFFF',
          bodyColor: '#EDEDED',
          padding: 10,
          displayColors: false,
          callbacks: {
            label: ctx => {
              const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
              const pct = total > 0 ? Math.round(ctx.raw / total * 100) : 0;
              return ' ' + ctx.label + ': ' + ctx.raw + ' (' + pct + '%)';
            },
            // Abre a fatia por obra: quantas ações cada obra tem naquele status
            // e o que são (indicador · ação · responsável · prazo).
            afterBody: items => {
              if (!items.length) return '';
              const status = items[0].label;
              const doStatus = acoes.filter(r => mesmoStatus(r.status, status));
              if (!doStatus.length) return '';

              const porObra = {};
              doStatus.forEach(r => {
                const o = r.obra || '(sem obra)';
                (porObra[o] = porObra[o] || []).push(r);
              });

              const linhas = ['', '── por obra ──'];
              Object.keys(porObra).sort((a, b) => a.localeCompare(b, 'pt-BR')).forEach(o => {
                const its = porObra[o];
                linhas.push(o + ': ' + its.length + (its.length !== 1 ? ' ações' : ' ação'));
                its.slice(0, 3).forEach(r => {
                  const det = [r.indicador, r.acao].filter(Boolean).join(' · ');
                  if (det) linhas.push('   • ' + (det.length > 58 ? det.slice(0, 57) + '…' : det));
                  const rp = [r.responsavel, r.prazo ? 'prazo ' + r.prazo : ''].filter(Boolean).join(' · ');
                  if (rp) linhas.push('     ' + rp);
                });
                if (its.length > 3) linhas.push('   … e mais ' + (its.length - 3));
              });
              return linhas;
            }
          }
        }
      }
    },
    plugins: [paDonutLabels]
  });
}

// Classifica a ocorrência num assunto a partir de palavras-chave do texto — a
// fonte (Dim_reclamacoes_vizinhos) não tem coluna de assunto preenchida.
// Usado pelo tooltip do gráfico e pelo painel de maiores incidências.
const REC_ASSUNTOS = [
  ['Ruído',            /bar[uú]lho|ru[ií]do|som/i],
  ['Poeira',           /poeira|p[oó]\b/i],
  ['Sujeira',          /sujeir|sujidad|limpez|cimento|res[ií]duo/i],
  ['Dano a terceiros', /dano|danific|ve[ií]culo|carro|parabris|conserto|reparad/i],
  ['Águas pluviais',   /pluvia|[aá]gua|transbord/i],
  ['Queda de objeto',  /caiu|queda|objeto/i],
];
function assuntoDe(txt) {
  const t = String(txt || '');
  for (const [nome, re] of REC_ASSUNTOS) if (re.test(t)) return nome;
  return 'Outros';
}

// Reclamações de Vizinhos por Obra.
// Conta a partir de D.reclamacoes_lista (linhas reais) em vez de D.rec_por_obra,
// para que o gráfico e o painel de detalhes ao lado usem exatamente a mesma base.
// Obras em REC_OBRAS_OCULTAS não entram (pedido da área).
function recFiltradas(obraF, anoF) {
  let lista = (D.reclamacoes_lista || []).filter(r => !REC_OBRAS_OCULTAS.includes(r.obra));
  if (obraF) lista = lista.filter(r => r.obra === obraF);
  const per = anoF !== undefined ? anoF : periodoGlobal();
  if (per && per !== 'all') {
    const mesesSet = mesesDoPeriodo(per, (D.reclamacoes_lista || []).map(r => r.mes_ano));
    lista = lista.filter(r => mesesSet.has(r.mes_ano));
  }
  return lista;
}

function renderRecPorObra(obraF, anoF) {
  if (anoF === undefined) anoF = periodoGlobal();
  const lista = recFiltradas(obraF, anoF);

  // Conta por obra; obra vazia vira "(sem obra informada)" pra não sumir do total
  const SEM_OBRA = '(sem obra informada)';
  const cont = {};
  lista.forEach(r => {
    const o = r.obra || SEM_OBRA;
    cont[o] = (cont[o] || 0) + 1;
  });
  const entries = Object.entries(cont).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'pt-BR'));

  const badge = document.getElementById('rec-obra-badge');
  if (!entries.length) {
    if (chartReg['chart-rec-obra']) { chartReg['chart-rec-obra'].destroy(); delete chartReg['chart-rec-obra']; }
    if (badge) badge.textContent = '0 reclamações';
    renderRecDetalhe(obraF, anoF);
    return;
  }

  const total = entries.reduce((s, [, v]) => s + v, 0);
  if (badge) {
    badge.textContent = obraF
      ? total + (total !== 1 ? ' reclamações · ' : ' reclamação · ') + obraF
      : total + (total !== 1 ? ' reclamações · ' : ' reclamação · ') + entries.length + (entries.length !== 1 ? ' obras' : ' obra');
  }

  const colors = entries.map(([o]) => obraF && o === obraF ? '#00484C' : '#006B6B');

  makeChart('chart-rec-obra', {
    type: 'bar',
    data: {
      labels: entries.map(([o]) => o),
      datasets: [{
        label: 'Reclamações',
        data: entries.map(([, v]) => v),
        backgroundColor: colors,
        borderRadius: 4,
        maxBarThickness: 22,
      }]
    },
    options: {
      responsive: true,
      indexAxis: 'y',
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            // Só quantidade e assunto — sem data e sem o texto da ocorrência
            // (esses ficam na tabela do painel ao lado).
            label: ctx => ' ' + ctx.raw + (ctx.raw !== 1 ? ' reclamações' : ' reclamação'),
            afterBody: items => {
              if (!items.length) return '';
              const o = items[0].label;
              const freq = {};
              lista.filter(r => (r.obra || SEM_OBRA) === o)
                   .forEach(r => { const a = assuntoDe(r.ocorrencia); freq[a] = (freq[a] || 0) + 1; });
              const ord = Object.entries(freq).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'pt-BR'));
              return ord.length ? [''].concat(ord.map(([a, c]) => a + ': ' + c)) : '';
            }
          }
        }
      },
      layout: { padding: { right: 22 } },
      scales: {
        x: { display: false, beginAtZero: true },
        y: { ticks: { font: { size: 11 } }, grid: { display: false } }
      }
    },
    plugins: [barLabels(v => String(v), { id: 'recobra' })]
  });

  renderRecDetalhe(obraF, anoF);
}

// Detalhe das reclamações, em dois painéis lado a lado:
//   #rec-rank    -> maiores incidências por assunto
//   #rec-detalhe -> assunto e ação por obra
function renderRecDetalhe(obraF, anoF) {
  if (anoF === undefined) anoF = periodoGlobal();
  const elTab  = document.getElementById('rec-detalhe');
  const bRank  = document.getElementById('rec-rank-badge');
  const bTab   = document.getElementById('rec-detalhe-badge');
  if (!elTab && !document.getElementById('chart-rec-rank')) return;
  const lista = recFiltradas(obraF, anoF);

  if (!lista.length) {
    const vazio = '<div class="empty-msg">Sem reclamações no recorte</div>';
    if (chartReg['chart-rec-rank']) { chartReg['chart-rec-rank'].destroy(); delete chartReg['chart-rec-rank']; }
    if (elTab) elTab.innerHTML = vazio;
    if (bRank) bRank.textContent = '—';
    if (bTab) bTab.textContent = '—';
    return;
  }

  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  // ─── Painel 1: maiores incidências por assunto ───
  const freq = {};
  lista.forEach(r => { const a = assuntoDe(r.ocorrencia); freq[a] = (freq[a] || 0) + 1; });
  const ranking = Object.entries(freq).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'pt-BR'));
  const maxF = ranking.length ? ranking[0][1] : 1;
  const totalR = lista.length;

  if (bRank) bRank.textContent = ranking.length + (ranking.length !== 1 ? ' assuntos' : ' assunto') + (obraF ? ' · ' + obraF : '');
  if (document.getElementById('chart-rec-rank')) {
    // Card na largura toda: ~30px por assunto, mínimo 320px — mostra todos os
    // assuntos sem espremer as barras
    const rankWrap = document.getElementById('rec-rank-wrap');
    if (rankWrap) rankWrap.style.height = Math.max(320, ranking.length * 30 + 30) + 'px';
    makeChart('chart-rec-rank', {
      type: 'bar',
      data: {
        labels: ranking.map(([a]) => a),
        datasets: [{
          label: 'Reclamações',
          data: ranking.map(([, c]) => c),
          backgroundColor: '#006B6B',
          borderRadius: 4,
          maxBarThickness: 22,
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        indexAxis: 'y',
        layout: { padding: { right: 28 } },
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: ctx => {
                const pct = Math.round(ctx.raw / totalR * 100);
                return ' ' + ctx.raw + (ctx.raw !== 1 ? ' reclamações' : ' reclamação') + ' (' + pct + '% do total)';
              }
            }
          }
        },
        scales: {
          x: { display: false, beginAtZero: true },
          y: { ticks: { font: { size: 11 } }, grid: { display: false } }
        }
      },
      // Rótulo só com a quantidade (o percentual segue no tooltip)
      plugins: [barLabels(v => String(v), { id: 'recrank' })]
    });
  }

  // ─── Painel 2: assunto e ação por obra ───
  // Ordenado por OBRA (alfabética) e, dentro da obra, por data — as reclamações
  // da mesma obra ficam em sequência, com o nome só na primeira linha do grupo.
  if (bTab) {
    const nObras = new Set(lista.map(r => r.obra || '—')).size;
    bTab.textContent = totalR + (totalR !== 1 ? ' registros · ' : ' registro · ') + nObras + (nObras !== 1 ? ' obras' : ' obra');
  }
  if (elTab) {
    const ordenada = lista.slice().sort((a, b) => {
      const oa = a.obra || 'zzz', ob = b.obra || 'zzz';
      const c = oa.localeCompare(ob, 'pt-BR');
      if (c) return c;
      return String(a.mes_ano || '').localeCompare(String(b.mes_ano || ''));
    });

    let h = '<table class="rec-table"><thead><tr>'
      + '<th>Obra</th><th>Mês</th><th>Assunto</th><th>Ocorrência</th><th>Status / ação</th>'
      + '</tr></thead><tbody>';
    let obraAnt = null;
    ordenada.forEach(r => {
      const o = r.obra || '—';
      const novoGrupo = o !== obraAnt;
      h += '<tr' + (novoGrupo ? ' class="rec-grupo"' : '') + '>'
        + '<td class="rec-obra">' + (novoGrupo ? esc(o) : '') + '</td>'
        + '<td>' + esc(r.mes_label || '—') + '</td>'
        + '<td><span class="rec-tag">' + esc(assuntoDe(r.ocorrencia)) + '</span></td>'
        + '<td>' + esc(r.ocorrencia || '—') + '</td>'
        + '<td>' + esc(r.status || '—') + '</td>'
        + '</tr>';
      obraAnt = o;
    });
    h += '</tbody></table>';
    elTab.innerHTML = h;
  }
}

function renderIpaObras(obraF, anoF) {
  if (!anoF) anoF = (document.getElementById('filter-periodo')||{}).value || 'all';

  // Conjunto de meses do período selecionado (mesmo padrão usado no heatmap de conformidade)
  const allMeses = D.conf_meses || [];
  // Janela de calendário única (ver mesesDoPeriodo/janelaPeriodo)
  let mesesSet = (anoF && anoF !== 'all') ? mesesDoPeriodo(anoF, allMeses) : null;

  const tituloEl = document.getElementById('ipa-obras-title');
  // Paleta nova baseada no IPA:
  //   ≥90 → #006B6B (verde teal — excelente)
  //   ≥80 → #336B87 (azul acinzentado — bom)
  //   ≥70 → #C5C639 (amarelo lima — atenção)
  //   <70 → #CE8686 (rosa muted — crítico)
  const corIpa = v => v >= 90 ? '#006B6B' : v >= 80 ? '#336B87' : v >= 70 ? '#C5C639' : '#CE8686';
  const tooltipLabel = ctx => ` IPA: ${ctx.raw} — ${ctx.raw >= 90 ? '🟢 Excelente' : ctx.raw >= 80 ? '🟡 Bom' : ctx.raw >= 70 ? '🟠 Atenção' : '🔴 Crítico'}`;

  let labels, vals;
  if (obraF) {
    // Obra selecionada: histórico mensal (até 12 meses) dessa obra
    let serie = (D.ipa_mensal && D.ipa_mensal[obraF]) || [];
    if (mesesSet) serie = serie.filter(s => mesesSet.has(s.mes));
    if (serie.length) {
      labels = serie.map(s => s.mes_label);
      vals = serie.map(s => s.ipa);
      if (tituloEl) tituloEl.textContent = `IPA — histórico de ${labels.length} meses (${obraF})`;
    } else {
      // Fallback (ex.: dashboard ainda não regerado com o histórico mensal): 1 coluna com o IPA agregado
      labels = [obraF];
      vals = [(D.ipa[obraF] && D.ipa[obraF].ipa) || 0];
      if (tituloEl) tituloEl.textContent = 'IPA por obra — Índice de Performance Ambiental';
    }
  } else {
    // "Todas as obras": uma coluna por obra, respeitando o período selecionado.
    // Sem recorte de período ("Todos"), usa o IPA agregado (histórico completo).
    labels = D.obras || [];
    vals = labels.map(o => {
      const serieObra = (D.ipa_mensal && D.ipa_mensal[o]) || [];
      const filtrada = mesesSet ? serieObra.filter(s => mesesSet.has(s.mes)) : [];
      if (!filtrada.length) return (D.ipa[o] && D.ipa[o].ipa) || 0;
      return +(filtrada.reduce((a, s) => a + s.ipa, 0) / filtrada.length).toFixed(1);
    });
    if (tituloEl) tituloEl.textContent = 'IPA por obra — Índice de Performance Ambiental';
  }

  const colors = vals.map(corIpa);
  makeChart('chart-ipa-obras', {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'IPA', data: vals, backgroundColor: colors, borderRadius: 5 },
        { label: 'Meta 80', data: labels.map(() => 80), type: 'line', borderColor: '#C5C639', borderDash: [5,3], pointRadius: 0, borderWidth: 1.5, fill: false },
        { label: 'Meta 90', data: labels.map(() => 90), type: 'line', borderColor: '#006B6B', borderDash: [5,3], pointRadius: 0, borderWidth: 1.5, fill: false },
      ]
    },
    options: {
      responsive: true,
      layout: { padding: { top: 18 } },
      plugins: { legend: { position: 'top' }, tooltip: { callbacks: { label: tooltipLabel } } },
      scales: { y: { min: 0, max: 100, ticks: { stepSize: 10 }, title: { display: true, text: 'Pontuação', font: { size: 12.1 } } } }
    },
    plugins: [barValueLabels(v => String(v))]
  });
}

// ═══════════════════════════════════════════════════════════
//  Ruído — medições: histórico mensal (mesmo padrão do IPA por obra)
// ═══════════════════════════════════════════════════════════
function renderRuidoObra(obraF, anoF) {
  if (!anoF) anoF = (document.getElementById('filter-periodo')||{}).value || 'all';

  const allMeses = D.conf_meses || [];
  // Janela de calendário única (ver mesesDoPeriodo/janelaPeriodo)
  let mesesSet = (anoF && anoF !== 'all') ? mesesDoPeriodo(anoF, allMeses) : null;

  // Paleta: >85 → rosa (crítico, acima do limite Itajaí) · >60 → amarelo lima (acima BC) · ≤60 → verde teal (ok)
  const corDb = db => db > 85 ? '#CE8686' : db > 60 ? '#C5C639' : '#006B6B';

  let labels = [], vals = [];
  if (obraF) {
    // Obra selecionada: histórico mensal (até 12 meses) de medições dessa obra
    let serie = (D.ruido_mensal && D.ruido_mensal[obraF]) || [];
    if (mesesSet) serie = serie.filter(s => mesesSet.has(s.mes));
    if (serie.length) {
      labels = serie.map(s => s.mes_label);
      vals = serie.map(s => s.db);
    } else {
      // Fallback (sem histórico mensal, ex.: dashboard ainda não regerado): 1 coluna com a média agregada
      const agg = D.ruido_por_obra && D.ruido_por_obra[obraF];
      if (agg && agg.media_db != null) { labels = [obraF]; vals = [agg.media_db]; }
    }
  } else {
    // "Todas as obras": uma coluna por obra, média do período selecionado
    const todasObras = D.obras || [];
    todasObras.forEach(o => {
      const serieObra = (D.ruido_mensal && D.ruido_mensal[o]) || [];
      const filtrada = mesesSet ? serieObra.filter(s => mesesSet.has(s.mes)) : serieObra;
      let v = null;
      if (filtrada.length) v = +(filtrada.reduce((a, s) => a + s.db, 0) / filtrada.length).toFixed(1);
      else {
        const agg = D.ruido_por_obra && D.ruido_por_obra[o];
        if (agg && agg.media_db != null) v = agg.media_db;
      }
      if (v != null) { labels.push(o); vals.push(v); }
    });
  }

  if (!labels.length) {
    makeChart('chart-obra-ruido', {
      type: 'bar',
      data: { labels: [obraF || '—'], datasets: [{ data: [0], backgroundColor: '#D3D1C7', borderRadius: 6 }] },
      options: { responsive: true, plugins: { legend: { display: false }, subtitle: { display: true, text: 'Sem dados de ruído', font: { size: 12.1 } } }, scales: { y: { min: 0, max: 100 } } }
    });
    return;
  }

  makeChart('chart-obra-ruido', {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'dB médio', data: vals, backgroundColor: vals.map(corDb), borderRadius: 6, maxBarThickness: 80, order: 3 },
        { label: 'Limite Itajaí (85 dB)', data: labels.map(() => 85), type: 'line', borderColor: '#CE8686', borderDash: [6,3], pointRadius: 0, borderWidth: 2, fill: false, order: 1 },
        { label: 'Limite BC (60 dB)',    data: labels.map(() => 60), type: 'line', borderColor: '#336B87', borderDash: [6,3], pointRadius: 0, borderWidth: 2, fill: false, order: 2 }
      ]
    },
    options: {
      responsive: true,
      layout: { padding: { top: 18 } },
      plugins: {
        legend: { position: 'top', labels: { boxWidth: 14, padding: 10, font: { size: 11 } } },
        tooltip: {
          callbacks: {
            label: ctx => {
              if (ctx.datasetIndex === 0) {
                const v = ctx.raw;
                let status = '🟢 OK (≤ 60 dB)';
                if (v > 85) status = '🔴 Acima Itajaí (>85)';
                else if (v > 60) status = '🟡 Acima BC (>60)';
                return [` dB médio: ${v}`, ` ${status}`];
              }
              return ` ${ctx.dataset.label}`;
            }
          }
        }
      },
      scales: { y: { min: 40, max: 95, title: { display: true, text: 'dB', font: { size: 12.1 } } } }
    },
    plugins: [barValueLabels(v => String(v))]
  });
}

// ═══════════════════════════════════════════════════════════
//  SGA — pessoas treinadas por ano (gráfico de barras)
// ═══════════════════════════════════════════════════════════
function renderSGAAnual(anoF) {
  if (anoF === undefined) anoF = periodoGlobal();
  const granular = D.sga_por_mes || [];
  const j = janelaPeriodo(anoF);
  let data, anosCheios = [];

  if (!j || !granular.length) {
    // "Todo o período" (ou base antiga, sem o dataset granular): série completa
    data = (D.sga_por_ano || []).slice();
  } else {
    // Dim_Acoes_SGA tem o mês para 2025/2026, mas os anos antigos (2021–2024)
    // vêm consolidados com Mês = 'Todos'. As linhas com mês são recortadas
    // normalmente; as consolidadas só entram quando a janela cobre o ano
    // INTEIRO — recortar um total anual por mês seria inventar precisão.
    const acc = {};
    granular.forEach(r => {
      let entra = false;
      if (r.anual) {
        entra = j.de <= r.ano + '-01' && j.ate >= r.ano + '-12';
        if (entra && !anosCheios.includes(r.ano)) anosCheios.push(r.ano);
      } else {
        entra = !!r.mes && r.mes >= j.de && r.mes <= j.ate;
      }
      if (entra) acc[r.ano] = (acc[r.ano] || 0) + (r.pessoas || 0);
    });
    data = Object.keys(acc).sort().map(a => ({ ano: Number(a), pessoas: acc[a] }));
  }

  const badgeSga = document.getElementById('sga-badge');
  if (!data.length) {
    if (chartReg['chart-sga-anual']) { chartReg['chart-sga-anual'].destroy(); delete chartReg['chart-sga-anual']; }
    if (badgeSga) badgeSga.textContent = 'sem eventos no recorte';
    return;
  }
  const labels = data.map(d => String(d.ano));
  const valores = data.map(d => d.pessoas);
  const total = valores.reduce((s, v) => s + v, 0);
  if (badgeSga) {
    // Com recorte ativo o badge deixa de dizer "acumulado desde 2021" (que
    // seria falso) e mostra a janela; avisa quando algum ano entrou fechado.
    badgeSga.textContent = !j
      ? `total acumulado: ${numBR(total, 0)} pessoas`
      : `${numBR(total, 0)} pessoas · ${faixaPeriodo(anoF)}`
        + (anosCheios.length ? ` · ${anosCheios.join(', ')} sem detalhe mensal` : '');
  }
  const palette = ['#97C459', '#639922', '#3B6D11', '#27500A', '#185FA5', '#0C447C'];
  const colors = labels.map((_, i) => palette[i % palette.length]);

  // Plugin pra rótulo no topo
  const labelPlugin = {
    id: 'sgaAnualLabels',
    afterDatasetsDraw(chart) {
      const { ctx } = chart;
      const ds = chart.getDatasetMeta(0);
      ctx.save();
      ctx.font = '600 12px "Segoe UI", Inter, system-ui, sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'bottom';
      ds.data.forEach((bar, i) => {
        const v = chart.data.datasets[0].data[i];
        if (v == null) return;
        ctx.fillStyle = '#173404';
        ctx.fillText(v.toLocaleString('pt-BR'), bar.x, bar.y - 4);
      });
      ctx.restore();
    }
  };

  makeChart('chart-sga-anual', {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Pessoas treinadas',
        data: valores,
        backgroundColor: colors,
        borderRadius: 4,
        maxBarThickness: 60,
      }]
    },
    options: {
      responsive: true,
      layout: { padding: { top: 18 } },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: ctx => {
              const pct = total > 0 ? Math.round(ctx.raw / total * 100) : 0;
              return ` ${ctx.raw.toLocaleString('pt-BR')} pessoas (${pct}% do acumulado)`;
            }
          }
        }
      },
      scales: {
        x: { grid: { display: false } },
        y: {
          beginAtZero: true,
          title: { display: true, text: 'Pessoas treinadas', font: { size: 12.1 } },
          ticks: { callback: v => v.toLocaleString('pt-BR') },
          grid: { color: 'rgba(0,0,0,0.04)' }
        }
      }
    },
    plugins: [labelPlugin]
  });
  // O badge é preenchido no início da função (respeita o recorte de período).
}

// ═══════════════════════════════════════════════════════════
//  Consumo de Energia — barras + meta média móvel 6m
//  (kWh absoluto + intensidade kWh/m² lado a lado)
// ═══════════════════════════════════════════════════════════
function renderConsumoEnergiaAll() {
  renderConsumoEnergia();
  renderConsumoEnergiaM2();
}

function _drawEnergiaChart(canvasId, badgeId, dados, valueKey, formatter, unit, obraF) {
  if (!dados.length) {
    const badge = document.getElementById(badgeId);
    if (badge) badge.textContent = obraF ? `sem dados · ${obraF}` : 'sem dados';
    if (chartReg[canvasId]) { chartReg[canvasId].destroy(); delete chartReg[canvasId]; }
    return;
  }

  const labels  = dados.map(d => d.mes_label);
  const valores = dados.map(d => d[valueKey]);
  const meta    = dados.map(d => d.meta_6m);

  // Cor das barras "Realizado" por tipo de consumo:
  //   Energia (kWh, kWh/m²) → #006B6B (verde teal)
  //   Água    (m³, m³/m²)   → #005E70 (azul petróleo)
  const barColor = /m³/.test(unit) ? '#005E70' : '#006B6B';

  // Plugin com posicionamento anti-colisão
  // Realizado: acima da barra (centralizado).
  // Meta: à direita do ponto, OFFSET HORIZONTAL = meia barra (sai fora da coluna).
  //       Último mês: rótulo à esquerda do ponto pra não cair fora do canvas.
  const dataLabelPlugin = {
    id: 'energiaLabels',
    afterDatasetsDraw(chart) {
      const { ctx, data, chartArea } = chart;
      const ds = chart.getDatasetMeta(0);
      const dsMeta = chart.getDatasetMeta(1);
      const n = ds.data.length;

      ctx.save();

      // ── 1) Rótulos das barras (Realizado): acima do topo, com pill branco
      //    pra legibilidade quando a linha passa perto do topo da barra
      ctx.font = '600 12px "Segoe UI", Inter, system-ui, sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ds.data.forEach((bar, i) => {
        const v = data.datasets[0].data[i];
        if (v == null) return;
        const txt = v.toLocaleString('pt-BR');
        const yLabel = bar.y - 14;
        const w = ctx.measureText(txt).width + 8;
        // pill branco semitransparente
        ctx.fillStyle = 'rgba(255,255,255,0.92)';
        ctx.fillRect(bar.x - w/2, yLabel - 8, w, 16);
        ctx.fillStyle = '#0C447C';
        ctx.fillText(txt, bar.x, yLabel);
      });

      // ── 2) Rótulos da linha (Meta 6m): lateral à barra, sem colidir com o do topo
      ctx.font = '500 11px "Segoe UI", Inter, system-ui, sans-serif';
      ctx.textBaseline = 'middle';
      dsMeta.data.forEach((pt, i) => {
        const v = data.datasets[1].data[i];
        if (v == null) return;
        const barEl = ds.data[i];
        const barW = (barEl && barEl.width) ? barEl.width : 32;
        const isLast = i === n - 1;
        // Padding lateral + meia-barra → garante que sai fora da coluna
        const dx = barW / 2 + 8;
        const x = isLast ? (pt.x - dx) : (pt.x + dx);
        ctx.textAlign = isLast ? 'right' : 'left';
        const txt = v.toLocaleString('pt-BR');
        // fundo claro pra contraste sobre as gridlines
        const w = ctx.measureText(txt).width + 6;
        const xBg = isLast ? x - w : x;
        ctx.fillStyle = 'rgba(255,255,255,0.85)';
        ctx.fillRect(xBg, pt.y - 7, w, 14);
        ctx.fillStyle = '#173404';
        ctx.fillText(txt, x + (isLast ? -3 : 3), pt.y);
      });

      ctx.restore();
    }
  };

  // Reescreve o dataLabelPlugin para usar o formatter passado
  const _origPlugin = dataLabelPlugin;
  const flexLabelPlugin = {
    id: 'energiaLabels_' + canvasId,
    afterDatasetsDraw(chart) {
      const { ctx } = chart;
      const ds = chart.getDatasetMeta(0);
      const dsMeta = chart.getDatasetMeta(1);
      const n = ds.data.length;
      ctx.save();
      ctx.font = '600 12px "Segoe UI", Inter, system-ui, sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ds.data.forEach((bar, i) => {
        const v = chart.data.datasets[0].data[i];
        if (v == null) return;
        const txt = formatter(v);
        const yLabel = bar.y - 14;
        const w = ctx.measureText(txt).width + 8;
        ctx.fillStyle = 'rgba(255,255,255,0.92)';
        ctx.fillRect(bar.x - w/2, yLabel - 8, w, 16);
        ctx.fillStyle = '#0C447C';
        ctx.fillText(txt, bar.x, yLabel);
      });
      ctx.font = '500 11px "Segoe UI", Inter, system-ui, sans-serif';
      dsMeta.data.forEach((pt, i) => {
        const v = chart.data.datasets[1].data[i];
        if (v == null) return;
        const barEl = ds.data[i];
        const barW = (barEl && barEl.width) ? barEl.width : 32;
        const isLast = i === n - 1;
        const dx = barW / 2 + 8;
        const x = isLast ? (pt.x - dx) : (pt.x + dx);
        ctx.textAlign = isLast ? 'right' : 'left';
        const txt = formatter(v);
        const w = ctx.measureText(txt).width + 6;
        const xBg = isLast ? x - w : x;
        ctx.fillStyle = 'rgba(255,255,255,0.85)';
        ctx.fillRect(xBg, pt.y - 7, w, 14);
        ctx.fillStyle = '#173404';
        ctx.fillText(txt, x + (isLast ? -3 : 3), pt.y);
      });
      ctx.restore();
    }
  };

  makeChart(canvasId, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Realizado', data: valores, backgroundColor: barColor, borderRadius: 3, maxBarThickness: 56, order: 2 },
        { label: 'Meta média 6 meses', data: meta, type: 'line', borderColor: '#173404', backgroundColor: '#173404', borderWidth: 2, pointRadius: 4, pointBackgroundColor: '#173404', tension: 0.25, fill: false, order: 1 }
      ]
    },
    options: {
      responsive: true,
      layout: { padding: { top: 28, right: 60 } },
      plugins: {
        legend: { position: 'bottom', labels: { boxWidth: 14, padding: 18, font: { size: 12 } } },
        tooltip: {
          callbacks: {
            label: ctx => {
              const v = ctx.raw;
              if (ctx.datasetIndex === 0) {
                const m = meta[ctx.dataIndex];
                const diff = v - m;
                const sinal = diff >= 0 ? '+' : '';
                return [` Realizado: ${formatter(v)} ${unit}`, ` vs meta: ${sinal}${formatter(diff)} ${unit}`];
              }
              return ` Meta 6m: ${formatter(v)} ${unit}`;
            }
          }
        }
      },
      scales: {
        x: { grid: { display: false } },
        y: { beginAtZero: true, ticks: { callback: v => formatter(v) }, grid: { color: 'rgba(0,0,0,0.04)' } }
      }
    },
    plugins: [flexLabelPlugin]
  });

  // Badge: comportamento muda se a métrica é total (kWh) ou média (kWh/m²)
  const badge = document.getElementById(badgeId);
  if (!badge) return;
  const obraTxt = obraF ? ` · ${obraF}` : '';
  // Unidades com "/m" no nome (kWh/m², m³/m²...) são intensidades → média.
  // Demais (kWh, m³, ...) são totais.
  const isIntensity = /\/m/.test(unit);
  if (!isIntensity) {
    const totalReal = valores.reduce((a,b) => a+b, 0);
    const totalMeta = meta.reduce((a,b) => a+b, 0);
    const variacao = totalMeta > 0 ? Math.round((totalReal - totalMeta) / totalMeta * 100) : 0;
    const sinal = variacao > 0 ? '+' : '';
    badge.textContent = `total: ${formatter(totalReal)} ${unit} · ${sinal}${variacao}% vs meta${obraTxt}`;
  } else {
    const mediaReal = valores.reduce((a,b) => a+b, 0) / valores.length;
    const mediaMeta = meta.reduce((a,b) => a+b, 0) / meta.length;
    const variacao = mediaMeta > 0 ? Math.round((mediaReal - mediaMeta) / mediaMeta * 100) : 0;
    const sinal = variacao > 0 ? '+' : '';
    badge.textContent = `média: ${formatter(mediaReal)} ${unit} · ${sinal}${variacao}% vs meta${obraTxt}`;
  }
}

// Agrega o dataset raw por mês, respeitando filtro de obra
// kWh: soma; kwh_m2: lê DIRETO da coluna Luz m² (média entre obras se "Todas")
function _computeEnergiaSeries(obraF, setorF) {
  let rows = D.consumo_energia_raw || [];
  if (obraF) rows = rows.filter(r => r.obra === obraF);
  if (setorF) rows = rows.filter(r => r.setor === setorF);
  if (!rows.length) return [];

  const map = {};
  rows.forEach(r => {
    if (!map[r.mes_ano]) {
      map[r.mes_ano] = {
        mes_ano: r.mes_ano, mes_label: r.mes_label,
        kwh: 0, m2Sum: 0, m2Count: 0
      };
    }
    map[r.mes_ano].kwh += r.kwh;
    if (r.kwh_m2 != null) {
      map[r.mes_ano].m2Sum += r.kwh_m2;
      map[r.mes_ano].m2Count += 1;
    }
  });

  const series = Object.keys(map).sort().map(k => {
    const x = map[k];
    return {
      mes_ano: x.mes_ano,
      mes_label: x.mes_label,
      kwh: x.kwh,
      kwh_m2: x.m2Count > 0 ? +(x.m2Sum / x.m2Count).toFixed(2) : null,
    };
  });

  // Médias móveis 6m
  series.forEach((s, i) => {
    const win = series.slice(Math.max(0, i - 5), i + 1);
    s.meta_6m = Math.round(win.reduce((a, b) => a + b.kwh, 0) / win.length);
    const winM2 = win.filter(x => x.kwh_m2 != null);
    s.meta_6m_m2 = winM2.length
      ? +(winM2.reduce((a, b) => a + b.kwh_m2, 0) / winM2.length).toFixed(2)
      : null;
  });

  return series;
}

function _filtraPeriodo(arr, periodoF) {
  // Janela de calendário única — antes era slice(-6)/slice(-12) por posição,
  // que dava uma janela diferente da dos outros gráficos (água termina em
  // 2026-03, conformidade em 2026-05).
  if (periodoF && periodoF !== 'all') {
    return arr.filter(d => noPeriodo(d.mes_ano, periodoF));
  }
  return arr;
}

function renderConsumoEnergia() {
  const obraF    = (document.getElementById('filter-obra-global')||{}).value || '';
  const setorF   = (document.getElementById('filter-energia-setor')||{}).value || '';
  const periodoF = periodoEfetivo('filter-energia-periodo', '12m');
  // Com obra OU setor filtrado, recalcula a partir do raw; sem nenhum dos dois,
  // usa o agregado pré-computado (mesmo resultado, sem recomputar 610 linhas)
  let dados;
  if (obraF || setorF) {
    dados = _filtraPeriodo(_computeEnergiaSeries(obraF, setorF), periodoF);
  } else {
    dados = _filtraPeriodo((D.consumo_energia || []).slice(), periodoF);
  }
  const fmt = v => Math.round(v).toLocaleString('pt-BR');
  _drawEnergiaChart('chart-consumo-energia', 'energia-badge', dados, 'kwh', fmt, 'kWh', obraF);
}

// ═══════════════════════════════════════════════════════════
//  Consumo de Água — barras + meta média móvel 6m
//  (m³ absoluto + intensidade m³/m² lado a lado)
// ═══════════════════════════════════════════════════════════
function _computeAguaSeries(obraF, setorF) {
  let rows = D.consumo_agua_raw || [];
  if (obraF) rows = rows.filter(r => r.obra === obraF);
  if (setorF) rows = rows.filter(r => r.setor === setorF);
  if (!rows.length) return [];

  const map = {};
  rows.forEach(r => {
    if (!map[r.mes_ano]) {
      map[r.mes_ano] = {
        mes_ano: r.mes_ano, mes_label: r.mes_label,
        m3: 0, m3m2Sum: 0, m3m2Count: 0
      };
    }
    map[r.mes_ano].m3 += r.m3;
    if (r.m3_m2 != null) {
      map[r.mes_ano].m3m2Sum += r.m3_m2;
      map[r.mes_ano].m3m2Count += 1;
    }
  });

  const series = Object.keys(map).sort().map(k => {
    const x = map[k];
    return {
      mes_ano: x.mes_ano,
      mes_label: x.mes_label,
      m3: x.m3,
      m3_m2: x.m3m2Count > 0 ? +(x.m3m2Sum / x.m3m2Count).toFixed(2) : null,
    };
  });

  // Médias móveis 6m
  series.forEach((s, i) => {
    const win = series.slice(Math.max(0, i - 5), i + 1);
    const nz = win.filter(x => x.m3 > 0);
    s.meta_6m = nz.length ? Math.round(nz.reduce((a, b) => a + b.m3, 0) / nz.length) : 0;
    const winM2 = win.filter(x => x.m3_m2 != null);
    s.meta_6m_m2 = winM2.length
      ? +(winM2.reduce((a, b) => a + b.m3_m2, 0) / winM2.length).toFixed(2)
      : null;
  });
  return series;
}

function renderConsumoAguaAll() {
  renderConsumoAgua();
  renderConsumoAguaM2();
}

function renderConsumoAgua() {
  const obraF    = (document.getElementById('filter-obra-global')||{}).value || '';
  const setorF   = (document.getElementById('filter-agua-setor')||{}).value || '';
  const periodoF = periodoEfetivo('filter-agua-periodo', '12m');
  let dados;
  if (obraF || setorF) {
    dados = _filtraPeriodo(_computeAguaSeries(obraF, setorF), periodoF);
  } else {
    dados = _filtraPeriodo((D.consumo_agua || []).slice(), periodoF);
  }
  const fmt = v => Math.round(v).toLocaleString('pt-BR');
  _drawEnergiaChart('chart-consumo-agua', 'agua-badge', dados, 'm3', fmt, 'm³', obraF);
}

function renderConsumoAguaM2() {
  const obraF    = (document.getElementById('filter-obra-global')||{}).value || '';
  const setorF   = (document.getElementById('filter-agua-setor')||{}).value || '';
  const periodoF = periodoEfetivo('filter-agua-periodo', '12m');
  let dados;
  if (obraF || setorF) {
    dados = _filtraPeriodo(_computeAguaSeries(obraF, setorF), periodoF)
      .filter(d => d.m3_m2 != null)
      .map(d => ({ ...d, m3_m2: +d.m3_m2.toFixed(2), meta_6m: d.meta_6m_m2 }));
  } else {
    dados = _filtraPeriodo((D.consumo_agua_m2 || []).slice(), periodoF);
  }
  const fmt = v => Number(v).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  _drawEnergiaChart('chart-consumo-agua-m2', 'agua-m2-badge', dados, 'm3_m2', fmt, 'm³/m²', obraF);
}

function renderConsumoEnergiaM2() {
  const obraF    = (document.getElementById('filter-obra-global')||{}).value || '';
  const setorF   = (document.getElementById('filter-energia-setor')||{}).value || '';
  const periodoF = periodoEfetivo('filter-energia-periodo', '12m');
  let dados;
  if (obraF || setorF) {
    // Para intensidade filtrada, usa o campo recalculado kwh_m2 + meta_6m_m2
    dados = _filtraPeriodo(_computeEnergiaSeries(obraF, setorF), periodoF)
      .filter(d => d.kwh_m2 != null)
      .map(d => ({ ...d, kwh_m2: +d.kwh_m2.toFixed(2), meta_6m: d.meta_6m_m2 }));
  } else {
    dados = _filtraPeriodo((D.consumo_energia_m2 || []).slice(), periodoF);
  }
  const fmt = v => Number(v).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  _drawEnergiaChart('chart-consumo-energia-m2', 'energia-m2-badge', dados, 'kwh_m2', fmt, 'kWh/m²', obraF);
}

// ═══════════════════════════════════════════════════════════
//  Impressões — histórico + setores (com filtros)
// ═══════════════════════════════════════════════════════════
function labelMesAno(ma) {
  if (!ma || typeof ma !== 'string') return ma || '';
  const p = ma.split('-');
  const mn = ['','Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  return (mn[parseInt(p[1])] || p[1]) + '/' + (p[0]||'').slice(2);
}

function initImpFilters() {
  const sel = document.getElementById('filter-imp-setor');
  if (!sel) return;
  // Limpa opções extras (mantém o "Todos os setores")
  while (sel.options.length > 1) sel.remove(1);
  (D.imp_setores_list || []).forEach(s => {
    const opt = document.createElement('option');
    opt.value = s; opt.textContent = s;
    sel.appendChild(opt);
  });
}

function clearImpFilters() {
  const s = document.getElementById('filter-imp-setor');
  const p = document.getElementById('filter-imp-periodo');
  if (s) s.value = '';
  if (p) p.value = 'all';
  renderImpressoes();
}

function renderImpressoes() {
  const setorF   = (document.getElementById('filter-imp-setor')||{}).value || '';
  // Filtro do card tem precedência; em 'all' segue o filtro global de Período
  const periodoF = periodoEfetivo('filter-imp-periodo', 'all');

  const histGlobal = D.imp_historico || [];
  const setorMes   = D.imp_setor_mes || [];
  const setoresTop = D.imp_setores || [];

  // Conjunto de meses disponíveis (de ambas as fontes)
  const allMonths = [...new Set([
    ...histGlobal.map(h => h.mes_ano),
    ...setorMes.map(s => s.mes_ano),
  ].filter(Boolean))].sort();

  let mesesFiltrados;
  if (periodoF && periodoF !== 'all') mesesFiltrados = allMonths.filter(m => noPeriodo(m, periodoF));
  else mesesFiltrados = allMonths;
  const mesesSet = new Set(mesesFiltrados);

  // ─── Dados do gráfico ─────────────────────────────────────
  let labels = [], valores = [];
  if (setorF) {
    // Filtra por setor + período → agrega por mês
    const monthly = {};
    setorMes.filter(s => s.setor === setorF && mesesSet.has(s.mes_ano)).forEach(s => {
      monthly[s.mes_ano] = (monthly[s.mes_ano] || 0) + s.paginas;
    });
    const keys = Object.keys(monthly).sort();
    labels = keys.map(labelMesAno);
    valores = keys.map(k => monthly[k]);
  } else {
    // "Todos os setores" → usa histórico global filtrado por período
    const filt = histGlobal.filter(h => mesesSet.has(h.mes_ano));
    labels = filt.map(h => h.mes);
    valores = filt.map(h => h.total);
  }

  // Recalcula média 6m com base nos dados filtrados (últimos 6 valores do recorte)
  let media6m = 0;
  if (valores.length >= 6) {
    media6m = Math.round(valores.slice(-6).reduce((a,b)=>a+b,0) / 6);
  } else if (valores.length > 0) {
    media6m = Math.round(valores.reduce((a,b)=>a+b,0) / valores.length);
  }

  // Badge meta
  const badge = document.getElementById('imp-media-badge');
  if (badge) {
    if (media6m > 0) {
      const label = valores.length >= 6 ? 'média 6m' : `média ${valores.length}m`;
      badge.textContent = `meta: ${media6m.toLocaleString('pt-BR')} pág (${label})${setorF ? ' · ' + setorF : ''}`;
    } else {
      badge.textContent = 'sem dados';
    }
  }

  // ─── Gráfico de linha ─────────────────────────────────────
  if (valores.length) {
    const mediaArr = valores.map(() => media6m);
    const ptColors = valores.map(v => v > media6m ? '#CE8686' : '#185FA5');

    makeChart('chart-imp-historico', {
      type: 'line',
      data: {
        labels,
        datasets: [
          {
            label: 'Total de páginas',
            data: valores,
            borderColor: '#185FA5',
            backgroundColor: 'rgba(24,95,165,0.08)',
            fill: true,
            tension: 0.3,
            pointRadius: 5,
            pointBackgroundColor: ptColors,
            pointBorderColor: ptColors,
            borderWidth: 2,
          },
          {
            label: 'Média 6 meses (meta)',
            data: mediaArr,
            borderColor: '#EF9F27',
            borderDash: [6, 4],
            borderWidth: 2,
            pointRadius: 0,
            fill: false,
            tension: 0,
          }
        ]
      },
      options: {
        responsive: true,
        plugins: {
          legend: { position: 'top', labels: { boxWidth: 12, padding: 14 } },
          tooltip: {
            callbacks: {
              label: ctx => {
                if (ctx.datasetIndex === 0) {
                  const v = ctx.raw;
                  const diff = v - media6m;
                  const sinal = diff > 0 ? '+' : '';
                  return [
                    ` Páginas: ${v.toLocaleString('pt-BR')}`,
                    ` vs meta: ${sinal}${diff.toLocaleString('pt-BR')}`,
                  ];
                }
                return ` Média 6m: ${ctx.raw.toLocaleString('pt-BR')}`;
              }
            }
          }
        },
        layout: { padding: { top: 26, left: 8, right: 8 } },
        scales: {
          // Sem eixo Y: o valor de cada mês já vai no rótulo, em cima do ponto
          y: { display: false, beginAtZero: false, grace: '12%' },
          x: { grid: { display: false }, ticks: { font: { size: 11 } } }
        }
      },
      plugins: [barLabels(v => v.toLocaleString('pt-BR'),
        { id: 'imphist', dataset: 0, color: '#0C447C', font: '600 11px "Segoe UI", Inter, system-ui, sans-serif' })]
    });
  } else {
    // Sem dados: limpa o canvas
    if (chartReg['chart-imp-historico']) { chartReg['chart-imp-historico'].destroy(); delete chartReg['chart-imp-historico']; }
  }

  // ─── Tabela de setores (respeita filtro de período; destaca setor selecionado) ───
  const tblEl = document.getElementById('imp-setores-table');
  if (!tblEl) return;

  // Agrega imp_setor_mes pelo período filtrado
  const setoresAgg = {};
  setorMes.filter(s => mesesSet.has(s.mes_ano)).forEach(s => {
    setoresAgg[s.setor] = (setoresAgg[s.setor] || 0) + s.paginas;
  });
  const setoresArr = Object.entries(setoresAgg)
    .sort(([,a],[,b]) => b - a)
    .slice(0, 10)
    .map(([setor, paginas]) => ({setor, paginas}));
  const totalAgg = setoresArr.reduce((sum, s) => sum + s.paginas, 0);
  setoresArr.forEach(s => { s.pct = totalAgg > 0 ? +(s.paginas / totalAgg * 100).toFixed(1) : 0; });

  if (!setoresArr.length) {
    tblEl.innerHTML = '<div class="empty-msg">Sem dados de setores para o filtro</div>';
    const badgeS0 = document.getElementById('imp-setores-badge');
    if (badgeS0) badgeS0.textContent = 'sem dados';
    return;
  }

  const maxPag = setoresArr[0].paginas;
  let html = `<table class="mini-table" style="width:100%">
    <thead><tr>
      <th>Setor</th>
      <th style="text-align:right">Páginas</th>
      <th style="text-align:right">%</th>
      <th style="min-width:80px">Participação</th>
    </tr></thead><tbody>`;
  setoresArr.forEach((s, i) => {
    const barW = Math.round(s.paginas / maxPag * 100);
    let rowColor = 'transparent';
    let setorLabel = s.setor;
    let weight = '400';
    if (setorF && s.setor === setorF) {
      // Setor selecionado tem destaque verde
      rowColor = 'var(--fg7)';
      setorLabel = '<strong style="color:var(--fg2)">▸ ' + s.setor + '</strong>';
      weight = '600';
    } else if (!setorF) {
      if (i === 0) rowColor = 'var(--vr7)';
      else if (i === 1) rowColor = 'var(--am6)';
    }
    html += `<tr style="background:${rowColor}">
      <td style="font-weight:${weight}">${setorLabel}</td>
      <td style="text-align:right; font-weight:500">${s.paginas.toLocaleString('pt-BR')}</td>
      <td style="text-align:right; color:var(--gz3)">${s.pct}%</td>
      <td>
        <div class="prog-wrap">
          <div class="prog-bar ${setorF && s.setor === setorF ? 'prog-green' : 'prog-blue'}" style="width:${barW}%"></div>
        </div>
      </td>
    </tr>`;
  });
  html += '</tbody></table>';
  tblEl.innerHTML = html;

  // Badge da tabela com total do período
  const badgeS = document.getElementById('imp-setores-badge');
  if (badgeS) {
    badgeS.textContent = `top ${setoresArr.length} · ${totalAgg.toLocaleString('pt-BR')} pág${periodoF !== 'all' ? ' · ' + (periodoF === '6m' ? 'últ. 6m' : periodoF === '12m' ? 'últ. 12m' : periodoF) : ''}`;
  }
}

// ═══════════════════════════════════════════════════════════
//  PAGE 2 — Obra Individual
// ═══════════════════════════════════════════════════════════
function initObraSelect() {
  // Filtro próprio da aba removido — a obra agora vem 100% do filtro global.
}

function renderObra(obraF, anoF) {
  // A obra é controlada exclusivamente pelo filtro global no topo
  const obra = obraF || (document.getElementById('filter-obra-global')||{}).value || '';
  const isAll = !obra;  // sem obra selecionada → modo "Todas as obras" (agregado)
  if (!anoF) anoF = (document.getElementById('filter-periodo')||{}).value || 'all';

  // Conjunto de meses do período selecionado (mesmo padrão do heatmap/IPA/Ruído)
  const allMeses = D.conf_meses || [];
  // Janela de calendário única (ver mesesDoPeriodo/janelaPeriodo)
  let mesesSet = (anoF && anoF !== 'all') ? mesesDoPeriodo(anoF, allMeses) : null;

  const g = D.global || {};
  const obrasEscopo = isAll ? (D.obras || []) : [obra];

  // IPA: média da(s) obra(s) no período selecionado; sem filtro (ou sem série
  // mensal disponível) cai pro IPA agregado (histórico completo) de cada obra.
  const ipaVals = obrasEscopo.map(o => {
    const serie = mesesSet ? ((D.ipa_mensal && D.ipa_mensal[o]) || []).filter(s => mesesSet.has(s.mes)) : [];
    if (serie.length) return serie.reduce((a,s) => a+s.ipa, 0) / serie.length;
    return (D.ipa[o] && D.ipa[o].ipa) || 0;
  });
  const ipaVal = ipaVals.length ? +(ipaVals.reduce((a,b)=>a+b,0)/ipaVals.length).toFixed(1) : 0;

  // Conformidade: média das células do heatmap (obra × mês) dentro do período;
  // sem filtro cai pro % agregado já calculado (conf_atual / conf_media global).
  const confCells = [];
  obrasEscopo.forEach(o => Object.entries(D.conf_heatmap[o] || {}).forEach(([m, v]) => {
    if (mesesSet && mesesSet.has(m) && v && v.perc != null) confCells.push(v);
  }));
  let confPerc, confMeta;
  if (mesesSet && confCells.length) {
    confPerc = +(confCells.reduce((a,v)=>a+v.perc,0)/confCells.length).toFixed(1);
    // Meta 0 = não lançada: fica fora da média (mesma regra do Consolidado)
    const metasLanc = confCells.filter(v => v.meta > 0);
    confMeta = metasLanc.length ? +(metasLanc.reduce((a,v)=>a+v.meta,0)/metasLanc.length).toFixed(1) : 90;
  } else if (isAll) {
    confPerc = g.conf_media || 0; confMeta = 90;
  } else {
    const confAgg = D.conf_atual[obra] || {};
    confPerc = confAgg.perc || 0; confMeta = confAgg.meta || 90;
  }

  // NCs e Reclamações: soma no período selecionado; sem filtro (ou sem
  // série mensal disponível) cai pro total histórico já calculado.
  const _somaMensal = (mensalDict, fallbackTotal) => {
    if (mesesSet) {
      let soma = 0, achou = false;
      obrasEscopo.forEach(o => {
        const serie = ((mensalDict && mensalDict[o]) || []).filter(s => mesesSet.has(s.mes));
        if (serie.length) { achou = true; soma += serie.reduce((a,s) => a+s.count, 0); }
      });
      if (achou) return soma;
    }
    return fallbackTotal;
  };
  const rec = _somaMensal(D.rec_mensal, isAll ? (g.total_rec || 0) : (D.rec_por_obra[obra] || 0));

  const al = isAll ? {} : (D.al_por_obra[obra] || {});
  const res = isAll ? {} : (D.res_por_obra[obra] || {});

  // 6 KPIs — mesmos cards e mesma ordem da aba Consolidado
  const kpObra = kpisDoPeriodo(isAll ? '' : obra, anoF);
  const kgM2Obra = (kpObra && kpObra.kg_m2 != null) ? kpObra.kg_m2
                 : (isAll ? g.media_kg_m2 : res.kg_m2);
  document.getElementById('kpis-obra').innerHTML = kpiCardsHtml({
    ipaLabel: isAll ? 'IPA global' : 'IPA da obra',
    ipa: ipaVal,
    conf: confPerc, confMeta: confMeta,
    pa: paPctDe(isAll ? '' : obra),
    emerg: emergFiltradas(isAll ? '' : obra, anoF).length, emergSub: emergSubDe(anoF),
    kgM2: kgM2Obra,
    resMeta: isAll ? undefined : res.meta,
    resSub: isAll ? 'kg/m² médio · Senna 50 · demais 36,6' : `kg/m² · meta: ${res.meta || 36.6} kg/m²`,
    rec: rec, recLimite: isAll ? 5 : 3,
  });

  // Tabela conformidade (respeita filtro de período)
  // Modo "todas as obras": média das obras por mês
  let heatmap;
  if (isAll) {
    heatmap = {};
    const obrasList = D.obras || [];
    (D.conf_meses || []).forEach(m => {
      const vals = [];
      const metas = [];
      obrasList.forEach(o => {
        const cell = (D.conf_heatmap[o] || {})[m];
        if (cell && cell.perc != null) vals.push(cell.perc);
        if (cell && cell.meta > 0) metas.push(cell.meta);
      });
      if (vals.length) {
        heatmap[m] = {
          perc: +(vals.reduce((a,b)=>a+b,0) / vals.length).toFixed(1),
          meta: metas.length ? +(metas.reduce((a,b)=>a+b,0) / metas.length).toFixed(1) : 90,
        };
      }
    });
  } else {
    heatmap = D.conf_heatmap[obra] || {};
  }
  let meses = D.conf_meses || [];
  if (anoF && anoF !== 'all') meses = meses.filter(m => noPeriodo(m, anoF));
  let confHtml = '';
  if (meses.length) {
    const labelMes = m => { const p = m.split('-'); const mn = ['','Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez']; return (mn[parseInt(p[1])]||p[1])+"'"+p[0].slice(2); };
    confHtml = '<table class="conf-table"><thead><tr><th>Mês</th><th>Realizado</th><th>Meta</th><th>Status</th></tr></thead><tbody>';
    meses.slice(-8).forEach(m => {
      const cell = heatmap[m] || {};
      const p = cell.perc; const me = cell.meta || 90;
      const cls = p != null ? (p >= me ? 'sem-green' : p >= me*0.9 ? 'sem-amber' : 'sem-red') : 'sem-gray';
      confHtml += `<tr><td>${labelMes(m)}</td><td>${p != null ? p.toFixed(1)+'%' : '—'}</td><td>${me != null ? me.toFixed(1)+'%' : '—'}</td><td><span class="semaforo ${cls}"></span></td></tr>`;
    });
    confHtml += '</tbody></table>';
  } else {
    confHtml = '<div class="empty-msg">Sem dados históricos</div>';
  }
  document.getElementById('obra-conf-table').innerHTML = confHtml;

  // Ruído — medições: histórico mensal (até 12 meses) da obra selecionada;
  // "todas as obras": uma coluna por obra, recalculada pelo período selecionado
  renderRuidoObra(isAll ? '' : obra, anoF);

  // Resíduos donut por classe — usa o dataset granular (obra × mês × classe)
  // quando disponível, reagindo a obra + período; sem esse dataset (ex.:
  // gerar_dashboard.py rodado sozinho, sem pbix) ou sem registro no filtro,
  // cai pro breakdown agregado (histórico completo) já calculado.
  const rawRes = D.residuos_raw || [];
  let resClasse = {};
  if (rawRes.length) {
    rawRes.forEach(r => {
      if (!isAll && r.obra !== obra) return;
      if (mesesSet && !mesesSet.has(r.mes_ano)) return;
      if (!r.classe) return;
      resClasse[r.classe] = (resClasse[r.classe] || 0) + (r.kg || 0);
    });
  }
  if (!Object.keys(resClasse).length) {
    resClasse = isAll
      ? (D.res_classe || {})
      : ((D.res_classe_por_obra && D.res_classe_por_obra[obra] && Object.keys(D.res_classe_por_obra[obra]).length)
          ? D.res_classe_por_obra[obra] : (D.res_classe || {}));
  }
  // Volume por tipo de resíduo — mesma função da aba Resíduos, com o recorte da obra.
  // Monta o filtRes com o mesmo critério usado lá (obra × período).
  let filtResObra = rawRes.filter(r => !mesesSet || mesesSet.has(r.mes_ano));
  if (!isAll) filtResObra = filtResObra.filter(r => r.obra === obra);
  const tipoSrcObra = isAll
    ? (D.res_tipo || {})
    : ((D.res_tipo_por_obra && D.res_tipo_por_obra[obra]) || {});
  renderResTipo('chart-obra-res-tipo', 'obra-res-tipo-badge', filtResObra, tipoSrcObra, rawRes.length > 0);

  const rcEntries = Object.entries(resClasse).filter(([k]) => k && k !== 'nan');
  if (rcEntries.length) {
    makeChart('chart-obra-res-donut', {
      type: 'doughnut',
      data: {
        labels: rcEntries.map(([k]) => 'Classe ' + k),
        datasets: [{ data: rcEntries.map(([,v]) => v), backgroundColor: ['#C5C639','#006B6B','#005E70','#336B87','#CE8686','#00484C','#85AFC5'], borderWidth: 2, borderColor: '#fff' }]
      },
      options: { responsive: true, plugins: { legend: { position: 'right', labels: { boxWidth: 8, padding: 6, font: { size: 10 } } } } },
      plugins: [donutValueLabels(v => numBR(v / 1000, 1) + ' t')]
    });
  }

  // Plano ação da obra (modo "todas": soma todas as obras)
  let paObra;
  if (isAll) {
    paObra = {};
    Object.values(D.pa_por_obra || {}).forEach(pa => {
      Object.entries(pa || {}).forEach(([k, v]) => {
        paObra[k] = (paObra[k] || 0) + (v || 0);
      });
    });
  } else {
    paObra = D.pa_por_obra[obra] || {};
  }
  const paEntries = Object.entries(paObra).filter(([,v]) => v > 0);
  // ATENÇÃO — mesmo caso do card de Plano de Ação no Consolidado: Dim_Plano_acao
  // não tem data de registro da ação (só Prazo, que é data futura de conclusão),
  // então este donut não filtra por período. Em vez de fingir que filtrou, o
  // badge avisa quando há período selecionado.
  const paBadgeEl = document.getElementById('obra-pa-badge');
  if (paBadgeEl) {
    const totalPaObra = paEntries.reduce((a, [, v]) => a + v, 0);
    const avisoPeriodoPa = (anoF && anoF !== 'all') ? ' · histórico completo (sem data na fonte)' : '';
    paBadgeEl.textContent = totalPaObra + (totalPaObra !== 1 ? ' ações' : ' ação') + avisoPeriodoPa;
  }
  if (paEntries.length) {
    const colorsMap2 = { 'Realizado': '#005E70', 'Em Andamento': '#89B1C5', 'Pendente': '#888780', 'Atrasado': '#CE8686' };
    makeChart('chart-obra-pa', {
      type: 'doughnut',
      data: {
        labels: paEntries.map(([k]) => k),
        datasets: [{ data: paEntries.map(([,v]) => v), backgroundColor: paEntries.map(([k]) => colorsMap2[k] || '#B4B2A9'), borderWidth: 2, borderColor: '#fff' }]
      },
      options: { responsive: true, plugins: { legend: { position: 'right', labels: { boxWidth: 8, padding: 6, font: { size: 10 } } } } },
      plugins: [donutValueLabels((v, total) => `${v} (${total > 0 ? Math.round(v / total * 100) : 0}%)`)]
    });
  }

  // IPA por obra — obra selecionada: histórico de 12 meses; "todas as obras":
  // comparativo entre obras, recalculado conforme o período selecionado
  renderIpaObras(isAll ? '' : obra, anoF);

  // Emergências Ambientais da obra (analítica)
  renderObraEmergencias(isAll ? '' : obra, anoF);

  // Auditorias Ambientais — movida da aba Resíduos; mesma função, sem mudança
  // de regra. O filtro de Período próprio do card (#filter-aud-mes) continua
  // com precedência sobre o global, como já era.
  renderAuditorias(isAll ? '' : obra, anoF);
}

function renderObraEmergencias(obra, anoF) {
  if (anoF === undefined) anoF = periodoGlobal();
  const all = D.emergencias_lista || [];
  let data = obra ? all.filter(r => r.obra === obra) : all.slice();
  // Ver comentário em renderEmergencias: usa `mes` (YYYY-MM), não mes_ano (com dia)
  if (anoF && anoF !== 'all') {
    data = data.filter(r => noPeriodo(r.mes || r.mes_ano, anoF));
  }
  const tblEl = document.getElementById('obra-emerg-table');
  const badge = document.getElementById('obra-emerg-badge');
  if (!tblEl) return;

  const escHtml = s => String(s == null ? '' : s).replace(/[&<>"']/g, c =>
    ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]));

  const scopeLabel = (obra || 'todas as obras') + (anoF && anoF !== 'all' ? ' · ' + labelPeriodo(anoF) : '');

  if (!data.length) {
    tblEl.innerHTML = `
      <thead><tr>
        <th class="col-mes">Mês/Ano</th>
        <th>Obra</th>
        <th>Fonte</th>
        <th>FMC</th>
        <th class="col-ocorr">Ocorrência</th>
        <th class="col-status">Status/Ações</th>
        <th>Categoria</th>
      </tr></thead>
      <tbody><tr class="empty-row"><td colspan="7">Nenhuma emergência registrada para ${escHtml(scopeLabel)}</td></tr></tbody>`;
    if (badge) badge.textContent = `0 registros · ${scopeLabel}`;
    return;
  }

  let html = `
    <thead><tr>
      <th class="col-mes">Mês/Ano</th>
      <th>Obra</th>
      <th>Fonte</th>
      <th>FMC</th>
      <th class="col-ocorr">Ocorrência</th>
      <th class="col-status">Status/Ações</th>
      <th>Categoria</th>
    </tr></thead>
    <tbody>`;
  data.forEach(r => {
    html += `<tr>
      <td class="col-mes">${escHtml(r.mes_label) || '—'}</td>
      <td>${r.obra ? `<span class="obra-tag">${escHtml(r.obra)}</span>` : '—'}</td>
      <td>${escHtml(r.fonte) || '—'}</td>
      <td class="col-fmc">${escHtml(r.fmc) || '—'}</td>
      <td class="col-ocorr">${escHtml(r.ocorrencia) || '—'}</td>
      <td class="col-status">${escHtml(r.status) || '—'}</td>
      <td>${escHtml(r.categoria) || '—'}</td>
    </tr>`;
  });
  html += '</tbody>';
  tblEl.innerHTML = html;
  if (badge) badge.textContent = `${data.length} registro${data.length !== 1 ? 's' : ''} · ${scopeLabel}`;
}

// ═══════════════════════════════════════════════════════════
//  PAGE 3 — Resíduos
// ═══════════════════════════════════════════════════════════
// ═══════════════════════════════════════════════════════════
//  Gráfico "Volume por tipo de resíduo" — compartilhado
// ═══════════════════════════════════════════════════════════
// Extraído de renderResiduos() para ser reaproveitado na aba Obra Individual.
// Recebe o recorte já filtrado (obra × período) em vez de refazer o filtro,
// para as duas abas mostrarem exatamente o mesmo critério.
//   canvasId / badgeId — onde desenhar
//   filtRes            — linhas de residuos_raw já recortadas
//   tipoSrc            — agregado por tipo (fallback do caminho legado via Excel)
//   hasRaw             — se existe dataset granular
function renderResTipo(canvasId, badgeId, filtRes, tipoSrc, hasRaw) {
  // Placeholder "Sem dados" — versão local do helper homônimo de
  // renderResiduos(), que não é visível daqui depois da extração.
  const semDados = id => makeChart(id, {
    type: 'bar',
    data: { labels: [''], datasets: [{ data: [0], backgroundColor: '#D3D1C7' }] },
    options: { responsive: true, plugins: { legend: { display: false }, subtitle: { display: true, text: 'Sem dados', font: { size: 12.1 } } }, scales: { y: { beginAtZero: true } } }
  });
// Tipo barras H (usa fonte filtrada) — cores conforme paleta por tipo
// Mapeamento: normalize p/ uppercase + sem acento + sem espaço extra
const CORES_TIPO = {
  // Azul petróleo
  'ENTULHO':                                              '#0E4A6B',
  // Escala da cor azul (mais escuro → mais claro)
  'GESSO':                                                '#185FA5',
  'PAPEL, PAPELAO':                                       '#378ADD',
  'PLASTICO':                                             '#85B7EB',
  'ISOPOR':                                               '#B5D4F4',
  // Escala de cor laranja
  'LA DE PET':                                            '#EF9F27',
  'TELA DE PROTECAO':                                     '#FAC775',
  // Marrom
  'MADEIRA':                                              '#6B4423',
  // Escalas de cinza
  'METAL, SUCATA':                                        '#D3D1C7',  // cinza claro
  'EFLUENTE CONTAMINADO (LIQUIDO)':                       '#888780',  // cinza médio
  'RESIDUO CONTAMINADO (SOLIDO)':                         '#5F5E5A',  // cinza escuro
  // Preto
  'NAO RECICLADO (FITAS ADESIVAS, LIXAS ENTRE OUTROS)':   '#2C2C2A',
};
const norm = s => String(s || '').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
const corPorTipo = tipo => CORES_TIPO[norm(tipo)] || '#888780';  // fallback cinza médio

// ─── Tipo × mês: 6 colunas (uma por mês) para cada um dos 8 tipos de maior volume ───
// Cada mês é um dataset → o Chart.js agrupa as 6 barras sob o rótulo do tipo.
// Escala sequencial azul, do mês mais antigo (claro) ao mais recente (escuro).
const CORES_MES_TIPO = ['#D6E7F8', '#B5D4F4', '#85B7EB', '#378ADD', '#185FA5', '#0E4A6B'];
const labelMesTipo = labelMesRes;
const badgeTipo = document.getElementById(badgeId);

// Últimos 6 meses COM dado dentro do recorte (obra × período)
const mesesComKg = [...new Set(filtRes.filter(r => (r.kg || 0) > 0).map(r => r.mes_ano))].sort();
const meses6 = mesesComKg.slice(-6);

if (hasRaw && meses6.length) {
  const noRecorte = filtRes.filter(r => meses6.includes(r.mes_ano) && (r.kg || 0) > 0);

  // Top 8 tipos por kg total nesses meses
  const totalPorTipo = {};
  noRecorte.forEach(r => {
    const t = String(r.tipo || '').trim();
    if (t && t.toLowerCase() !== 'nan') totalPorTipo[t] = (totalPorTipo[t] || 0) + r.kg;
  });
  const top8 = Object.entries(totalPorTipo).sort((a, b) => b[1] - a[1]).slice(0, 8).map(([t]) => t);

  // Matriz tipo × mês
  const kgTipoMes = {};
  noRecorte.forEach(r => {
    const t = String(r.tipo || '').trim();
    if (!top8.includes(t)) return;
    if (!kgTipoMes[t]) kgTipoMes[t] = {};
    kgTipoMes[t][r.mes_ano] = (kgTipoMes[t][r.mes_ano] || 0) + r.kg;
  });

  // Cores alinhadas à direita da escala: com menos de 6 meses, o mais recente segue o mais escuro
  const cores = CORES_MES_TIPO.slice(CORES_MES_TIPO.length - meses6.length);

  if (badgeTipo) badgeTipo.textContent = meses6.length > 1
    ? `${labelMesTipo(meses6[0])} – ${labelMesTipo(meses6[meses6.length-1])}`
    : labelMesTipo(meses6[0]);

  makeChart(canvasId, {
    type: 'bar',
    data: {
      labels: top8,
      datasets: meses6.map((m, i) => ({
        label: labelMesTipo(m),
        data: top8.map(t => Math.round((kgTipoMes[t] || {})[m] || 0)),
        backgroundColor: cores[i],
        borderRadius: 3,
      }))
    },
    options: {
      responsive: true,
      plugins: {
        legend: { position: 'top', labels: { boxWidth: 12, padding: 10 } },
        tooltip: {
          callbacks: {
            title: items => items.length ? items[0].label : '',
            label: ctx => ` ${ctx.dataset.label}: ${ctx.raw.toLocaleString('pt-BR')} kg`,
          }
        }
      },
      layout: { padding: { top: 18 } },
      scales: {
        x: {
          ticks: {
            autoSkip: false,
            maxRotation: 30,
            // Nomes de tipo são longos ("Não reciclado (fitas adesivas...)") — encurta no eixo;
            // o nome completo aparece no título do tooltip
            callback: function (v) {
              const t = String(this.getLabelForValue(v));
              return t.length > 20 ? t.slice(0, 19) + '…' : t;
            }
          }
        },
        y: { beginAtZero: true, title: { display: true, text: 'kg', font: { size: 12.1 } } }
      }
    },
    // Muitas barras estreitas (até 8 tipos × 6 meses): kg abreviado pra caber
    plugins: [barLabels(v => v >= 10000 ? numBR(v / 1000, 0) + 'k'
                           : v >= 1000  ? numBR(v / 1000, 1) + 'k'
                           : numBR(v, 0),
                        { id: 'restipo', font: '600 9px "Segoe UI", Inter, system-ui, sans-serif' })]
  });
} else if (hasRaw) {
  // Tem dado granular, mas o recorte (obra × período) está vazio — não cair no
  // agregado histórico, que ignoraria o filtro de período
  if (badgeTipo) badgeTipo.textContent = '—';
  semDados(canvasId);
} else {
  // Sem dado granular (caminho legado via Excel): mantém as barras H de total por tipo
  if (badgeTipo) badgeTipo.textContent = 'total do período';
  const tipoE = Object.entries(tipoSrc).filter(([k]) => k && k !== 'nan');
  if (tipoE.length) {
    makeChart(canvasId, {
      type: 'bar',
      data: {
        labels: tipoE.map(([k]) => k),
        datasets: [{
          label: 'kg',
          data: tipoE.map(([,v]) => v),
          backgroundColor: tipoE.map(([k]) => corPorTipo(k)),
          borderRadius: 4,
        }]
      },
      options: {
        responsive: true,
        indexAxis: 'y',
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: ctx => ` ${ctx.raw.toLocaleString('pt-BR')} kg`,
            }
          }
        },
        scales: { x: { beginAtZero: true, title: { display: true, text: 'kg', font: { size: 12.1 } } } }
      }
    });
  } else {
    semDados(canvasId);
  }
}
}

function renderResiduos(obraF, anoF) {
  const g = D.global;
  if (!obraF) obraF = (document.getElementById('filter-obra-global')||{}).value || '';
  if (!anoF)  anoF  = (document.getElementById('filter-periodo')||{}).value || 'all';

  // Fontes filtradas por obra (se houver) — usa agregados pré-computados (charts/tables)
  const resByObra = obraF ? (D.res_por_obra[obraF] || {}) : null;
  const classeSrc = obraF
    ? (D.res_classe_por_obra && D.res_classe_por_obra[obraF] ? D.res_classe_por_obra[obraF] : {})
    : (D.res_classe || {});
  const tipoSrc = obraF
    ? (D.res_tipo_por_obra && D.res_tipo_por_obra[obraF] ? D.res_tipo_por_obra[obraF] : {})
    : (D.res_tipo || {});

  // Conjunto de meses do filtro
  const allMonthsRes = (D.residuos_mensal || []).map(r => r.mes_ano).sort();
  let mesesSet;
  mesesSet = mesesDoPeriodo(anoF, allMonthsRes);

  // Fonte granular compartilhada (KPIs + gráficos por obra), recortada por obra × período
  const rawRes = D.residuos_raw || [];
  const hasRaw = rawRes.length > 0;
  let filtRes  = rawRes.filter(r => mesesSet.has(r.mes_ano));
  if (obraF) filtRes = filtRes.filter(r => r.obra === obraF);

  // ─── KPIs híbridos: estáticos quando "Todo + Todas"; senão recalcula do residuos_raw ───
  const isAllAll = (anoF === 'all' || !anoF) && !obraF;
  let totalKgF, mediaKmF;

  if (isAllAll || !hasRaw) {
    totalKgF = D.total_res_kg || 0;
    mediaKmF = g.media_kg_m2 || 0;
  } else {
    totalKgF = filtRes.reduce((s, r) => s + (r.kg || 0), 0);
    const kgM2Vals = filtRes.filter(r => r.kg_m2 != null && r.kg_m2 > 0).map(r => r.kg_m2);
    mediaKmF = kgM2Vals.length ? kgM2Vals.reduce((s, v) => s + v, 0) / kgM2Vals.length : 0;
  }

  // Classe CONAMA com maior geração no recorte (obra × período)
  let classeAgg;
  if (isAllAll || !hasRaw) {
    classeAgg = Object.entries(classeSrc).filter(([k]) => k && k !== 'nan');
  } else {
    const accCl = {};
    filtRes.forEach(r => {
      const c = String(r.classe || '').trim();
      if (c && c.toLowerCase() !== 'nan') accCl[c] = (accCl[c] || 0) + (r.kg || 0);
    });
    classeAgg = Object.entries(accCl);
  }
  classeAgg = classeAgg.filter(([, v]) => v > 0).sort((a, b) => b[1] - a[1]);
  const topClasse    = classeAgg.length ? classeAgg[0] : null;
  const topClasseKg  = topClasse ? topClasse[1] : 0;
  const topClassePct = (topClasse && totalKgF > 0) ? (topClasseKg / totalKgF * 100) : null;

  // Senna Tower consolidado: média das intensidades mensais (soma dos tipos no
  // mês) dentro do recorte — mesmo número do selo "acumulado" do card da Senna.
  // Com outra obra no filtro global, a Senna sai de cena (mesma regra do gráfico).
  let sennaKm = null;
  const sennaFora = !!(obraF && !SENNA_RE.test(obraF));
  if (!sennaFora) {
    const porMesSenna = {};
    rawRes.forEach(r => {
      if (!SENNA_RE.test(String(r.obra || '')) || !(r.kg_m2 > 0) || !mesesSet.has(r.mes_ano)) return;
      porMesSenna[r.mes_ano] = (porMesSenna[r.mes_ano] || 0) + r.kg_m2;
    });
    const vs = Object.values(porMesSenna);
    if (vs.length) sennaKm = vs.reduce((a, b) => a + b, 0) / vs.length;
  }
  const metaSenna = resMetaDe((D.obras || []).find(o => SENNA_RE.test(o)) || '') || 50;
  const metaSennaTxt = numBR(metaSenna, Number.isInteger(metaSenna) ? 0 : 1);

  // 4 KPI cards — refletem o recorte (obra × período)
  const periodoLbl = (anoF === '6m') ? 'últ. 6m'
                   : (anoF === '12m') ? 'últ. 12m'
                   : (anoF && anoF !== 'all') ? anoF
                   : 'todo o período';
  document.getElementById('kpis-residuos').innerHTML = `
    <div class="kpi green">
      <div class="kpi-label"><i class="ti ti-package"></i> Volume total</div>
      <div class="kpi-value">${numBR(totalKgF/1000,1)} t</div>
      <div class="kpi-sub">${numBR(totalKgF,0)} kg · ${obraF ? obraF : periodoLbl}</div>
    </div>
    <div class="kpi ${mediaKmF <= 50 ? 'green' : 'red'}">
      <div class="kpi-label"><i class="ti ti-ruler-measure"></i> Geração Resíduo</div>
      <div class="kpi-value">${numBR(mediaKmF,2)}</div>
      <div class="kpi-sub">kg/m² · ${obraF && obraF.toLowerCase().includes('senna') ? 'meta: 50' : 'meta: 36,6'} kg/m²</div>
    </div>
    <div class="kpi ${sennaKm == null ? 'gray' : sennaKm <= metaSenna ? 'green' : 'red'}">
      <div class="kpi-label"><i class="ti ti-building-skyscraper"></i> Senna Tower — consolidado</div>
      <div class="kpi-value">${sennaKm == null ? '—' : numBR(sennaKm, 2)}</div>
      <div class="kpi-sub">${sennaFora ? 'fora do filtro · ' + obraF : sennaKm == null ? 'sem lançamento no recorte' : 'kg/m² · meta: ' + metaSennaTxt + ' kg/m²'}</div>
    </div>
    <div class="kpi blue">
      <div class="kpi-label"><i class="ti ti-recycle"></i> Classes CONAMA</div>
      <div class="kpi-value">${topClasse ? 'Classe ' + topClasse[0] : '—'}</div>
      <div class="kpi-sub">${topClasse
        ? `maior geração · ${numBR(topClasseKg/1000,1)} t${topClassePct != null ? ' · ' + numBR(topClassePct,1) + '% do total' : ''}`
        : 'sem dados no recorte'}</div>
    </div>
  `;

  // ─── Série por obra, base dos dois gráficos: kg/m² ("Geração Resíduo") e ton ("Geração Resíduos (ton)") ───
  // "Todo o período + todas as obras" (ou sem residuos_raw) → agregados estáticos;
  // com filtro ativo → recalcula de residuos_raw, reagindo a obra × período
  const metaObra = o => ((D.res_por_obra || {})[o] || {}).meta
                     || (String(o).toLowerCase().includes('senna') ? 50 : 36.6);
  let serieObra;
  if (isAllAll || !hasRaw) {
    serieObra = (D.res_barras || [])
      .filter(d => !obraF || d.obra === obraF)
      .map(d => ({
        obra:  d.obra,
        kg_m2: d.kg_m2,
        meta:  d.meta || metaObra(d.obra),
        ton:   (((D.res_por_obra || {})[d.obra] || {}).total_kg || 0) / 1000,
      }));
  } else {
    const accObra = {};
    filtRes.forEach(r => {
      const o = r.obra;
      if (!o) return;
      if (!accObra[o]) accObra[o] = { kg: 0, km: [] };
      accObra[o].kg += (r.kg || 0);
      if (r.kg_m2 != null && r.kg_m2 > 0) accObra[o].km.push(r.kg_m2);
    });
    serieObra = Object.entries(accObra).map(([o, v]) => ({
      obra:  o,
      kg_m2: v.km.length ? v.km.reduce((a, x) => a + x, 0) / v.km.length : null,
      meta:  metaObra(o),
      ton:   v.kg / 1000,
    }));
  }

  const semDados = id => makeChart(id, { type: 'bar', data: { labels: [obraF||''], datasets: [{ data: [0], backgroundColor: '#D3D1C7' }] }, options: { responsive: true, plugins: { legend: { display: false }, subtitle: { display: true, text: 'Sem dados', font: { size: 12.1 } } }, scales: { y: { beginAtZero: true } } } });

  // ─── Eixo dos dois gráficos de volume ───
  // Sem obra filtrada → uma coluna por obra/empreendimento.
  // Com UMA obra filtrada → o eixo passa a ser temporal. Base 6 meses; o filtro de
  // Período manda quando pede outra coisa:
  //   "todo o período", "últ. 6 meses" e ano específico → 6 · "últ. 12 meses" → 12.
  // Meses sem lançamento no recorte não entram (por isso pode vir menos que o teto).
  const tetoMeses = (anoF === '12m') ? 12 : 6;
  const mesesComDadoObra = [...new Set(
    filtRes.filter(r => (r.kg || 0) > 0 || (r.kg_m2 || 0) > 0).map(r => r.mes_ano)
  )].sort();
  const mesesObra = mesesComDadoObra.slice(-tetoMeses);
  const porMes = !!(hasRaw && obraF && mesesObra.length);

  // ─── Geração Resíduo — kg/m², com a Senna Tower num card à parte ───
  // Mesmo padrão dos cards "Resíduos Gerados kg/m² por obra" + "... — Senna Tower"
  // da aba Consolidado (ver renderResBarras/renderResSenna, mesmos helpers
  // resGraficoMensal/resMetaDe/RES_COR_*): a Senna tem meta 50 (as demais 36,6) e
  // intensidade muito acima, então sempre fica num card separado — nunca escondida
  // atrás de um botão. Diferença desta aba: os dois cards continuam reagindo ao
  // filtro de Período (feature já existente aqui; o Consolidado não tem esse filtro).
  const tituloGer      = document.getElementById('res-ger-titulo');
  const badgeGer       = document.getElementById('res-ger-badge');
  const badgeGerSenna  = document.getElementById('res-ger-senna-badge');
  const acumGerSenna   = document.getElementById('res-ger-senna-acum');

  // kg/m² lançados por obra × mês dentro do recorte de obra × período
  const linhasKm = {};
  filtRes.forEach(r => {
    if (!(r.kg_m2 > 0)) return;
    const o = String(r.obra || '').trim();
    if (!o) return;
    linhasKm[o] = linhasKm[o] || {};
    (linhasKm[o][r.mes_ano] = linhasKm[o][r.mes_ano] || []).push(r.kg_m2);
  });

  const limparGerCard = (canvasId, badgeEl, msg, acumEl) => {
    if (chartReg[canvasId]) { chartReg[canvasId].destroy(); delete chartReg[canvasId]; }
    if (badgeEl) badgeEl.textContent = msg;
    if (acumEl) acumEl.hidden = true;
  };

  // ─── Card principal: obra filtrada (mês a mês) ou consolidado (barra por obra) ───
  if (hasRaw && obraF) {
    const serieO = linhasKm[obraF] || {};
    const meses = mesesObra.filter(m => (serieO[m] || []).length);
    if (tituloGer) tituloGer.textContent = `Geração Resíduo — ${obraF}`;
    if (!meses.length) limparGerCard('chart-res-intensidade', badgeGer, `sem lançamento · ${obraF}`);
    else resGraficoMensal('chart-res-intensidade', obraF, serieO, meses, badgeGer, false);
  } else {
    if (tituloGer) tituloGer.textContent = 'Geração Resíduo';
    const mediaDe = o => {
      const mensais = mesesObra.map(m => resIntensidadeMes((linhasKm[o] || {})[m])).filter(v => v != null);
      return mensais.length ? mensais.reduce((a, b) => a + b, 0) / mensais.length : null;
    };
    const serie = Object.keys(linhasKm)
      .filter(o => !SENNA_RE.test(o))
      .map(o => ({ obra: o, kg_m2: mediaDe(o) }))
      .filter(d => d.kg_m2 != null)
      .sort((a, b) => b.kg_m2 - a.kg_m2);
    if (!serie.length) {
      limparGerCard('chart-res-intensidade', badgeGer, 'sem lançamento no recorte');
    } else {
      makeChart('chart-res-intensidade', {
        type: 'bar',
        data: {
          labels: serie.map(d => d.obra),
          datasets: [
            { label: 'kg/m²', data: serie.map(d => +d.kg_m2.toFixed(2)), backgroundColor: RES_COR_BARRA, borderRadius: 4, order: 2 },
            { label: 'Meta (36,6)', data: serie.map(d => resMetaDe(d.obra)), type: 'line',
              stepped: 'middle', borderColor: RES_COR_META, borderWidth: 2, borderDash: [6, 4],
              pointRadius: 0, fill: false, order: 1 }
          ]
        },
        options: {
          responsive: true,
          layout: { padding: { top: 18 } },
          plugins: {
            legend: { display: true, position: 'top', labels: { boxWidth: 12, padding: 10 } },
            tooltip: {
              callbacks: {
                label: ctx => ctx.dataset.type === 'line'
                  ? ` Meta da obra: ${resMetaDe(ctx.label)} kg/m²`
                  : ` ${resFmtKg(ctx.raw)} kg/m² (média)`,
                afterBody: items => {
                  if (!items.length) return '';
                  const o = items[0].label;
                  const det = mesesObra.filter(m => (linhasKm[o] || {})[m])
                    .map(m => `${labelMesRes(m)}: ${resIntensidadeMes(linhasKm[o][m]).toFixed(1)}`);
                  return det.length ? [''].concat(det) : '';
                }
              }
            }
          },
          scales: { y: { beginAtZero: true, title: { display: true, text: 'kg/m²', font: { size: 12.1 } } } }
        },
        plugins: [barLabels(v => resFmtKg(v), { id: 'resger' })]
      });
      if (badgeGer) {
        const faixa = mesesObra.length > 1 ? `${labelMesRes(mesesObra[0])} – ${labelMesRes(mesesObra[mesesObra.length - 1])} · ` : (mesesObra.length ? `${labelMesRes(mesesObra[0])} · ` : '');
        badgeGer.textContent = `${faixa}meta 36,6 kg/m²`;
      }
    }
  }

  // ─── Card ao lado: Senna Tower, mês a mês, meta 50 — reage ao filtro de obra
  // (mesma régua: com outra obra filtrada, a Senna sai de cena) e de período ───
  if (obraF && !SENNA_RE.test(obraF)) {
    limparGerCard('chart-res-intensidade-senna', badgeGerSenna, `fora do filtro · ${obraF}`, acumGerSenna);
  } else {
    const obraSenna = Object.keys(linhasKm).find(o => SENNA_RE.test(o))
                   || (D.obras || []).find(o => SENNA_RE.test(o)) || '';
    if (!obraSenna) {
      limparGerCard('chart-res-intensidade-senna', badgeGerSenna, 'obra não encontrada', acumGerSenna);
    } else {
      const serieSenna = linhasKm[obraSenna] || {};
      const mesesSenna = mesesObra.filter(m => (serieSenna[m] || []).length);
      if (!mesesSenna.length) {
        limparGerCard('chart-res-intensidade-senna', badgeGerSenna, 'sem lançamento no período', acumGerSenna);
      } else {
        if (acumGerSenna) {
          const todos = Object.keys(serieSenna).map(m => resIntensidadeMes(serieSenna[m])).filter(v => v != null);
          if (todos.length) {
            acumGerSenna.textContent = `acumulado: ${resFmtKg(todos.reduce((a, b) => a + b, 0) / todos.length)} kg/m²`;
            acumGerSenna.hidden = false;
          } else acumGerSenna.hidden = true;
        }
        resGraficoMensal('chart-res-intensidade-senna', obraSenna, serieSenna, mesesSenna, badgeGerSenna, true);
      }
    }
  }

  // ─── Geração Resíduos (ton) — uma barra por classe CONAMA ───
  // Eixo por obra (ou por mês, com obra filtrada) e, dentro de cada um, uma barra
  // separada por classe, lado a lado — obra com 2 classes mostra 2 barras.
  // Paleta CONAMA 307: A reutilizáveis/recicláveis · B recicláveis · C sem tecnologia viável · D perigosos
  const conamaColors = { 'A': '#006B6B', 'B': '#336B87', 'C': '#C5C639', 'D': '#CE8686' };
  const classeKey = r => String(r.classe || '').trim().toUpperCase();
  let catsTon = [], tonMat = {};
  if (porMes) {
    catsTon = mesesObra.slice();
    catsTon.forEach(c => { tonMat[c] = {}; });
    filtRes.forEach(r => {
      const k = classeKey(r);
      if (!tonMat[r.mes_ano] || !k || k === 'NAN' || !(r.kg > 0)) return;
      tonMat[r.mes_ano][k] = (tonMat[r.mes_ano][k] || 0) + r.kg / 1000;
    });
  } else if (isAllAll || !hasRaw) {
    // View padrão (ou sem dado granular): agregados estáticos, com a composição
    // por classe vinda de res_classe_por_obra
    catsTon = (D.res_barras || []).filter(d => !obraF || d.obra === obraF).map(d => d.obra);
    catsTon.forEach(o => {
      tonMat[o] = {};
      Object.entries((D.res_classe_por_obra || {})[o] || {}).forEach(([k, v]) => {
        const c = String(k).trim().toUpperCase();
        if (c && c !== 'NAN' && v > 0) tonMat[o][c] = v / 1000;
      });
    });
  } else {
    // Período filtrado com todas as obras: recalcula por obra do raw
    filtRes.forEach(r => {
      const k = classeKey(r);
      if (!r.obra || !k || k === 'NAN' || !(r.kg > 0)) return;
      if (!tonMat[r.obra]) tonMat[r.obra] = {};
      tonMat[r.obra][k] = (tonMat[r.obra][k] || 0) + r.kg / 1000;
    });
    catsTon = Object.keys(tonMat);
  }
  const totalCatTon = c => Object.values(tonMat[c] || {}).reduce((a, b) => a + b, 0);
  catsTon = catsTon.filter(c => totalCatTon(c) > 0);
  if (!porMes) catsTon.sort((a, b) => totalCatTon(b) - totalCatTon(a));  // por mês mantém cronológico
  // Classes que aparecem no recorte, em ordem A → D (ordem das pilhas e da legenda)
  const classesTon = [...new Set(catsTon.flatMap(c => Object.keys(tonMat[c] || {})))].sort();

  if (catsTon.length && classesTon.length) {
    makeChart('chart-res-conama', {
      type: 'bar',
      data: {
        labels: catsTon.map(c => porMes ? labelMesRes(c) : c),
        datasets: classesTon.map(cl => ({
          label: 'Classe ' + cl,
          data: catsTon.map(c => Number(((tonMat[c] || {})[cl] || 0).toFixed(1))),
          backgroundColor: conamaColors[cl] || '#888780',
          borderRadius: 3,
        }))
      },
      options: {
        responsive: true,
        plugins: {
          legend: { position: 'top', labels: { boxWidth: 12, padding: 10 } },
          tooltip: {
            callbacks: {
              label: ctx => ` ${ctx.dataset.label}: ${numBR(ctx.raw, 1)} t`,
              footer: items => {
                // Soma as classes do grupo (obra ou mês) sob o cursor
                if (!items.length) return '';
                const i = items[0].dataIndex;
                const tot = items[0].chart.data.datasets.reduce((a, ds) => a + (ds.data[i] || 0), 0);
                return tot > items[0].raw ? `Total: ${numBR(tot, 1)} t` : '';
              },
            }
          }
        },
        layout: { padding: { top: 18 } },
        scales: {
          x: { grid: { display: false } },
          y: { beginAtZero: true, title: { display: true, text: 'ton', font: { size: 12.1 } } }
        }
      },
      plugins: [barLabels(v => numBR(v, 1), { id: 'reston' })]
    });
  } else {
    semDados('chart-res-conama');
  }

  // ─── Classe A e Classe B — tipos de resíduo por empreendimento ───
  // Um painel por tipo de resíduo, agrupados pela classe (A antes de B). Cada
  // painel tem ESCALA PRÓPRIA e barras horizontais com o nome do empreendimento
  // no eixo: num gráfico único o entulho (até ~120 t) achatava os tipos pequenos
  // (papel, plástico: 0,1–2 t) a ponto de não dar para ver qual obra era qual.
  // Toneladas, no recorte obra × período.
  {
    const badgeCT = document.getElementById('res-classe-tipo-badge');
    const boxCT   = document.getElementById('res-classe-tipo-paineis');
    // Painéis da renderização anterior: destrói os gráficos antes de refazer o HTML
    Object.keys(chartReg).filter(k => k.startsWith('chart-res-ct-')).forEach(k => {
      chartReg[k].destroy(); delete chartReg[k];
    });
    const tipoTon = {}, totCl = {};
    filtRes.forEach(r => {
      const cl = classeKey(r);
      const tp = String(r.tipo == null ? '' : r.tipo).trim();
      const o  = String(r.obra || '').trim();
      if (!cl || cl === 'NAN' || !tp || tp.toLowerCase() === 'nan' || !o || !(r.kg > 0)) return;
      const t = r.kg / 1000;
      const k = cl + '\u0000' + tp;
      if (!tipoTon[k]) tipoTon[k] = { cl, tp, total: 0, porObra: {} };
      tipoTon[k].total += t;
      tipoTon[k].porObra[o] = (tipoTon[k].porObra[o] || 0) + t;
      totCl[cl] = (totCl[cl] || 0) + t;
    });
    // Tipos: classe A → D; dentro da classe, maior tonelagem primeiro
    const tiposCT = Object.values(tipoTon).sort((x, y) => x.cl.localeCompare(y.cl) || y.total - x.total);
    const classesCT = Object.keys(totCl).sort();
    const escCT = v => String(v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
    // Toneladas pequenas ganham mais casas para não virar "0,0"
    const fmtTon = v => numBR(v, v < 1 ? 2 : 1) + ' t';

    if (!tiposCT.length) {
      if (badgeCT) badgeCT.textContent = hasRaw ? 'sem dados no recorte' : 'sem dado por tipo';
      if (boxCT) boxCT.innerHTML = '<div class="empty-msg">Sem dados no recorte</div>';
    } else if (boxCT) {
      if (badgeCT) badgeCT.textContent = classesCT.map(c => `Classe ${c}: ${numBR(totCl[c], 1)} t`).join(' · ');
      let html = '';
      classesCT.forEach(cl => {
        html += `<div class="ct-classe"><span class="ct-dot" style="background:${conamaColors[cl] || '#888780'}"></span>`
              + `Classe ${escCT(cl)} <span class="ct-tot">· ${numBR(totCl[cl], 1)} t</span></div><div class="ct-grid">`;
        tiposCT.forEach((x, i) => {
          if (x.cl !== cl) return;
          const nObras = Object.keys(x.porObra).length;
          html += `<div class="ct-painel"><div class="ct-painel-titulo"><span>${escCT(x.tp)}</span><span>${fmtTon(x.total)}</span></div>`
                + `<div style="position:relative; height:${Math.max(90, nObras * 22 + 24)}px"><canvas id="chart-res-ct-${i}"></canvas></div></div>`;
        });
        html += '</div>';
      });
      boxCT.innerHTML = html;

      tiposCT.forEach((x, i) => {
        const obrasX = Object.entries(x.porObra).sort((p, q) => q[1] - p[1]);
        makeChart('chart-res-ct-' + i, {
          type: 'bar',
          data: {
            labels: obrasX.map(([o]) => o),
            datasets: [{
              label: x.tp,
              data: obrasX.map(([, t]) => t),
              backgroundColor: conamaColors[x.cl] || '#888780',
              borderRadius: 3, maxBarThickness: 16,
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: 'y',
            layout: { padding: { right: 52 } },
            plugins: {
              legend: { display: false },
              tooltip: {
                callbacks: {
                  title: items => items.length ? `${items[0].label} · ${x.tp}` : '',
                  label: ctx => ` ${fmtTon(ctx.raw)} (${numBR(x.total > 0 ? ctx.raw / x.total * 100 : 0, 1)}% do tipo)`,
                }
              }
            },
            scales: {
              x: { display: false, beginAtZero: true },
              y: { grid: { display: false }, ticks: { autoSkip: false, font: { size: 11 } } }
            }
          },
          plugins: [barLabels(v => fmtTon(v), { id: 'resct' + i })]
        });
      });
    }
  }

  // Volume por tipo de resíduo — função compartilhada com a aba Obra Individual
  renderResTipo('chart-res-tipo', 'res-tipo-badge', filtRes, tipoSrc, hasRaw);

  // ─── Série mensal de resíduos (kg, kg/m², m³, m³/m²) ───
  const resMensal = D.residuos_mensal || [];
  const resMensalEl = document.getElementById('res-mensal-tabela');
  if (resMensalEl) {
    // Respeita o filtro global de período
    let resMensalF = resMensal.filter(r => mesesSet.has(r.mes_ano));
    if (!resMensalF.length) {
      resMensalEl.innerHTML = '<div class="empty-msg">Sem série mensal de resíduos no recorte</div>';
    } else {
      // máximo de 12 meses na tabela; se filtro for 6m, vão aparecer só 6
      const ultimos = resMensalF.slice(-12);
      // Ordem das colunas e legenda das unidades conforme definido pela área:
      //   m³ = volume · kg = peso dos resíduos
      //   m³/m² = volume por m² construído · kg/m² = peso por m² construído
      const subCab = t => `<div style="font-weight:400; font-size:9px; color:var(--gz4); margin-top:1px">${t}</div>`;
      let html = `<table class="mini-table" style="width:100%">
        <thead><tr>
          <th>Mês/Ano</th>
          <th style="text-align:right">m³${subCab('volume')}</th>
          <th style="text-align:right">kg${subCab('peso resíduos')}</th>
          <th style="text-align:right">m³/m²${subCab('volume/m² construído')}</th>
          <th style="text-align:right">kg/m²${subCab('peso/m² construído')}</th>
          <th style="text-align:center">Status</th>
        </tr></thead><tbody>`;
      // Meta efetiva: se filtrou Senna → 50; senão usa meta default (36,6)
      const metaSenna = (D.meta_residuos && D.meta_residuos.senna) || 50;
      const metaDefault = (D.meta_residuos && D.meta_residuos.default) || 36.6;
      const metaEfetiva = (obraF && obraF.toLowerCase().includes('senna')) ? metaSenna : metaDefault;
      ultimos.forEach(r => {
        const meta = metaEfetiva;
        // Semáforo: kg/m² ≤ meta verde; até 1.2× meta âmbar; acima vermelho
        let cls = 'sem-gray';
        if (r.kg_m2 != null) {
          if (r.kg_m2 <= meta) cls = 'sem-green';
          else if (r.kg_m2 <= meta * 1.2) cls = 'sem-amber';
          else cls = 'sem-red';
        }
        html += `<tr>
          <td style="font-weight:500">${r.mes_label}</td>
          <td style="text-align:right">${numBR(r.m3, 2)}</td>
          <td style="text-align:right">${numBR(r.kg, 0)}</td>
          <td style="text-align:right">${numBR(r.m3_m2, 3)}</td>
          <td style="text-align:right">${numBR(r.kg_m2, 2)}</td>
          <td style="text-align:center"><span class="semaforo ${cls}"></span></td>
        </tr>`;
      });
      html += '</tbody></table>';
      html += `<div style="font-size:10px; color:var(--gz4); margin-top:8px"><i class="ti ti-info-circle"></i> kg/m² e m³ vêm de Fat_soma_residuos e Fat_Residuos_m³_m² · Meta: ${metaEfetiva.toFixed(1).replace('.', ',')} kg/m² ${obraF && obraF.toLowerCase().includes('senna') ? '(Senna)' : '(demais obras · Senna usa 50)'}</div>`;
      resMensalEl.innerHTML = html;
    }
  }

  // ─── % Resíduos perigosos — matriz Meta × Realizado ───
  const resPer = D.residuos_perigosos || [];
  const resPerEl = document.getElementById('res-perigosos-tabela');
  if (resPerEl) {
    // Respeita o filtro global de período
    const resPerF = resPer.filter(r => mesesSet.has(r.mes_ano)).slice(-12);
    if (!resPerF.length) {
      resPerEl.innerHTML = '<div class="empty-msg">Sem meses no recorte selecionado</div>';
    } else {
      const meses = resPerF;
      let html = `<table class="mini-table" style="width:100%">
        <thead><tr>
          <th></th>
          ${meses.map(m => `<th style="text-align:center">${m.mes_label}</th>`).join('')}
        </tr></thead><tbody>
        <tr>
          <td style="font-weight:500">Meta</td>
          ${meses.map(m => `<td style="text-align:center; color:var(--fg2); font-weight:500">${m.meta}%</td>`).join('')}
        </tr>
        <tr>
          <td style="font-weight:500">Realizado</td>
          ${meses.map(m => `<td style="text-align:center; color:var(--gz4)">${m.realizado != null ? m.realizado + '%' : '—'}</td>`).join('')}
        </tr>
      </tbody></table>`;
      html += '<div style="font-size:10px; color:var(--gz4); margin-top:8px"><i class="ti ti-info-circle"></i> Coluna "destinação adequada" pendente no modelo · CONAMA 307 Classe D (perigosos)</div>';
      resPerEl.innerHTML = html;
    }
  }

  // Auditorias Ambientais foi movida para a aba Obra Individual (renderObra)

  // Tabela detalhe por obra (filtra por obra × período)
  // Antes só a obra filtrava — o período era ignorado e a tabela sempre
  // mostrava o agregado histórico de D.res_por_obra, mesmo com um período
  // selecionado no topo.
  let obras = D.obras || [];
  if (obraF) obras = obras.filter(o => o === obraF);
  // Meta no formato "50kg/m²" (sem decimal quando inteira: 50 · 36,6)
  const metaFmt = m => numBR(m, Number.isInteger(m) ? 0 : 1) + 'kg/m²';

  // "Todo o período" (ou sem dataset granular) → usa o agregado pronto,
  // sem recomputar; com período ativo → recalcula do recorte já filtrado (filtRes)
  const resPorObraRecorte = {};
  if (isAllAll || !hasRaw) {
    obras.forEach(o => { resPorObraRecorte[o] = D.res_por_obra[o] || {}; });
  } else {
    obras.forEach(o => {
      const linhas = filtRes.filter(r => r.obra === o);
      const kgM2Vals = linhas.filter(r => r.kg_m2 != null && r.kg_m2 > 0).map(r => r.kg_m2);
      resPorObraRecorte[o] = {
        kg_m2: kgM2Vals.length ? kgM2Vals.reduce((a, b) => a + b, 0) / kgM2Vals.length : null,
        meta: (D.res_por_obra[o] || {}).meta || metaObra(o),
        total_kg: linhas.reduce((a, r) => a + (r.kg || 0), 0),
      };
    });
  }

  let tblHtml = '<table class="mini-table"><thead><tr><th>Obra</th><th>Kg/m²</th><th>Meta</th><th>Status</th><th>Total kg</th></tr></thead><tbody>';
  obras.forEach(o => {
    const r = resPorObraRecorte[o] || {};
    const v = r.kg_m2; const m = r.meta || 50;
    const cls = v == null ? 'sem-gray' : v <= m ? 'sem-green' : v <= m * 1.2 ? 'sem-amber' : 'sem-red';
    tblHtml += `<tr><td>${o}</td><td>${numBR(v, 2)}</td><td>${metaFmt(m)}</td><td><span class="semaforo ${cls}"></span></td><td>${numBR(r.total_kg||0, 0)}</td></tr>`;
  });
  tblHtml += '</tbody></table>';
  document.getElementById('res-tabela-detalhe').innerHTML = tblHtml;
}

// ═══════════════════════════════════════════════════════════
//  Auditorias Ambientais — matriz Meta × Realizado
// ═══════════════════════════════════════════════════════════
function renderAuditorias(obraF, periodoF) {
  if (obraF === undefined) obraF = (document.getElementById('filter-obra-global')||{}).value || '';
  // O filtro de período próprio do card tem PRECEDÊNCIA sobre o global
  const cardPeriodo = (document.getElementById('filter-aud-mes')||{}).value || 'all';
  const periodoEfetivo = cardPeriodo !== 'all' ? cardPeriodo : (
    periodoF !== undefined ? periodoF : ((document.getElementById('filter-periodo')||{}).value || 'all')
  );

  let meses = (D.auditorias_meses || []).slice();
  let ocorr = (D.auditorias_ocorrencias || []).slice();

  // Filtro de obra → recomputa do raw
  if (obraF) {
    const raw = (D.auditorias_obra_mes || []).filter(r => r.obra === obraF);
    const map = {};
    raw.forEach(r => {
      if (!map[r.mes_ano]) {
        map[r.mes_ano] = { mes_ano: r.mes_ano, mes_label: r.mes_label, meta: 0, sim: 0, nao: 0, sem_resp: 0 };
      }
      map[r.mes_ano].meta += 1;
      if (r.sim) map[r.mes_ano].sim += 1;
      else if (r.nao) map[r.mes_ano].nao += 1;
      else map[r.mes_ano].sem_resp += 1;
    });
    Object.values(map).forEach(m => { m.realizado = m.sim; });
    meses = Object.values(map).sort((a, b) => a.mes_ano.localeCompare(b.mes_ano));
    ocorr = raw.filter(r => r.ocorrencia).reverse()
      .map(r => ({ obra: r.obra, mes_label: r.mes_label, texto: r.ocorrencia }))
      .slice(0, 5);
  }

  // Filtro de período (mesmas opções do filtro global: all / 6m / 12m / 2025 / 2026)
  if (periodoEfetivo && periodoEfetivo !== 'all') meses = meses.filter(m => noPeriodo(m.mes_ano, periodoEfetivo));

  // Filtrar ocorrências pelo mesmo recorte (pelos mes_ano que sobraram)
  if (meses.length && (D.auditorias_meses || []).length) {
    const mesesValidos = new Set(meses.map(m => m.mes_label));
    ocorr = ocorr.filter(o => mesesValidos.has(o.mes_label));
  }

  const tblEl = document.getElementById('aud-tabela');
  const ocorrEl = document.getElementById('aud-ocorrencias');
  const badge = document.getElementById('aud-badge');
  if (!tblEl || !ocorrEl) return;

  if (!meses.length) {
    tblEl.innerHTML = '<div class="empty-msg">Sem auditorias no recorte selecionado</div>';
    if (badge) badge.textContent = obraF ? `0 · ${obraF}` : '0 auditorias';
  } else {
    let html = '<table class="mini-table" style="width:100%"><thead><tr><th></th>';
    meses.forEach(m => { html += `<th style="text-align:center">${m.mes_label}</th>`; });
    html += '</tr></thead><tbody>';

    // Meta (planejadas)
    html += '<tr><td style="font-weight:500">Meta (planejadas)</td>';
    meses.forEach(m => {
      html += `<td style="text-align:center; color:var(--fg2); font-weight:500">${m.meta}</td>`;
    });
    html += '</tr>';

    // Realizado — Sim
    html += '<tr><td style="font-weight:500; color:var(--fg3)">✓ Realizado (Sim)</td>';
    meses.forEach(m => {
      html += `<td style="text-align:center; color:var(--fg3); font-weight:500">${m.sim != null ? m.sim : m.realizado || 0}</td>`;
    });
    html += '</tr>';

    // Realizado — Não
    html += '<tr><td style="font-weight:500; color:var(--vr3)">✗ Não realizado</td>';
    meses.forEach(m => {
      const v = m.nao != null ? m.nao : (m.meta - (m.realizado || 0));
      html += `<td style="text-align:center; color:var(--vr3); font-weight:500">${v}</td>`;
    });
    html += '</tr>';

    // % Conclusão (Sim/Meta)
    html += '<tr><td style="font-weight:500; color:var(--gz3); font-size:10px">% Conclusão</td>';
    meses.forEach(m => {
      const sim = m.sim != null ? m.sim : (m.realizado || 0);
      const pct = m.meta > 0 ? (sim / m.meta * 100) : 0;
      const sem = pct >= 100 ? 'sem-green' : pct >= 70 ? 'sem-amber' : 'sem-red';
      html += `<td style="text-align:center; color:var(--gz3); font-size:10px">${pct.toFixed(0)}% <span class="semaforo ${sem}" style="margin-left:4px"></span></td>`;
    });
    html += '</tr>';

    html += '</tbody></table>';
    tblEl.innerHTML = html;

    // Badge: totalizadores
    const tMeta = meses.reduce((s, m) => s + m.meta, 0);
    const tSim = meses.reduce((s, m) => s + (m.sim != null ? m.sim : m.realizado || 0), 0);
    const tPct = tMeta > 0 ? Math.round(tSim / tMeta * 100) : 0;
    if (badge) badge.textContent = obraF
      ? `${tSim}/${tMeta} · ${tPct}% · ${obraF}`
      : `${tSim}/${tMeta} · ${tPct}% no recorte`;
  }

  // ─── Principais ocorrências ───
  if (!ocorr.length) {
    ocorrEl.innerHTML = '<em style="color:var(--gz4)">— sem ocorrências registradas no recorte —</em>';
  } else {
    ocorrEl.innerHTML = ocorr.slice(0, 5).map(o =>
      `<div style="margin-bottom:6px; padding-bottom:6px; border-bottom:1px dashed var(--gz6)">
         <strong style="color:var(--fg2)">${o.obra}</strong>
         <span style="color:var(--gz4); font-size:10px"> · ${o.mes_label}</span>
         <div style="margin-top:2px">${o.texto}</div>
       </div>`
    ).join('');
  }
}

// ═══════════════════════════════════════════════════════════
//  Filtros
// ═══════════════════════════════════════════════════════════
function initFilters() {
  const sel = document.getElementById('filter-obra-global');
  (D.obras || []).forEach(o => {
    const opt = document.createElement('option');
    opt.value = o; opt.textContent = o;
    sel.appendChild(opt);
  });
  initFilterBarSombra();
  initSetorSelects();
}
// Popula os selects de setor dos blocos de água e energia.
// A lista vem do payload (setores presentes em água OU energia); sem ela
// (caminho legado via Excel, sem o dataset granular) os selects ficam só com
// "Todos os setores" e o filtro simplesmente não tem efeito.
function initSetorSelects() {
  const setores = D.consumo_setores_list || [];
  ['filter-energia-setor', 'filter-agua-setor'].forEach(id => {
    const sel = document.getElementById(id);
    if (!sel) return;
    setores.forEach(st => {
      const opt = document.createElement('option');
      opt.value = st; opt.textContent = st;
      sel.appendChild(opt);
    });
  });
}
// A filter-bar já é sticky por CSS (top: 52px, colada abaixo do header).
// Isto só liga/desliga a sombra (.is-stuck) quando ela está de fato colada,
// usando a sentinela de 1px que fica logo acima dela no HTML.
function initFilterBarSombra() {
  const barra = document.getElementById('filter-bar-consolidado');
  const sentinela = document.getElementById('filter-bar-sentinela');
  if (!barra || !sentinela || typeof IntersectionObserver === 'undefined') return;
  new IntersectionObserver(
    ([entry]) => barra.classList.toggle('is-stuck', !entry.isIntersecting),
    { rootMargin: '-52px 0px 0px 0px', threshold: 0 }
  ).observe(sentinela);
}
function applyFilters() {
  const obraF  = document.getElementById('filter-obra-global').value;
  const anoF   = document.getElementById('filter-periodo').value;

  // Heatmap filtrado
  renderHeatmap(obraF, anoF);
  renderTopNcs(obraF, anoF);

  // Gráficos filtrados por obra × período
  renderResBarras(obraF, anoF);
  renderResSenna(obraF, anoF);
  // renderIpaObras movido para a aba Obra Individual (renderObra)
  renderRecPorObra(obraF, anoF);
  renderMonitoramentoRuido(obraF, anoF);
  renderEmergencias(obraF, anoF);
  renderPaStatus(obraF, anoF);   // não filtra por período — ver comentário na função
  renderSGAAnual(anoF);          // só reage às opções de ano (fonte é anual)
  // renderAuditorias movida para a aba Obra Individual (chamada dentro de renderObra)

  // Energia + Água + Impressões: filtro do card tem precedência; em branco,
  // seguem o filtro global de obra × período
  renderConsumoEnergiaAll();
  renderConsumoAguaAll();
  renderImpressoes();

  // KPIs: se obra selecionada, troca para valores daquela obra
  updateKpisFiltered(obraF, anoF);

  // Propaga filtros para as outras páginas
  // Página "Obra individual" lê direto do filtro global (sem select próprio)
  renderObra(obraF, anoF);

  // Página "Resíduos" filtra por obra/período
  renderResiduos(obraF, anoF);

  // Badge no heatmap — mostra a JANELA real (ex.: "últimos 6 meses · Dez/25 – Mai/26"),
  // para ficar explícito qual recorte de calendário está valendo em toda a tela.
  const badge = document.getElementById('heatmap-badge');
  if (badge) {
    let periodoLabel = labelPeriodo(anoF);
    if (anoF === '6m' || anoF === '12m') periodoLabel += ' · ' + faixaPeriodo(anoF);
    badge.textContent = obraF
      ? (anoF !== 'all' ? obraF + ' · ' + periodoLabel : obraF)
      : periodoLabel;
  }
}

// Recalcula os KPIs do Consolidado dentro do recorte de período.
// Devolve null quando não há período selecionado (usa os agregados prontos) ou
// quando falta série mensal, para nunca inventar número.
function kpisDoPeriodo(obraF, anoF) {
  if (!anoF || anoF === 'all') return null;
  const obras = obraF ? [obraF] : (D.obras || []);
  const out = {};

  // IPA: média dos meses do recorte na série ipa_mensal
  const mesesIpa = mesesDoPeriodo(anoF, Object.values(D.ipa_mensal || {}).flat().map(x => x.mes));
  const ipaVals = [];
  obras.forEach(o => {
    const serie = (D.ipa_mensal || {})[o] || [];
    const noRec = serie.filter(x => mesesIpa.has(x.mes) && x.ipa != null);
    if (noRec.length) ipaVals.push(noRec.reduce((s, x) => s + x.ipa, 0) / noRec.length);
  });
  if (ipaVals.length) out.ipa = Math.round(ipaVals.reduce((s, v) => s + v, 0) / ipaVals.length * 10) / 10;

  // Conformidade: média das células do heatmap dentro do recorte
  const mesesConf = mesesDoPeriodo(anoF, D.conf_meses || []);
  const confVals = [], metaVals = [];
  obras.forEach(o => {
    const linha = (D.conf_heatmap || {})[o] || {};
    Object.entries(linha).forEach(([m, c]) => {
      if (!mesesConf.has(m) || !c || c.perc == null) return;
      confVals.push(c.perc);
      if (c.meta > 0) metaVals.push(c.meta);  // meta 0 = não lançada, fica fora
    });
  });
  if (confVals.length) {
    out.conf = Math.round(confVals.reduce((s, v) => s + v, 0) / confVals.length * 10) / 10;
    if (metaVals.length) out.conf_meta = Math.round(metaVals.reduce((s, v) => s + v, 0) / metaVals.length * 10) / 10;
  }

  // Resíduos kg/m²: soma das linhas do recorte (mesmo critério da aba Resíduos)
  const raw = D.residuos_raw || [];
  if (raw.length) {
    const mesesRes = mesesDoPeriodo(anoF, raw.map(r => r.mes_ano));
    let linhas = raw.filter(r => mesesRes.has(r.mes_ano) && r.kg_m2 > 0);
    if (obraF) linhas = linhas.filter(r => r.obra === obraF);
    out.kg_m2 = linhas.length
      ? Math.round(linhas.reduce((s, r) => s + r.kg_m2, 0) * 100) / 100
      : 0;
  }

  // Reclamações: contagem real no recorte
  out.rec = recFiltradas(obraF, anoF).length;
  return out;
}

function updateKpisFiltered(obraF, anoF) {
  if (anoF === undefined) anoF = periodoGlobal();
  const kp = kpisDoPeriodo(obraF, anoF);
  const g = D.global;
  const emerg = emergFiltradas(obraF, anoF).length;
  if (!obraF) {
    // Sem filtro de obra: com período selecionado os KPIs vêm do recorte; sem ele, dos agregados
    document.getElementById('kpis-global').innerHTML = kpiCardsHtml({
      ipaLabel: 'IPA global',
      ipa:   (kp && kp.ipa   != null) ? kp.ipa   : g.ipa,
      conf:  (kp && kp.conf  != null) ? kp.conf  : g.conf_media,
      confMeta: null,
      pa:    g.pa_no_prazo_pct,
      emerg, emergSub: emergSubDe(anoF),
      kgM2:  (kp && kp.kg_m2 != null) ? kp.kg_m2 : g.media_kg_m2,
      resMeta: undefined, resSub: 'kg/m² médio · Senna 50 · demais 36,6',
      rec:   (kp && kp.rec   != null) ? kp.rec   : g.total_rec, recLimite: 5,
    });
    return;
  }

  // Com filtro de obra — período selecionado troca os agregados da obra pelo recorte
  const ipa_d  = D.ipa[obraF] || {};
  const conf_d = D.conf_atual[obraF] || {};
  const resAg  = D.res_por_obra[obraF] || {};
  document.getElementById('kpis-global').innerHTML = kpiCardsHtml({
    ipaLabel: 'IPA — ' + obraF,
    ipa:   (kp && kp.ipa != null) ? kp.ipa : (ipa_d.ipa || 0),
    conf:  (kp && kp.conf != null) ? kp.conf : conf_d.perc,
    // Com período selecionado e nenhuma meta lançada nele, vale a regra padrão
    // de 90% (igual à Obra Individual), e não a meta do histórico completo
    confMeta: (kp ? kp.conf_meta : conf_d.meta) || 90,
    pa:    paPctDe(obraF),
    emerg, emergSub: emergSubDe(anoF),
    kgM2:  (kp && kp.kg_m2 != null) ? kp.kg_m2 : resAg.kg_m2,
    resMeta: resAg.meta, resSub: `kg/m² · meta: ${resAg.meta || 36.6} kg/m²`,
    rec:   (kp && kp.rec != null) ? kp.rec : (D.rec_por_obra[obraF] || 0), recLimite: 3,
  });
}

// ═══════════════════════════════════════════════════════════
//  Dicionário de dados e termos
// ═══════════════════════════════════════════════════════════
// A aba é conteúdo estático, com duas exceções lidas dos dados para não
// envelhecer sozinhas: as metas de resíduos e a contagem de verbetes.

// Busca sem acento e sem caixa — "residuo" acha "Resíduo".
function dicNorm(s) {
  return (s || '').toString().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
}

// Meta sem decimal quando é inteira (50), com 1 casa quando não é (36,6).
function dicMeta(v) {
  if (v == null) return '—';
  return Number.isInteger(Number(v)) ? numBR(v, 0) : numBR(v, 1);
}

function renderDicionario() {
  const metas = (D && D.meta_residuos) || {};
  const elDef = document.getElementById('dic-meta-default');
  const elSen = document.getElementById('dic-meta-senna');
  if (elDef && metas.default != null) elDef.textContent = dicMeta(metas.default);
  if (elSen && metas.senna != null) elSen.textContent = dicMeta(metas.senna);
  filtrarDicionario();
}

function irParaDic(id) {
  const el = document.getElementById(id);
  if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function filtrarDicionario() {
  const campo = document.getElementById('dic-busca');
  const q = dicNorm(campo ? campo.value : '').trim();
  let visiveis = 0, total = 0;

  document.querySelectorAll('#page-dicionario .dic-sec').forEach(sec => {
    const linhas = sec.querySelectorAll('tr.dic-row');
    let n = 0;
    linhas.forEach(tr => {
      const ok = !q || dicNorm(tr.textContent).includes(q);
      tr.classList.toggle('dic-hidden', !ok);
      if (ok) n++;
    });
    total += linhas.length;
    visiveis += n;
    // Seção inteira sai de cena quando nenhum verbete dela casa com a busca.
    sec.classList.toggle('dic-hidden', !!q && n === 0);
    const vazio = sec.querySelector('.dic-vazio');
    if (vazio) vazio.classList.add('dic-hidden');
  });

  const semRes = document.getElementById('dic-sem-resultado');
  if (semRes) semRes.classList.toggle('dic-hidden', !(q && visiveis === 0));

  const cont = document.getElementById('dic-contador');
  if (cont) cont.textContent = q ? (visiveis + ' de ' + total + ' verbetes') : (total + ' verbetes');
}

// ═══════════════════════════════════════════════════════════
//  Init
// ═══════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => {
  initFilters();
  initObraSelect();
  initImpFilters();
  renderConsolidado();
  renderResiduos();
  renderDicionario();
});
</script>
</body>
</html>
"""

DATA_DATE = TODAY.strftime("%d/%m/%Y")
HTML = HTML.replace('__DATA_JSON__', DATA_JSON)
HTML = HTML.replace('__DATA_DATE__', DATA_DATE)

print(f"Salvando em:\n  {OUTPUT_PATH}")
with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
    f.write(HTML)

size_kb = os.path.getsize(OUTPUT_PATH) // 1024
print(f"Dashboard gerado com sucesso! ({size_kb} KB)")
print(f"\nAbra no navegador:\n  {OUTPUT_PATH}")
