# -*- coding: utf-8 -*-
"""
Dashboard FG Viva+ / Sustentabilidade — Gerador HTML a partir do .pbix
Lê o modelo semântico do Power BI (Fat_* e Dim_*) e gera dashboard_fg.html
"""
import pandas as pd
import numpy as np
import json, os, sys, re, warnings, base64
from datetime import datetime, date
from pbixray import PBIXRay
warnings.filterwarnings('ignore')

# ════════════════════════════════════════════════════════════════════════════
#  CONFIGURAÇÃO DE CAMINHOS
#  ─────────────────────────
#  PBIX_PATH    → modelo Power BI (fonte de dados — atualizado via PBI Service)
#  OUTPUT_PATH  → arquivo HTML gerado (mesma pasta do script)
#  TEMPLATE_PATH → gerar_dashboard.py (template HTML — mesma pasta do script)
#
#  Estrutura recomendada em produção (servidor 10.1.1.151):
#    C:\Apps\Dash_ambiental_qualidade\
#      ├── gerar_dashboard.py            (template HTML/CSS/JS)
#      ├── gerar_dashboard_pbix.py       (este arquivo — lê PBIX e gera HTML)
#      ├── logo-fg.png                   (logo embutida em base64)
#      ├── dashboard_fg.html             (saída — abrir no navegador)
#      └── Atualizar_Dashboard_PBI.bat   (entry point do usuário)
# ════════════════════════════════════════════════════════════════════════════
# Detecta se está rodando como .exe (PyInstaller) ou como script Python normal.
# Quando empacotado:
#   - sys.executable → caminho do .exe na U:\ (onde os usuários clicam)
#   - sys._MEIPASS   → pasta temporária com os arquivos embutidos (template, logo)
if getattr(sys, 'frozen', False):
    SCRIPT_DIR = os.path.dirname(sys.executable)   # onde o .exe vive (U:\...\POWER BI\)
    BUNDLE_DIR = sys._MEIPASS                       # arquivos embutidos no .exe
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
    BUNDLE_DIR = SCRIPT_DIR

# Fonte de dados — prioridade:
#   1) Copia LOCAL na mesma pasta do script (produção no servidor 10.1.1.151)
#      → independente de mapeamento de unidade de rede
#   2) U:\ (fallback para execução manual de quem tem acesso ao share)
#   3) OneDrive do desenvolvedor (ambiente de desenvolvimento)
_PBIX_CANDIDATES = [
    os.path.join(SCRIPT_DIR, "AMBIENTAL_QUALIDADE_202602_v1_DEV.pbix"),
    r"U:\Qualidade\AMBIENTAL\GESTÃO AMBIENTAL\INDICADORES AMBIENTAIS\POWER BI\AMBIENTAL_QUALIDADE_202602_v1_DEV.pbix",
    r"C:\Users\edueslei.souza\OneDrive - FG BRAZIL HOLDING LTDA\Área de Trabalho\Projetos em adamento\AMBIENTAL_QUALIDADE_202602_v1_DEV.pbix",
]
PBIX_PATH = next((p for p in _PBIX_CANDIDATES if os.path.exists(p)), _PBIX_CANDIDATES[0])

# Saída fica AO LADO do .exe (ou do script, em dev). Template vem do bundle.
OUTPUT_PATH   = os.path.join(SCRIPT_DIR, "dashboard_fg.html")
TEMPLATE_PATH = os.path.join(BUNDLE_DIR, "gerar_dashboard.py")
TODAY = datetime.today()

# ─── Utilitários ──────────────────────────────────────────────────────────────
OBRA_MAP = {
    'BLUE COAST': 'Blue Coast', 'BLUE VIEW': 'Blue View',
    'BOREAL': 'Boreal', 'SAPPHIRE': 'Sapphire', 'TITANIUM': 'Titanium',
    'IMPERIUM': 'Imperium', 'ICONIC': 'Iconic', 'HORIZON': 'Horizon',
    'SENNA': 'Senna', 'SCENARIUM': 'Scenarium', 'GRAND PLACE': 'Grand Place',
    'GARDEN PARK': 'Garden Park', 'GRANDPLACE': 'Grand Place',
    # Dim_reclamacoes_vizinhos grava "Garden"; todas as outras tabelas usam
    # "Garden Park". Sem esta linha as reclamações dessa obra não casavam com
    # OBRAS e sumiam do gráfico "Reclamações de Vizinhos por Obra".
    'GARDEN': 'Garden Park',
}
def norm_obra(s):
    if pd.isna(s): return ''
    s = str(s).strip()
    key = re.sub(r'\s+', ' ', s.upper().split('–')[0].split('-')[0]).strip()
    return OBRA_MAP.get(key, s.strip().title() if s.isupper() else s.strip())

# Meta de resíduos kg/m² por obra: Senna Tower = 50, demais = 36,6
META_RESIDUOS_DEFAULT = 36.6
META_RESIDUOS_SENNA = 50.0
def meta_obra(obra):
    if obra and 'senna' in str(obra).lower():
        return META_RESIDUOS_SENNA
    return META_RESIDUOS_DEFAULT

def num_col(df, col):
    """Coerce to numeric. Para colunas já numéricas (int/float/Int64/Float64),
    retorna direto. Para strings, trata formato brasileiro (1.601,12)."""
    if col not in df.columns:
        return pd.Series(dtype=float, index=df.index)
    s = df[col]
    # Se já é numérico, retorna direto (sem mexer)
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors='coerce')
    # Senão (object/string), aplica conversão BR
    s_str = s.astype(str).str.strip()
    s_clean = s_str.str.replace(r'\.(?=\d{3}(\D|$))', '', regex=True)
    s_clean = s_clean.str.replace(',', '.', regex=False)
    return pd.to_numeric(s_clean, errors='coerce')

class DateEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, (datetime, date)): return obj.strftime('%Y-%m-%d')
        if isinstance(obj, (np.integer,)): return int(obj)
        if isinstance(obj, (np.floating,)): return None if np.isnan(obj) else float(obj)
        if isinstance(obj, np.ndarray): return obj.tolist()
        return super().default(obj)

# ─── Carrega modelo PBIX ──────────────────────────────────────────────────────
print("Lendo .pbix...")
m = PBIXRay(PBIX_PATH)

def T(name):
    """Get table by name (handles encoding variations)."""
    try:
        return m.get_table(name).copy()
    except Exception as e:
        print(f"  Aviso: tabela '{name}' nao encontrada: {e}")
        return pd.DataFrame()

# Fato/Dim
dim_obras = T('Dim_Obras')
dim_obras.columns = [c.strip() for c in dim_obras.columns]
fat_conf = T('Fat_Conformidade_LAI_EIV')
fat_conf.columns = [c.strip() for c in fat_conf.columns]
fat_res = T('Fat_soma_residuos')
fat_res.columns = [c.strip() for c in fat_res.columns]
fat_ruido = T('Fat_Ruido')
fat_ruido.columns = [c.strip() for c in fat_ruido.columns]
fat_al = T('Fat_consumo_agua_luz')
fat_al.columns = [c.strip() for c in fat_al.columns]
dim_aud = T('Dim_Auditoria_Ambientais')
dim_aud.columns = [c.strip() for c in dim_aud.columns]
dim_pa = T('Dim_Plano_acao')
dim_pa.columns = [c.strip() for c in dim_pa.columns]
dim_nc = T('Dim_NC_LAI_EIV')
dim_nc.columns = [c.strip() for c in dim_nc.columns]
dim_rec = T('Dim_reclamacoes_vizinhos')
dim_rec.columns = [c.strip() for c in dim_rec.columns]
dim_emerg = T('Dim_emergencias_ambientais')
dim_emerg.columns = [c.strip() for c in dim_emerg.columns]
dim_sga = T('Dim_Acoes_SGA')
dim_sga.columns = [c.strip() for c in dim_sga.columns]
fat_imp = T('Fat_Impressoes')
fat_imp.columns = [c.strip() for c in fat_imp.columns]
fat_av = T('Fat_avanco_fisico')
fat_av.columns = [c.strip() for c in fat_av.columns]
fat_metas = T('Fat_Metas_LAI_EIV')
fat_metas.columns = [c.strip() for c in fat_metas.columns]
fat_imp_setor = T('Fat_Impressoes_por_setor')
fat_imp_setor.columns = [c.strip() for c in fat_imp_setor.columns] if not fat_imp_setor.empty else fat_imp_setor.columns
fat_res_m3 = T('Fat_Residuos_m³_m²')
fat_res_m3.columns = [c.strip() for c in fat_res_m3.columns] if not fat_res_m3.empty else fat_res_m3.columns

print(f"  Dim_Obras: {len(dim_obras)} linhas")
print(f"  Fat_Conformidade_LAI_EIV: {len(fat_conf)} linhas")
print(f"  Fat_soma_residuos: {len(fat_res)} linhas")
print(f"  Fat_Ruido: {len(fat_ruido)} linhas")
print(f"  Fat_consumo_agua_luz: {len(fat_al)} linhas")

# ─── Normalizar tabelas ───────────────────────────────────────────────────────
print("Normalizando dados...")

# Lista canônica de obras a partir de Dim_Obras
OBRAS_DIM = [norm_obra(o) for o in dim_obras['Obra'].dropna().tolist()] if 'Obra' in dim_obras.columns else []

def prep(df, obra_col='Obra', date_col='Data'):
    if df.empty: return df
    if obra_col in df.columns:
        df['obra'] = df[obra_col].apply(norm_obra)
    if date_col in df.columns:
        df['mes_ano'] = pd.to_datetime(df[date_col], errors='coerce')
    return df

fat_conf = prep(fat_conf)
fat_res = prep(fat_res)
fat_ruido = prep(fat_ruido)
fat_al = prep(fat_al)
dim_aud = prep(dim_aud, date_col='Mês/Ano' if 'Mês/Ano' in dim_aud.columns else 'Data')
dim_pa['obra'] = dim_pa['Obra'].apply(norm_obra) if 'Obra' in dim_pa.columns else ''
dim_nc = prep(dim_nc)
dim_rec = prep(dim_rec, date_col='Mês/Ano' if 'Mês/Ano' in dim_rec.columns else 'Data')
dim_emerg = prep(dim_emerg, date_col='Mês/Ano' if 'Mês/Ano' in dim_emerg.columns else 'Data')
fat_imp = prep(fat_imp, obra_col=None) if 'Obra' not in fat_imp.columns else prep(fat_imp)
fat_av = prep(fat_av)

# Coerce numéricos
if 'PERCENTUAL' in fat_conf.columns: fat_conf['perc'] = num_col(fat_conf, 'PERCENTUAL')
if 'META' in fat_conf.columns: fat_conf['meta'] = num_col(fat_conf, 'META')
if 'Kg' in fat_res.columns: fat_res['kg'] = num_col(fat_res, 'Kg')
if 'Kg/m²' in fat_res.columns: fat_res['kg_m2'] = num_col(fat_res, 'Kg/m²')
if 'Meta Kg' in fat_res.columns: fat_res['meta_kg'] = num_col(fat_res, 'Meta Kg')
if 'm²' in fat_res.columns: fat_res['area_m2'] = num_col(fat_res, 'm²')

# ─── Aviso de obra-mês com ÁREA INVÁLIDA (NÃO descarta) ──────────────────────
# Na fonte, Kg/m² = Kg ÷ m² (confere em 100% das linhas) e a área é o avanço
# físico DO MÊS. Quando a área não é lançada e fica 1,0, a divisão não acontece
# e o Kg/m² sai igual ao Kg — ex.: Senna Tower 04/2026, madeira 1.601,12 kg
# virou 1.601,12 kg/m². Nesse mês a obra teve avanço físico ZERO
# (Fat_avanco_fisico), então o 1,0 foi um conserto manual pra evitar divisão
# por zero.
#
# DECISÃO DA ÁREA (14/09/2026): essas linhas PERMANECEM na base, a pedido — o
# acumulado do Senna Tower precisa mostrar 130,16 kg/m² e o mês de abril precisa
# aparecer no gráfico dele. O impacto de manter: o card "Resíduos gerados" vai
# de 13,61 para 25,89 kg/m², o score de resíduos do Senna cai de 100 para 38,4 e
# o IPA global vai de 79,3 para 78,4. Só logamos o aviso, para a inconsistência
# não passar despercebida por quem mantém a planilha.
# Para voltar a descartar, acrescentar:  fat_res = fat_res[~_area_ruim].copy()
if 'area_m2' in fat_res.columns and 'kg_m2' in fat_res.columns:
    _area_ruim = fat_res['area_m2'].notna() & (fat_res['area_m2'] <= 1) & (fat_res['kg_m2'].fillna(0) > 0)
    if _area_ruim.any():
        _desc = fat_res.loc[_area_ruim]
        _chaves = sorted({(str(r.get('obra', '')), str(r.get('mes_ano'))[:7]) for _, r in _desc.iterrows()})
        print(f"  AVISO: {int(_area_ruim.sum())} linha(s) de residuos com area m2 <= 1 (nao lancada).")
        print("         MANTIDAS na base a pedido da area; o Kg/m2 delas sai igual ao Kg:")
        for _o, _m in _chaves:
            print(f"    - {_o} / {_m}")
if 'Classe' in fat_res.columns:
    fat_res['classe'] = fat_res['Classe'].astype(str).str.strip()
    fat_res.loc[fat_res['classe'].str.lower().isin(['nan', '<na>', 'none', 'null', '']), 'classe'] = ''
if 'Resíduos Detalhados' in fat_res.columns:
    fat_res['tipo'] = fat_res['Resíduos Detalhados'].astype(str).str.strip()
    # Trata <NA>, nan, none, null, vazio → string vazia (será descartado nas agregações)
    fat_res.loc[fat_res['tipo'].str.lower().isin(['nan', '<na>', 'none', 'null', '']), 'tipo'] = ''
if 'dB' in fat_ruido.columns: fat_ruido['db'] = num_col(fat_ruido, 'dB')
if 'Ponto' in fat_ruido.columns: fat_ruido['ponto'] = fat_ruido['Ponto'].astype(str)
if 'Luz Kwh m²/m²' in fat_al.columns: fat_al['kwh_m2'] = num_col(fat_al, 'Luz Kwh m²/m²')
if 'Água m³/m²' in fat_al.columns: fat_al['m3_m2'] = num_col(fat_al, 'Água m³/m²')
if 'Realizado' in dim_aud.columns: dim_aud['realizado'] = dim_aud['Realizado'].astype(str).str.strip()
if 'Status' in dim_pa.columns: dim_pa['status'] = dim_pa['Status'].astype(str).str.strip()
if 'Prazo' in dim_pa.columns: dim_pa['prazo'] = pd.to_datetime(dim_pa['Prazo'], errors='coerce')
if 'Indicador' in dim_pa.columns: dim_pa['indicador'] = dim_pa['Indicador'].astype(str).str.strip()
if 'NC LAI' in dim_nc.columns: dim_nc['nc_lai'] = dim_nc['NC LAI'].astype(str)
if 'NC EIV' in dim_nc.columns: dim_nc['nc_eiv'] = dim_nc['NC EIV'].astype(str)
if 'Categoria' in dim_rec.columns: dim_rec['categoria'] = dim_rec['Categoria'].astype(str).str.strip()
if 'Categoria' in dim_emerg.columns: dim_emerg['categoria'] = dim_emerg['Categoria'].astype(str).str.strip()
if 'Pessoas' in dim_sga.columns: dim_sga['pessoas'] = num_col(dim_sga, 'Pessoas')
if 'Total de pág' in fat_imp.columns: fat_imp['total_pag'] = num_col(fat_imp, 'Total de pág')
if 'Custos' in fat_imp.columns: fat_imp['custo'] = num_col(fat_imp, 'Custos')
if '% P&B' in fat_imp.columns: fat_imp['pct_pb'] = num_col(fat_imp, '% P&B')
if 'Avanço %' in fat_av.columns: fat_av['avanco_pct'] = num_col(fat_av, 'Avanço %')
if 'Total m²' in fat_av.columns: fat_av['total_m2'] = num_col(fat_av, 'Total m²')

# Fat_Residuos_m³_m² (volume + intensidade m³/m²)
fat_res_m3 = prep(fat_res_m3)
for c in fat_res_m3.columns:
    cl = c.lower()
    if 'resíduos m³' in cl and '/' not in cl: fat_res_m3['res_m3'] = num_col(fat_res_m3, c)
    elif 'avanço físico' in cl and 'm²' in cl: fat_res_m3['area_m2'] = num_col(fat_res_m3, c)
    elif 'resíduos m³/m²' in cl or 'm³/m²' in cl: fat_res_m3['res_m3_m2'] = num_col(fat_res_m3, c)

OBRAS = sorted(set([o for o in (OBRAS_DIM + list(fat_conf.get('obra', pd.Series()).unique())) if o and str(o) not in ('', 'nan')]))
print(f"  Obras: {OBRAS}")
print("Processando KPIs...")

# ─── Scores e IPA ─────────────────────────────────────────────────────────────
def score_conf(o):
    d = fat_conf[fat_conf['obra'] == o] if 'obra' in fat_conf.columns else pd.DataFrame()
    if d.empty or 'perc' not in d.columns: return 0.5
    p = d['perc'].mean(); mt = d['meta'].replace(0, np.nan).mean() if 'meta' in d.columns else 0.9
    if pd.isna(mt) or mt == 0: mt = 0.9
    return float(min(p/mt, 1.0)) if not pd.isna(p) else 0.5

def score_res(o):
    d = fat_res[fat_res['obra'] == o] if 'obra' in fat_res.columns else pd.DataFrame()
    if d.empty or 'kg_m2' not in d.columns: return 0.5
    v = d['kg_m2'].replace(0, np.nan).mean()
    # Meta: Senna=50, demais=36,6 (não usa coluna Meta Kg do modelo)
    mt = meta_obra(o)
    if pd.isna(v) or v == 0: return 0.5
    return float(min(mt/v, 1.0))

def score_ruido(o):
    d = fat_ruido[fat_ruido['obra'] == o] if 'obra' in fat_ruido.columns else pd.DataFrame()
    if 'db' not in d.columns or d.empty: return 0.5
    d2 = d.dropna(subset=['db'])
    if d2.empty: return 0.5
    return float((d2['db'] <= 70).sum() / len(d2))

def score_aud(o):
    d = dim_aud[dim_aud['obra'] == o] if 'obra' in dim_aud.columns else pd.DataFrame()
    if 'realizado' not in d.columns or d.empty: return 0.5
    t = len(d); s = d['realizado'].str.lower().isin(['sim', 'yes', 's']).sum()
    return float(s/t) if t > 0 else 0.5

def score_pa(o):
    d = dim_pa[dim_pa['obra'] == o] if 'obra' in dim_pa.columns else pd.DataFrame()
    if 'status' not in d.columns or d.empty: return 0.5
    t = len(d); np_ = d['status'].str.lower().isin(['realizado', 'em andamento']).sum()
    return float(np_/t) if t > 0 else 0.5

# Pesos do IPA definidos pela área (out/2026): conformidade 45%, auditoria 35%,
# resíduos 20% — fecha 100%. Ruído e plano de ação saíram da composição (os
# scores continuam calculados e publicados em D.ipa, só não pesam mais).
IPA_PESOS = {'conf': 0.45, 'aud': 0.35, 'res': 0.20}

def calc_ipa(o):
    return round((score_conf(o)*IPA_PESOS['conf'] + score_aud(o)*IPA_PESOS['aud'] + score_res(o)*IPA_PESOS['res'])*100, 1)

def ipa_class(v):
    if v >= 90: return 'excelente'
    if v >= 80: return 'bom'
    if v >= 70: return 'atencao'
    return 'critico'

ipa_data = {}
for o in OBRAS:
    val = calc_ipa(o)
    ipa_data[o] = {
        'ipa': val, 'classe': ipa_class(val),
        'score_conf': round(score_conf(o)*100,1), 'score_res': round(score_res(o)*100,1),
        'score_ruido': round(score_ruido(o)*100,1), 'score_aud': round(score_aud(o)*100,1),
        'score_pa': round(score_pa(o)*100,1),
    }
global_ipa = round(np.mean([v['ipa'] for v in ipa_data.values()]), 1) if ipa_data else 0

# ─── Heatmap conformidade ────────────────────────────────────────────────────
conf_heatmap = {}; conf_meses = []
if 'mes_ano' in fat_conf.columns and 'perc' in fat_conf.columns:
    d = fat_conf.dropna(subset=['mes_ano', 'perc'])
    meses = sorted(d['mes_ano'].dt.to_period('M').unique())[-12:]
    conf_meses = [str(p) for p in meses]
    for o in OBRAS:
        row = {}
        dd = d[d['obra'] == o]
        for p in meses:
            md = dd[dd['mes_ano'].dt.to_period('M') == p]
            pv = md['perc'].mean() if not md.empty else None
            mv = md['meta'].replace(0, np.nan).mean() if not md.empty and 'meta' in md.columns else None
            row[str(p)] = {
                'perc': round(float(pv)*100,1) if pv is not None and not pd.isna(pv) else None,
                'meta': round(float(mv)*100,1) if mv is not None and not pd.isna(mv) else None,
            }
        conf_heatmap[o] = row

conf_atual = {}
for o in OBRAS:
    d = fat_conf[fat_conf['obra'] == o] if 'obra' in fat_conf.columns else pd.DataFrame()
    if not d.empty and 'perc' in d.columns:
        p = d['perc'].mean(); mt = d['meta'].replace(0, np.nan).mean() if 'meta' in d.columns else 0.9
        conf_atual[o] = {
            'perc': round(float(p)*100,1) if not pd.isna(p) else None,
            'meta': round(float(mt if not pd.isna(mt) else 0.9)*100,1),
        }
    else:
        conf_atual[o] = {'perc': None, 'meta': 90.0}

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
    if 'obra' not in fat_conf.columns or 'mes_ano' not in fat_conf.columns: return None
    d = fat_conf[(fat_conf['obra'] == o) & (fat_conf['mes_ano'].dt.to_period('M') == mp)]
    if d.empty or 'perc' not in d.columns: return None
    p = d['perc'].mean(); mt = d['meta'].mean() if 'meta' in d.columns else 0.9
    if pd.isna(mt) or mt == 0: mt = 0.9
    return float(min(p/mt, 1.0)) if not pd.isna(p) else None

def _score_res_mes(o, mp):
    if 'obra' not in fat_res.columns or 'mes_ano' not in fat_res.columns: return None
    d = fat_res[(fat_res['obra'] == o) & (fat_res['mes_ano'].dt.to_period('M') == mp)]
    if d.empty or 'kg_m2' not in d.columns: return None
    v = d['kg_m2'].replace(0, np.nan).mean()
    if pd.isna(v) or v == 0: return None
    mt = meta_obra(o)
    return float(min(mt/v, 1.0))

def _score_ruido_mes(o, mp):
    if 'obra' not in fat_ruido.columns or 'mes_ano' not in fat_ruido.columns or 'db' not in fat_ruido.columns:
        return None
    d = fat_ruido[(fat_ruido['obra'] == o) & (fat_ruido['mes_ano'].dt.to_period('M') == mp)].dropna(subset=['db'])
    if d.empty: return None
    return float((d['db'] <= 70).sum() / len(d))

def _score_aud_mes(o, mp):
    if 'obra' not in dim_aud.columns or 'mes_ano' not in dim_aud.columns or 'realizado' not in dim_aud.columns:
        return None
    d = dim_aud[(dim_aud['obra'] == o) & (dim_aud['mes_ano'].dt.to_period('M') == mp)]
    if d.empty: return None
    t = len(d); s = d['realizado'].str.lower().isin(['sim', 'yes', 's']).sum()
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
    sc_agg, sr_agg, sru_agg = score_conf(o), score_res(o), score_ruido(o)
    sa_agg, spa_agg = score_aud(o), score_pa(o)
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
    if 'obra' not in fat_ruido.columns or 'mes_ano' not in fat_ruido.columns or 'db' not in fat_ruido.columns:
        return None
    d = fat_ruido[(fat_ruido['obra'] == o) & (fat_ruido['mes_ano'].dt.to_period('M') == mp)].dropna(subset=['db'])
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

# ─── Resíduos ────────────────────────────────────────────────────────────────
res_por_obra = {}
for o in OBRAS:
    d = fat_res[fat_res['obra'] == o] if 'obra' in fat_res.columns else pd.DataFrame()
    if d.empty:
        res_por_obra[o] = {'kg_m2': None, 'meta': meta_obra(o), 'total_kg': 0}
        continue
    km = d['kg_m2'].replace(0, np.nan).sum() if 'kg_m2' in d.columns else None
    tk = d['kg'].sum() if 'kg' in d.columns else 0
    # Meta: Senna=50, demais=36,6
    mt = meta_obra(o)
    res_por_obra[o] = {
        'kg_m2': round(float(km),2) if km is not None and not pd.isna(km) else None,
        'meta': round(float(mt),1),
        'total_kg': round(float(tk),0) if not pd.isna(tk) else 0,
    }

res_classe = {}
if 'classe' in fat_res.columns and 'kg' in fat_res.columns:
    df_c = fat_res[fat_res['classe'] != '']
    if not df_c.empty:
        for k, v in df_c.groupby('classe')['kg'].sum().items():
            s = str(k).strip()
            if s and s.lower() not in ('nan', '<na>', 'none', 'null'):
                res_classe[s] = round(float(v), 0)

res_tipo = {}
if 'tipo' in fat_res.columns and 'kg' in fat_res.columns:
    df_t = fat_res[fat_res['tipo'] != '']
    if not df_t.empty:
        for k, v in df_t.groupby('tipo')['kg'].sum().nlargest(8).items():
            s = str(k).strip()
            if s and s.lower() not in ('nan', '<na>', 'none', 'null'):
                res_tipo[s] = round(float(v), 0)

# Breakdown por obra (para filtros)
res_classe_por_obra = {}
res_tipo_por_obra = {}
INVALID_VALS = ('nan', '<na>', 'none', 'null', '')
if 'obra' in fat_res.columns and 'kg' in fat_res.columns:
    for o in OBRAS:
        d = fat_res[fat_res['obra'] == o]
        rc = {}
        if 'classe' in d.columns and not d.empty:
            for k, v in d.groupby('classe')['kg'].sum().items():
                s = str(k).strip()
                if s and s.lower() not in INVALID_VALS:
                    rc[s] = round(float(v), 0)
        res_classe_por_obra[o] = rc
        rt = {}
        if 'tipo' in d.columns and not d.empty:
            d_t = d[d['tipo'] != '']
            if not d_t.empty:
                for k, v in d_t.groupby('tipo')['kg'].sum().nlargest(8).items():
                    s = str(k).strip()
                    if s and s.lower() not in INVALID_VALS:
                        rt[s] = round(float(v), 0)
        res_tipo_por_obra[o] = rt

total_res_kg = fat_res['kg'].sum() if 'kg' in fat_res.columns else 0
media_kg_m2 = fat_res['kg_m2'].replace(0, np.nan).sum() if 'kg_m2' in fat_res.columns else 0

# ─── Dataset granular de resíduos (para filtro de período no client) ──────────
# Inclui TODAS as linhas (não filtra kg > 0) — necessário pra contar obras/classes/tipos
# que têm registros no recorte, mesmo sem kg preenchido ainda
residuos_raw = []
if not fat_res.empty and 'mes_ano' in fat_res.columns:
    df_rr = fat_res.dropna(subset=['mes_ano']).copy()
    df_rr['kg'] = df_rr['kg'].fillna(0) if 'kg' in df_rr.columns else 0
    df_rr['kg_m2_n'] = df_rr['kg_m2'].fillna(0) if 'kg_m2' in df_rr.columns else 0
    for col in ['classe', 'tipo', 'obra']:
        if col not in df_rr.columns: df_rr[col] = ''
        df_rr[col] = df_rr[col].astype(str).str.strip()
        df_rr[col] = df_rr[col].where(~df_rr[col].str.lower().isin(['nan', 'none']), '')
    df_rr = df_rr[df_rr['obra'] != '']
    if not df_rr.empty:
        gp = df_rr.groupby([
            df_rr['mes_ano'].dt.to_period('M'), 'obra', 'classe', 'tipo'
        ], dropna=False).agg(
            kg=('kg', 'sum'),
            kg_m2=('kg_m2_n', 'sum'),
        ).reset_index()
        gp.columns = ['mes_periodo', 'obra', 'classe', 'tipo', 'kg', 'kg_m2']
        for _, r in gp.iterrows():
            residuos_raw.append({
                'mes_ano': str(r['mes_periodo']),
                'obra': r['obra'],
                'classe': r['classe'],
                'tipo': r['tipo'],
                'kg': round(float(r['kg']), 0),
                'kg_m2': round(float(r['kg_m2']), 4) if not pd.isna(r['kg_m2']) else 0,
            })

# ─── Série mensal de resíduos (kg, kg/m², m³, m³/m²) ───────────────────────
# kg/kg_m2: Fat_soma_residuos · m3/m3_m2: Fat_Residuos_m³_m²
residuos_mensal = []
MESES_PT_ABR_RES = ['','jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez']

# Agrupa kg por mês
kg_by_month = pd.Series(dtype=float)
kg_m2_by_month = pd.Series(dtype=float)
if 'kg' in fat_res.columns and 'mes_ano' in fat_res.columns:
    df_rm = fat_res.dropna(subset=['mes_ano']).copy()
    df_rm = df_rm[df_rm['kg'].fillna(0) > 0]
    if not df_rm.empty:
        kg_by_month = df_rm.groupby(df_rm['mes_ano'].dt.to_period('M'))['kg'].sum()
        if 'kg_m2' in df_rm.columns:
            df_rm_pos = df_rm[df_rm['kg_m2'].fillna(0) > 0]
            kg_m2_by_month = df_rm_pos.groupby(df_rm_pos['mes_ano'].dt.to_period('M'))['kg_m2'].sum()

# Agrupa m³ por mês (de Fat_Residuos_m³_m²)
m3_by_month = pd.Series(dtype=float)
m3_m2_by_month = pd.Series(dtype=float)
if not fat_res_m3.empty and 'mes_ano' in fat_res_m3.columns:
    df_rm3 = fat_res_m3.dropna(subset=['mes_ano']).copy()
    if 'res_m3' in df_rm3.columns:
        df_pos = df_rm3[df_rm3['res_m3'].fillna(0) > 0]
        if not df_pos.empty:
            m3_by_month = df_pos.groupby(df_pos['mes_ano'].dt.to_period('M'))['res_m3'].sum()
    if 'res_m3_m2' in df_rm3.columns:
        df_pos2 = df_rm3[df_rm3['res_m3_m2'].fillna(0) > 0]
        if not df_pos2.empty:
            m3_m2_by_month = df_pos2.groupby(df_pos2['mes_ano'].dt.to_period('M'))['res_m3_m2'].mean()

# União de meses (qualquer fonte com dado)
all_months_res = sorted(set(list(kg_by_month.index) + list(m3_by_month.index)))
for mp in all_months_res:
    item = {
        'mes_ano': str(mp),
        'mes_label': f"{MESES_PT_ABR_RES[mp.month]}/{str(mp.year)[-2:]}",
        'kg': int(round(float(kg_by_month.get(mp, 0)))) if mp in kg_by_month.index else None,
        'kg_m2': round(float(kg_m2_by_month.get(mp)), 2) if mp in kg_m2_by_month.index and not pd.isna(kg_m2_by_month.get(mp)) else None,
        'm3': round(float(m3_by_month.get(mp, 0)), 2) if mp in m3_by_month.index else None,
        'm3_m2': round(float(m3_m2_by_month.get(mp)), 3) if mp in m3_m2_by_month.index and not pd.isna(m3_m2_by_month.get(mp)) else None,
        'meta_kg_m2': META_RESIDUOS_DEFAULT,
    }
    residuos_mensal.append(item)

# ─── % Resíduos perigosos com destinação adequada (matriz Meta vs Realizado) ──
# CONAMA 307: Classe D = perigosos. "Destinação adequada" não consta como coluna
# no modelo, então Realizado fica pendente até o campo ser criado no PBI.
residuos_perigosos = []
# Pega os últimos 6 meses da série acima pra alinhar
ultimos_meses = residuos_mensal[-6:] if len(residuos_mensal) > 6 else residuos_mensal
for item in ultimos_meses:
    residuos_perigosos.append({
        'mes_ano':   item['mes_ano'],
        'mes_label': item['mes_label'],
        'meta':      100,
        'realizado': None,
    })

# ─── Ruído ────────────────────────────────────────────────────────────────────
# Visão grouped bar: dB médio por obra × mês (últimos N meses)
ruido_obra_mes = []
if not fat_ruido.empty and 'db' in fat_ruido.columns and 'obra' in fat_ruido.columns and 'mes_ano' in fat_ruido.columns:
    df_r = fat_ruido.dropna(subset=['obra', 'db', 'mes_ano']).copy()
    df_r = df_r[(df_r['db'] > 0) & (df_r['obra'].astype(str).str.strip() != '')]
    if not df_r.empty:
        g = df_r.groupby([df_r['mes_ano'].dt.to_period('M'), 'obra'])['db'].mean().reset_index()
        g.columns = ['mes_periodo', 'obra', 'db_medio']
        _MESES_ABR = ['','jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez']
        for _, r in g.iterrows():
            mp = r['mes_periodo']
            ruido_obra_mes.append({
                'mes_ano': str(mp),
                'mes_label': f"{str(mp.year)[-2:]}-{_MESES_ABR[mp.month]}",
                'obra': r['obra'],
                'db': int(round(float(r['db_medio']))),
            })

ruido_por_obra = {}
if 'db' in fat_ruido.columns and 'obra' in fat_ruido.columns:
    dv = fat_ruido.dropna(subset=['db','obra'])
    for o in OBRAS:
        d = dv[dv['obra'] == o]
        if d.empty:
            ruido_por_obra[o] = {'media_db': None, 'pct_conforme': None, 'n': 0}
            continue
        ruido_por_obra[o] = {
            'media_db': round(float(d['db'].mean()),1),
            'pct_conforme': round(float((d['db'] <= 70).mean()*100),1),
            'n': len(d),
        }

# ─── Água e Luz ───────────────────────────────────────────────────────────────
al_por_obra = {}
for o in OBRAS:
    d = fat_al[fat_al['obra'] == o] if 'obra' in fat_al.columns else pd.DataFrame()
    kw = d['kwh_m2'].replace(0, np.nan).mean() if 'kwh_m2' in d.columns and not d.empty else None
    m3 = d['m3_m2'].replace(0, np.nan).mean() if 'm3_m2' in d.columns and not d.empty else None
    al_por_obra[o] = {
        'kwh_m2': round(float(kw),3) if kw is not None and not pd.isna(kw) else None,
        'm3_m2': round(float(m3),3) if m3 is not None and not pd.isna(m3) else None,
    }

# Rótulo de setor de Fat_consumo_agua_luz. A coluna `nome_setor` mistura nomes
# de obra em caixa alta (BLUE COAST, SENNA) com locais que não são obra
# (Sede, Depósito, Sala 02, Plantão norte, ITC Busness SL 612, Tour Royalle).
# Normaliza a grafia para casar com as demais tabelas e devolve 'Sem setor'
# quando a célula está vazia, para o filtro nunca perder linha.
def setor_col(df):
    if 'nome_setor' not in df.columns:
        return pd.Series(['Sem setor'] * len(df), index=df.index)
    s = df['nome_setor'].astype(str).str.strip()
    s = s.where(~s.str.lower().isin(['', 'nan', 'none', '<na>']), 'Sem setor')
    return s.map(lambda v: norm_obra(v) if v != 'Sem setor' else v)

# ─── Consumo de Água mensal (todas as obras) ─────────────────────────────────
consumo_agua = []
consumo_agua_m2 = []
consumo_agua_raw = []

# ─── Consumo de Energia mensal (todas as obras) ──────────────────────────────
consumo_energia = []
consumo_energia_m2 = []
consumo_energia_raw = []  # linhas por obra × mês para filtro client-side
MESES_PT_ABR = ['','jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez']
if 'Luz (kwh)' in fat_al.columns and 'mes_ano' in fat_al.columns:
    df_e = fat_al.copy()
    df_e['kwh'] = pd.to_numeric(df_e['Luz (kwh)'], errors='coerce')
    # Intensidade vem DIRETO da coluna 'Luz m²' do modelo (sem cálculo)
    if 'Luz m²' in df_e.columns:
        df_e['intensidade_col'] = pd.to_numeric(df_e['Luz m²'], errors='coerce')
    else:
        df_e['intensidade_col'] = np.nan

    # ───── União de meses com QUALQUER dado (kWh > 0 OU Luz m² > 0) ─────
    # Isso garante que mar/26 apareça mesmo sem kWh, desde que tenha Luz m²
    df_any = df_e.dropna(subset=['mes_ano']).copy()
    df_any['_has'] = (df_any['kwh'].fillna(0) > 0) | (df_any['intensidade_col'].fillna(0) > 0)
    df_master = df_any[df_any['_has']]
    if not df_master.empty:
        all_months = sorted(df_master['mes_ano'].dt.to_period('M').unique())

        # Sums/means com base no df_e completo, indexados por mês
        kwh_by_month = (
            df_e.dropna(subset=['mes_ano']).groupby(df_e['mes_ano'].dt.to_period('M'))['kwh']
            .sum(min_count=1)
        )
        def _mean_pos(s):
            valid = s[s > 0]
            return float(valid.mean()) if len(valid) else None
        int_by_month = (
            df_e.dropna(subset=['mes_ano']).groupby(df_e['mes_ano'].dt.to_period('M'))['intensidade_col']
            .apply(_mean_pos)
        )

        # Constrói série de kWh sobre TODOS os meses (zero quando sem dado)
        for mp in all_months:
            mes_label = f"{MESES_PT_ABR[mp.month]}/{str(mp.year)[-2:]}"
            v = kwh_by_month.get(mp)
            kwh_v = int(round(float(v))) if v is not None and not pd.isna(v) and v > 0 else 0
            consumo_energia.append({
                'mes_ano': str(mp),
                'mes_label': mes_label,
                'kwh': kwh_v,
                'meta_6m': 0,
            })
            iv = int_by_month.get(mp)
            if iv is not None and not pd.isna(iv) and iv > 0:
                consumo_energia_m2.append({
                    'mes_ano': str(mp),
                    'mes_label': mes_label,
                    'kwh_m2': round(float(iv), 2),
                    'meta_6m': 0,
                })

        # Meta rolling 6m (kWh: ignora meses zero pra meta não diluir)
        for i in range(len(consumo_energia)):
            win = consumo_energia[max(0, i-5):i+1]
            nz = [w['kwh'] for w in win if w['kwh'] > 0]
            consumo_energia[i]['meta_6m'] = int(round(sum(nz) / len(nz))) if nz else 0
        for i in range(len(consumo_energia_m2)):
            win = consumo_energia_m2[max(0, i-5):i+1]
            consumo_energia_m2[i]['meta_6m'] = round(sum(w['kwh_m2'] for w in win) / len(win), 2)

    # Raw por obra × setor × mês (mantém qualquer mês com kWh > 0 OU Luz m² > 0)
    # O setor entra no grão para alimentar o filtro de setores do dashboard:
    # `nome_setor` traz tanto obras (BLUE COAST, SENNA...) quanto locais que não
    # são obra (Sede, Depósito, Sala 02, Plantão norte...), estes últimos com
    # Obra = 'Outros'. Sem o setor no grão não dá para separar um do outro.
    if 'obra' in df_e.columns:
        df_raw = df_e.dropna(subset=['mes_ano', 'obra'])
        df_raw = df_raw[(df_raw['obra'] != '') & (df_raw['obra'].str.lower() != 'nan')]
        df_raw['setor_lbl'] = setor_col(df_raw)
        df_raw_agg = df_raw.groupby([df_raw['mes_ano'].dt.to_period('M'), 'obra', 'setor_lbl']).agg(
            kwh=('kwh', 'sum'),
            luz_m2=('intensidade_col', 'mean'),
        ).reset_index()
        df_raw_agg.columns = ['mes_periodo', 'obra', 'setor', 'kwh', 'luz_m2']
        for _, r in df_raw_agg.iterrows():
            mp = r['mes_periodo']
            kwh_v = float(r['kwh']) if not pd.isna(r['kwh']) else 0
            lm2 = float(r['luz_m2']) if not pd.isna(r['luz_m2']) else None
            if kwh_v <= 0 and (lm2 is None or lm2 <= 0): continue
            consumo_energia_raw.append({
                'mes_ano': str(mp),
                'mes_label': f"{MESES_PT_ABR[mp.month]}/{str(mp.year)[-2:]}",
                'obra': r['obra'],
                'setor': r['setor'],
                'kwh': int(round(kwh_v)),
                'kwh_m2': round(lm2, 2) if lm2 is not None else None,
            })

# ─── Água (m³) mensal + Intensidade m³/m² ─────────────────────────────────────
if 'Água (m³)' in fat_al.columns and 'mes_ano' in fat_al.columns:
    df_a = fat_al.copy()
    df_a['m3'] = pd.to_numeric(df_a['Água (m³)'], errors='coerce')
    # Intensidade lida DIRETO da coluna `Água m³/m²` (sem cálculo)
    if 'Água m³/m²' in df_a.columns:
        df_a['m3_m2_col'] = pd.to_numeric(df_a['Água m³/m²'], errors='coerce')
    else:
        df_a['m3_m2_col'] = np.nan

    # União de meses com m³ > 0 OU intensidade > 0
    df_any_a = df_a.dropna(subset=['mes_ano']).copy()
    df_any_a['_has'] = (df_any_a['m3'].fillna(0) > 0) | (df_any_a['m3_m2_col'].fillna(0) > 0)
    df_master_a = df_any_a[df_any_a['_has']]

    if not df_master_a.empty:
        all_months_a = sorted(df_master_a['mes_ano'].dt.to_period('M').unique())
        m3_by_month = (
            df_a.dropna(subset=['mes_ano']).groupby(df_a['mes_ano'].dt.to_period('M'))['m3']
            .sum(min_count=1)
        )
        def _mean_pos_a(s):
            valid = s[s > 0]
            return float(valid.mean()) if len(valid) else None
        m3m2_by_month = (
            df_a.dropna(subset=['mes_ano']).groupby(df_a['mes_ano'].dt.to_period('M'))['m3_m2_col']
            .apply(_mean_pos_a)
        )

        for mp in all_months_a:
            mes_label = f"{MESES_PT_ABR[mp.month]}/{str(mp.year)[-2:]}"
            # m³ total
            v = m3_by_month.get(mp)
            m3_v = int(round(float(v))) if v is not None and not pd.isna(v) and v > 0 else 0
            consumo_agua.append({
                'mes_ano': str(mp),
                'mes_label': mes_label,
                'm3': m3_v,
                'meta_6m': 0,
            })
            # Intensidade m³/m² (só inclui meses com valor positivo)
            iv = m3m2_by_month.get(mp)
            if iv is not None and not pd.isna(iv) and iv > 0:
                consumo_agua_m2.append({
                    'mes_ano': str(mp),
                    'mes_label': mes_label,
                    'm3_m2': round(float(iv), 2),
                    'meta_6m': 0,
                })

        # Rolling meta — m³ (ignora meses zero)
        for i in range(len(consumo_agua)):
            win = consumo_agua[max(0, i-5):i+1]
            nz = [w['m3'] for w in win if w['m3'] > 0]
            consumo_agua[i]['meta_6m'] = int(round(sum(nz) / len(nz))) if nz else 0
        # Rolling meta — m³/m²
        for i in range(len(consumo_agua_m2)):
            win = consumo_agua_m2[max(0, i-5):i+1]
            consumo_agua_m2[i]['meta_6m'] = round(sum(w['m3_m2'] for w in win) / len(win), 2)

    # Raw por obra × setor × mês (m³ + m³/m²) — mesmo grão do dataset de energia
    if 'obra' in df_a.columns:
        df_raw_a = df_a.dropna(subset=['mes_ano', 'obra'])
        df_raw_a = df_raw_a[(df_raw_a['obra'] != '') & (df_raw_a['obra'].str.lower() != 'nan')]
        df_raw_a['setor_lbl'] = setor_col(df_raw_a)
        df_raw_a_agg = df_raw_a.groupby([df_raw_a['mes_ano'].dt.to_period('M'), 'obra', 'setor_lbl']).agg(
            m3=('m3', 'sum'),
            m3_m2=('m3_m2_col', 'mean'),
        ).reset_index()
        df_raw_a_agg.columns = ['mes_periodo', 'obra', 'setor', 'm3', 'm3_m2']
        for _, r in df_raw_a_agg.iterrows():
            mp = r['mes_periodo']
            m3_v = float(r['m3']) if not pd.isna(r['m3']) else 0
            m3m2_v = float(r['m3_m2']) if not pd.isna(r['m3_m2']) else None
            if m3_v <= 0 and (m3m2_v is None or m3m2_v <= 0): continue
            consumo_agua_raw.append({
                'mes_ano': str(mp),
                'mes_label': f"{MESES_PT_ABR[mp.month]}/{str(mp.year)[-2:]}",
                'obra': r['obra'],
                'setor': r['setor'],
                'm3': int(round(m3_v)),
                'm3_m2': round(m3m2_v, 2) if m3m2_v is not None else None,
            })

# ─── Auditoria ────────────────────────────────────────────────────────────────
aud_por_obra = {}
# Considera só linhas com Realizado PREENCHIDO (Sim ou Não), ignora <NA>/null/vazio
if 'obra' in dim_aud.columns and 'realizado' in dim_aud.columns:
    _real_str_all = dim_aud['realizado'].astype(str).str.lower().str.strip()
    _mask_preenchido_all = _real_str_all.isin(['sim', 'não', 'nao', 'yes', 'no'])
    dim_aud_validas = dim_aud[_mask_preenchido_all]
    for o in OBRAS:
        d = dim_aud_validas[dim_aud_validas['obra'] == o]
        t = len(d)
        s = d['realizado'].astype(str).str.lower().str.strip().isin(['sim', 'yes']).sum() if t > 0 else 0
        aud_por_obra[o] = {'total': t, 'realizadas': int(s), 'pct': round(s/t*100, 1) if t > 0 else 0}
    total_aud = int(_mask_preenchido_all.sum())
    aud_sim = int((_real_str_all == 'sim').sum())
else:
    total_aud = 0
    aud_sim = 0

# ─── Auditorias Ambientais — matriz Meta vs Realizado por mês ────────────────
# IMPORTANTE: a coluna `Mês/Ano` em Dim_Auditoria_Ambientais tem dias variando
# por obra (01, 02, 03...) e NÃO reflete o mês real. O mês correto vem das
# colunas TEXTO `Mês` + `Ano`. Usar essas duas pra derivar o período.
_MES_NUM = {
    'JANEIRO': 1, 'FEVEREIRO': 2, 'MARCO': 3, 'MARÇO': 3, 'ABRIL': 4,
    'MAIO': 5, 'JUNHO': 6, 'JULHO': 7, 'AGOSTO': 8, 'SETEMBRO': 9,
    'OUTUBRO': 10, 'NOVEMBRO': 11, 'DEZEMBRO': 12,
}
auditorias_meses = []
auditorias_ocorrencias = []
auditorias_obra_mes = []
if not dim_aud.empty:
    df_au = dim_aud.copy()
    # Encontra colunas reais (case insensitive)
    col_mes_txt = None; col_ano = None; col_real = None; col_ocorr = None
    for c in df_au.columns:
        cl = c.lower().strip()
        if cl == 'mês' or cl == 'mes': col_mes_txt = c
        elif cl == 'ano': col_ano = c
        elif cl == 'realizado' and col_real is None: col_real = c
        elif 'ocorer' in cl or 'ocorrência' in cl or 'ocorrencia' in cl or 'principais' in cl:
            col_ocorr = c

    if col_mes_txt and col_ano:
        df_au['_mes_str'] = df_au[col_mes_txt].astype(str).str.strip().str.upper()
        df_au['_mes_num'] = df_au['_mes_str'].map(_MES_NUM)
        df_au['_ano'] = pd.to_numeric(df_au[col_ano], errors='coerce')
        # Período YYYY-MM derivado das colunas texto
        df_au['_periodo'] = df_au.apply(
            lambda r: f"{int(r['_ano'])}-{int(r['_mes_num']):02d}"
            if pd.notna(r['_ano']) and pd.notna(r['_mes_num']) else None, axis=1)

        if col_real:
            df_au['_real_str'] = df_au[col_real].astype(str).str.lower().str.strip()
            df_au['_sim'] = df_au['_real_str'] == 'sim'
            df_au['_nao'] = df_au['_real_str'].isin(['não', 'nao'])
            df_au['_sem_resp'] = ~(df_au['_sim'] | df_au['_nao'])
        else:
            df_au['_sim'] = False; df_au['_nao'] = False; df_au['_sem_resp'] = True

        # Raw por (mes, obra) — alimenta filtro de obra no JS
        for _, r in df_au.iterrows():
            if pd.isna(r['_periodo']): continue
            ocorr_txt = ''
            if col_ocorr:
                v = r.get(col_ocorr)
                if v is not None and pd.notna(v):
                    vs = str(v).strip()
                    if vs.lower() not in ('nan', 'none', '<na>', ''): ocorr_txt = vs
            mes_num = int(r['_mes_num'])
            ano_v = int(r['_ano'])
            auditorias_obra_mes.append({
                'mes_ano':   r['_periodo'],
                'mes_label': f"{MESES_PT_ABR[mes_num]}/{str(ano_v)[-2:]}",
                'obra':      str(r.get('obra', '') or '').strip(),
                'sim':       bool(r['_sim']),
                'nao':       bool(r['_nao']),
                'sem_resp':  bool(r['_sem_resp']),
                'ocorrencia': ocorr_txt,
            })

        # Agregação global por período
        df_au_valid = df_au.dropna(subset=['_periodo'])
        agg = df_au_valid.groupby('_periodo').agg(
            meta=('_sim', 'count'),
            sim=('_sim', 'sum'),
            nao=('_nao', 'sum'),
            sem_resp=('_sem_resp', 'sum'),
        ).reset_index()
        agg = agg.sort_values('_periodo')
        for _, r in agg.iterrows():
            periodo = r['_periodo']
            ano_str, mes_str = periodo.split('-')
            mes_num = int(mes_str)
            auditorias_meses.append({
                'mes_ano':   periodo,
                'mes_label': f"{MESES_PT_ABR[mes_num]}/{ano_str[-2:]}",
                'meta':      int(r['meta']),
                'sim':       int(r['sim']),
                'nao':       int(r['nao']),
                'sem_resp':  int(r['sem_resp']),
                'realizado': int(r['sim']),  # compatibilidade retro
            })

        # Últimas ocorrências (top 5 mais recentes, por período)
        if col_ocorr:
            df_oc = df_au_valid.copy()
            df_oc = df_oc[df_oc[col_ocorr].notna()]
            df_oc['_oc_str'] = df_oc[col_ocorr].astype(str).str.strip()
            df_oc = df_oc[~df_oc['_oc_str'].str.lower().isin(['nan', 'none', '<na>', ''])]
            df_oc = df_oc.sort_values('_periodo', ascending=False)
            for _, r in df_oc.iterrows():
                periodo = r['_periodo']
                ano_str, mes_str = periodo.split('-')
                mes_num = int(mes_str)
                auditorias_ocorrencias.append({
                    'obra':      str(r.get('obra', '') or '—').strip(),
                    'mes_label': f"{MESES_PT_ABR[mes_num]}/{ano_str[-2:]}",
                    'texto':     r['_oc_str'],
                })
                if len(auditorias_ocorrencias) >= 5: break

# ─── Plano de Ação ────────────────────────────────────────────────────────────
pa_status = {}
if 'status' in dim_pa.columns:
    for s in dim_pa['status'].unique():
        s_clean = str(s).strip() if s is not None else ''
        if s_clean and s_clean.lower() not in ('nan', '<na>', 'none', 'null'):
            pa_status[s_clean] = int((dim_pa['status'] == s).sum())

pa_atrasadas = 0
if 'prazo' in dim_pa.columns and 'status' in dim_pa.columns:
    pa_atrasadas = len(dim_pa[(dim_pa['prazo'] < pd.Timestamp(TODAY)) & (~dim_pa['status'].str.lower().isin(['realizado','nan']))])

total_pa = len(dim_pa); pa_no_prazo_pct = 0
if total_pa > 0 and 'status' in dim_pa.columns:
    n = dim_pa['status'].str.strip().str.lower().isin(['realizado','em andamento']).sum()
    pa_no_prazo_pct = round(float(n/total_pa)*100,1)

pa_por_obra = {}
if 'obra' in dim_pa.columns and 'status' in dim_pa.columns:
    for o in OBRAS:
        d = dim_pa[dim_pa['obra'] == o]
        pa_por_obra[o] = {s: int(d['status'].str.strip().str.lower().eq(s.lower()).sum()) for s in ['Realizado','Em Andamento','Pendente','Atrasado']}

# Lista das ações — alimenta o tooltip do gráfico "Plano de ação Auditorias
# distribuição de status", que abre cada fatia por obra sem precisar de um
# segundo gráfico. Descarta as linhas em branco da planilha.
plano_acao_lista = []
if not dim_pa.empty:
    def _pa_str(v):
        if v is None or (isinstance(v, float) and pd.isna(v)): return ''
        s = str(v).strip()
        return '' if s.lower() in ('nan', 'none', '<na>') else s

    for _, r in dim_pa.iterrows():
        pz = r.get('prazo') if 'prazo' in r else None
        item = {
            'obra':        _pa_str(r.get('obra', '')),
            'status':      _pa_str(r.get('status', '')),
            'indicador':   _pa_str(r.get('indicador', '')),
            'analise':     _pa_str(r.get('Análise', '')),
            'acao':        _pa_str(r.get('Ação', '')),
            'responsavel': _pa_str(r.get('Responsável', '')),
            'prazo':       pz.strftime('%d/%m/%Y') if isinstance(pz, pd.Timestamp) and pd.notna(pz) else '',
        }
        if not any([item['obra'], item['status'], item['indicador'], item['acao']]):
            continue
        plano_acao_lista.append(item)

# ─── NC de resíduos: 1 ocorrência por obra-mês em que o kg/m² realizado
# ultrapassa a meta (Senna Tower 50 kg/m² · demais obras 36,6 kg/m² — meta_obra(),
# mesma usada no IPA e na tabela "Resíduos por Obra").
# NÃO entra mais no card "NC registradas LAI/EIV": esse card voltou a ser
# exclusivamente LAI/EIV (contagem por assunto — ver bloco NC abaixo). O cálculo
# fica mantido e é exportado em `res_nc_por_obra`/`res_nc_total` para uso futuro.
# Importante: aqui usa SOMA dos Kg/m² das linhas do mês (uma linha por tipo de
# resíduo), não média — cada linha já é kg daquele tipo ÷ área, então a
# intensidade real da obra no mês é a soma dessas frações (= kg total ÷ área).
# O resto do dashboard (KPI Resíduos, tabela, IPA) ainda usa média — problema
# em aberto, aguardando decisão — ver memória "kgm2-agregacao-media-vs-soma".
# Corrigido só aqui a pedido da usuária, pra não subestimar as NCs de resíduos.
def _residuos_kgm2_mes(o, mp):
    if 'obra' not in fat_res.columns or 'mes_ano' not in fat_res.columns or 'kg_m2' not in fat_res.columns:
        return None
    d = fat_res[(fat_res['obra'] == o) & (fat_res['mes_ano'].dt.to_period('M') == mp)]
    if d.empty: return None
    return float(d['kg_m2'].sum(skipna=True))

_meses_res_todos = sorted({item['mes_ano'] for item in residuos_mensal})
res_nc_mensal = {}  # obra -> {mes_str: True/False (excedeu meta naquele mês)}
for o in OBRAS:
    res_nc_mensal[o] = {}
    for mp_str in _meses_res_todos:
        v = _residuos_kgm2_mes(o, pd.Period(mp_str, freq='M'))
        res_nc_mensal[o][mp_str] = (v is not None and v > meta_obra(o))
res_nc_por_obra = {o: sum(1 for excedeu in res_nc_mensal[o].values() if excedeu) for o in OBRAS}
res_nc_total = sum(res_nc_por_obra.values())

# ─── NC LAI/EIV — contagem por ASSUNTO ───────────────────────────────────────
# Fonte única: Dim_NC_LAI_EIV. Cada LINHA é uma auditoria de obra e pode apontar
# VÁRIOS assuntos, separados por "|", em duas colunas: `NC LAI` e `NC EIV`.
# A medida conta cada ASSUNTO individualmente (não a linha), somando LAI + EIV.
#   ex: "Armazenamento e separação de resíduos | Ruído | Placa da LAI" = 3 NCs
# "-" / vazio significam "nenhuma NC apontada" naquela coluna e não contam.
def _assuntos_nc(valor):
    if valor is None: return []
    s = str(valor).strip()
    if s.lower() in ('', '-', 'nan', 'none', '<na>'): return []
    return [a.strip() for a in s.split('|') if a.strip() and a.strip() != '-']

def _assuntos_linha(row):
    return _assuntos_nc(row.get('nc_lai')) + _assuntos_nc(row.get('nc_eiv'))

nc_por_obra = {o: 0 for o in OBRAS}
_nc_assunto_freq = {}
total_ncs = 0
# Lista granular (obra × mês × assunto) — permite o gráfico "Top não
# conformidades recorrentes" reagir aos filtros de obra e período no client.
nc_assuntos_lista = []
if not dim_nc.empty:
    for _, _r in dim_nc.iterrows():
        _assuntos = _assuntos_linha(_r)
        if not _assuntos: continue
        _ma = _r.get('mes_ano') if 'mes_ano' in _r else None
        _mes = _ma.strftime('%Y-%m') if isinstance(_ma, pd.Timestamp) and pd.notna(_ma) else ''
        for _a in _assuntos:
            nc_assuntos_lista.append({
                'obra': str(_r.get('obra', '') or ''),
                'mes': _mes,
                'assunto': _a,
                'fonte': 'LAI' if _a in _assuntos_nc(_r.get('nc_lai')) else 'EIV',
            })
        total_ncs += len(_assuntos)
        _o = _r.get('obra', '')
        if _o in nc_por_obra: nc_por_obra[_o] += len(_assuntos)
        for _a in _assuntos:
            _nc_assunto_freq[_a] = _nc_assunto_freq.get(_a, 0) + 1

# Top NCs recorrentes: mesmo critério (por assunto), pra bater com o card
top_ncs = [{'nc': a, 'count': c}
           for a, c in sorted(_nc_assunto_freq.items(), key=lambda kv: (-kv[1], kv[0]))[:10]]

# ─── Reclamações ──────────────────────────────────────────────────────────────
rec_por_obra = {o: int((dim_rec.get('obra', pd.Series()) == o).sum()) for o in OBRAS}
# total_rec é atribuído depois de montar reclamacoes_lista: Dim_reclamacoes_vizinhos
# tem linhas em branco (grade da planilha) e len(dim_rec) as contava como
# reclamação, fazendo o card divergir do gráfico.
rec_categoria = {}
if 'categoria' in dim_rec.columns:
    for k, v in dim_rec['categoria'].value_counts().head(5).items():
        if str(k) not in ('nan','NaN',''): rec_categoria[str(k)] = int(v)

# Lista detalhada de reclamações — alimenta o tooltip do gráfico
# "Reclamações de Vizinhos por Obra" (assunto/ocorrência, status/ações e canal).
# Mesmo padrão de emergencias_lista: descarta linhas totalmente vazias.
reclamacoes_lista = []
if not dim_rec.empty:
    _rc_ocorr  = next((c for c in dim_rec.columns if 'ocorr' in c.lower()), None)
    _rc_status = next((c for c in dim_rec.columns if 'status' in c.lower()), None)
    _rc_canal  = next((c for c in dim_rec.columns if 'canal' in c.lower()), None)
    _rc_cod    = next((c for c in dim_rec.columns if 'chamado' in c.lower()), None)

    def _rc_str(v):
        if v is None or (isinstance(v, float) and pd.isna(v)): return ''
        s = str(v).strip()
        return '' if s.lower() in ('nan', 'none', '<na>') else s

    for _, r in dim_rec.iterrows():
        ma = r.get('mes_ano') if 'mes_ano' in r else None
        item = {
            'mes_ano':    ma.strftime('%Y-%m') if isinstance(ma, pd.Timestamp) and pd.notna(ma) else '',
            'mes_label':  ma.strftime('%m/%Y') if isinstance(ma, pd.Timestamp) and pd.notna(ma) else '',
            'obra':       _rc_str(r.get('obra', '')),
            'ocorrencia': _rc_str(r.get(_rc_ocorr, '')) if _rc_ocorr else '',
            'status':     _rc_str(r.get(_rc_status, '')) if _rc_status else '',
            'categoria':  _rc_str(r.get('categoria', '')),
            'canal':      _rc_str(r.get(_rc_canal, '')) if _rc_canal else '',
            'cod':        _rc_str(r.get(_rc_cod, '')) if _rc_cod else '',
        }
        if not any([item['obra'], item['ocorrencia'], item['status'], item['canal']]):
            continue
        reclamacoes_lista.append(item)
    reclamacoes_lista.sort(key=lambda x: x['mes_ano'] or '', reverse=True)

# Card "Reclamações de Vizinhos" = reclamações REAIS (linhas preenchidas),
# a mesma base do gráfico "Reclamações de Vizinhos por Obra". Antes era
# len(dim_rec), que incluía as linhas vazias da planilha.
total_rec = len(reclamacoes_lista)

# ─── NC e Reclamações mensal por obra (pra filtro de período reagir na aba Obra Individual) ─
# Diferente do ruído, aqui 0 é uma contagem real (nenhuma NC/reclamação naquele mês),
# não "sem dado" — por isso todo mês do período entra na série, mesmo com count=0.
def _count_mes(df, o, mp):
    if 'obra' not in df.columns or 'mes_ano' not in df.columns: return None
    return int(len(df[(df['obra'] == o) & (df['mes_ano'].dt.to_period('M') == mp)]))

# NC: conta ASSUNTOS (LAI + EIV) no mês, mesmo critério do card — não linhas.
def _nc_assuntos_mes(o, mp):
    if 'obra' not in dim_nc.columns or 'mes_ano' not in dim_nc.columns: return None
    d = dim_nc[(dim_nc['obra'] == o) & (dim_nc['mes_ano'].dt.to_period('M') == mp)]
    return int(sum(len(_assuntos_linha(r)) for _, r in d.iterrows()))

nc_mensal = {}
rec_mensal = {}
for o in OBRAS:
    serie_nc, serie_rec = [], []
    for mp_str in conf_meses:
        mp = pd.Period(mp_str, freq='M')
        cnt_nc = _nc_assuntos_mes(o, mp)
        cnt_rec = _count_mes(dim_rec, o, mp)
        if cnt_nc is not None:
            serie_nc.append({'mes': mp_str, 'mes_label': _mes_label_ipa(mp_str), 'count': cnt_nc})
        if cnt_rec is not None:
            serie_rec.append({'mes': mp_str, 'mes_label': _mes_label_ipa(mp_str), 'count': cnt_rec})
    nc_mensal[o] = serie_nc
    rec_mensal[o] = serie_rec

# ─── Emergências ──────────────────────────────────────────────────────────────
emerg_por_obra = {o: int((dim_emerg.get('obra', pd.Series()) == o).sum()) for o in OBRAS}
total_emerg = len(dim_emerg)

# Lista detalhada para a tabela (Dim_emergencias_ambientais → tabela visual)
emergencias_lista = []
if not dim_emerg.empty:
    def _find_col(df, *cands):
        for c in df.columns:
            cl = c.lower().strip()
            for cand in cands:
                if cand.lower() in cl:
                    return c
        return None
    col_fonte  = _find_col(dim_emerg, 'fonte')
    col_fmc    = _find_col(dim_emerg, 'fmc')
    col_ocorr  = _find_col(dim_emerg, 'ocorerência', 'ocorrência', 'ocorrencia', 'ocoreren')
    col_status = _find_col(dim_emerg, 'status')
    col_categ  = _find_col(dim_emerg, 'categoria')

    def _str(v):
        if v is None or (isinstance(v, float) and pd.isna(v)): return ''
        s = str(v).strip()
        return '' if s.lower() in ('nan', 'none') else s

    def _fmc_str(v):
        if v is None or (isinstance(v, float) and pd.isna(v)): return ''
        try:
            if isinstance(v, (int, float)) and not pd.isna(v):
                return str(int(v))
        except: pass
        return _str(v)

    for _, r in dim_emerg.iterrows():
        ma = r.get('mes_ano') if 'mes_ano' in r else None
        mes_label = ma.strftime('%m/%Y') if isinstance(ma, pd.Timestamp) and pd.notna(ma) else ''
        item = {
            # mes_ano aqui tem DIA (a data exata da ocorrência importa e é o que
            # ordena a tabela). `mes` é o mesmo valor truncado no mês, para o
            # filtro de Período comparar com as demais fontes, que são mensais.
            'mes_ano':    ma.strftime('%Y-%m-%d') if isinstance(ma, pd.Timestamp) and pd.notna(ma) else '',
            'mes':        ma.strftime('%Y-%m') if isinstance(ma, pd.Timestamp) and pd.notna(ma) else '',
            'mes_label':  mes_label,
            'obra':       _str(r.get('obra', '')),
            'fonte':      _str(r.get(col_fonte, '')) if col_fonte else '',
            'fmc':        _fmc_str(r.get(col_fmc, '')) if col_fmc else '',
            'ocorrencia': _str(r.get(col_ocorr, '')) if col_ocorr else '',
            'status':     _str(r.get(col_status, '')) if col_status else '',
            'categoria':  _str(r.get(col_categ, '')) if col_categ else '',
        }
        # ignora linhas completamente vazias
        if not any([item['obra'], item['fonte'], item['fmc'], item['ocorrencia'], item['status'], item['categoria']]):
            continue
        emergencias_lista.append(item)

    emergencias_lista.sort(key=lambda x: x['mes_ano'] or '', reverse=True)

# ─── SGA ──────────────────────────────────────────────────────────────────────
sga_total = int(dim_sga['pessoas'].sum()) if 'pessoas' in dim_sga.columns else 0

# SGA — pessoas treinadas por ano (gráfico de barras)
sga_por_ano = []
if not dim_sga.empty and 'pessoas' in dim_sga.columns and 'Ano' in dim_sga.columns:
    df_sga = dim_sga.copy()
    df_sga['_ano'] = pd.to_numeric(df_sga['Ano'], errors='coerce')
    df_sga = df_sga.dropna(subset=['_ano', 'pessoas'])
    agg = df_sga.groupby('_ano')['pessoas'].sum().reset_index().sort_values('_ano')
    for _, r in agg.iterrows():
        sga_por_ano.append({'ano': int(r['_ano']), 'pessoas': int(r['pessoas'])})

# SGA granular (ano × mês) — permite o gráfico reagir ao filtro de Período.
# A coluna `Mês` traz o nome do mês em português para 2025/2026, mas os anos
# antigos (2021–2024) vêm consolidados com Mês = 'Todos', sem detalhe mensal.
# Esses ficam marcados com `anual: True` e só entram quando o recorte cobre o
# ano inteiro — recortar um total anual por mês seria inventar precisão.
_MES_PT_NUM = {
    'janeiro': 1, 'fevereiro': 2, 'marco': 3, 'março': 3, 'abril': 4,
    'maio': 5, 'junho': 6, 'julho': 7, 'agosto': 8, 'setembro': 9,
    'outubro': 10, 'novembro': 11, 'dezembro': 12,
}
sga_por_mes = []
if not dim_sga.empty and 'pessoas' in dim_sga.columns and 'Ano' in dim_sga.columns:
    _df = dim_sga.copy()
    _df['_ano'] = pd.to_numeric(_df['Ano'], errors='coerce')
    _df = _df.dropna(subset=['_ano', 'pessoas'])
    for _, r in _df.iterrows():
        _ano = int(r['_ano'])
        _mes_txt = str(r.get('Mês', '') or '').strip().lower()
        _num = _MES_PT_NUM.get(_mes_txt)
        sga_por_mes.append({
            'ano': _ano,
            'mes': f"{_ano}-{_num:02d}" if _num else None,
            'anual': _num is None,          # linha 'Todos' — ano fechado, sem mês
            'pessoas': int(r['pessoas']),
        })
    _n_anual = sum(1 for x in sga_por_mes if x['anual'])
    print(f"  SGA: {len(sga_por_mes)} linhas ({_n_anual} sem detalhe mensal)")

# ─── Impressões ───────────────────────────────────────────────────────────────
imp_total = int(fat_imp['total_pag'].sum()) if 'total_pag' in fat_imp.columns else 0
imp_custo = round(float(fat_imp['custo'].sum()),2) if 'custo' in fat_imp.columns else 0
imp_pct_pb = round(float(fat_imp['pct_pb'].mean()*100),1) if 'pct_pb' in fat_imp.columns else 0

# Histórico mensal de impressões
imp_historico = []
imp_media_6m = 0
if 'total_pag' in fat_imp.columns and 'mes_ano' in fat_imp.columns:
    df_ih = fat_imp.dropna(subset=['mes_ano', 'total_pag']).sort_values('mes_ano')
    MESES_PT = ['','Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez']
    for _, row in df_ih.iterrows():
        mes = row['mes_ano']
        mes_label = f"{MESES_PT[mes.month]}/{str(mes.year)[-2:]}" if pd.notna(mes) else '?'
        imp_historico.append({
            'mes': mes_label,
            'mes_ano': mes.strftime('%Y-%m'),
            'total': int(row['total_pag']),
            'custo': round(float(row['custo']), 2) if 'custo' in row and pd.notna(row.get('custo')) else 0,
            'pct_pb': round(float(row['pct_pb']) * 100, 1) if 'pct_pb' in row and pd.notna(row.get('pct_pb')) else 0,
        })
    # Média dos últimos 6 meses (meta = não ultrapassar essa média)
    if len(imp_historico) >= 6:
        imp_media_6m = int(round(np.mean([h['total'] for h in imp_historico[-6:]])))
    elif imp_historico:
        imp_media_6m = int(round(np.mean([h['total'] for h in imp_historico])))

# Top setores de impressão
imp_setores = []
imp_setor_mes = []
imp_setores_list = []
if not fat_imp_setor.empty and 'Nome_setor' in fat_imp_setor.columns and 'Páginas' in fat_imp_setor.columns:
    fat_imp_setor['paginas_num'] = pd.to_numeric(fat_imp_setor['Páginas'], errors='coerce')
    gs = fat_imp_setor.groupby('Nome_setor')['paginas_num'].sum().nlargest(10)
    total_s = gs.sum()
    for setor, qtde in gs.items():
        if str(setor) not in ('nan', 'NaN', ''):
            imp_setores.append({
                'setor': str(setor),
                'paginas': int(qtde),
                'pct': round(float(qtde / total_s * 100), 1) if total_s > 0 else 0,
            })

    # Dataset granular por setor × mês para filtro client-side
    df_sm = fat_imp_setor.copy()
    if 'Data' in df_sm.columns:
        df_sm['mes_ano_dt'] = pd.to_datetime(df_sm['Data'], errors='coerce')
        df_sm = df_sm.dropna(subset=['mes_ano_dt', 'paginas_num', 'Nome_setor'])
        df_sm['setor_clean'] = df_sm['Nome_setor'].astype(str).str.strip()
        df_sm = df_sm[(df_sm['setor_clean'] != '') & (df_sm['setor_clean'].str.lower() != 'nan')]
        # Agrega por (mes_ano, setor) para evitar múltiplas linhas
        df_agg = df_sm.groupby([df_sm['mes_ano_dt'].dt.to_period('M'), 'setor_clean'])['paginas_num'].sum().reset_index()
        df_agg.columns = ['mes_periodo', 'setor', 'paginas']
        for _, r in df_agg.iterrows():
            imp_setor_mes.append({
                'mes_ano': str(r['mes_periodo']),
                'setor': r['setor'],
                'paginas': int(r['paginas']),
            })
        imp_setores_list = sorted(set(s['setor'] for s in imp_setor_mes))

# ─── Avanço físico por obra ───────────────────────────────────────────────────
avanco_por_obra = {}
if 'obra' in fat_av.columns:
    for o in OBRAS:
        d = fat_av[fat_av['obra'] == o].sort_values('mes_ano') if 'mes_ano' in fat_av.columns else fat_av[fat_av['obra'] == o]
        if d.empty:
            avanco_por_obra[o] = {'avanco_pct': None, 'total_m2': None}
            continue
        last = d.iloc[-1]
        avanco_por_obra[o] = {
            'avanco_pct': round(float(last['avanco_pct'])*100,1) if 'avanco_pct' in d.columns and not pd.isna(last.get('avanco_pct')) else None,
            'total_m2': round(float(last['total_m2']),0) if 'total_m2' in d.columns and not pd.isna(last.get('total_m2')) else None,
        }

# ─── Conformidade evolução ────────────────────────────────────────────────────
conf_evolucao = []
if 'mes_ano' in fat_conf.columns and 'perc' in fat_conf.columns:
    g = fat_conf.dropna(subset=['mes_ano','perc']).groupby(fat_conf['mes_ano'].dt.to_period('M'))['perc'].mean()
    for p, v in g.tail(12).items():
        conf_evolucao.append({'mes': str(p), 'perc': round(float(v)*100,1)})

# ─── Barras por obra ──────────────────────────────────────────────────────────
conf_barras = []
for o in OBRAS:
    d = fat_conf[fat_conf['obra'] == o] if 'obra' in fat_conf.columns else pd.DataFrame()
    if not d.empty and 'perc' in d.columns:
        p = d['perc'].mean(); mt = d['meta'].replace(0, np.nan).mean() if 'meta' in d.columns else 0.9
        if not pd.isna(p):
            conf_barras.append({'obra': o, 'perc': round(float(p)*100,1), 'meta': round(float(mt if not pd.isna(mt) else 0.9)*100,1)})
conf_barras.sort(key=lambda x: x['perc'], reverse=True)

ruido_barras = [{'obra': o, 'db': v['media_db'], 'pct': v['pct_conforme']} for o, v in ruido_por_obra.items() if v['media_db'] is not None]
ruido_barras.sort(key=lambda x: x['db'] or 0, reverse=True)

res_barras = [{'obra': o, 'kg_m2': v['kg_m2'], 'meta': v['meta']} for o, v in res_por_obra.items() if v['kg_m2'] is not None]
res_barras.sort(key=lambda x: x['kg_m2'] or 0, reverse=True)

# ─── Mês de referência do filtro de Período ──────────────────────────────────
# "Últimos 6/12 meses" precisa significar A MESMA janela de calendário em todos
# os gráficos. Antes cada gráfico pegava os N últimos meses DA SUA PRÓPRIA
# fonte, e como cada tabela termina num mês diferente (conformidade em 2026-05,
# água em 2026-03, resíduos com linhas zeradas até 2026-07, auditorias com
# planejamento até 2026-10), a mesma seleção mostrava sete janelas diferentes
# na mesma tela.
#
# A âncora é o mês mais recente com dado REAL, o que exclui de propósito:
#   - meses de grade zerada (resíduos Mai–Jul/2026: linhas existem, valor é 0)
#   - meses de planejamento futuro (auditorias até Out/2026)
#   - qualquer mês posterior ao mês de geração do dashboard
def _ultimo_mes_real(pares):
    # pares: lista de (mes_str, valor) — só entra mês com valor de verdade
    ms = [m for m, v in pares if m and v not in (None, 0) and not pd.isna(v)]
    return max(ms) if ms else None

_TETO_MES = TODAY.strftime('%Y-%m')   # nunca aponta para o futuro
_candidatos = []

# Conformidade: mês com percentual lançado
for _o, _linha in conf_heatmap.items():
    _candidatos.append(_ultimo_mes_real([(m, c.get('perc')) for m, c in _linha.items()]))
# Resíduos: mês com kg lançado (descarta a grade zerada)
_candidatos.append(_ultimo_mes_real([(r['mes_ano'], r.get('kg')) for r in residuos_raw]))
# Ruído, água, energia, impressões, reclamações, emergências
_candidatos.append(_ultimo_mes_real([(r['mes_ano'], r.get('db')) for r in ruido_obra_mes]))
_candidatos.append(_ultimo_mes_real([(r['mes_ano'], r.get('m3')) for r in consumo_agua_raw]))
_candidatos.append(_ultimo_mes_real([(r['mes_ano'], r.get('kwh')) for r in consumo_energia_raw]))
_candidatos.append(_ultimo_mes_real([(r['mes_ano'], r.get('paginas')) for r in imp_setor_mes]))
_candidatos.append(_ultimo_mes_real([(r['mes_ano'], 1) for r in reclamacoes_lista]))
_candidatos.append(_ultimo_mes_real([(r.get('mes'), 1) for r in emergencias_lista]))

_candidatos = [m for m in _candidatos if m and m <= _TETO_MES]
PERIODO_REF = max(_candidatos) if _candidatos else _TETO_MES
print(f"  Mes de referencia do filtro de Periodo: {PERIODO_REF}")

# ─── Compilar JSON ────────────────────────────────────────────────────────────
DATA = {
    'obras': OBRAS,
    # Âncora do filtro de Período: "últimos 6/12 meses" conta a partir daqui,
    # igual para todos os gráficos (ver bloco acima).
    'periodo_ref': PERIODO_REF,
    'global': {
        'ipa': global_ipa, 'ipa_classe': ipa_class(global_ipa),
        'conf_media': round(np.mean([v['perc'] for v in conf_atual.values() if v['perc'] is not None]),1) if conf_atual else 0,
        'total_ncs': total_ncs,
        'media_kg_m2': round(float(media_kg_m2),2) if not pd.isna(media_kg_m2) else 0,
        'total_rec': total_rec, 'pa_no_prazo_pct': pa_no_prazo_pct,
        'total_emerg': total_emerg, 'sga_pessoas': sga_total,
        'total_auds': total_aud, 'aud_realizadas': int(aud_sim),
    },
    'ipa': ipa_data, 'ipa_mensal': ipa_mensal, 'ruido_mensal': ruido_mensal, 'conf_heatmap': conf_heatmap, 'conf_meses': conf_meses,
    'conf_atual': conf_atual, 'conf_barras': conf_barras, 'conf_evolucao': conf_evolucao,
    'obras_risco': obras_risco, 'top_ncs': top_ncs, 'nc_por_obra': nc_por_obra,
    'nc_assuntos_lista': nc_assuntos_lista,
    'nc_mensal': nc_mensal, 'rec_mensal': rec_mensal,
    # NC de resíduos (obra-mês acima da meta kg/m²) — calculado mas NÃO somado
    # ao card "NC registradas LAI/EIV"; exportado para uso futuro.
    'res_nc_por_obra': res_nc_por_obra, 'res_nc_total': res_nc_total,
    'res_por_obra': res_por_obra, 'res_classe': res_classe, 'res_tipo': res_tipo,
    'res_classe_por_obra': res_classe_por_obra, 'res_tipo_por_obra': res_tipo_por_obra,
    'res_barras': res_barras, 'total_res_kg': round(float(total_res_kg),0),
    'residuos_mensal': residuos_mensal, 'residuos_perigosos': residuos_perigosos,
    'residuos_raw': residuos_raw,
    'auditorias_meses': auditorias_meses,
    'auditorias_ocorrencias': auditorias_ocorrencias,
    'auditorias_obra_mes': auditorias_obra_mes,
    'sga_por_ano': sga_por_ano,
    'sga_por_mes': sga_por_mes,
    'meta_residuos': {
        'default': META_RESIDUOS_DEFAULT,
        'senna': META_RESIDUOS_SENNA,
        'por_obra': {o: meta_obra(o) for o in OBRAS},
    },
    'ruido_por_obra': ruido_por_obra, 'ruido_barras': ruido_barras,
    'ruido_obra_mes': ruido_obra_mes,
    'al_por_obra': al_por_obra, 'aud_por_obra': aud_por_obra,
    'pa_status': pa_status, 'pa_atrasadas': pa_atrasadas, 'pa_por_obra': pa_por_obra,
    'plano_acao_lista': plano_acao_lista,
    'rec_por_obra': rec_por_obra, 'rec_categoria': rec_categoria,
    'reclamacoes_lista': reclamacoes_lista,
    'emerg_por_obra': emerg_por_obra,
    'emergencias_lista': emergencias_lista,
    'imp_total': imp_total, 'imp_custo': imp_custo, 'imp_pct_pb': imp_pct_pb,
    'imp_historico': imp_historico, 'imp_media_6m': imp_media_6m, 'imp_setores': imp_setores,
    'imp_setor_mes': imp_setor_mes, 'imp_setores_list': imp_setores_list,
    'consumo_energia': consumo_energia,
    'consumo_energia_m2': consumo_energia_m2,
    'consumo_energia_raw': consumo_energia_raw,
    'consumo_agua': consumo_agua,
    'consumo_agua_m2': consumo_agua_m2,
    'consumo_agua_raw': consumo_agua_raw,
    # Setores que aparecem em água OU energia — alimenta o filtro de setores
    'consumo_setores_list': sorted({r['setor'] for r in consumo_agua_raw if r.get('setor')}
                                   | {r['setor'] for r in consumo_energia_raw if r.get('setor')}),
    'avanco_por_obra': avanco_por_obra,
}
DATA_JSON = json.dumps(DATA, cls=DateEncoder, ensure_ascii=False)
print(f"  IPA global: {global_ipa}")

# ─── Reusa template HTML do gerador Excel ─────────────────────────────────────
print("Lendo template HTML do gerador Excel...")
with open(TEMPLATE_PATH, 'r', encoding='utf-8') as f:
    src = f.read()

# Extrai o HTML template (entre HTML = r""" e a próxima """)
start = src.find('HTML = r"""') + len('HTML = r"""')
end = src.find('"""', start)
HTML = src[start:end]

DATA_DATE = TODAY.strftime("%d/%m/%Y")
HTML = HTML.replace('__DATA_JSON__', DATA_JSON).replace('__DATA_DATE__', DATA_DATE)

# ── Embed da logo ────────────────────────────────────────────────────────────
# Procura primeiro nos arquivos embutidos no .exe (BUNDLE_DIR), depois ao lado
# do .exe/script (SCRIPT_DIR). Se nada for encontrado, usa o ícone folha.
LOGO_FILES = ['logo-fg.png', 'logo-fg.jpg', 'logo-fg.svg', 'logo-fg.webp']
logo_html = '<i class="ti ti-leaf" style="color:#fff; font-size:20px"></i>'
_logo_found = None
for base_dir in [BUNDLE_DIR, SCRIPT_DIR]:
    for fname in LOGO_FILES:
        fpath = os.path.join(base_dir, fname)
        if os.path.exists(fpath):
            _logo_found = (fname, fpath)
            break
    if _logo_found:
        break
if _logo_found:
    fname, fpath = _logo_found
    ext = fname.split('.')[-1].lower()
    mime = {'png':'image/png','jpg':'image/jpeg','jpeg':'image/jpeg','svg':'image/svg+xml','webp':'image/webp'}.get(ext, 'image/png')
    with open(fpath, 'rb') as f:
        b64 = base64.b64encode(f.read()).decode('ascii')
    logo_html = f'<img src="data:{mime};base64,{b64}" alt="FG" style="height:35px; width:auto; display:block;">'
    print(f"  Logo embutida: {fname}")
else:
    print(f"  Logo nao encontrada (procurando logo-fg.png) - usando icone folha")

HTML = HTML.replace(
    '<div class="brand-logo"><i class="ti ti-leaf" style="color:#fff"></i></div>',
    f'<div class="brand-logo">{logo_html}</div>'
)

# Atualiza o subtítulo
# (Removido: substituição de "Indicadores Ambientais 2025" — texto agora vem direto do template)
HTML = HTML.replace('Dados lidos diretamente do Excel', 'Dados lidos do modelo Power BI (.pbix)')

print(f"Salvando em:\n  {OUTPUT_PATH}")
with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
    f.write(HTML)
size_kb = os.path.getsize(OUTPUT_PATH) // 1024
print(f"Dashboard gerado com sucesso! ({size_kb} KB)")
print(f"\nAbra no navegador:\n  {OUTPUT_PATH}")

# Quando rodando como .exe (PyInstaller), abre o dashboard automaticamente
# para o usuário não precisar dar mais nenhum clique.
if getattr(sys, 'frozen', False):
    print("\nAbrindo no navegador padrao...")
    try:
        os.startfile(OUTPUT_PATH)
    except Exception as e:
        print(f"  Nao foi possivel abrir automaticamente: {e}")
        print(f"  Abra manualmente: {OUTPUT_PATH}")
    # Pequena pausa para o usuário ver as mensagens antes do console fechar
    import time
    time.sleep(4)
