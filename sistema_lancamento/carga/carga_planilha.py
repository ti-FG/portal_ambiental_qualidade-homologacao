"""
carga_planilha.py — carga inicial do histórico a partir da planilha, 21/10/2026

Lê as tabelas nomeadas da "Controle Ambiental 2025_oficial_db.xlsx" (usando o
leitor do ETL do dashboard, C:\\Apps\\Dash_ambiental_qualidade\\etl) e escreve
db/carga_inicial.sql com os cadastros e os lançamentos no formato do sistema.

Não grava em banco nenhum: só gera o SQL. Quem aplica é o hml/memdb.mjs (na
homologação) ou a TI (no 152). Cada decisão de saneamento sai no relatório
(carga/relatorio_carga.txt) para a área conferir.

Uso:
    python carga_planilha.py [caminho_da_planilha]
"""
import datetime as dt
import json
import math
import os
import re
import sys
from collections import OrderedDict, defaultdict

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)
# No repositório o ETL e a planilha ficam na pasta acima de sistema_lancamento;
# DASH_DIR sobrescreve; o caminho antigo da homologação é o último recurso.
DASH = os.environ.get('DASH_DIR') or os.path.dirname(RAIZ)
if not os.path.isdir(os.path.join(DASH, 'etl')):
    DASH = r'C:\Apps\Dash_ambiental_qualidade'
ETL = os.path.join(DASH, 'etl')
PLANILHA = os.path.join(DASH, 'Controle Ambiental 2025_oficial_db.xlsx')
sys.path.insert(0, ETL)

from openpyxl import load_workbook  # noqa: E402
from etl_excel_para_db import ler_tabela_nomeada  # noqa: E402
from tabelas import NORMALIZACAO_OBRA  # noqa: E402

INICIO_HISTORICO = dt.date(2024, 1, 1)
QUEM = 'carga_planilha'
DUPLA_MAR25 = dt.date(2025, 3, 1)   # único mês em que consumo aceita 2 linhas da mesma unidade

relatorio = []


def nota(msg):
    relatorio.append(msg)


def vazio(v):
    # NaT (data vazia do pandas) também é vazio
    return v is None or type(v).__name__ == "NaTType" or (isinstance(v, float) and math.isnan(v)) or (isinstance(v, str) and not v.strip())


def num(v):
    if vazio(v):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def competencia(v):
    if vazio(v):
        return None
    if hasattr(v, 'date'):
        v = v.date()
    if isinstance(v, dt.date):
        return dt.date(v.year, v.month, 1)
    return None


def sql(v):
    if v is None:
        return 'NULL'
    if isinstance(v, bool):
        return 'true' if v else 'false'
    if isinstance(v, (int, float)):
        return repr(round(v, 6)) if isinstance(v, float) else str(v)
    if isinstance(v, (dt.date, dt.time)):
        return "'%s'" % v.isoformat()
    return "'" + str(v).replace("'", "''") + "'"


# ── obras ───────────────────────────────────────────────────────────────────
# Obras já conhecidas (nome que o sistema usa). Obra NOVA na planilha entra
# sozinha por descobrir_obras(), com o nome escrito no Avanço físico.
OBRAS = ['Blue Coast', 'Blue View', 'Boreal', 'Garden Park', 'Grand Place', 'Horizon',
         'Iconic', 'Imperium', 'North Tower', 'Sapphire', 'Scenarium', 'Senna', 'Titanium']


def chave_obra(texto):
    """Mesma obra, grafias diferentes: ignora maiúsculas, espaços repetidos e o
    sufixo "Tower" (NORTH = North Tower, CENTRAL = Central Tower, Senna Tower = Senna)."""
    return re.sub(r'\s+TOWER$', '', ' '.join(str(texto).split()).upper())


_POR_CHAVE = {}


def _indexar():
    _POR_CHAVE.clear()
    _POR_CHAVE.update({chave_obra(o): o for o in OBRAS})


_indexar()


def obra_oficial(texto):
    if vazio(texto):
        return None
    t = str(texto).strip()
    t = NORMALIZACAO_OBRA.get(t, t)
    return _POR_CHAVE.get(chave_obra(t))


def descobrir_obras(tabelas):
    """Acrescenta a OBRAS as obras da planilha que o sistema ainda não conhece.
    O nome vem do Avanço físico (escrito por extenso, ex.: "Central Tower")."""
    novas = []
    for nome, col in (('Avanco_fisico', 'Obra'), ('Metas_LAI_EIV', 'Obra')):
        for v in tabelas[nome][col]:
            if vazio(v) or obra_oficial(v):
                continue
            o = ' '.join(str(NORMALIZACAO_OBRA.get(str(v).strip(), v)).split())
            o = o.title() if o.isupper() else o
            OBRAS.append(o)
            novas.append(o)
            _indexar()
    return novas


LISTA_NC = []
CONGELADAS = {}


def congelar(ws):
    """Aba de consulta (relatório antigo): cada linha com os valores das células,
    até a última coluna usada. Datas viram dd/mm/aaaa; erros (#REF!) ficam como texto."""
    linhas = []
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        vals = list(row)
        while vals and vals[-1] is None:
            vals.pop()
        if not vals:
            continue
        conv = []
        for v in vals:
            if hasattr(v, 'strftime'):
                v = v.strftime('%d/%m/%Y')
            elif isinstance(v, float) and v.is_integer():
                v = int(v)
            elif v is not None and not isinstance(v, (int, float, str)):
                v = str(v)
            conv.append(v)
        linhas.append((i, conv))
    return linhas

# ── abas que eram listas (auditoria, plano, emergências, ...) ────────────────
ABAS_LISTA = ['Auditoria_Ambientais', 'Plano_acao', 'emergencias_ambientais', 'reclamacoes_vizinhos', 'Eco_urbana',
              'Acoes_SGA', 'Impressoes', 'Impressoes_por_setor', 'Impressoes_por_pessoa']
_MES_N = {'janeiro': 1, 'fevereiro': 2, 'marco': 3, 'março': 3, 'abril': 4, 'maio': 5, 'junho': 6, 'julho': 7,
          'agosto': 8, 'setembro': 9, 'outubro': 10, 'novembro': 11, 'dezembro': 12}


def mes_num(texto):
    return _MES_N.get(str(texto).strip().lower()) if not vazio(texto) else None


def texto(v):
    return None if vazio(v) else ' '.join(str(v).split())


def data(v):
    if vazio(v):
        return None
    if hasattr(v, 'date'):
        return v.date()
    return v if isinstance(v, dt.date) else None


def pct(v):
    """'17,19%' (texto) ou 0,1719 (número) -> 0,1719; 11,36 (pessoa) fica como está."""
    if vazio(v):
        return None
    if isinstance(v, str):
        s = v.strip().replace('%', '').replace('.', '').replace(',', '.')
        try:
            return float(s) / 100.0 if '%' in v else float(s)
        except ValueError:
            return None
    return num(v)


def carregar_abas_lista(T, out, obra_ou_nota, OID):
    Q = sql(QUEM)
    ins = lambda tab, cols, vals: out.append('INSERT INTO %s (%s, origem, criado_por, atualizado_por) VALUES (%s, \'planilha\', %s, %s);'
                                             % (tab, ', '.join(cols), ', '.join(vals), Q, Q))

    # Auditoria Ambientais — o mês vem de Mês/Ano (é o que o painel usa)
    n, contradiz = 0, []
    for i, r in T['Auditoria_Ambientais'].iterrows():
        o = obra_ou_nota('Auditoria_Ambientais', r['Obra '])
        d = data(r['Mês/Ano'])
        if not o or not d:
            continue
        c = dt.date(d.year, d.month, 1)
        if (num(r['Ano']) and int(num(r['Ano'])) != d.year) or (mes_num(r['Mês']) and mes_num(r['Mês']) != d.month):
            contradiz.append('%s %s/%s × Mês/Ano %s' % (o, texto(r['Mês']), int(num(r['Ano']) or 0), d.strftime('%d/%m/%Y')))
        real = texto(r['Realizado '])
        ins('auditoria_ambiental', ['obra_id', 'competencia', 'meta', 'realizado', 'ocorrencias'],
            [OID % sql(o), sql(c), sql(pct(r['Meta '])), sql(None if real is None else real.lower().startswith('s')),
             sql(texto(r['Principais ocorrencias da última auditória']))])
        n += 1
    nota('Auditoria Ambientais: %d auditorias.' % n + (' Ano/Mês CONTRADIZEM o Mês/Ano em: %s (valeu o Mês/Ano).' % '; '.join(contradiz) if contradiz else ''))

    # Plano de Ação
    n = 0
    for _, r in T['Plano_acao'].iterrows():
        o = obra_ou_nota('Plano_acao', r['Obra'])
        if not o:
            continue
        ins('plano_acao', ['obra_id', 'indicador', 'analise', 'acao', 'responsavel', 'prazo', 'status'],
            [OID % sql(o), sql(texto(r['Indicador '])), sql(texto(r['Análise '])), sql(texto(r['Ação'])),
             sql(texto(r['Responsável '])), sql(data(r['Prazo'])), sql(texto(r['Status']))])
        n += 1
    nota('Plano de Ação: %d ações.' % n)

    # Emergências Ambientais — "Emegência" (erro de digitação) vira "Emergência"
    n, vazias = 0, 0
    for _, r in T['emergencias_ambientais'].iterrows():
        o = obra_ou_nota('emergencias_ambientais', r['Obra '])
        d = data(r['Mês/Ano'])
        if not o or not d:
            continue
        if all(vazio(r[c]) for c in ('N° FMC', 'Ocorerência', 'Status/Ações')):
            vazias += 1
            continue
        cat = texto(r['Categoria '])
        cat = {'Emegência': 'Emergência'}.get(cat, cat)
        ins('emergencia', ['obra_id', 'data', 'fonte', 'n_fmc', 'ocorrencia', 'status_acoes', 'categoria'],
            [OID % sql(o), sql(d), sql(texto(r['Fonte '])), sql(int(num(r['N° FMC'])) if num(r['N° FMC']) else None),
             sql(texto(r['Ocorerência'])), sql(texto(r['Status/Ações'])), sql(cat)])
        n += 1
    nota('Emergências Ambientais: %d registros; %d linha(s) só com obra e data, sem ocorrência/FMC/status, ignoradas.' % (n, vazias))

    # Reclamações vizinhos — "Blue Coast – 01 - Ouvidoria" = obra + canal
    n = 0
    for _, r in T['reclamacoes_vizinhos'].iterrows():
        bruto = texto(r['Obra']) or ''
        partes = [p.strip() for p in re.split(r'\s*[-–]\s*', bruto) if p.strip()]
        canal = partes[-1] if len(partes) > 1 else None
        o = obra_ou_nota('reclamacoes_vizinhos', partes[0] if partes else None)
        d = data(r['Mês/Ano'])
        if not o or not d:
            continue
        cod = r['Cod. Chamado']
        ins('reclamacao', ['obra_id', 'canal', 'data', 'fonte', 'cod_chamado', 'ocorrencia', 'status_acoes', 'categoria'],
            [OID % sql(o), sql(canal), sql(d), sql(texto(r['Fonte'])), sql(None if vazio(cod) else str(int(cod)) if isinstance(cod, float) else str(cod)),
             sql(texto(r['Ocorrência'])), sql(texto(r['Status/Ações'])), sql(texto(r['Categoria ']))])
        n += 1
    nota('Reclamações vizinhos: %d chamados (o canal Ouvidoria/Relacionamento, que vinha no texto da Obra, virou campo próprio).' % n)

    # Eco urbana
    n = 0
    for _, r in T['Eco_urbana'].iterrows():
        o = obra_ou_nota('Eco_urbana', r['Obra'])
        c = competencia(r['Mês/Ano'])
        if not o or not c:
            continue
        ins('eco_urbana', ['obra_id', 'competencia', 'ponto_1', 'db', 'ponto2', 'db3'],
            [OID % sql(o), sql(c), sql(texto(r['Ponto_1'])), sql(num(r['db'])), sql(texto(r['Ponto2'])), sql(num(r['db3']))])
        n += 1
    nota('Eco urbana: %d medições.' % n)

    # Ações SGA — "Todos" (anos fechados) = mês 0
    n = 0
    for _, r in T['Acoes_SGA'].iterrows():
        ano = num(r['Ano'])
        m = 0 if str(r['Mês']).strip().lower() == 'todos' else mes_num(r['Mês'])
        if ano is None or m is None:
            continue
        ins('acoes_sga', ['ano', 'mes', 'pessoas'], [str(int(ano)), str(m), sql(int(num(r['Pessoas'])) if num(r['Pessoas']) is not None else None)])
        n += 1
    nota('Ações SGA: %d linhas (as contas de calculadora, ex. =297+2+10+32, entram pelo resultado).' % n)

    # Impressões — só o que era digitado; o resto o sistema calcula
    n = 0
    for _, r in T['Impressoes'].iterrows():
        c = competencia(r['Mês/Ano'])
        if not c and num(r['Ano']) and mes_num(r['Mês']):
            c = dt.date(int(num(r['Ano'])), mes_num(r['Mês']), 1)
        tot = num(r['Total de pág '])
        if not c or tot is None:
            continue
        cor = num(r['% Colorido'])
        if cor is None and num(r['% P&B']) is not None:
            cor = 1 - num(r['% P&B'])
        if cor is None and num(r['Pag Colorido']) is not None and tot:
            cor = num(r['Pag Colorido']) / tot
        dupla = num(r['Pág Dupla'])
        if dupla is None and num(r['Pág Simples']) is not None:
            dupla = 1 - num(r['Pág Simples'])
        inteiro = lambda v: None if num(v) is None else int(round(num(v)))
        ins('impressoes', ['competencia', 'total_pag', 'pct_colorido', 'pag_dupla', 'economia_pb', 'economia_colorido', 'simples_duplas', 'economia_folhas'],
            [sql(c), str(int(tot)), sql(cor), sql(dupla), sql(inteiro(r['Economia preto e branco'])), sql(inteiro(r['Economia Colorido'])),
             sql(num(r[' impressões simples poderiam ter sido feitas duplas'])), sql(inteiro(r['Economia de folhas']))])
        n += 1
    nota('Impressões: %d meses.' % n)

    # Impressões por setor / por pessoa (export do NDD Print)
    for tab, aba, cols in (('impressoes_setor', 'Impressoes_por_setor', [('nome', 'Nome')]),
                           ('impressoes_pessoa', 'Impressoes_por_pessoa', [('usuario', 'Usuário'), ('conta_atual', 'Conta atual')])):
        n, sem_data = 0, 0
        for _, r in T[aba].iterrows():
            c = competencia(r['Mês/Ano'])
            if not c:
                sem_data += 1
                continue
            inteiro = lambda v: None if num(v) is None else int(round(num(v)))
            ins(tab, ['competencia'] + [c_ for c_, _ in cols] + ['pb', 'cor', 'paginas', 'pct'],
                [sql(c)] + [sql(texto(r[h]) or '-') for _, h in cols] + [sql(inteiro(r['P&B'])), sql(inteiro(r['Cor'])), sql(inteiro(r['Páginas'])), sql(pct(r['%']))])
            n += 1
        nota('%s: %d linhas%s.' % (aba.replace('_', ' '), n, ', %d sem Mês/Ano ignoradas' % sem_data if sem_data else ''))


def lista_nc(wb):
    """Lista padronizada de tipos de NC ao lado da tabela NC_LAI_EIV (coluna
    com cabeçalho "Tipo de Não Conformidades..."), na ordem da planilha."""
    for ws in wb.worksheets:
        if 'NC_LAI_EIV' not in ws.tables:
            continue
        for col in ws.iter_cols(min_row=1, max_row=1):
            cab = col[0].value
            if isinstance(cab, str) and cab.strip().lower().startswith('tipo de não conformidade'):
                vals = []
                for (v,) in ws.iter_rows(min_row=2, min_col=col[0].column, max_col=col[0].column, values_only=True):
                    if vazio(v):
                        break
                    vals.append(str(v).strip())
                return vals
    return []


def main():
    caminho = sys.argv[1] if len(sys.argv) > 1 else PLANILHA
    wb = load_workbook(caminho, data_only=True)
    avisos = []
    T = {n: ler_tabela_nomeada(wb, n, avisos) for n in
         ['SOMA_RESIDUOS', 'Avanco_fisico', 'Conformidade_LAI_EIV', 'Metas_LAI_EIV',
          'Ruido', 'agua_luz', 'NC_LAI_EIV'] + ABAS_LISTA}
    global LISTA_NC
    LISTA_NC = lista_nc(wb)
    global CONGELADAS
    CONGELADAS = {ws.title.strip(): congelar(ws) for ws in wb.worksheets if ws.title.strip().lower().startswith("histórico de res")}
    wb.close()
    novas = descobrir_obras(T)
    if novas:
        nota('OBRA NOVA na planilha, cadastrada no sistema: %s (confirmar área total e meta no cadastro)' % ', '.join(novas))
    nota('Planilha: %s' % caminho)
    nota('Gerado em: %s' % dt.datetime.now().strftime('%d/%m/%Y %H:%M'))
    for a in avisos:
        nota('aviso do leitor: ' + a)

    sem_obra = defaultdict(set)

    def obra_ou_nota(tabela, texto):
        o = obra_oficial(texto)
        if o is None:
            sem_obra[tabela].add(str(texto))
        return o

    out = ['-- carga_inicial.sql — GERADO por carga/carga_planilha.py; não editar à mão',
           '-- %s' % dt.datetime.now().isoformat(timespec='seconds'),
           'SET search_path TO portal_ambiental_qualidade;', 'BEGIN;',
           # recarregável: zera cadastros e lançamentos, mantém usuários, sessões e auditoria de acesso
           'TRUNCATE obra, tipo_residuo, unidade_consumidora, meta_conformidade, assunto_nc, avanco_fisico, conformidade, '
           'conformidade_nc, medicao_ruido, residuo, consumo, log_alteracao, fechamento, auditoria_ambiental, plano_acao, emergencia, '
           'reclamacao, eco_urbana, acoes_sga, impressoes, impressoes_setor, impressoes_pessoa, aba_congelada RESTART IDENTITY CASCADE;']

    # ── cadastro de obras (área total vem do Avanço físico) ───────────────────
    av = T['Avanco_fisico']
    area_total = {}
    for _, r in av.iterrows():
        o = obra_ou_nota('Avanco_fisico', r['Obra'])
        if o and num(r['Total m²']):
            area_total[o] = num(r['Total m²'])
    meta_res = {'Senna': 50.0}  # demais: parâmetro padrão 36,6 (era o "Meta Kg" digitado)
    for o in OBRAS:
        out.append('INSERT INTO obra (nome, area_total_m2, meta_residuo_kg_m2, atualizado_por) VALUES (%s, %s, %s, %s);'
                   % (sql(o), sql(area_total.get(o)), sql(meta_res.get(o)), sql(QUEM)))
    sem_area = [o for o in OBRAS if o not in area_total]
    if sem_area:
        nota('Obras sem área total (não aparecem no Avanço físico): ' + ', '.join(sem_area))
    OID = "(SELECT id FROM obra WHERE nome = %s)"

    # ── tipos de resíduo ──────────────────────────────────────────────────────
    s = T['SOMA_RESIDUOS']
    tipos = OrderedDict()
    for _, r in s.iterrows():
        nome = r['Resíduos Detalhados']
        if vazio(nome):
            continue
        nome = str(nome).strip()
        dens = num(r['Densidade'])
        tipos.setdefault(nome, (str(r['Classe']).strip(), dens if dens else None))
    perigosos = {'EFLUENTE CONTAMINADO (LÍQUIDO)', 'RESÍDUO CONTAMINADO (SÓLIDO)'}
    ordem = 10
    for nome, (classe, dens) in sorted(tipos.items(), key=lambda kv: (kv[1][0], kv[0])):
        out.append('INSERT INTO tipo_residuo (nome, classe, densidade_kg_m3, perigoso, ordem, atualizado_por) '
                   'VALUES (%s, %s, %s, %s, %d, %s);' % (sql(nome), sql(classe), sql(dens), sql(nome in perigosos), ordem, sql(QUEM)))
        ordem += 10
    sem_dens = [n for n, (_, d) in tipos.items() if d is None]
    nota('Tipos de resíduo SEM densidade (na planilha davam Kg = 0; aqui o kg fica em branco até a área informar): '
         + '; '.join(sem_dens))
    for c in 'ABCD':
        out.append("INSERT INTO tipo_residuo (nome, classe, densidade_kg_m3, ativo, legado, ordem, atualizado_por) "
                   "VALUES (%s, '%s', NULL, false, true, %d, %s);" % (sql('Classe %s (sem detalhe)' % c), c, 900 + ord(c), sql(QUEM)))

    # ── lançamentos de resíduo ────────────────────────────────────────────────
    res = defaultdict(float)
    vazias = 0
    for _, r in s.iterrows():
        m3 = num(r['m³'])
        if m3 is None:
            vazias += 1
            continue
        o = obra_ou_nota('SOMA_RESIDUOS', r['Obra'])
        c = competencia(r['Mês/Ano'])
        if not o or not c:
            continue
        tipo = str(r['Resíduos Detalhados']).strip() if not vazio(r['Resíduos Detalhados']) else 'Classe %s (sem detalhe)' % str(r['Classe']).strip()
        res[(o, c, tipo)] += m3
    for (o, c, tipo), m3 in sorted(res.items()):
        out.append('INSERT INTO residuo (obra_id, competencia, tipo_residuo_id, volume_m3, origem, criado_por, atualizado_por) '
                   'VALUES (%s, %s, (SELECT id FROM tipo_residuo WHERE nome = %s), %s, \'planilha\', %s, %s);'
                   % (OID % sql(o), sql(c), sql(tipo), sql(m3), sql(QUEM), sql(QUEM)))
    legado = sum(1 for k in res if k[2].endswith('(sem detalhe)'))
    nota('Resíduos: %d lançamentos (obra × mês × tipo); %d linhas da planilha sem m³ (grade pré-criada) ignoradas; '
         '%d lançamentos até 2025 só por classe -> tipo "Classe X (sem detalhe)", sem densidade.' % (len(res), vazias, legado))

    # ── avanço físico ─────────────────────────────────────────────────────────
    avn = OrderedDict()
    for _, r in av.iterrows():
        o = obra_ou_nota('Avanco_fisico', r['Obra'])
        c = competencia(r['Mês/Ano'])
        p = num(r['Avanço %'])
        if o and c and p is not None:
            if (o, c) in avn:
                nota('Avanço físico duplicado %s %s: ficou o último (%.4f)' % (o, c, p))
            avn[(o, c)] = p
    for (o, c), p in avn.items():
        out.append("INSERT INTO avanco_fisico (obra_id, competencia, avanco_pct, origem, criado_por, atualizado_por) "
                   "VALUES (%s, %s, %s, 'planilha', %s, %s);" % (OID % sql(o), sql(c), sql(p), sql(QUEM), sql(QUEM)))
    nota('Avanço físico: %d lançamentos.' % len(avn))

    # ── metas de conformidade (formato largo -> vigências) ───────────────────
    meses = {'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4, 'mai': 5, 'jun': 6, 'jul': 7, 'ago': 8, 'set': 9, 'out': 10, 'nov': 11, 'dez': 12}
    n_metas = 0
    for _, r in T['Metas_LAI_EIV'].iterrows():
        o = obra_ou_nota('Metas_LAI_EIV', r['Obra'])
        if not o:
            continue
        anterior, primeira = None, True
        for col in T['Metas_LAI_EIV'].columns[1:]:
            m, a = col.split('/')
            v = num(r[col])
            # todo mês preenchido na planilha vira valor explícito (mesmo repetido):
            # a tela "Metas LAI EIV" mostra a grade igual à da planilha
            if v is None:
                continue
            mes_planilha = dt.date(2000 + int(a), meses[m], 1)
            # a primeira meta da obra também cobre o histórico anterior (era o PROCV em ago/25)
            for desde in ([INICIO_HISTORICO] if primeira and mes_planilha > INICIO_HISTORICO else []) + [mes_planilha]:
                out.append('INSERT INTO meta_conformidade (obra_id, vigente_desde, meta, atualizado_por) VALUES (%s, %s, %s, %s);'
                           % (OID % sql(o), sql(desde), sql(v), sql(QUEM)))
            n_metas += 1
            anterior, primeira = v, False
    nota('Metas de conformidade: %d vigências. O primeiro valor de cada obra vale desde %s (a planilha aplicava a meta de '
         'ago/25 a todo o histórico). Meses vazios mantêm a meta anterior — na planilha, ficavam sem meta ou com 0.'
         % (n_metas, INICIO_HISTORICO.strftime('%m/%Y')))
    sem_meta = [o for o in OBRAS if o not in {obra_oficial(x) for x in T['Metas_LAI_EIV']['Obra']}]
    if sem_meta:
        nota('Obras sem nenhuma meta de conformidade: ' + ', '.join(sem_meta))

    # ── conformidade ─────────────────────────────────────────────────────────
    conf = OrderedDict()
    for _, r in T['Conformidade_LAI_EIV'].iterrows():
        o = obra_ou_nota('Conformidade_LAI_EIV', r['OBRA'])
        c = competencia(r['Mês/Ano'])
        p = num(r['PERCENTUAL'])
        if o and c and p is not None:
            if p > 1:
                p = p / 100.0
            conf[(o, c)] = p
    for (o, c), p in conf.items():
        out.append("INSERT INTO conformidade (obra_id, competencia, percentual, origem, criado_por, atualizado_por) "
                   "VALUES (%s, %s, %s, 'planilha', %s, %s);" % (OID % sql(o), sql(c), sql(p), sql(QUEM), sql(QUEM)))
    nota('Conformidade: %d lançamentos (só o %%: a planilha não guardava itens conformes/avaliados).' % len(conf))

    # ── NCs ──────────────────────────────────────────────────────────────────
    # Dois formatos. Desde 21/10/2026 a planilha tem uma linha por NC (coluna
    # "NC LAI/EIV", sem separar LAI de EIV) e a lista padronizada de tipos ao
    # lado (coluna E): a lista vira o cadastro, na ordem da planilha. Antes,
    # eram as colunas "NC LAI" e "NC EIV" com vários assuntos separados por "|".
    assuntos = OrderedDict()
    ncs = []
    nc = T['NC_LAI_EIV']
    if 'NC LAI/EIV' in nc.columns:
        for a in LISTA_NC:
            assuntos[('LAI/EIV', a)] = True
        for _, r in nc.iterrows():
            o = obra_ou_nota('NC_LAI_EIV', r['Obra'])
            c = competencia(r['Mês/Ano'])
            a = None if vazio(r['NC LAI/EIV']) else str(r['NC LAI/EIV']).strip()
            if a:
                assuntos[('LAI/EIV', a)] = True
                if o and c:
                    ncs.append((o, c, 'LAI/EIV', a))
        fora = [a for (_, a) in assuntos if a not in LISTA_NC]
        if fora:
            nota('NC: assuntos usados que não estão na lista da coluna E (cadastrados mesmo assim): ' + '; '.join(fora))
    else:
        for _, r in nc.iterrows():
            o = obra_ou_nota('NC_LAI_EIV', r['Obra'])
            c = competencia(r['Mês/Ano'])
            for tipo, col in (('LAI', 'NC LAI'), ('EIV', 'NC EIV')):
                if vazio(r[col]):
                    continue
                for a in str(r[col]).split('|'):
                    a = a.strip()
                    if a and a != '-':
                        assuntos[(tipo, a)] = True
                        if o and c:
                            ncs.append((o, c, tipo, a))
    for i, (tipo, a) in enumerate(assuntos):
        out.append('INSERT INTO assunto_nc (tipo, nome, ordem) VALUES (%s, %s, %d) ON CONFLICT DO NOTHING;' % (sql(tipo), sql(a), (i + 1) * 10))
    for o, c, tipo, a in sorted(set(ncs)):
        out.append('INSERT INTO conformidade_nc (obra_id, competencia, assunto_id, criado_por) VALUES '
                   '(%s, %s, (SELECT id FROM assunto_nc WHERE tipo = %s AND nome = %s), %s) ON CONFLICT DO NOTHING;'
                   % (OID % sql(o), sql(c), sql(tipo), sql(a), sql(QUEM)))
    nota('NC LAI/EIV: %d assuntos cadastrados, %d apontamentos.' % (len(assuntos), len(set(ncs))))

    # ── ruído ────────────────────────────────────────────────────────────────
    n_ruido, zeradas = 0, 0
    for _, r in T['Ruido'].iterrows():
        o = obra_ou_nota('Ruido', r['Obra'])
        c = competencia(r['Mês/Ano'])
        d = num(r['dB'])
        if not d:
            zeradas += 1
            continue
        h = r['Horário']
        if isinstance(h, dt.datetime):
            h = None  # horário gravado como data na planilha
        elif isinstance(h, dt.time):
            h = h.replace(microsecond=0)
            if h == dt.time(0, 0):
                h = None
        else:
            h = None
        if o and c:
            out.append("INSERT INTO medicao_ruido (obra_id, competencia, ponto, horario, db, empresa, origem, criado_por, atualizado_por) "
                       "VALUES (%s, %s, %s, %s, %s, NULL, 'planilha', %s, %s);"
                       % (OID % sql(o), sql(c), sql(str(r['Ponto'] or 'P1').strip()), sql(h), sql(d), sql(QUEM), sql(QUEM)))
            n_ruido += 1
    nota('Ruído: %d medições; %d linhas com dB 0 ou vazio ignoradas (obra sem medição no mês).' % (n_ruido, zeradas))

    # ── unidades consumidoras e consumo ──────────────────────────────────────
    al = T['agua_luz']
    ucs = OrderedDict()
    for _, r in al.iterrows():
        if vazio(r['Obra']):
            continue
        nome = str(r['Obra']).strip()
        if nome not in ucs:
            ucs[nome] = obra_oficial(nome)
    out.append("INSERT INTO unidade_consumidora (nome, tipo, ordem, atualizado_por) VALUES ('Sede', 'administrativo', 500, %s);" % sql(QUEM))
    ordem = 10
    for nome, o in sorted(ucs.items(), key=lambda kv: (kv[1] is None, kv[0])):
        if nome == 'Sede':
            continue
        exibir = o if o else nome
        if o:
            out.append("INSERT INTO unidade_consumidora (nome, tipo, obra_id, ordem, atualizado_por) VALUES (%s, 'obra', %s, %d, %s);"
                       % (sql(exibir), OID % sql(o), ordem, sql(QUEM)))
        else:
            pai = "(SELECT id FROM unidade_consumidora WHERE nome = 'Sede')" if nome.startswith('Sala ') else 'NULL'
            out.append("INSERT INTO unidade_consumidora (nome, tipo, agrupa_em_id, ordem, atualizado_por) VALUES (%s, 'administrativo', %s, %d, %s);"
                       % (sql(exibir), pai, 500 + ordem if nome.startswith('Sala ') else 600 + ordem, sql(QUEM)))
        ordem += 10
    nome_uc = {n: (o or n) for n, o in ucs.items()}
    cons = OrderedDict()
    sem_setor = 0
    for _, r in al.iterrows():
        c = competencia(r['Mês/Ano'])
        luz, agua = num(r['Luz (kwh)']), num(r['Água (m³)'])
        if not c:
            continue
        if vazio(r['Obra']):
            if luz is not None or agua is not None:
                sem_setor += 1
            continue
        nome = str(r['Obra']).strip()
        k = (nome_uc.get(nome, nome), c)
        linha = 1
        if k + (1,) in cons:
            # 30/10/2026: a planilha tem 2 linhas da Assistência técnica em mar/25 e a área mandou guardar as duas
            # (exceção de histórico, também no CHECK de consumo); em outros meses vale 1 linha: fica a última
            if c == DUPLA_MAR25:
                linha = 2
                nota('Consumo com 2 linhas na planilha, guardadas as duas: %s %s' % k)
            else:
                nota('Consumo duplicado %s %s: ficou o último' % k)
        cons[k + (linha,)] = (luz, agua)
    # 30/10/2026: a Sede entra como está digitada na planilha (a tela mostra a soma das salas ao lado) e
    # linha em branco fica NULL, como na planilha
    for (u, c, linha), (luz, agua) in cons.items():
        out.append("INSERT INTO consumo (uc_id, competencia, linha, luz_kwh, agua_m3, origem, criado_por, atualizado_por) VALUES "
                   "((SELECT id FROM unidade_consumidora WHERE nome = %s), %s, %d, %s, %s, 'planilha', %s, %s);"
                   % (sql(u), sql(c), linha, sql(luz), sql(agua), sql(QUEM), sql(QUEM)))
    if sem_setor:
        nota('Água e luz: %d linha(s) com consumo e SEM obra/setor na coluna Obra foram ignoradas (conferir na planilha).' % sem_setor)
    nota('Água e luz: %d unidades consumidoras, %d lançamentos (inclui linhas em branco, com luz e água NULL, e a Sede '
         'como digitada na planilha).' % (len(ucs), len(cons)))

    # ── obra/unidade ativa: teve lançamento nos 3 últimos meses com dado ─────
    ultima = defaultdict(lambda: dt.date(1900, 1, 1))
    for (o, c, _t) in res:
        ultima[o] = max(ultima[o], c)
    for (o, c) in list(avn) + list(conf):
        ultima[o] = max(ultima[o], c)
    corte_mes = max(ultima.values())
    corte = dt.date(corte_mes.year - (1 if corte_mes.month <= 3 else 0), (corte_mes.month - 4) % 12 + 1, 1)
    inativas = sorted(o for o in OBRAS if ultima[o] < corte)
    for o in inativas:
        out.append('UPDATE obra SET ativa = false, observacao = %s WHERE nome = %s;'
                   % (sql('Sem lançamento desde %s (carga da planilha)' % ultima[o].strftime('%m/%Y')
                          if ultima[o].year > 1900 else 'Sem lançamento na planilha'), sql(o)))
    nota('Obras marcadas INATIVAS (sem lançamento desde antes de %s; reativar no cadastro se ainda estiverem em obra): %s'
         % (corte.strftime('%m/%Y'), ', '.join('%s (último %s)' % (o, ultima[o].strftime('%m/%Y')) for o in inativas) or 'nenhuma'))
    ult_uc = defaultdict(lambda: dt.date(1900, 1, 1))
    for (u, c, _), (luz, agua) in cons.items():
        if luz is not None or agua is not None:     # linha em branco não conta como "teve conta no mês"
            ult_uc[u] = max(ult_uc[u], c)
    corte_uc = max(ult_uc.values())
    corte_uc = dt.date(corte_uc.year - (1 if corte_uc.month <= 6 else 0), (corte_uc.month - 7) % 12 + 1, 1)
    ucs_inativas = sorted(u for u in nome_uc.values() if u != 'Sede' and (ult_uc[u] < corte_uc or u in inativas))
    for u in ucs_inativas:
        out.append('UPDATE unidade_consumidora SET ativa = false WHERE nome = %s;' % sql(u))
    nota('Unidades consumidoras INATIVAS (obra inativa ou sem conta há 6+ meses): %s' % (', '.join(ucs_inativas) or 'nenhuma'))

    carregar_abas_lista(T, out, obra_ou_nota, OID)
    for aba, linhas in CONGELADAS.items():
        for n_linha, vals in linhas:
            out.append("INSERT INTO aba_congelada (aba, linha, valores) VALUES (%s, %d, %s::jsonb);" % (sql(aba), n_linha, sql(json.dumps(vals, ensure_ascii=False))))
        nota("%s: %d linhas trazidas como consulta, do jeito que estão na planilha (relatório antigo, sem formato de tabela)." % (aba, len(linhas)))

    for tabela, textos in sem_obra.items():
        nota('%s: grafias de obra NÃO reconhecidas (linhas ignoradas): %s' % (tabela, ', '.join(sorted(textos))))

    out.append('COMMIT;')
    destino = os.path.join(RAIZ, 'db', 'carga_inicial.sql')
    with open(destino, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(out) + '\n')
    with open(os.path.join(AQUI, 'relatorio_carga.txt'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(relatorio) + '\n')
    print('\n'.join(relatorio))
    print('\n-> %s (%d comandos)' % (destino, len(out)))


if __name__ == '__main__':
    main()
