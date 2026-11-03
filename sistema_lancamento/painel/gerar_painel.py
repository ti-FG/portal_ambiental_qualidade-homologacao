"""
gerar_painel.py — gera o dashboard a partir do BANCO do sistema, 21/10/2026

Roda o gerador OFICIAL (C:\\Apps\\Dash_ambiental_qualidade\\gerar_dashboard_pbix.py)
sem alterar uma linha dele: o código é lido, o leitor do .pbix (PBIXRay) é
trocado por um que monta as mesmas tabelas do modelo a partir do banco, e a
saída vai para www/painel/ deste sistema. Mesmos gráficos, mesmos cálculos,
mesmo HTML do painel de hoje — só a origem dos dados muda.

Origem de cada tabela do modelo:
  - lançadas no sistema (banco): Dim_Obras, Fat_Conformidade_LAI_EIV,
    Fat_Metas_LAI_EIV, Dim_NC_LAI_EIV, Fat_Ruido, Fat_soma_residuos,
    Fat_Residuos_m³_m², Fat_avanco_fisico, Fat_consumo_agua_luz
  - listas (remontadas com os cabeçalhos da planilha e passadas pelo Power
    Query do ETL): auditorias, plano de ação, emergências, reclamações, SGA, impressões, Eco urbana

Uso:  python gerar_painel.py            (lê PG* do ambiente ou de www/.env)
Saída: www/painel/painel.html  + www/painel/painel.json (quando e de onde)
"""
import datetime as dt
import json
import os
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)
# Pasta do dashboard (gerar_dashboard_pbix.py, etl/, planilha). No repositório é a
# pasta acima de sistema_lancamento; DASH_DIR sobrescreve; o caminho antigo da
# homologação só vale se o gerador não estiver ali.
DASH = os.environ.get('DASH_DIR') or os.path.dirname(RAIZ)
if not os.path.exists(os.path.join(DASH, 'gerar_dashboard_pbix.py')):
    DASH = r'C:\Apps\Dash_ambiental_qualidade'
GERADOR = os.path.join(DASH, 'gerar_dashboard_pbix.py')
PLANILHA = os.path.join(DASH, 'Controle Ambiental 2025_oficial_db.xlsx')
SAIDA_DIR = os.path.join(RAIZ, 'www', 'painel')
SAIDA = os.path.join(SAIDA_DIR, 'painel.html')
sys.path.insert(0, os.path.join(DASH, 'etl'))

import warnings  # noqa: E402
import pandas as pd  # noqa: E402
warnings.filterwarnings("ignore", message="pandas only supports SQLAlchemy")
import psycopg2  # noqa: E402

MESES = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho',
         'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
DO_BANCO = ['Dim_Obras', 'Fat_Conformidade_LAI_EIV', 'Fat_Metas_LAI_EIV', 'Dim_NC_LAI_EIV',
            'Fat_Ruido', 'Fat_soma_residuos', 'Fat_Residuos_m³_m²', 'Fat_avanco_fisico',
            'Fat_consumo_agua_luz']


def ler_env():
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
    return env


def consultar(con, sql):
    return pd.read_sql_query(sql, con)


def mes_cols(df, col='competencia'):
    d = pd.to_datetime(df[col])
    df['Ano'] = d.dt.year.astype('Int64')
    df['Mês'] = d.dt.month.map(lambda m: MESES[m - 1])
    df['Data'] = d.dt.date
    return df


def tabelas_do_banco(env):
    S = env.get('APP_SCHEMA', 'portal_ambiental_qualidade')
    con = psycopg2.connect(host=env.get('PGHOST'), port=int(env.get('PGPORT', 5432)),
                           dbname=env.get('PGDATABASE'), user=env.get('PGUSER'),
                           password=env.get('PGPASSWORD') or None, connect_timeout=10)
    try:
        con.cursor().execute(f'SET search_path TO {S}')
        t = {}
        obras = consultar(con, 'SELECT nome AS "Obra" FROM obra ORDER BY nome')
        t['Dim_Obras'] = obras

        # avanço físico e área do mês (o denominador único)
        av = consultar(con, 'SELECT obra AS "Obra", competencia, avanco_pct AS "Avanço %", '
                            'area_mes_m2 AS "Avanço m²", area_total_m2 AS "Total m²" FROM v_area_mes')
        av = mes_cols(av)
        t['Fat_avanco_fisico'] = av[['Obra', 'Ano', 'Mês', 'Data', 'Avanço %', 'Avanço m²', 'Total m²']]

        # conformidade com a meta VIGENTE (META e META2 iguais: a vigência substitui as duas)
        cf = consultar(con, 'SELECT obra AS "Obra", competencia, percentual AS "PERCENTUAL", meta AS "META" FROM v_conformidade')
        cf = mes_cols(cf)
        cf['MÊS'] = cf['Mês']
        cf['META2'] = cf['META']
        t['Fat_Conformidade_LAI_EIV'] = cf[['Obra', 'Ano', 'MÊS', 'Data', 'PERCENTUAL', 'META', 'META2']]

        # metas no formato largo do modelo (ago/25..dez/26), meta vigente em cada mês
        meses_metas = pd.date_range('2025-08-01', '2026-12-01', freq='MS')
        mt = consultar(con, 'SELECT o.nome AS "Obra", m.vigente_desde, m.meta FROM meta_conformidade m JOIN obra o ON o.id = m.obra_id')
        largo = []
        for o in obras['Obra']:
            linha = {'Obra': o}
            vig = mt[mt['Obra'] == o].sort_values('vigente_desde')
            for d in meses_metas:
                v = vig[pd.to_datetime(vig['vigente_desde']) <= d]['meta']
                linha['%s/%s' % (['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'][d.month - 1], str(d.year)[2:])] = v.iloc[-1] if len(v) else None
            largo.append(linha)
        t['Fat_Metas_LAI_EIV'] = pd.DataFrame(largo)

        # NCs: "assunto | assunto" por tipo, "-" quando não há
        nc = consultar(con, '''SELECT o.nome AS "Obra", c.competencia AS "Data",
                coalesce(string_agg(a.nome, ' | ' ORDER BY a.nome) FILTER (WHERE a.tipo IN ('LAI', 'LAI/EIV')), '-') AS "NC LAI",
                coalesce(string_agg(a.nome, ' | ' ORDER BY a.nome) FILTER (WHERE a.tipo = 'EIV'), '-') AS "NC EIV"
              FROM conformidade_nc c JOIN obra o ON o.id = c.obra_id JOIN assunto_nc a ON a.id = c.assunto_id
             GROUP BY o.nome, c.competencia''')
        nc['Data'] = pd.to_datetime(nc['Data']).dt.date
        t['Dim_NC_LAI_EIV'] = nc[['Obra', 'NC LAI', 'NC EIV', 'Data']]

        ru = consultar(con, '''SELECT o.nome AS "Obra", r.competencia, r.ponto AS "Ponto", r.horario AS "Horário", r.db AS "dB"
              FROM medicao_ruido r JOIN obra o ON o.id = r.obra_id''')
        ru = mes_cols(ru)
        t['Fat_Ruido'] = ru[['Ano', 'Mês', 'Data', 'Obra', 'Ponto', 'Horário', 'dB']]

        # resíduos: uma linha por obra × mês × tipo; legado (só classe) sem tipo
        rs = consultar(con, 'SELECT obra AS "Obra", competencia, classe AS "Classe", tipo, legado, volume_m3 AS "m³", '
                            'densidade_kg_m3 AS "Densidade", kg AS "Kg", area_mes_m2 AS "m²", kg_m2 AS "Kg/m²", meta_kg_m2 AS "Meta Kg" FROM v_residuo')
        rs = mes_cols(rs)
        rs['Resíduos Detalhados'] = rs['tipo'].where(~rs['legado'].astype(bool), None)
        t['Fat_soma_residuos'] = rs[['Ano', 'Mês', 'Data', 'Obra', 'Classe', 'm³', 'Resíduos Detalhados', 'Kg',
                                     'Densidade', 'm²', 'Kg/m²', 'Meta Kg']]

        rm = consultar(con, '''SELECT r.obra AS "Obra", r.competencia, sum(r.volume_m3) AS "Resíduos m³", max(r.area_mes_m2) AS "Avanço Físico m²"
              FROM v_residuo r GROUP BY r.obra, r.competencia''')
        rm = mes_cols(rm)
        rm['Resíduos m³/m²'] = rm['Resíduos m³'] / rm['Avanço Físico m²'].where(rm['Avanço Físico m²'] > 0)
        t['Fat_Residuos_m³_m²'] = rm[['Ano', 'Mês', 'Data', 'Obra', 'Resíduos m³', 'Avanço Físico m²', 'Resíduos m³/m²']]

        # água e luz. A Sede não entra: é a soma das salas, que já entram
        # (na planilha ela entrava junto e o administrativo contava em dobro).
        # O gerador lê a intensidade de energia na coluna "Luz m²" e a de água
        # em "Água m³/m²": as duas recebem consumo ÷ área do mês da obra.
        al = consultar(con, '''SELECT v.unidade, v.tipo, v.obra, v.competencia, v.luz_kwh, v.agua_m3, v.area_mes_m2
              FROM v_consumo v JOIN unidade_consumidora u ON u.id = v.uc_id
             WHERE NOT EXISTS (SELECT 1 FROM unidade_consumidora f WHERE f.agrupa_em_id = u.id)''')
        al = mes_cols(al)
        eh_obra = al['tipo'] == 'obra'
        al['nome_setor'] = al['unidade'].where(~eh_obra, al['obra'].str.upper())
        al['Obra'] = al['obra'].where(eh_obra, 'Outros')
        area = al['area_mes_m2'].where(eh_obra & (al['area_mes_m2'] > 0))
        al['Luz (kwh)'] = al['luz_kwh']
        al['Água (m³)'] = al['agua_m3']
        al['Luz m²'] = al['luz_kwh'] / area
        al['Luz Kwh m³/m²'] = al['Luz m²']
        al['Água m²'] = area
        al['Água m³/m²'] = al['agua_m3'] / area
        t['Fat_consumo_agua_luz'] = al[['Ano', 'Mês', 'Data', 'nome_setor', 'Luz (kwh)', 'Água (m³)', 'Obra',
                                        'Luz m²', 'Luz Kwh m³/m²', 'Água m²', 'Água m³/m²']]
        t.update(abas_lista_do_banco(con, list(obras['Obra'])))
        ultima = consultar(con, "SELECT max(atualizado_em) AT TIME ZONE 'America/Sao_Paulo' AS u FROM (" + ' UNION ALL '.join(
            'SELECT atualizado_em FROM %s' % x for x in ('residuo', 'conformidade', 'consumo', 'avanco_fisico', 'medicao_ruido', 'auditoria_ambiental',
                                                         'plano_acao', 'emergencia', 'reclamacao', 'eco_urbana', 'acoes_sga', 'impressoes',
                                                         'impressoes_setor', 'impressoes_pessoa')) + ') x')['u'].iloc[0]
        return t, ultima
    finally:
        con.close()


def abas_lista_do_banco(con, obras):
    """Auditoria, Plano de Ação, Emergências, Reclamações, Eco urbana, Ações SGA e
    Impressões: o banco guarda as colunas da aba; aqui a aba é remontada com os
    cabeçalhos ORIGINAIS da planilha e passa pela mesma transformação do Power
    Query (tabelas.py do ETL). Assim o modelo sai igual ao do pbix."""
    from tabelas import TABELAS
    from modelo_semantico import aplicar_tipos
    q = lambda sql: consultar(con, sql)
    nome_mes = lambda d: MESES[d.month - 1]
    precos = {r['chave']: float(r['valor']) for r in q("SELECT chave, valor FROM parametro WHERE chave LIKE 'preco_pagina_%'").to_dict('records')}
    p_pb, p_cor = precos.get('preco_pagina_pb', 0.05), precos.get('preco_pagina_cor', 0.30)
    brutas = {}

    a = q('SELECT o.nome, x.competencia, x.meta, x.realizado, x.ocorrencias FROM auditoria_ambiental x JOIN obra o ON o.id = x.obra_id ORDER BY x.id')
    brutas['Auditoria_Ambientais'] = pd.DataFrame({
        'Obra ': a['nome'], 'Ano': [d.year for d in a['competencia']], 'Mês': [nome_mes(d).upper() for d in a['competencia']],
        'Mês/Ano': a['competencia'], 'Meta ': a['meta'], 'Realizado ': a['realizado'].map(lambda v: None if v is None or pd.isna(v) else ('Sim' if v else 'Não')),
        'Principais ocorrencias da última auditória': a['ocorrencias']})

    p = q('SELECT o.nome, x.indicador, x.analise, x.acao, x.responsavel, x.prazo, x.status FROM plano_acao x JOIN obra o ON o.id = x.obra_id ORDER BY x.id')
    brutas['Plano_acao'] = pd.DataFrame({'Obra': p['nome'], 'Indicador ': p['indicador'], 'Análise ': p['analise'], 'Ação': p['acao'],
                                         'Responsável ': p['responsavel'], 'Prazo': p['prazo'], 'Status': p['status']})

    e = q('SELECT o.nome, x.data, x.fonte, x.n_fmc, x.ocorrencia, x.status_acoes, x.categoria FROM emergencia x JOIN obra o ON o.id = x.obra_id ORDER BY x.data, x.id')
    brutas['emergencias_ambientais'] = pd.DataFrame({'Obra ': e['nome'], 'Mês/Ano': e['data'], 'Fonte ': e['fonte'], 'N° FMC': e['n_fmc'],
                                                     'Ocorerência': e['ocorrencia'], 'Status/Ações': e['status_acoes'], 'Categoria ': e['categoria']})

    r = q('SELECT o.nome, x.canal, x.data, x.fonte, x.cod_chamado, x.ocorrencia, x.status_acoes, x.categoria FROM reclamacao x JOIN obra o ON o.id = x.obra_id ORDER BY x.data, x.id')
    brutas['reclamacoes_vizinhos'] = pd.DataFrame({
        'Mês/Ano': r['data'], 'Obra': [n if not c else '%s - %s' % (n, c) for n, c in zip(r['nome'], r['canal'])],  # o Power Query separa pelo "-"
        'Fonte': r['fonte'], 'Cod. Chamado': r['cod_chamado'], 'Ocorrência': r['ocorrencia'], 'Status/Ações': r['status_acoes'], 'Categoria ': r['categoria']})

    u = q('SELECT o.nome, x.competencia, x.ponto_1, x.db, x.ponto2, x.db3 FROM eco_urbana x JOIN obra o ON o.id = x.obra_id ORDER BY x.id')
    brutas['Eco_urbana'] = pd.DataFrame({'Ano': [d.year for d in u['competencia']], 'Mês': [nome_mes(d) for d in u['competencia']], 'Mês/Ano': u['competencia'],
                                         'Obra': u['nome'], 'Ponto_1': u['ponto_1'], 'db': u['db'], 'Ponto2': u['ponto2'], 'db3': u['db3']})

    s = q('SELECT ano, mes, pessoas FROM acoes_sga ORDER BY ano, mes')
    brutas['Acoes_SGA'] = pd.DataFrame({'Ano': s['ano'], 'Mês': ['Todos' if m == 0 else MESES[int(m) - 1] for m in s['mes']], 'Pessoas': s['pessoas']})

    i = q('SELECT * FROM impressoes ORDER BY competencia')
    tot, cor, dup = i['total_pag'].astype(float), i['pct_colorido'].astype(float), i['pag_dupla'].astype(float)
    brutas['Impressoes'] = pd.DataFrame({
        'Ano': [d.year for d in i['competencia']], 'Mês': [nome_mes(d).lower() for d in i['competencia']], 'Mês/Ano': i['competencia'],
        # 2 casas, como na planilha (o gerador lê '1.234' com 3 casas como milhar)
        'Total de pág ': i['total_pag'], 'Custos': (tot * ((1 - cor) * p_pb + cor * p_cor)).round(2), '% P&B': 1 - cor, 'Pag P&B': tot * (1 - cor),
        '% Colorido': cor, 'Pag Colorido': tot * cor, 'Pág Simples': 1 - dup, 'Pág Dupla': dup,
        'Economia preto e branco': i['economia_pb'], 'Economia Colorido': i['economia_colorido'],
        'Economia Total': i['economia_pb'].astype(float).fillna(0) + i['economia_colorido'].astype(float).fillna(0),
        ' impressões simples poderiam ter sido feitas duplas': i['simples_duplas'], 'Economia de folhas': i['economia_folhas'], 'Economia mínima': (tot * 0.2).round(2)})

    st = q('SELECT competencia, nome, pb, cor, paginas, pct FROM impressoes_setor ORDER BY competencia, id')
    brutas['Impressoes_por_setor'] = pd.DataFrame({'Mês/Ano': st['competencia'], 'Nome': st['nome'], 'P&B': st['pb'], 'Cor': st['cor'], 'Páginas': st['paginas'], '%': st['pct']})
    ps = q('SELECT competencia, usuario, conta_atual, pb, cor, paginas, pct FROM impressoes_pessoa ORDER BY competencia, id')
    brutas['Impressoes_por_pessoa'] = pd.DataFrame({'Mês/Ano': ps['competencia'], 'Usuário': ps['usuario'], 'Conta atual': ps['conta_atual'],
                                                    'P&B': ps['pb'], 'Cor': ps['cor'], 'Páginas': ps['paginas'], '%': ps['pct']})

    saida = {}
    for t in TABELAS:
        if t.fonte in brutas:
            df = aplicar_tipos(t.construir(brutas[t.fonte].astype(object).where(pd.notna(brutas[t.fonte]), None)), t.colunas)
            saida[t.nome] = casar_obras(df, obras)
        elif t.fonte is None:  # Dim_Calendario: gerada
            saida[t.nome] = aplicar_tipos(t.construir(pd.DataFrame()), t.colunas)
    return saida


def casar_obras(df, obras):
    """O Power Query renomeia "Senna" para "Senna Tower" (e similares); o cadastro
    tem um nome só por obra. Sem casar, a obra perde auditorias e plano de ação no IPA."""
    por_chave = {o.upper(): o for o in obras}
    por_chave.update({'SENNA TOWER': por_chave.get('SENNA', 'Senna'), 'GARDEN': por_chave.get('GARDEN PARK', 'Garden Park')})
    df = df.copy()
    for col in df.columns:
        if col.strip() == 'Obra':
            df[col] = df[col].map(lambda v: por_chave.get(str(v).strip().upper(), v) if isinstance(v, str) else v)
    return df


def tabelas_da_planilha(obras):
    """O que ainda não tem tela no sistema: pelo ETL, como o pbix faria.

    O Power Query renomeia "Senna" para "Senna Tower" em algumas tabelas; o
    cadastro do sistema tem um nome só por obra. Sem casar os nomes, as
    auditorias e o plano de ação da obra não contam no IPA dela.
    """
    import etl_excel_para_db as etl
    avisos = []
    dados = etl.extrair(PLANILHA, avisos)
    etl.sanear_obras(dados, avisos)
    por_chave = {o.upper(): o for o in obras}
    por_chave.update({'SENNA TOWER': por_chave.get('SENNA', 'Senna')})
    saida = {}
    for nome, df in dados.items():
        if nome in DO_BANCO:
            continue
        df = df.copy()
        for col in df.columns:
            if col.strip() == 'Obra':
                df[col] = df[col].map(lambda v: por_chave.get(str(v).strip().upper(), v) if isinstance(v, str) else v)
        saida[nome] = df
    return saida


class ModeloDoBanco:
    """Faz o papel do PBIXRay para o gerador: get_table(nome) -> DataFrame."""
    def __init__(self, tabelas):
        self.tabelas = tabelas

    def __call__(self, _caminho):  # o gerador faz PBIXRay(PBIX_PATH)
        return self

    def get_table(self, nome):
        if nome not in self.tabelas:
            raise KeyError(nome)
        return self.tabelas[nome]


def remendar(src, antigo, novo, obrigatorio=True):
    n = src.count(antigo)
    if obrigatorio and n != 1:
        raise RuntimeError('o gerador mudou: trecho esperado %d vez(es), achado %d: %r' % (1, n, antigo[:60]))
    return src.replace(antigo, novo)


def aplicar_tema(arquivo):
    """Injeta painel/tema_juridico.css no fim do <head> do HTML gerado (layout do Sistema Jurídico).
    Sem o arquivo de tema, ou se o HTML não tiver </head>, o painel sai como o gerador fez."""
    css = os.path.join(AQUI, 'tema_juridico.css')
    if not os.path.exists(css):
        return
    with open(css, encoding='utf-8') as f:
        tema = f.read()
    with open(arquivo, encoding='utf-8') as f:
        html = f.read()
    if '</head>' not in html or 'id="tema-juridico"' in html:
        return
    # a nota do filtro passa a dizer quando o painel foi atualizado (a data de geração saiu da tela do sistema)
    import datetime
    html = html.replace('Dados lidos do sistema de Lançamentos Ambientais',
                        'Sistema de Lançamentos Ambientais · atualizado em ' + datetime.datetime.now().strftime('%d/%m/%Y %H:%M'), 1)
    marca = '<script>if (window.parent !== window || /embutido/.test(location.search)) document.documentElement.classList.add("embutido");</script>\n'
    html = html.replace('</head>', '<style id="tema-juridico">\n' + tema + '</style>\n' + marca + '</head>', 1)
    with open(arquivo, 'w', encoding='utf-8') as f:
        f.write(html)


def main():
    t0 = time.time()
    env = ler_env()
    # --referencia ARQ: tudo da planilha (ETL), para comparar com o painel do banco
    saida = SAIDA
    if len(sys.argv) > 2 and sys.argv[1] == '--referencia':
        import etl_excel_para_db as etl
        avisos = []
        tabelas = etl.extrair(PLANILHA, avisos)
        etl.sanear_obras(tabelas, avisos)
        etl.corrigir_intensidade_agua(tabelas, avisos)
        banco, ultima, planilha, saida = {}, None, tabelas, sys.argv[2]
    else:
        banco, ultima = tabelas_do_banco(env)
        print('banco: %s' % ', '.join('%s (%d)' % (k, len(v)) for k, v in banco.items()))
        planilha = {}  # desde 21/10/2026 nenhuma tabela vem mais da planilha
        tabelas = banco

    with open(GERADOR, encoding='utf-8') as f:
        src = f.read()
    src = remendar(src, 'from pbixray import PBIXRay', 'PBIXRay = __MODELO__')
    src = remendar(src, 'OUTPUT_PATH   = os.path.join(SCRIPT_DIR, "dashboard_fg.html")', 'OUTPUT_PATH   = __SAIDA__')
    src = remendar(src, "HTML.replace('Dados lidos diretamente do Excel', 'Dados lidos do modelo Power BI (.pbix)')",
                   "HTML.replace('Dados lidos diretamente do Excel', 'Dados lidos do sistema de Lançamentos Ambientais')", obrigatorio=False)

    os.makedirs(os.path.dirname(saida), exist_ok=True)
    tmp = saida + '.tmp'
    g = {'__name__': '__gerador__', '__file__': GERADOR, '__MODELO__': ModeloDoBanco(tabelas), '__SAIDA__': tmp}
    exec(compile(src, GERADOR, 'exec'), g)
    aplicar_tema(tmp)
    os.replace(tmp, saida)  # troca atômica: quem está vendo o painel nunca pega arquivo pela metade
    if saida != SAIDA:
        print('referência gerada -> %s' % saida)
        return

    info = {'gerado_em': dt.datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
            'ultimo_lancamento': ultima.strftime('%d/%m/%Y %H:%M') if ultima is not None and not pd.isna(ultima) else None,
            'segundos': round(time.time() - t0, 1),
            'do_banco': sorted(tabelas), 'da_planilha': sorted(planilha)}
    with open(os.path.join(SAIDA_DIR, 'painel.json'), 'w', encoding='utf-8') as f:
        json.dump(info, f, ensure_ascii=False)
    print('painel gerado em %.1fs -> %s' % (info['segundos'], SAIDA))


if __name__ == '__main__':
    main()
