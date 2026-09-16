# -*- coding: utf-8 -*-
"""
As 19 tabelas do modelo semantico, uma a uma.

Cada entrada descreve:
  nome     - nome no modelo semantico (identico ao do .pbix)
  tabela   - nome fisico no banco (snake_case ASCII)
  fonte    - tabela nomeada (ListObject) do Excel; None se a tabela e gerada
  colunas  - [(nome no modelo, tipo logico)] na MESMA ORDEM do .pbix
  chave    - colunas que identificam a linha no negocio (viram indice, nao PK:
             a fonte Excel nao garante unicidade)
  construir- funcao(df_bruto_do_excel) -> df com as colunas do modelo, antes
             da coercao de tipo (feita depois por aplicar_tipos)
  m        - resumo dos passos do Power Query que estao sendo reproduzidos

O codigo M completo de referencia esta no proprio .pbix e pode ser relido com:
    from pbixray import PBIXRay; PBIXRay(<pbix>).power_query
"""
from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass, field
from typing import Callable, Optional

import pandas as pd

from modelo_semantico import (
    DATE, DATETIME, DECIMAL, FLOAT, INT, PERCENT, TEXT, TIME,
    DIAS_PT, MESES_PT, OBRAS_VALIDAS_MAIUSCULO, OBRAS_VALIDAS_TITULO,
    classificar_obra, duplicar, nome_fisico, renomear, substituir_texto,
)


@dataclass
class Tabela:
    nome: str
    fonte: Optional[str]
    colunas: list
    construir: Callable
    chave: list = field(default_factory=list)
    m: str = ''
    descricao: str = ''

    @property
    def tabela(self):
        return nome_fisico(self.nome)

    @property
    def colunas_fisicas(self):
        return [nome_fisico(c) for c, _ in self.colunas]


# Substituicoes de nome de obra que o pbix faz em varias tabelas.
# Sao SUBSTRING (Replacer.ReplaceText) e a ordem importa.
_SUBST_CONFORMIDADE = [
    ('BLUE COAST', 'Blue Coast'), ('BOREAL', 'Boreal'),
    ('GRAND PLACE', 'Grand Place'), ('HORIZON', 'Horizon'),
    ('ICONIC', 'Iconic'), ('IMPERIUM', 'Imperium'),
    ('SAPPHIRE', 'Sapphire'), ('SCENARIUM', 'Scenarium'),
    ('TITANIUM', 'Titanium'), ('BLUE VIEW', 'Blue View'),
    ('GARDEN PARK', 'Garden Park'), ('SENNA', 'Senna Tower'),
]
_SUBST_AGUA_LUZ = [
    ('ICONIC', 'Iconic'), ('IMPERIUM', 'Imperium'), ('SAPPHIRE', 'Sapphire'),
    ('SCENARIUM', 'Scenarium'), ('TITANIUM', 'Titanium'),
    ('BLUE VIEW', 'Blue View'), ('BLUE COAST', 'Blue Coast'),
    ('GRAND PLACE', 'Grand Place'), ('HORIZON', 'Horizon'), ('BOREAL', 'Boreal'),
]
_SUBST_AUDITORIA = [
    ('BLUE COAST', 'Blue Coast'), ('BLUE VIEW', 'Blue View'),
    ('GARDEN PARK', 'Garden Park'), ('GRAND PLACE', 'Grand Place'),
    ('HORIZON', 'Horizon'), ('ICONIC', 'Iconic'), ('IMPERIUM', 'Imperium'),
    ('SAPPHIRE', 'Sapphire'), ('SCENARIUM', 'Scenarium'),
    ('SENNA', 'Senna Tower'), ('TITANIUM', 'Titanium'),
]
_SUBST_EMERGENCIAS = [
    ('SENNA', 'Senna Tower'), ('SAPPHIRE', 'Sapphire'),
    ('TITANIUM', 'Titanium Tower'), ('BLUE VIEW', 'Blue View'),
    ('Titanium Tower', 'Titanium'),
]
_SUBST_IMP_SETOR = [('Obra ' + o, o) for o in OBRAS_VALIDAS_TITULO]
_MESES_METAS = ['ago/25', 'set/25', 'out/25', 'nov/25', 'dez/25',
                'jan/26', 'fev/26', 'mar/26', 'abr/26', 'mai/26', 'jun/26',
                'jul/26', 'ago/26', 'set/26', 'out/26', 'nov/26', 'dez/26']


def _aplicar(df, subs, coluna):
    for de, para in subs:
        df = substituir_texto(df, coluna, de, para)
    return df


# --- Fatos ------------------------------------------------------------------
def _fat_consumo_agua_luz(df):
    df = renomear(df, {'Mês/Ano': 'Data', 'Obra': 'nome_setor'})
    df = duplicar(df, 'nome_setor', 'Obra')
    df = classificar_obra(df, 'Obra', OBRAS_VALIDAS_MAIUSCULO, maiusculo=True)
    return _aplicar(df, _SUBST_AGUA_LUZ, 'Obra')


def _fat_avanco_fisico(df):
    df = renomear(df, {'Mês/Ano': 'Data'})
    return substituir_texto(df, 'Obra', 'Senna', 'Senna Tower')


def _fat_conformidade(df):
    df = renomear(df, {'Mês/Ano': 'Data', 'OBRA': 'Obra'})
    return _aplicar(df, _SUBST_CONFORMIDADE, 'Obra')


def _fat_impressoes(df):
    return renomear(df, {'Mês/Ano': 'Data'})


def _fat_impressoes_por_pessoa(df):
    df = renomear(df, {'Mês/Ano': 'Data'})
    df = duplicar(df, 'Conta atual', 'Obra')
    df = substituir_texto(df, 'Obra', 'Obra ', '')
    return classificar_obra(df, 'Obra', OBRAS_VALIDAS_TITULO)


def _fat_impressoes_por_setor(df):
    df = renomear(df, {'Mês/Ano': 'Data', 'Nome': 'Nome_setor'})
    df = duplicar(df, 'Nome_setor', 'Obra')
    df = _aplicar(df, _SUBST_IMP_SETOR, 'Obra')
    df = classificar_obra(df, 'Obra', OBRAS_VALIDAS_TITULO)
    # Table.Sort({{"%", Order.Descending}}) - so muda a ordem das linhas.
    ordem = pd.to_numeric(df['%'], errors='coerce')
    return df.assign(_ord=ordem).sort_values('_ord', ascending=False,
                                             kind='mergesort').drop(columns='_ord')


def _fat_metas(df):
    return substituir_texto(df, 'Obra', 'Senna', 'Senna Tower')


def _fat_residuos_m3_m2(df):
    df = renomear(df, {'Mês/Ano': 'Data'})
    return substituir_texto(df, 'Obra', 'Senna', 'Senna Tower')


def _fat_ruido(df):
    df = renomear(df, {'Mês/Ano': 'Data'})
    return substituir_texto(df, 'Obra', 'Senna', 'Senna Tower')


def _fat_soma_residuos(df):
    df = renomear(df, {'Mês/Ano': 'Data'})
    return substituir_texto(df, 'Obra', 'Senna', 'Senna Tower')


def _fat_eco_urbana(df):
    return df


# --- Dimensoes --------------------------------------------------------------
def _dim_nc(df):
    return renomear(df, {'Mês/Ano': 'Data'})


def _dim_obras(df):
    df = df[['Obra']].copy()
    return substituir_texto(df, 'Obra', 'Senna', 'Senna Tower')


def _dim_acoes_sga(df):
    return df


def _dim_auditoria(df):
    return _aplicar(df, _SUBST_AUDITORIA, 'Obra ')


def _dim_plano_acao(df):
    # Atencao: o valor procurado tem espaco no fim ("Senna "), como no .pbix.
    return substituir_texto(df, 'Obra', 'Senna ', 'Senna Tower')


def _dim_emergencias(df):
    return _aplicar(df, _SUBST_EMERGENCIAS, 'Obra ')


def _dim_reclamacoes(df):
    """Split de "Blue Coast - 01 - Ouvidoria" em Obra + Canal.

    Table.SplitColumn(Splitter.SplitTextByDelimiter("-"), {"Obra.1","Obra.2"}).
    O separador e o hifen ASCII; o travessao "–" dos codigos ("- 01") NAO
    separa, por isso "Blue Coast – 01 - Ouvidoria" rende 2 partes.
    """
    def _parte(v, i):
        if v is None or v is pd.NA or (isinstance(v, float) and pd.isna(v)):
            return None
        p = str(v).split('-')
        return p[i] if len(p) > i else None
    df['Obra.1'] = df['Obra'].map(lambda v: _parte(v, 0))
    df['Canal'] = df['Obra'].map(lambda v: _parte(v, 1))
    df = df.drop(columns=['Obra'])
    df = substituir_texto(df, 'Obra.1', 'Blue Coast – 01 ', 'Blue Coast')
    df = substituir_texto(df, 'Obra.1', 'Imperium Tower – 02 ', 'Imperium Tower')
    df = renomear(df, {'Obra.1': 'Obra'})
    for de, para in [('Imperium Tower', 'Imperium'), ('Horizon Tower ', 'Horizon'),
                     ('Titanium Tower ', 'Titanium'), ('Garden Park ', 'Garden')]:
        df = substituir_texto(df, 'Obra', de, para)
    return df


def _dim_calendario(_df=None, hoje=None):
    """Tabela de datas gerada pelo proprio M (nao vem do Excel).

    Intervalo = 1o de janeiro de (ano atual - 5) ate 31 de dezembro de
    (ano atual + 5), exatamente como o M faz com DateTime.LocalNow().
    """
    hoje = hoje or _dt.date.today()
    ini = _dt.date(hoje.year - 5, 1, 1)
    fim = _dt.date(hoje.year + 5, 12, 31)
    datas = pd.date_range(ini, fim, freq='D')
    return pd.DataFrame({
        'Data': datas.date,
        'Ano': datas.year,
        'MesNum': datas.month,
        'Mês': [MESES_PT[m - 1] for m in datas.month],
        'Mês/Ano (key)': [d.replace(day=1) for d in datas.date],
        'Trimestre': ['T%d' % q for q in datas.quarter],
        'Dia': datas.day,
        'Dia da Semana': [DIAS_PT[w] for w in datas.weekday],
        'ChaveData': datas.year * 10000 + datas.month * 100 + datas.day,
        'Mês Abrev': [MESES_PT[m - 1][:3] for m in datas.month],
        'Mês/Ano Abrev': ['%02d-%s' % (y % 100, MESES_PT[m - 1][:3])
                          for y, m in zip(datas.year, datas.month)],
    })


# --- Catalogo ---------------------------------------------------------------
TABELAS = [
    Tabela(
        nome='Dim_Obras', fonte='Metas_LAI_EIV', chave=['Obra'],
        colunas=[('Obra', TEXT)],
        construir=_dim_obras,
        m='RemoveColumns(meses) + ReplaceText Obra "Senna"->"Senna Tower"',
        descricao='Lista canonica de obras. Vem da coluna Obra de Metas_LAI_EIV.'),

    Tabela(
        nome='Dim_Calendario', fonte=None, chave=['Data'],
        colunas=[('Data', DATE), ('Ano', INT), ('MesNum', INT), ('Mês', TEXT),
                 ('Mês/Ano (key)', DATE), ('Trimestre', TEXT), ('Dia', INT),
                 ('Dia da Semana', TEXT), ('ChaveData', INT),
                 ('Mês Abrev', TEXT), ('Mês/Ano Abrev', TEXT)],
        construir=_dim_calendario,
        m='List.Dates(ano atual -5 .. ano atual +5) + colunas derivadas',
        descricao='Calendario gerado; nao depende do Excel.'),

    Tabela(
        nome='Dim_Acoes_SGA', fonte='Acoes_SGA', chave=['Ano', 'Mês'],
        colunas=[('Ano', INT), ('Mês', TEXT), ('Pessoas', INT)],
        construir=_dim_acoes_sga,
        m='So TransformColumnTypes.',
        descricao='Pessoas alcancadas por acoes do SGA, por ano/mes.'),

    Tabela(
        nome='Dim_Auditoria_Ambientais', fonte='Auditoria_Ambientais',
        chave=['Obra ', 'Mês/Ano'],
        colunas=[('Obra ', TEXT), ('Ano', INT), ('Mês', TEXT),
                 ('Mês/Ano', DATE), ('Meta ', FLOAT), ('Realizado ', TEXT),
                 ('Principais ocorrencias da última auditória', TEXT)],
        construir=_dim_auditoria,
        m='TransformColumnTypes + 11 ReplaceText na coluna "Obra " (com espaco).',
        descricao='Auditorias ambientais por obra/mes. "Realizado " e Sim/Nao.'),

    Tabela(
        nome='Dim_NC_LAI_EIV', fonte='NC_LAI_EIV', chave=['Obra', 'Data'],
        colunas=[('Obra', TEXT), ('NC LAI', TEXT), ('NC EIV', TEXT),
                 ('Data', DATE)],
        construir=_dim_nc,
        m='TransformColumnTypes + Rename "Mes/Ano"->"Data".',
        descricao='Nao-conformidades de LAI e EIV descritas em texto livre.'),

    Tabela(
        nome='Dim_Plano_acao', fonte='Plano_acao', chave=['Obra', 'Ação'],
        colunas=[('Obra', TEXT), ('Indicador ', TEXT), ('Análise ', TEXT),
                 ('Ação', TEXT), ('Responsável ', TEXT), ('Prazo', DATE),
                 ('Status', TEXT)],
        construir=_dim_plano_acao,
        m='TransformColumnTypes + ReplaceText "Senna " -> "Senna Tower".',
        descricao='Plano de acao das auditorias. Status: Realizado/Em Andamento/...'),

    Tabela(
        nome='Dim_emergencias_ambientais', fonte='emergencias_ambientais',
        chave=['Obra ', 'Mês/Ano', 'N° FMC'],
        colunas=[('Obra ', TEXT), ('Mês/Ano', DATE), ('Fonte ', TEXT),
                 ('N° FMC', INT), ('Ocorerência', TEXT), ('Status/Ações', TEXT),
                 ('Categoria ', TEXT)],
        construir=_dim_emergencias,
        m='TransformColumnTypes + 5 ReplaceText encadeados em "Obra ".',
        descricao='Emergencias ambientais (vazamentos, derrames).'),

    Tabela(
        nome='Dim_reclamacoes_vizinhos', fonte='reclamacoes_vizinhos',
        chave=['Mês/Ano', 'Cod. Chamado'],
        colunas=[('Mês/Ano', DATE), ('Fonte', TEXT), ('Cod. Chamado', INT),
                 ('Ocorrência', TEXT), ('Status/Ações', TEXT),
                 ('Categoria ', TEXT), ('Obra', TEXT), ('Canal', TEXT)],
        construir=_dim_reclamacoes,
        m='SplitColumn "-" em Obra/Canal + ReplaceText encadeados.',
        descricao='Reclamacoes de vizinhos; Canal = Ouvidoria/Relacionamento.'),

    Tabela(
        nome='Fat_Conformidade_LAI_EIV', fonte='Conformidade_LAI_EIV',
        chave=['Obra', 'Data'],
        colunas=[('Obra', TEXT), ('Ano', INT), ('MÊS', TEXT), ('Data', DATE),
                 ('PERCENTUAL', PERCENT), ('META', PERCENT), ('META2', PERCENT)],
        construir=_fat_conformidade,
        m='Rename OBRA->Obra, "Mes/Ano"->Data + 12 ReplaceText.',
        descricao='% de conformidade LAI/EIV por obra/mes, com meta.'),

    Tabela(
        nome='Fat_Eco_urbana', fonte='Eco_urbana', chave=['Obra', 'Mês/Ano'],
        colunas=[('Ano', INT), ('Mês', TEXT), ('Mês/Ano', DATE), ('Obra', TEXT),
                 ('db', INT), ('Ponto2', TEXT), ('db3', FLOAT),
                 ('Ponto_1', TEXT)],
        construir=_fat_eco_urbana,
        m='So TransformColumnTypes.',
        descricao='Medicoes de ruido de terceiros (Ecourbana/Eliz).'),

    Tabela(
        nome='Fat_Impressoes', fonte='Impressoes', chave=['Data'],
        colunas=[('Ano', INT), ('Mês', TEXT), ('Data', DATE),
                 ('Total de pág ', INT), ('Custos', DECIMAL), ('% P&B', PERCENT),
                 ('Pag P&B', FLOAT), ('% Colorido', PERCENT),
                 ('Pag Colorido', FLOAT), ('Pág Simples', FLOAT),
                 ('Pág Dupla', FLOAT), ('Economia preto e branco', INT),
                 ('Economia Colorido', INT), ('Economia Total', INT),
                 (' impressões simples poderiam ter sido feitas duplas', FLOAT),
                 ('Economia de folhas', INT), ('Economia mínima', DECIMAL)],
        construir=_fat_impressoes,
        m='TransformColumnTypes + Rename "Mes/Ano"->Data.',
        descricao='Consumo mensal de impressao da empresa.'),

    Tabela(
        nome='Fat_Impressoes_por_pessoa', fonte='Impressoes_por_pessoa',
        chave=['Data', 'Usuário'],
        colunas=[('Data', DATE), ('Usuário', TEXT), ('Conta atual', TEXT),
                 ('P&B', INT), ('Cor', INT), ('Páginas', INT), ('%', PERCENT),
                 ('Obra', TEXT)],
        construir=_fat_impressoes_por_pessoa,
        m='Duplica "Conta atual"->Obra, tira prefixo "Obra ", resto vira "Outros".',
        descricao='Impressoes por usuario, com a obra deduzida da conta.'),

    Tabela(
        nome='Fat_Impressoes_por_setor', fonte='Impressoes_por_setor',
        chave=['Data', 'Nome_setor'],
        colunas=[('Data', DATE), ('Nome_setor', TEXT), ('P&B', INT),
                 ('Cor', INT), ('Páginas', INT), ('%', PERCENT), ('Obra', TEXT)],
        construir=_fat_impressoes_por_setor,
        m='Duplica Nome_setor->Obra, 11 ReplaceText "Obra X"->"X", resto "Outros".',
        descricao='Impressoes por setor/obra.'),

    Tabela(
        nome='Fat_Metas_LAI_EIV', fonte='Metas_LAI_EIV', chave=['Obra'],
        colunas=[('Obra', TEXT)] + [(m, PERCENT) for m in _MESES_METAS],
        construir=_fat_metas,
        m='TransformColumnTypes + ReplaceText "Senna"->"Senna Tower".',
        descricao='Meta de conformidade por obra, uma COLUNA por mes (formato '
                  'largo herdado do Excel - ver ALERTAS_MODELO).'),

    Tabela(
        nome='Fat_Residuos_m³_m²', fonte='Residuos_m³_m²',
        chave=['Obra', 'Data'],
        colunas=[('Ano', INT), ('Mês', TEXT), ('Data', DATE), ('Obra', TEXT),
                 ('Resíduos m³', FLOAT), ('Avanço Físico m²', FLOAT),
                 ('Resíduos m³/m²', FLOAT)],
        construir=_fat_residuos_m3_m2,
        m='Rename "Mes/Ano"->Data + ReplaceText "Senna"->"Senna Tower".',
        descricao='Volume de residuos (m3) e intensidade por m2 construido.'),

    Tabela(
        nome='Fat_Ruido', fonte='Ruido', chave=['Obra', 'Data', 'Ponto'],
        colunas=[('Ano', INT), ('Mês', TEXT), ('Data', DATE), ('Obra', TEXT),
                 ('Ponto', TEXT), ('Horário', TIME), ('dB', FLOAT)],
        construir=_fat_ruido,
        m='Rename "Mes/Ano"->Data + ReplaceText "Senna"->"Senna Tower".',
        descricao='Medicoes de ruido por obra/ponto. Limite de referencia: 70 dB.'),

    Tabela(
        nome='Fat_avanco_fisico', fonte='Avanco_fisico', chave=['Obra', 'Data'],
        colunas=[('Obra', TEXT), ('Ano', INT), ('Mês', TEXT), ('Data', DATE),
                 ('Avanço %', PERCENT), ('Avanço m²', FLOAT), ('Total m²', FLOAT)],
        construir=_fat_avanco_fisico,
        m='Rename "Mes/Ano"->Data, ReplaceText Senna, ReplaceErrorValues(Data->null).',
        descricao='Avanco fisico da obra; base do denominador m2 dos residuos.'),

    Tabela(
        nome='Fat_consumo_agua_luz', fonte='agua_luz',
        chave=['nome_setor', 'Data'],
        colunas=[('Ano', INT), ('Mês', TEXT), ('Data', DATE),
                 ('nome_setor', TEXT), ('Luz (kwh)', FLOAT), ('Água (m³)', FLOAT),
                 ('Obra', TEXT), ('Luz m²', FLOAT), ('Luz Kwh m³/m²', FLOAT),
                 ('Água m²', FLOAT), ('Água m³/m²', INT)],
        construir=_fat_consumo_agua_luz,
        m='Rename Obra->nome_setor, duplica pra Obra, classifica em MAIUSCULO '
          'contra lista de 11, resto "Outros", depois 10 ReplaceText.',
        descricao='Consumo de agua e luz por setor/obra (ver ALERTAS_MODELO).'),

    Tabela(
        nome='Fat_soma_residuos', fonte='SOMA_RESIDUOS',
        chave=['Obra', 'Data', 'Classe', 'Resíduos Detalhados'],
        colunas=[('Ano', INT), ('Mês', TEXT), ('Data', DATE), ('Obra', TEXT),
                 ('Classe', TEXT), ('m³', INT), ('Resíduos Detalhados', TEXT),
                 ('Kg', TEXT), ('Densidade', FLOAT), ('m²', FLOAT),
                 ('Kg/m²', FLOAT), ('Meta Kg', FLOAT)],
        construir=_fat_soma_residuos,
        m='Rename "Mes/Ano"->Data + ReplaceText "Senna"->"Senna Tower". '
          'Kg fica como TEXTO, como no pbix.',
        descricao='Residuos por obra/mes/classe/tipo. Kg = m3 x Densidade; '
                  'Kg/m2 = Kg / m2.'),
]

TABELAS_POR_NOME = {t.nome: t for t in TABELAS}


# --- Relacionamentos do modelo (do .pbix) -----------------------------------
# Copia fiel dos 26 relacionamentos de PBIXRay(...).relationships. Todos estao
# ATIVOS no modelo. `cardinalidade` e `filtro` sao conceitos do Power BI que
# nao existem em SQL - ficam registrados na tabela `_relacionamentos_modelo`
# para quem for reconstruir o modelo em outra ferramenta.
@dataclass
class Relacionamento:
    origem: str
    coluna_origem: str
    destino: str
    coluna_destino: str
    cardinalidade: str = 'M:1'
    filtro: str = 'Single'      # Single | Both (filtro cruzado no Power BI)
    nota: str = ''

    @property
    def nome_fk(self):
        return 'fk_%s_%s' % (nome_fisico(self.origem),
                             nome_fisico(self.coluna_origem))


_R = Relacionamento
RELACIONAMENTOS = [
    # --- para a dimensao de datas ---
    _R('Fat_Conformidade_LAI_EIV', 'Data', 'Dim_Calendario', 'Data'),
    _R('Fat_Residuos_m³_m²', 'Data', 'Dim_Calendario', 'Data'),
    _R('Fat_consumo_agua_luz', 'Data', 'Dim_Calendario', 'Data'),
    _R('Fat_Eco_urbana', 'Mês/Ano', 'Dim_Calendario', 'Data'),
    _R('Fat_Impressoes', 'Data', 'Dim_Calendario', 'Data', filtro='Both',
       nota='No .pbix este relacionamento esta gravado ao contrario '
            '(Dim_Calendario -> Fat_Impressoes, M:1). Provavel engano de quem '
            'montou o modelo; aqui fica na direcao correta fato -> dimensao.'),
    _R('Fat_Impressoes_por_setor', 'Data', 'Dim_Calendario', 'Data'),
    _R('Fat_Impressoes_por_pessoa', 'Data', 'Dim_Calendario', 'Data'),
    _R('Fat_soma_residuos', 'Data', 'Dim_Calendario', 'Data'),
    _R('Fat_Ruido', 'Data', 'Dim_Calendario', 'Data'),
    _R('Fat_avanco_fisico', 'Data', 'Dim_Calendario', 'Data'),
    _R('Dim_NC_LAI_EIV', 'Data', 'Dim_Calendario', 'Data'),
    # --- para a dimensao de obras ---
    _R('Fat_avanco_fisico', 'Obra', 'Dim_Obras', 'Obra'),
    _R('Fat_Conformidade_LAI_EIV', 'Obra', 'Dim_Obras', 'Obra'),
    _R('Fat_Impressoes_por_pessoa', 'Obra', 'Dim_Obras', 'Obra'),
    _R('Fat_Impressoes_por_setor', 'Obra', 'Dim_Obras', 'Obra'),
    _R('Fat_Metas_LAI_EIV', 'Obra', 'Dim_Obras', 'Obra', cardinalidade='1:1',
       filtro='Both'),
    _R('Fat_Residuos_m³_m²', 'Obra', 'Dim_Obras', 'Obra'),
    _R('Fat_Ruido', 'Obra', 'Dim_Obras', 'Obra'),
    _R('Fat_soma_residuos', 'Obra', 'Dim_Obras', 'Obra'),
    _R('Fat_Eco_urbana', 'Obra', 'Dim_Obras', 'Obra', filtro='Both'),
    _R('Dim_NC_LAI_EIV', 'Obra', 'Dim_Obras', 'Obra'),
    _R('Fat_consumo_agua_luz', 'Obra', 'Dim_Obras', 'Obra'),
    _R('Dim_Auditoria_Ambientais', 'Obra ', 'Dim_Obras', 'Obra'),
    _R('Dim_Plano_acao', 'Obra', 'Dim_Obras', 'Obra'),
    _R('Dim_emergencias_ambientais', 'Obra ', 'Dim_Obras', 'Obra'),
    _R('Dim_reclamacoes_vizinhos', 'Obra', 'Dim_Obras', 'Obra'),
]

# Lado "1" de cada relacionamento: precisa de chave unica para a FK existir.
CHAVES_DIMENSAO = {'Dim_Obras': ['Obra'], 'Dim_Calendario': ['Data']}


# --- Saneamento de obra (NAO vem do pbix) -----------------------------------
# Tudo daqui pra baixo e DESVIO PROPOSITAL do modelo do pbix, feito para o dado
# fechar num banco relacional. Desligue com --fiel-ao-pbix (e o que o
# validar_vs_pbix.py usa, para a prova de fidelidade continuar valendo).
#
# Sao dois problemas diferentes, e a diferenca importa:
#   - obra que EXISTE no dado mas nao em Dim_Obras  -> cadastrar (abaixo)
#   - a MESMA obra escrita de outro jeito           -> normalizar (mais abaixo)
# Confundir os dois criaria obra duplicada ("Garden" e "Garden Park" como se
# fossem duas).

# Obras ausentes de Metas_LAI_EIV, que e de onde Dim_Obras sai.
OBRAS_COMPLEMENTARES = [
    ('Outros',
     'Nao e obra: e o rotulo que o proprio Power Query poe em tudo que nao '
     'esta na lista de obras validas (Financeiro, Marketing, Sede, Deposito, '
     'plantoes de venda...). Precisa existir na dimensao para as 3 tabelas de '
     'impressoes/consumo fecharem.'),
    ('North Tower',
     'Obra real, sem linha em Metas_LAI_EIV - por isso nunca entrou em '
     'Dim_Obras. O certo e a area cadastrar a meta dela na planilha; ate la '
     'entra por aqui para nao derrubar 65 linhas de resIduos, ruido e avanco.'),
]

# Grafias diferentes da MESMA obra ja existente. Nao sao obras novas.
# Aplicado DEPOIS das substituicoes do pbix, em toda coluna que aponta para
# Dim_Obras. A contagem e de 16/09/2026.
NORMALIZACAO_OBRA = {
    # North Tower escrita de 4 jeitos em 4 tabelas diferentes
    'NORTH': 'North Tower',            # Conformidade (4), emergencias (3)
    'North': 'North Tower',            # Ruido (4), avanco fisico (1)
    'NORTH TOWER': 'North Tower',      # auditorias (1)
    # maiusculas que o M de emergencias_ambientais nao trata
    'ICONIC': 'Iconic',                # emergencias (3)
    'IMPERIUM': 'Imperium',            # emergencias (1)
    'GARDEN': 'Garden Park',           # emergencias (1)
    # nome curto / com sufixo "Tower" que sobrou do texto original
    'Garden': 'Garden Park',           # plano de acao (3), reclamacoes (2)
    'Iconic Tower': 'Iconic',          # reclamacoes (2)
    'Blue View Tower': 'Blue View',    # reclamacoes (1)
    # erro de digitacao na planilha (aba Eco urbana)
    'Gran Place': 'Grand Place',       # eco urbana (1)
}

# Medida DAX que existe no modelo (o dashboard nao a usa; recalcula em Python).
MEDIDAS_DAX = {
    '%_Conformidade_LAI_EIV':
        'DIVIDE(SUMX(Fat_Conformidade_LAI_EIV[PERCENTUAL]), '
        'COUNTROWS(Fat_Conformidade_LAI_EIV))',
}

# --- Defeitos conhecidos do modelo, reproduzidos de proposito ---------------
# Defeitos do modelo do pbix que continuam SEM CORRECAO por decisao pendente
# da area - mudam o valor de um indicador, entao nao sao mexidos sem sinal
# verde explicito (mesmo padrao ja usado para [kg/m2 - ver kgm2-agregacao]).
ALERTAS_MODELO = [
    ('Fat_soma_residuos', 'Kg',
     'Kg e TEXTO no modelo, com virgula decimal pt-BR ("1200,8399999999999"). '
     'Quem consumir o banco precisa converter. Num banco relacional o natural '
     'seria NUMERIC - ver secao "Proximos passos" do README.'),
    ('Fat_Metas_LAI_EIV', '(formato)',
     'Tabela larga: uma coluna por mes (ago/25..dez/26). Toda competencia nova '
     'exige ALTER TABLE. O natural no banco e despivotar para '
     '(obra, competencia, meta).'),
]

# Estes JA FORAM CORRIGIDOS por padrao (ver sanear_obras() e
# corrigir_intensidade_agua() em etl_excel_para_db.py). Nao sao mais um
# problema no caminho normal - so reaparecem com --fiel-ao-pbix.
ALERTAS_MODELO_CORRIGIDOS = [
    ('Fat_consumo_agua_luz', 'Água m³/m²',
     'APROVADO PELA AREA em 16/09/2026. O Power Query convertia esta coluna '
     'para Int64, truncando a intensidade de agua (0,206 m3/m2 virava 0) - '
     '72 das 521 linhas calculaveis (14%) zeravam so por arredondamento, nao '
     'porque o consumo foi zero. Corrigido: recalculada como '
     '"Água (m³)" / "Água m²" (as duas ja existiam decimais na mesma tabela; '
     'a conta bate 100% onde as duas existem). Onde a area e 0/nula, o '
     'resultado fica nulo, nao zero.'),
    ('Fat_consumo_agua_luz', 'Obra',
     'A lista de obras validas nao incluia SENNA nem NORTH TOWER: o consumo '
     'dessas obras caia em "Outros". Corrigido: "Outros" agora e uma obra '
     'valida em Dim_Obras (nao e obra de verdade, e o valor que o proprio '
     'Power Query atribui a setor que nao e obra).'),
    ('Dim_reclamacoes_vizinhos', 'Obra',
     'So havia ReplaceText para Imperium/Horizon/Titanium/Garden. "Iconic '
     'Tower", "North Tower" e "Blue View Tower" nao casavam com Dim_Obras. '
     'Corrigido: normalizados para a grafia oficial (Iconic, North Tower, '
     'Blue View).'),
    ('Dim_Obras', '(cobertura)',
     'Dim_Obras saia de Metas_LAI_EIV e tinha so 12 obras; North Tower '
     'aparecia em SOMA_RESIDUOS e Auditoria_Ambientais mas nao estava em '
     'Dim_Obras. Corrigido: North Tower acrescentada.'),
]
