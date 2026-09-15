# -*- coding: utf-8 -*-
"""
Modelo semantico Ambiental/Qualidade - helpers de transformacao.

Este modulo reproduz, em Python, o que o Power Query do
`AMBIENTAL_QUALIDADE_202602_v1_DEV.pbix` faz com a planilha
`Controle Ambiental 2025_oficial_db.xlsx`.

O codigo M de cada tabela foi extraido do proprio .pbix
(`PBIXRay.power_query`) e esta transcrito em `tabelas.py`, para auditoria.

Regra de ouro: **fidelidade ao pbix**, inclusive nos defeitos. Onde o Power
Query faz algo questionavel (ex.: truncar `Agua m3/m2` para inteiro), a mesma
coisa e feita aqui e o problema fica registrado em ALERTAS_MODELO (tabelas.py),
nao corrigido silenciosamente.
"""
from __future__ import annotations

import datetime as _dt
import re
import unicodedata
from decimal import Decimal

import numpy as np
import pandas as pd

# --- Tipos logicos ----------------------------------------------------------
INT = 'int'            # Int64.Type                     -> BIGINT
FLOAT = 'float'        # type number / Percentage.Type  -> DOUBLE PRECISION
DECIMAL = 'decimal'    # Currency.Type                  -> NUMERIC(19,4)
TEXT = 'text'          # type text                      -> NVARCHAR
PERCENT = 'percent'    # Percentage.Type                -> DOUBLE PRECISION
DATE = 'date'          # type date                      -> DATE
TIME = 'time'          # type datetime (so hora)        -> TIME
DATETIME = 'datetime'  # type datetime                  -> TIMESTAMP

# --- Listas de obras usadas pelo Power Query --------------------------------
# Sao listas DIFERENTES entre si no .pbix (caixa e conteudo) e continuam
# diferentes aqui de proposito. Unifica-las mudaria os numeros do dashboard.
OBRAS_VALIDAS_TITULO = [
    'Blue Coast', 'Blue View', 'Boreal', 'Garden Park', 'Grand Place',
    'Horizon', 'Iconic', 'Imperium', 'Sapphire', 'Scenarium', 'Titanium',
]
OBRAS_VALIDAS_MAIUSCULO = [
    'BLUE COAST', 'BLUE VIEW', 'BOREAL', 'GARDEN PARK', 'GRAND PLACE',
    'HORIZON', 'ICONIC', 'IMPERIUM', 'SAPPHIRE', 'SCENARIUM', 'TITANIUM',
]

MESES_PT = ['janeiro', 'fevereiro', 'marco', 'abril', 'maio', 'junho', 'julho',
            'agosto', 'setembro', 'outubro', 'novembro', 'dezembro']
MESES_PT[2] = 'março'
DIAS_PT = ['segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira',
           'sexta-feira', 'sábado', 'domingo']

# --- Nome fisico das colunas no banco ---------------------------------------
# Regra mecanica + excecoes onde o resultado sairia ilegivel ("% P&B" -> pct_peb).
_NOME_FISICO_EXCECOES = {'% P&B': 'pct_pb', 'Pag P&B': 'pag_pb', 'P&B': 'pb'}

# O modelo do pbix nao guarda espaco no FIM de valor de texto: a planilha tem
# 'Emegencia ' e o modelo tem 'Emegencia' (espaco no INICIO e preservado, como
# em ' Ouvidoria'). Reproduzimos isso; desligar so faz sentido pra depurar.
TRIM_TEXTO_FINAL = True


def nome_fisico(nome):
    """Converte um nome do modelo (com acento, %, m3, /) em identificador SQL."""
    if nome in _NOME_FISICO_EXCECOES:
        return _NOME_FISICO_EXCECOES[nome]
    s = nome.strip()
    s = s.replace('³', '3').replace('²', '2').replace('%', 'pct')
    s = s.replace('&', 'e').replace('°', '').replace('º', '')
    s = unicodedata.normalize('NFKD', s)
    s = ''.join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r'[^0-9A-Za-z]+', '_', s).strip('_').lower()
    if not s:
        s = 'col'
    if s[0].isdigit():
        s = 'c_' + s
    return s


# --- Operacoes equivalentes as do Power Query -------------------------------
def _txt(v):
    """Valor como texto, preservando nulo (equivale a `type text` do PQ)."""
    if v is None or v is pd.NA:
        return None
    if isinstance(v, float) and np.isnan(v):
        return None
    return v if isinstance(v, str) else str(v)


def substituir_texto(df, coluna, de, para):
    """Table.ReplaceValue(..., Replacer.ReplaceText, {coluna}).

    Replacer.ReplaceText troca SUBSTRING, nao o valor inteiro - por isso a
    ordem das substituicoes importa (o pbix usa isso de proposito em
    Dim_emergencias_ambientais: TITANIUM -> 'Titanium Tower' -> 'Titanium').
    """
    if coluna not in df.columns:
        return df
    df[coluna] = df[coluna].map(
        lambda v: None if _txt(v) is None else _txt(v).replace(de, para))
    return df


def classificar_obra(df, coluna, validos, maiusculo=False, outros='Outros'):
    """Table.TransformColumns: mantem o que esta na lista, resto vira 'Outros'."""
    def _c(v):
        t = _txt(v)
        if t is None:
            return outros
        if maiusculo:
            return t.upper() if t.upper() in validos else outros
        return t if t in validos else outros
    df[coluna] = df[coluna].map(_c)
    return df


def duplicar(df, origem, destino):
    """Table.DuplicateColumn."""
    df[destino] = df[origem]
    return df


def renomear(df, mapa):
    """Table.RenameColumns."""
    return df.rename(columns=mapa)


# --- Coercao de tipos -------------------------------------------------------
def _parece_ptbr(s):
    return bool(re.fullmatch(r'-?\d{1,3}(\.\d{3})*(,\d+)?|-?\d+,\d+', s.strip()))


def _para_int(v):
    """Int64.Type - arredonda pro inteiro mais proximo (metade longe do zero)."""
    f = _para_float(v)
    if f is None or np.isnan(f):
        return pd.NA
    return int(np.floor(f + 0.5)) if f >= 0 else int(np.ceil(f - 0.5))


def _para_float(v):
    if v is None or v is pd.NA:
        return np.nan
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, str):
        s = v.strip()
        if s == '' or s.startswith('#'):
            return np.nan
        if _parece_ptbr(s):
            s = s.replace('.', '').replace(',', '.')
        try:
            return float(s)
        except ValueError:
            return np.nan
    if isinstance(v, (_dt.datetime, _dt.date, _dt.time)):
        return np.nan
    try:
        return float(v)
    except (TypeError, ValueError):
        return np.nan


def _para_decimal(v):
    """Currency.Type - decimal fixo de 4 casas."""
    f = _para_float(v)
    if f is None or np.isnan(f):
        return None
    return Decimal(repr(round(f, 4)))


def _fim(s):
    """Remove espaco no fim do texto, como o modelo do pbix faz."""
    if not TRIM_TEXTO_FINAL:
        return s
    s = s.rstrip()
    return s if s else None


def _para_percent(v):
    """Percentage.Type - aceita numero (0,17) e texto do Excel ('17,19%')."""
    if isinstance(v, str):
        s = v.strip()
        if s.endswith('%'):
            f = _para_float(s[:-1])
            return np.nan if np.isnan(f) else f / 100.0
    return _para_float(v)


def _para_texto_pq(v):
    """`type text` do Power Query aplicado sobre um numero.

    O pbix roda em pt-BR: o separador decimal vira virgula e o numero sai em
    formato round-trip - por isso valores como '1200,8399999999999' aparecem
    na coluna Kg de Fat_soma_residuos.
    """
    if v is None or v is pd.NA:
        return None
    if isinstance(v, str):
        return _fim(v)
    if isinstance(v, bool):
        return 'TRUE' if v else 'FALSE'
    if isinstance(v, (int, np.integer)):
        return str(int(v))
    if isinstance(v, (float, np.floating, Decimal)):
        f = float(v)
        if np.isnan(f):
            return None
        return str(int(f)) if f.is_integer() else repr(f).replace('.', ',')
    if isinstance(v, _dt.datetime):
        return v.strftime('%d/%m/%Y %H:%M:%S')
    if isinstance(v, _dt.date):
        return v.strftime('%d/%m/%Y')
    return str(v)


def _para_data(v):
    if v is None or v is pd.NA:
        return None
    if isinstance(v, _dt.datetime):
        return v.date()
    if isinstance(v, _dt.date):
        return v
    if isinstance(v, str):
        s = v.strip()
        # '#VALOR!', '#N/D' -> Table.ReplaceErrorValues({{"Data", null}})
        if s == '' or s.startswith('#'):
            return None
        d = pd.to_datetime(s, errors='coerce', dayfirst=True)
        return None if pd.isna(d) else d.date()
    if isinstance(v, (int, float, np.number)):
        if isinstance(v, float) and np.isnan(v):
            return None
        return (_dt.datetime(1899, 12, 30) + _dt.timedelta(days=float(v))).date()
    return None


def _para_hora(v):
    if v is None or v is pd.NA:
        return None
    if isinstance(v, _dt.datetime):
        return v.time().replace(microsecond=0)
    if isinstance(v, _dt.time):
        return v.replace(microsecond=0)
    if isinstance(v, (int, float, np.number)):
        if isinstance(v, float) and np.isnan(v):
            return None
        seg = int(round((float(v) % 1.0) * 86400))
        return (_dt.datetime(2000, 1, 1) + _dt.timedelta(seconds=seg)).time()
    if isinstance(v, str) and v.strip():
        d = pd.to_datetime(v.strip(), errors='coerce')
        return None if pd.isna(d) else d.time().replace(microsecond=0)
    return None


def _para_datetime(v):
    if v is None or v is pd.NA:
        return None
    if isinstance(v, _dt.datetime):
        return v
    if isinstance(v, _dt.date):
        return _dt.datetime(v.year, v.month, v.day)
    d = pd.to_datetime(v, errors='coerce')
    return None if pd.isna(d) else d.to_pydatetime()


_COERCAO = {
    INT: (_para_int, 'Int64'),
    FLOAT: (_para_float, 'Float64'),
    PERCENT: (_para_percent, 'Float64'),
    DECIMAL: (_para_decimal, 'object'),
    TEXT: (_para_texto_pq, 'string'),
    DATE: (_para_data, 'object'),
    TIME: (_para_hora, 'object'),
    DATETIME: (_para_datetime, 'datetime64[ns]'),
}


def aplicar_tipos(df, colunas):
    """Coage cada coluna ao tipo declarado e devolve na ordem do modelo."""
    saida = {}
    for nome, tipo in colunas:
        fn, dtype = _COERCAO[tipo]
        if nome in df.columns:
            serie = df[nome]
        else:
            serie = pd.Series([None] * len(df), index=df.index)
        saida[nome] = pd.Series([fn(v) for v in serie], index=df.index, dtype=dtype)
    return pd.DataFrame(saida, index=df.index)[[c for c, _ in colunas]]
