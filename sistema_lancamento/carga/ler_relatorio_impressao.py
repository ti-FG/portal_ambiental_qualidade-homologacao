# -*- coding: utf-8 -*-
"""
Lê o relatório de impressão exportado do sistema de impressão (arquivo .xls ou
.xlsx) e devolve JSON no stdout. Não grava nada: quem grava é o servidor.

Dois relatórios, reconhecidos pelo título da planilha:
  "Relatório por conta"     -> setores (Impressões por setor) + total (Impressões)
  "Relatório por usuários"  -> pessoas (Impressões por pessoa)

Regras combinadas com a Qualidade (30/10/2026):
  - por setor: só a linha da conta (as linhas "Usuários de ..." repetem o mesmo
    número e ficam de fora); "(Usuários sem conta)" entra, como já entrava;
  - por pessoa: ficam de fora o domínio "Estações Locais" e o usuário FG\\Unknown;
    vale tudo a partir do domínio FG;
  - colunas lidas: P&B, Cor, Páginas e % (o custo do relatório é ignorado: o
    sistema calcula com o preço por página de Cadastros -> Parâmetros).

Uso: python ler_relatorio_impressao.py arquivo.xls
"""
import calendar
import json
import os
import re
import sys
import unicodedata

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'libs'))


def sem_acento(s):
    return ''.join(c for c in unicodedata.normalize('NFD', str(s)) if unicodedata.category(c) != 'Mn').lower().strip()


def ler_linhas(caminho):
    """Lista de linhas, cada uma lista de células (texto, número ou '')."""
    with open(caminho, 'rb') as f:
        cab = f.read(8)
    if cab[:2] == b'PK':  # .xlsx
        import openpyxl
        wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
        ws = wb.worksheets[0]
        return [['' if v is None else v for v in row] for row in ws.iter_rows(values_only=True)]
    import xlrd  # .xls (formato antigo, o que o sistema de impressão exporta)
    wb = xlrd.open_workbook(caminho)
    s = wb.sheet_by_index(0)
    return [s.row_values(r) for r in range(s.nrows)]


def eh_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def pct_texto(v):
    """'22,72%' -> 0.2272 (fração)."""
    m = re.fullmatch(r'\s*(-?\d+(?:[.,]\d+)?)\s*%\s*', str(v))
    return float(m.group(1).replace(',', '.')) / 100 if m else None


def inteiro(v):
    return int(round(v)) if eh_num(v) else None


def periodo_de(linhas):
    for r in linhas[:20]:
        for v in r:
            m = re.search(r'(\d{2})/(\d{2})/(\d{4}).*?(\d{2})/(\d{2})/(\d{4})', str(v))
            if m:
                d1 = (int(m.group(3)), int(m.group(2)), int(m.group(1)))
                d2 = (int(m.group(6)), int(m.group(5)), int(m.group(4)))
                return d1, d2
    return None, None


def tipo_de(linhas):
    for r in linhas[:15]:
        for v in r:
            t = sem_acento(v)
            if t.startswith('relatorio por conta'):
                return 'conta'
            if t.startswith('relatorio por usuario'):
                return 'usuarios'
    # sem título: o cabeçalho da tabela diz qual é
    for r in linhas[:30]:
        textos = [sem_acento(v) for v in r if isinstance(v, str)]
        if 'nome' in textos and 'p&b' in textos:
            return 'conta'
        if 'usuario' in textos and 'p&b' in textos:
            return 'usuarios'
    return None


def partes(r):
    """Separa a linha em (coluna do 1º texto, texto, números em ordem, % em texto)."""
    col = txt = None
    nums, pct = [], None
    for c, v in enumerate(r):
        if isinstance(v, str) and v.strip():
            p = pct_texto(v)
            if p is not None:
                pct = p
            elif txt is None:
                col, txt = c, v.strip()
        elif eh_num(v):
            nums.append(v)
    return col, txt, nums, pct


def main(caminho):
    linhas = ler_linhas(caminho)
    tipo = tipo_de(linhas)
    avisos = []
    if tipo is None:
        raise SystemExit('Não reconheci o arquivo: esperava o "Relatório por conta" ou o "Relatório por usuários".')

    d1, d2 = periodo_de(linhas)
    if not d1:
        raise SystemExit('Não achei o período (dd/mm/aaaa até dd/mm/aaaa) no topo do relatório.')
    if (d1[0], d1[1]) != (d2[0], d2[1]):
        raise SystemExit('O relatório cobre mais de um mês (%02d/%02d/%d a %02d/%02d/%d). Exporte um mês por vez.' % (d1[2], d1[1], d1[0], d2[2], d2[1], d2[0]))
    ultimo = calendar.monthrange(d1[0], d1[1])[1]
    r = {
        'tipo': tipo,
        'competencia': '%d-%02d' % (d1[0], d1[1]),
        'periodo': {'de': '%d-%02d-%02d' % (d1[0], d1[1], d1[2]), 'ate': '%d-%02d-%02d' % (d2[0], d2[1], d2[2])},
        'parcial': not (d1[2] == 1 and d2[2] >= ultimo),
        'total': None, 'setores': [], 'pessoas': [], 'ignorados': {}, 'avisos': avisos,
    }

    if tipo == 'conta':
        iniciou = False
        for lin in linhas:
            col, txt, nums, pct = partes(lin)
            if txt is None:
                continue
            t = sem_acento(txt)
            if not iniciou:
                iniciou = t == 'nome' and col is not None
                continue
            if t.startswith('total geral'):
                if len(nums) >= 3:
                    r['total'] = {'pb': inteiro(nums[0]), 'cor': inteiro(nums[1]), 'paginas': inteiro(nums[2])}
                break
            if len(nums) < 3:
                continue
            if t.startswith('usuarios de '):  # repete o número da conta acima
                continue
            pb, cor, pag = inteiro(nums[0]), inteiro(nums[1]), inteiro(nums[2])
            if pct is None and len(nums) >= 5:  # % numérico (fração) no lugar do texto
                pct = nums[4]
            r['setores'].append({'nome': txt, 'pb': pb, 'cor': cor, 'paginas': pag, 'pct': pct})
        if not r['total']:
            avisos.append('O relatório não trouxe a linha "Total geral": o total do mês foi somado das contas.')
            r['total'] = {k: sum(x[k] or 0 for x in r['setores']) for k in ('pb', 'cor', 'paginas')}
        soma = sum(x['paginas'] or 0 for x in r['setores'])
        if soma != r['total']['paginas']:
            avisos.append('A soma das contas (%d páginas) não bate com o "Total geral" do relatório (%d).' % (soma, r['total']['paginas']))
        for x in r['setores']:
            if x['pct'] is None and r['total']['paginas']:
                x['pct'] = round((x['paginas'] or 0) / r['total']['paginas'], 6)
        nomes = {}
        for x in r['setores']:
            nomes.setdefault(sem_acento(x['nome']), []).append(x['nome'])
        dup = [v[0] for v in nomes.values() if len(v) > 1]
        if dup:
            avisos.append('Nome repetido no relatório (as páginas foram somadas numa linha só): ' + ', '.join(dup))
            somados = {}
            for x in r['setores']:
                k = sem_acento(x['nome'])
                if k in somados:
                    for c in ('pb', 'cor', 'paginas'):
                        somados[k][c] = (somados[k][c] or 0) + (x[c] or 0)
                    somados[k]['pct'] = (somados[k]['pct'] or 0) + (x['pct'] or 0)
                else:
                    somados[k] = dict(x)
            r['setores'] = list(somados.values())
    else:
        iniciou = False
        dominio = None
        estacoes = {'linhas': 0, 'paginas': 0}
        unknown = {'linhas': 0, 'paginas': 0}
        for lin in linhas:
            col, txt, nums, pct = partes(lin)
            if txt is None:
                continue
            t = sem_acento(txt)
            if not iniciou:
                iniciou = t == 'usuario'
                continue
            if t == 'dominio':
                textos = [str(v).strip() for v in lin if isinstance(v, str) and v.strip()]
                dominio = textos[1] if len(textos) > 1 else ''
                continue
            if t.startswith('total geral'):
                if len(nums) >= 3:
                    r['total'] = {'pb': inteiro(nums[0]), 'cor': inteiro(nums[1]), 'paginas': inteiro(nums[2])}
                break
            if t.startswith('total') or len(nums) < 3:
                continue
            textos = [str(v).strip() for v in lin if isinstance(v, str) and v.strip()]
            conta = textos[1] if len(textos) > 1 else None
            pb, cor, pag = inteiro(nums[0]), inteiro(nums[1]), inteiro(nums[2])
            if pct is None and len(nums) >= 5:
                pct = nums[4]  # no relatório por usuário o % vem como número já em pontos percentuais
            if dominio is not None and sem_acento(dominio).startswith('estaco'):
                estacoes['linhas'] += 1
                estacoes['paginas'] += pag or 0
                continue
            if t.endswith('\\unknown'):
                unknown['linhas'] += 1
                unknown['paginas'] += pag or 0
                continue
            r['pessoas'].append({'usuario': txt, 'conta_atual': conta, 'pb': pb, 'cor': cor, 'paginas': pag, 'pct': pct})
        r['ignorados'] = {'estacoes_locais': estacoes, 'unknown': unknown}
        if not r['total']:
            avisos.append('O relatório não trouxe a linha "Total geral".')
        # mesmo usuário na mesma conta duas vezes: soma numa linha só
        vistos = {}
        for x in r['pessoas']:
            k = (sem_acento(x['usuario']), sem_acento(x['conta_atual'] or ''))
            if k in vistos:
                for c in ('pb', 'cor', 'paginas', 'pct'):
                    vistos[k][c] = (vistos[k][c] or 0) + (x[c] or 0)
                avisos.append('Usuário repetido no relatório (somado numa linha só): ' + x['usuario'])
            else:
                vistos[k] = dict(x)
        r['pessoas'] = list(vistos.values())
        if not r['pessoas']:
            raise SystemExit('Não encontrei usuários no relatório.')

    if tipo == 'conta' and not r['setores']:
        raise SystemExit('Não encontrei contas (setores) no relatório.')
    sys.stdout.buffer.write(json.dumps(r, ensure_ascii=False).encode('utf-8'))  # utf-8 mesmo com o stdout em cp1252


if __name__ == '__main__':
    if len(sys.argv) != 2:
        raise SystemExit('uso: ler_relatorio_impressao.py arquivo.xls')
    try:
        main(sys.argv[1])
    except SystemExit:
        raise
    except Exception as e:  # arquivo corrompido ou de outro tipo
        raise SystemExit('Não consegui ler o arquivo (%s).' % e)
