# -*- coding: utf-8 -*-
"""
De onde vem a planilha e como saber se ela mudou.

Tres problemas praticos que este modulo resolve, e que sao a diferenca entre
uma automacao que funciona e uma que falha calada:

1. **`U:` nao existe para servico nem para tarefa agendada.** Unidade mapeada e
   por sessao de usuario. Uma tarefa agendada rodando como conta de servico
   (ou o proprio servico Windows) enxerga `U:\\...` como "caminho nao
   encontrado". Por isso o caminho de rede tem que ser **UNC**
   (`\\\\servidor\\share\\...`), nunca a letra mapeada.

2. **A planilha fica aberta no Excel o dia inteiro.** O Excel nao trava o
   .xlsx para leitura, mas se a carga cair exatamente durante um "Salvar", o
   arquivo lido pode vir truncado. Por isso copiamos para uma pasta local
   antes de ler, com tentativa e repeticao, e validamos que o zip abre.

3. **Recarregar sem necessidade e desperdicio e risco.** Guardamos a impressao
   digital (sha256 + tamanho + data de modificacao) de cada carga; se a
   planilha nao mudou, a carga e pulada.
"""
from __future__ import annotations

import datetime as _dt
import hashlib
import os
import shutil
import tempfile
import time
import zipfile


class OrigemIndisponivel(Exception):
    """A planilha nao pode ser lida (caminho inacessivel, arquivo corrompido)."""


def resolver(caminho_informado=None, candidatos=()):
    """Descobre qual planilha usar.

    Caminho EXPLICITO (--excel ou a variavel DASHFG_PLANILHA) e obrigatorio:
    se nao existir, isso e ERRO, nunca cai para outro candidato. Cair de volta
    numa copia local seria pior que falhar - a carga terminaria "com sucesso"
    sobre dado velho, e ninguem perceberia.

    So a lista interna de candidatos e tentada em ordem.
    """
    explicito = caminho_informado or os.environ.get('DASHFG_PLANILHA')
    if explicito:
        de_onde = '--excel' if caminho_informado else 'DASHFG_PLANILHA'
        if os.path.exists(explicito):
            return explicito
        raise OrigemIndisponivel(
            'A planilha indicada em %s nao existe ou nao esta acessivel:\n'
            '    %s%s' % (de_onde, explicito, _dica(explicito)))

    for c in candidatos:
        if c and os.path.exists(c):
            return c

    detalhe = '\n'.join('    - %s' % t for t in candidatos) or '    (nenhum)'
    raise OrigemIndisponivel(
        'Planilha nao encontrada. Caminhos tentados:\n%s%s'
        % (detalhe, _dica(*candidatos)))


def _dica(*caminhos):
    """Aviso especifico para o erro mais comum: unidade mapeada em servico."""
    mapeada = next((c for c in caminhos
                    if c and len(c) > 2 and c[1:3] == ':\\' and c[0].upper() != 'C'),
                   None)
    if not mapeada:
        return ''
    return ('\n\n  ATENCAO: "%s" usa unidade de rede mapeada (%s).\n'
            '  Tarefa agendada e servico Windows NAO enxergam unidade mapeada -\n'
            '  ela so existe dentro da sessao do usuario que a mapeou.\n'
            '  Use o caminho UNC (\\\\servidor\\share\\...) em DASHFG_PLANILHA\n'
            '  ou em --excel, e garanta que a conta da tarefa tem leitura no share.'
            % (mapeada, mapeada[:2]))


def impressao_digital(caminho, blocos=1 << 20):
    """(sha256, tamanho, data de modificacao) do arquivo."""
    h = hashlib.sha256()
    with open(caminho, 'rb') as fh:
        for bloco in iter(lambda: fh.read(blocos), b''):
            h.update(bloco)
    st = os.stat(caminho)
    return (h.hexdigest(), st.st_size,
            _dt.datetime.fromtimestamp(st.st_mtime).isoformat(timespec='seconds'))


def copiar_local(origem, destino_dir=None, tentativas=3, espera=20):
    """Copia a planilha para disco local e confere que o arquivo abre.

    Se a copia cair no meio de um "Salvar" do Excel, o .xlsx sai invalido -
    nesse caso espera e tenta de novo, em vez de carregar dado truncado.
    """
    destino_dir = destino_dir or os.path.join(tempfile.gettempdir(),
                                              'etl_ambiental')
    os.makedirs(destino_dir, exist_ok=True)
    destino = os.path.join(destino_dir, os.path.basename(origem))

    erro = None
    for n in range(1, tentativas + 1):
        try:
            shutil.copy2(origem, destino)
            _conferir_xlsx(destino)
            return destino
        except Exception as e:                                # noqa: BLE001
            erro = e
            if n < tentativas:
                print('  copia da planilha falhou (tentativa %d/%d): %s'
                      % (n, tentativas, e))
                print('  aguardando %ds - a planilha pode estar sendo salva...'
                      % espera)
                time.sleep(espera)
    raise OrigemIndisponivel(
        'Nao foi possivel copiar a planilha de %s apos %d tentativa(s): %s'
        % (origem, tentativas, erro))


def _conferir_xlsx(caminho):
    """Um .xlsx e um zip; se nao abrir, a copia saiu pela metade."""
    with zipfile.ZipFile(caminho) as z:
        ruim = z.testzip()
        if ruim is not None:
            raise OrigemIndisponivel('arquivo corrompido em %s' % ruim)
        if 'xl/workbook.xml' not in z.namelist():
            raise OrigemIndisponivel('nao parece um .xlsx (sem xl/workbook.xml)')


def eh_rede(caminho):
    """True para UNC (\\\\servidor\\share) ou unidade mapeada."""
    if caminho.startswith('\\\\') or caminho.startswith('//'):
        return True
    unidade = os.path.splitdrive(os.path.abspath(caminho))[0].upper()
    return bool(unidade) and unidade not in ('C:',)
