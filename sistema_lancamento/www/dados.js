// =============================================================================
// dados.js — tudo o que o sistema lê e grava no schema portal_ambiental_qualidade
// (db/schema.sql). SQL à mão, sem ORM, como no FG Engenharia e no Projetos.
//
// Regra de gravação dos lançamentos (padrão do Jurídico): cada célula alterada
// vira uma linha em log_alteracao com o valor de antes e o de depois. Linha
// que não mudou não é regravada. Célula esvaziada = exclusão, também logada.
// Mês fechado (tabela fechamento): só admin altera, e com motivo.
// =============================================================================
'use strict';

const { Pool, types } = require('pg');
const crypto = require('crypto');

types.setTypeParser(1700, v => (v === null ? null : Number(v))); // numeric -> número
types.setTypeParser(1082, v => v);                                // date -> 'AAAA-MM-DD'

function erroUso(msg, status) { const e = new Error(msg); e.uso = true; e.status = status || 400; return e; }

// "2026-09" ou "2026-09-01" -> "2026-09-01"
function competencia(v) {
  const m = /^(\d{4})-(\d{2})(?:-01)?$/.exec(String(v || ''));
  if (!m || +m[2] < 1 || +m[2] > 12) throw erroUso('competência inválida (use AAAA-MM)');
  return `${m[1]}-${m[2]}-01`;
}

function numero(v, { min, max, inteiro } = {}) {
  if (v === null || v === undefined || v === '') return null;
  const n = typeof v === 'number' ? v : Number(String(v).replace(/\./g, '').replace(',', '.'));
  if (!Number.isFinite(n)) throw erroUso('valor numérico inválido: ' + v);
  if (inteiro && !Number.isInteger(n)) throw erroUso('valor deve ser inteiro: ' + v);
  if (min !== undefined && n < min) throw erroUso(`valor abaixo do mínimo (${min}): ${v}`);
  if (max !== undefined && n > max) throw erroUso(`valor acima do máximo (${max}): ${v}`);
  return n;
}

// "8:13" -> "08:13"; vazio -> null
function hora(v) {
  if (v === null || v === undefined || String(v).trim() === '') return null;
  const m = /^(\d{1,2}):(\d{2})$/.exec(String(v).trim());
  if (!m || +m[1] > 23 || +m[2] > 59) throw erroUso('horário inválido (HH:MM): ' + v);
  return m[1].padStart(2, '0') + ':' + m[2];
}

// ── Lançamentos editáveis em grade: chave + campos ─────────────────────────
// `rotulo` monta a chave legível do histórico a partir dos cadastros.
const FONTES = {
  avanco: {
    tabela: 'avanco_fisico', chave: ['obra_id'], campos: {
      avanco_pct: v => numero(v, { min: 0, max: 1 }),
    },
    // o valor vem do Prevision (Integrações): ninguém digita nem altera, só exclui a linha
    soExclusao: 'O Avanço físico vem do Prevision: não se digita nem se altera aqui. Só é possível excluir uma linha que não faça sentido.',
  },
  conformidade: {
    tabela: 'conformidade', chave: ['obra_id'], campos: {
      itens_conformes: v => numero(v, { min: 0, inteiro: true }),
      itens_avaliados: v => numero(v, { min: 1, inteiro: true }),
      percentual: v => numero(v, { min: 0, max: 1 }),
      observacao: v => (v === null || v === undefined || String(v).trim() === '' ? null : String(v).trim().slice(0, 500)),
    },
    // % sai dos itens quando eles vêm (era "=53/57" digitado na célula)
    ajustar(l) {
      if (l.itens_conformes !== null && l.itens_avaliados !== null) {
        if (l.itens_conformes > l.itens_avaliados) throw erroUso('itens conformes maiores que os avaliados');
        l.percentual = Math.round((l.itens_conformes / l.itens_avaliados) * 10000) / 10000;
      }
      if (l.percentual === null && (l.itens_conformes !== null || l.itens_avaliados !== null || l.observacao)) {
        throw erroUso('informe o percentual ou os itens conformes e avaliados');
      }
    },
  },
  residuo: {
    tabela: 'residuo', chave: ['obra_id', 'tipo_residuo_id'], campos: {
      volume_m3: v => numero(v, { min: 0 }),
    },
    // "Classe B (sem detalhe)" e afins são só do histórico: lançamento novo exige o resíduo detalhado
    async validar(cli, S, k, atual) {
      if (atual) return;
      const r = await cli.query(`SELECT nome, legado FROM ${S}.tipo_residuo WHERE id = $1`, [k.tipo_residuo_id]);
      if (r.rows[0] && r.rows[0].legado) throw erroUso(`"${r.rows[0].nome}" é um registro do histórico e não recebe lançamento novo: escolha o resíduo detalhado.`);
    },
  },
  consumo: {
    tabela: 'consumo', chave: ['uc_id'], campos: {
      luz_kwh: v => numero(v, { min: 0 }),
      agua_m3: v => numero(v, { min: 0 }),
    },
    // unidade que agrupa outras (Sede = salas) é soma: não se digita
    async validar(cli, S, k) {
      const r = await cli.query(`SELECT nome FROM ${S}.unidade_consumidora u
        WHERE id = $1 AND EXISTS (SELECT 1 FROM ${S}.unidade_consumidora f WHERE f.agrupa_em_id = u.id)`, [k.uc_id]);
      if (r.rowCount) throw erroUso(`${r.rows[0].nome} é a soma das unidades dela: lance nas unidades`);
    },
  },
};

// ── Cadastros editáveis (só admin) ─────────────────────────────────────────
const texto = (max = 200) => v => {
  const s = v === null || v === undefined ? '' : String(v).trim();
  return s ? s.slice(0, max) : null;
};
const booleano = v => v === true || v === 'true' || v === 1;
const CADASTROS = {
  obra: {
    tabela: 'obra', chave: 'id', obrigatorios: ['nome'], campos: {
      nome: texto(120), ativa: booleano,
      area_total_m2: v => numero(v, { min: 0 }),
      meta_residuo_kg_m2: v => numero(v, { min: 0 }),
      observacao: texto(500),
    },
  },
  tipo_residuo: {
    tabela: 'tipo_residuo', chave: 'id', obrigatorios: ['nome', 'classe'], campos: {
      nome: texto(120),
      classe: v => { const c = String(v || '').trim().toUpperCase(); if (!/^[ABCD]$/.test(c)) throw erroUso('classe deve ser A, B, C ou D'); return c; },
      densidade_kg_m3: v => numero(v, { min: 0 }),
      perigoso: booleano, ativo: booleano,
      ordem: v => numero(v, { inteiro: true }) ?? 100,
    },
  },
  unidade_consumidora: {
    tabela: 'unidade_consumidora', chave: 'id', obrigatorios: ['nome', 'tipo'], campos: {
      nome: texto(120),
      tipo: v => { if (!['obra', 'administrativo'].includes(v)) throw erroUso('tipo deve ser obra ou administrativo'); return v; },
      obra_id: v => numero(v, { inteiro: true }),
      agrupa_em_id: v => numero(v, { inteiro: true }),
      ativa: booleano,
      ordem: v => numero(v, { inteiro: true }) ?? 100,
    },
  },
  assunto_nc: {
    tabela: 'assunto_nc', chave: 'id', obrigatorios: ['tipo', 'nome'], campos: {
      tipo: v => { if (!['LAI/EIV', 'LAI', 'EIV'].includes(v)) throw erroUso('tipo deve ser LAI/EIV, LAI ou EIV'); return v; },
      nome: texto(200), ativo: booleano,
    },
  },
};

// ── Abas que são listas ────────────────────────────────────────────────────
// filtro: 'mes' (linhas da competência), 'ano' (linhas do ano), 'todos'.
// datado: a competência sai da data do próprio registro (emergência, reclamação).
const dataBr = v => {
  if (v === null || v === undefined || String(v).trim() === '') return null;
  const s = String(v).trim();
  let m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (m) return s;
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s);
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  throw erroUso('data inválida (dd/mm/aaaa): ' + v);
};
const inteiro = (min, max) => v => numero(v, { inteiro: true, min, max });
const fracao = v => numero(v, { min: 0, max: 1 });
const simNao = v => (v === null || v === undefined || v === '' ? null : v === true || v === 'true' || v === 'Sim');
const LISTAS = {
  auditoria: {
    tabela: 'auditoria_ambiental', filtro: 'todos', mesLinha: true, ordem: 'competencia, obra_id, id', obrigatorios: ['obra_id', 'competencia'],
    colunas: ['obra_id', 'competencia', 'meta', 'realizado', 'ocorrencias'],
    campos: { obra_id: inteiro(1), ano: inteiro(2000, 2100), mes: inteiro(1, 12), meta: fracao, realizado: simNao, ocorrencias: texto(2000) },
    // ano e mês vêm da própria linha (a tela mostra todas, como a planilha)
    preparar(l) { l.competencia = l.ano && l.mes ? `${l.ano}-${String(l.mes).padStart(2, '0')}-01` : null; delete l.ano; delete l.mes; },
    rotulos: { obra_id: 'Obra', competencia: 'Ano e Mês' },
  },
  plano_acao: {
    tabela: 'plano_acao', filtro: 'todos', ordem: "(status ILIKE 'realizado%'), prazo NULLS LAST, obra_id, id", obrigatorios: ['obra_id'],
    colunas: ['obra_id', 'indicador', 'analise', 'acao', 'responsavel', 'prazo', 'status'],
    campos: { obra_id: inteiro(1), indicador: texto(120), analise: texto(500), acao: texto(500), responsavel: texto(120), prazo: dataBr, status: texto(60) },
    rotulos: { obra_id: 'Obra' },
  },
  emergencia: {
    tabela: 'emergencia', filtro: 'todos', datado: true, ordem: 'data, id', obrigatorios: ['obra_id', 'data'],
    colunas: ['obra_id', 'data', 'fonte', 'n_fmc', 'ocorrencia', 'status_acoes', 'categoria'],
    campos: { obra_id: inteiro(1), data: dataBr, fonte: texto(60), n_fmc: inteiro(0), ocorrencia: texto(500), status_acoes: texto(1000), categoria: texto(60) },
    rotulos: { obra_id: 'Obra', data: 'Mês/Ano' },
  },
  reclamacao: {
    tabela: 'reclamacao', filtro: 'todos', datado: true, ordem: 'data, id', obrigatorios: ['obra_id', 'data'], extraLeitura: ['sf_case_id'],
    // os chamados vêm do Salesforce (Integrações): ninguém digita nem altera, só exclui a linha
    soExclusao: 'As reclamações vêm do Salesforce: não se digita nem se altera aqui. Só é possível excluir uma linha que não faça sentido.',
    colunas: ['obra_id', 'canal', 'data', 'fonte', 'cod_chamado', 'ocorrencia', 'status_acoes', 'categoria'],
    campos: { obra_id: inteiro(1), canal: texto(60), data: dataBr, fonte: texto(60), cod_chamado: texto(40), ocorrencia: texto(1000), status_acoes: texto(1000), categoria: texto(60) },
    rotulos: { obra_id: 'Obra', data: 'Mês/Ano' },
  },
  eco_urbana: {
    tabela: 'eco_urbana', filtro: 'todos', mesLinha: true, ordem: 'competencia, obra_id, id', obrigatorios: ['obra_id', 'competencia'],
    colunas: ['obra_id', 'competencia', 'ponto_1', 'db', 'ponto2', 'db3'],
    // ano e mês vêm da própria linha (a tela mostra todas, como a planilha)
    preparar(l) { l.competencia = l.ano && l.mes ? `${l.ano}-${String(l.mes).padStart(2, '0')}-01` : null; delete l.ano; delete l.mes; },
    campos: { obra_id: inteiro(1), ano: inteiro(2000, 2100), mes: inteiro(1, 12), ponto_1: texto(20), db: v => numero(v, { min: 0, max: 199 }), ponto2: texto(20), db3: v => numero(v, { min: 0, max: 199 }) },
    rotulos: { obra_id: 'Obra' },
  },
  acoes_sga: {
    tabela: 'acoes_sga', filtro: 'todos', ordem: 'ano, mes', obrigatorios: ['ano', 'mes'],
    colunas: ['ano', 'mes', 'pessoas'],
    campos: { ano: inteiro(2000, 2100), mes: inteiro(0, 12), pessoas: inteiro(0) },
    rotulos: { ano: 'Ano', mes: 'Mês' },
  },
  impressoes: {
    tabela: 'impressoes', filtro: 'todos', mesLinha: true, ordem: 'competencia', obrigatorios: ['competencia'],
    colunas: ['competencia', 'total_pag', 'pct_colorido', 'pag_dupla', 'economia_pb', 'economia_colorido', 'simples_duplas', 'economia_folhas'],
    campos: { ano: inteiro(2000, 2100), mes: inteiro(1, 12), total_pag: inteiro(0), pct_colorido: fracao, pag_dupla: fracao, economia_pb: inteiro(), economia_colorido: inteiro(), simples_duplas: v => numero(v), economia_folhas: inteiro() },
    // ano e mês vêm da própria linha (a tela mostra todas, como a planilha)
    preparar(l) { l.competencia = l.ano && l.mes ? `${l.ano}-${String(l.mes).padStart(2, '0')}-01` : null; delete l.ano; delete l.mes; },
    rotulos: { competencia: 'Ano e Mês' },
  },
  impressoes_setor: {
    tabela: 'impressoes_setor', filtro: 'mes', comOrigem: true, unico: ['nome'], ordem: 'paginas DESC NULLS LAST, nome', obrigatorios: ['nome'],
    colunas: ['competencia', 'nome', 'pb', 'cor', 'paginas', 'pct'],
    campos: { nome: texto(120), pb: inteiro(0), cor: inteiro(0), paginas: inteiro(0), pct: v => numero(v) },
    rotulos: { nome: 'Nome' },
  },
  impressoes_pessoa: {
    tabela: 'impressoes_pessoa', filtro: 'mes', comOrigem: true, unico: ['usuario', 'conta_atual'], ordem: 'paginas DESC NULLS LAST, usuario', obrigatorios: ['usuario'],
    colunas: ['competencia', 'usuario', 'conta_atual', 'pb', 'cor', 'paginas', 'pct'],
    campos: { usuario: texto(200), conta_atual: texto(120), pb: inteiro(0), cor: inteiro(0), paginas: inteiro(0), pct: v => numero(v) },
    rotulos: { usuario: 'Usuário' },
  },
};
// mês da linha, para a trava de mês fechado (null = sem mês: Plano de Ação, "Todos")
function compDe(L, row, F) {
  if (L.datado) return row.data ? String(row.data).slice(0, 7) + '-01' : null;
  if (row.competencia) return String(row.competencia).slice(0, 10);
  if (L.tabela === 'acoes_sga') return row.mes && row.ano ? `${row.ano}-${String(row.mes).padStart(2, '0')}-01` : null;
  return F.comp || null;
}
function chaveLinha(L, row, F, id) {
  const k = { id };
  if (row.obra_id !== undefined && row.obra_id !== null) k.obra_id = row.obra_id;
  const c = compDe(L, row, F);
  if (c) k.competencia = c;
  return k;
}
const iguais = (a, b) => (a ?? null) === (b ?? null) || (a != null && b != null && String(a) === String(b));

function criarDados(env) {
  const S = env.APP_SCHEMA || 'portal_ambiental_qualidade';
  if (!/^[a-z_][a-z0-9_]*$/.test(S)) throw new Error('APP_SCHEMA inválido: ' + S);
  const pool = new Pool({
    host: env.PGHOST, port: Number(env.PGPORT || 5432), database: env.PGDATABASE,
    user: env.PGUSER, password: env.PGPASSWORD || undefined, max: Number(env.DB_POOL_MAX || 4),
  });
  pool.on('error', e => console.error('[db] ' + e.message));
  const q = (sql, p) => pool.query(sql, p);
  const quando = c => `to_char(${c} AT TIME ZONE 'America/Sao_Paulo', 'DD/MM/YYYY HH24:MI')`;

  async function emTransacao(fn) {
    const cli = await pool.connect();
    try {
      await cli.query('BEGIN');
      const r = await fn(cli);
      await cli.query('COMMIT');
      return r;
    } catch (e) {
      await cli.query('ROLLBACK').catch(() => {});
      throw e;
    } finally {
      cli.release();
    }
  }

  async function auditar(login, acao, detalhe) {
    try { await q(`INSERT INTO ${S}.auditoria (login, acao, detalhe) VALUES ($1, $2, $3)`, [login, acao, detalhe || null]); }
    catch (e) { console.error('[auditoria] ' + e.message); }
  }

  // ── Cadastros ────────────────────────────────────────────────────────────
  async function cadastros() {
    const [obras, tipos, ucs, assuntos, metas, parametros, fechamentos] = await Promise.all([
      q(`SELECT id, nome, ativa, area_total_m2, meta_residuo_kg_m2, observacao FROM ${S}.obra ORDER BY nome`),
      q(`SELECT id, nome, classe, densidade_kg_m3, perigoso, ativo, legado, ordem FROM ${S}.tipo_residuo ORDER BY legado, classe, ordem, nome`),
      q(`SELECT id, nome, tipo, obra_id, agrupa_em_id, ativa, ordem FROM ${S}.unidade_consumidora ORDER BY tipo DESC, ordem, nome`),
      q(`SELECT id, tipo, nome, ativo, ordem FROM ${S}.assunto_nc ORDER BY (tipo = 'LAI/EIV') DESC, tipo, ordem, nome`),
      q(`SELECT obra_id, vigente_desde, meta FROM ${S}.meta_conformidade ORDER BY obra_id, vigente_desde`),
      q(`SELECT chave, valor, descricao FROM ${S}.parametro ORDER BY chave`),
      q(`SELECT competencia, fechado_por, ${quando('fechado_em')} AS fechado_em FROM ${S}.fechamento ORDER BY competencia DESC`),
    ]);
    return {
      obras: obras.rows, tipos: tipos.rows, ucs: ucs.rows, assuntos: assuntos.rows,
      metas: metas.rows, parametros: Object.fromEntries(parametros.rows.map(p => [p.chave, p])),
      fechamentos: fechamentos.rows,
    };
  }

  // nomes para a chave legível do histórico
  async function nomes(cli) {
    const r = await cli.query(`SELECT 'obra' AS t, id, nome FROM ${S}.obra UNION ALL
      SELECT 'tipo', id, nome FROM ${S}.tipo_residuo UNION ALL
      SELECT 'uc', id, nome FROM ${S}.unidade_consumidora UNION ALL
      SELECT 'assunto', id, tipo || ': ' || nome FROM ${S}.assunto_nc`);
    const m = { obra: {}, tipo: {}, uc: {}, assunto: {} };
    r.rows.forEach(x => { m[x.t][x.id] = x.nome; });
    return m;
  }
  function chaveLegivel(n, k) {
    const c = {};
    if (k.obra_id !== undefined) c.obra = n.obra[k.obra_id] || k.obra_id;
    if (k.competencia) c.competencia = k.competencia.slice(0, 7);
    if (k.tipo_residuo_id !== undefined) c.tipo = n.tipo[k.tipo_residuo_id] || k.tipo_residuo_id;
    if (k.uc_id !== undefined) c.unidade = n.uc[k.uc_id] || k.uc_id;
    if (k.assunto_id !== undefined) c.assunto = n.assunto[k.assunto_id] || k.assunto_id;
    if (k.id !== undefined) c.id = k.id;
    return c;
  }
  async function logar(cli, login, tabela, chave, acao, antes, depois, motivo) {
    await cli.query(`INSERT INTO ${S}.log_alteracao (login, tabela, chave, acao, antes, depois, motivo) VALUES ($1,$2,$3,$4,$5,$6,$7)`,
      [login, tabela, chave, acao, antes ? JSON.stringify(antes) : null, depois ? JSON.stringify(depois) : null, motivo || null]);
  }

  // o que uma pessoa excluiu de propósito (linha do Prevision ou chamado do Salesforce) não volta na sincronização
  async function marcarExcluida(cli, fonte, chave, login) {
    await cli.query(`INSERT INTO ${S}.integracao_excluida (fonte, chave, excluido_por) VALUES ($1, $2, $3) ON CONFLICT DO NOTHING`, [fonte, chave, login]);
  }

  async function mesFechado(cli, comp) {
    const r = await cli.query(`SELECT 1 FROM ${S}.fechamento WHERE competencia = $1`, [comp]);
    return r.rowCount > 0;
  }
  // editor lança em mês aberto; mês fechado só admin, com motivo
  // comp nulo = registro sem mês (Plano de Ação, "Todos" das Ações SGA): só o nível conta
  async function conferirTrava(cli, comp, sessao, motivo) {
    if (!['editor', 'admin'].includes(sessao.nivel)) throw erroUso('seu acesso é só de consulta', 403);
    if (comp && await mesFechado(cli, comp)) {
      if (sessao.nivel !== 'admin') throw erroUso('mês fechado: só um administrador pode alterar', 403);
      if (!motivo || !String(motivo).trim()) throw erroUso('mês fechado: informe o motivo da correção');
    }
  }

  // ── Tudo de uma competência (alimenta todas as telas de lançamento) ──────
  async function mes(compBruta) {
    const comp = competencia(compBruta);
    const p = [comp];
    const [avanco, conf, ncs, ruido, res, cons, fech, anterior] = await Promise.all([
      q(`SELECT v.obra_id, v.avanco_pct, v.area_total_m2, v.area_mes_m2, a.origem
           FROM ${S}.v_area_mes v LEFT JOIN ${S}.avanco_fisico a ON a.obra_id = v.obra_id AND a.competencia = v.competencia
          WHERE v.competencia = $1`, p),
      q(`SELECT c.obra_id, c.itens_conformes, c.itens_avaliados, c.percentual, c.observacao, c.origem,
                c.atualizado_por, ${quando('c.atualizado_em')} AS atualizado_em
           FROM ${S}.conformidade c WHERE c.competencia = $1`, p),
      q(`SELECT obra_id, assunto_id FROM ${S}.conformidade_nc WHERE competencia = $1`, p),
      q(`SELECT id, obra_id, ponto, to_char(horario, 'HH24:MI') AS horario, db, empresa, origem
           FROM ${S}.medicao_ruido WHERE competencia = $1 ORDER BY obra_id, horario NULLS LAST, id`, p),
      q(`SELECT r.obra_id, r.tipo_residuo_id, r.volume_m3, r.origem, r.atualizado_por, ${quando('r.atualizado_em')} AS atualizado_em
           FROM ${S}.residuo r WHERE r.competencia = $1`, p),
      q(`SELECT uc_id, linha, luz_kwh, agua_m3, origem, atualizado_por, ${quando('atualizado_em')} AS atualizado_em
           FROM ${S}.consumo WHERE competencia = $1 ORDER BY uc_id, linha`, p),
      q(`SELECT fechado_por, ${quando('fechado_em')} AS fechado_em FROM ${S}.fechamento WHERE competencia = $1`, p),
      // mês anterior, para o "igual ao mês passado?" e para comparar
      q(`SELECT 'consumo' AS f, uc_id AS id, luz_kwh AS a, agua_m3 AS b FROM ${S}.consumo WHERE linha = 1 AND competencia = ($1::date - interval '1 month')::date
         UNION ALL SELECT 'avanco', obra_id, avanco_pct, NULL FROM ${S}.avanco_fisico WHERE competencia = ($1::date - interval '1 month')::date`, p),
    ]);
    // meta de conformidade vigente para cada obra nesta competência
    const metas = await q(`SELECT DISTINCT ON (obra_id) obra_id, meta FROM ${S}.meta_conformidade
                            WHERE vigente_desde <= $1 ORDER BY obra_id, vigente_desde DESC`, p);
    return {
      competencia: comp,
      fechado: fech.rows[0] || null,
      avanco: avanco.rows, conformidade: conf.rows, ncs: ncs.rows, ruido: ruido.rows,
      residuo: res.rows, consumo: cons.rows,
      metas: Object.fromEntries(metas.rows.map(m => [m.obra_id, m.meta])),
      anterior: anterior.rows,
    };
  }

  // ── Panorama: quantas obras têm cada lançamento, mês a mês ───────────────
  async function panorama(meses) {
    const n = Math.min(Math.max(Number(meses) || 12, 1), 36);
    const r = await q(`
      WITH m AS (SELECT generate_series(date_trunc('month', now())::date - ($1::int - 1) * interval '1 month',
                                        date_trunc('month', now())::date, interval '1 month')::date AS competencia)
      SELECT m.competencia,
        (SELECT count(*) FROM ${S}.avanco_fisico x WHERE x.competencia = m.competencia)::int AS avanco,
        (SELECT count(*) FROM ${S}.conformidade x WHERE x.competencia = m.competencia)::int AS conformidade,
        (SELECT count(DISTINCT obra_id) FROM ${S}.medicao_ruido x WHERE x.competencia = m.competencia)::int AS ruido,
        (SELECT count(DISTINCT obra_id) FROM ${S}.residuo x WHERE x.competencia = m.competencia)::int AS residuo,
        (SELECT count(*) FROM ${S}.consumo x WHERE x.competencia = m.competencia AND x.linha = 1 AND (x.luz_kwh IS NOT NULL OR x.agua_m3 IS NOT NULL))::int AS consumo,
        EXISTS (SELECT 1 FROM ${S}.fechamento f WHERE f.competencia = m.competencia) AS fechado
      FROM m ORDER BY m.competencia DESC`, [n]);
    return r.rows;
  }

  // ── Gravação em grade ────────────────────────────────────────────────────
  async function salvarGrade(fonte, compBruta, linhas, sessao, motivo) {
    const F = FONTES[fonte];
    if (!F) throw erroUso('fonte desconhecida: ' + fonte, 404);
    const comp = competencia(compBruta);
    if (!Array.isArray(linhas) || linhas.length === 0) throw erroUso('nada para gravar');
    if (linhas.length > 500) throw erroUso('linhas demais numa gravação');
    const cols = Object.keys(F.campos);
    // consumo: o sistema só mexe na linha 1 (a 2ª linha de mar/25 da Assistência técnica é histórico da planilha)
    const chaves = [...F.chave, 'competencia', ...(F.tabela === 'consumo' ? ['linha'] : [])];

    return emTransacao(async cli => {
      await conferirTrava(cli, comp, sessao, motivo);
      const n = await nomes(cli);
      const resumo = { incluidas: 0, alteradas: 0, excluidas: 0, iguais: 0 };
      for (const bruta of linhas) {
        const k = { competencia: comp };
        if (F.tabela === 'consumo') k.linha = 1;
        for (const c of F.chave) {
          const v = numero(bruta[c], { inteiro: true });
          if (v === null) throw erroUso('linha sem ' + c);
          k[c] = v;
        }
        const l = {};
        for (const c of cols) l[c] = F.campos[c](bruta[c]);
        if (F.ajustar) F.ajustar(l);
        const vazia = cols.every(c => l[c] === null);
        if (F.soExclusao && !vazia) throw erroUso(F.soExclusao, 403);

        const onde = chaves.map((c, i) => `${c} = $${i + 1}`).join(' AND ');
        const valsChave = chaves.map(c => k[c]);
        const atual = (await cli.query(`SELECT ${cols.join(', ')} FROM ${S}.${F.tabela} WHERE ${onde} FOR UPDATE`, valsChave)).rows[0];
        if (F.validar && !vazia) await F.validar(cli, S, k, atual);
        const legivel = chaveLegivel(n, k);

        if (!atual) {
          if (vazia) continue;
          const todas = [...chaves, ...cols, 'criado_por', 'atualizado_por'];
          const vals = [...valsChave, ...cols.map(c => l[c]), sessao.login, sessao.login];
          await cli.query(`INSERT INTO ${S}.${F.tabela} (${todas.join(', ')}) VALUES (${todas.map((_, i) => '$' + (i + 1)).join(', ')})`, vals);
          await logar(cli, sessao.login, F.tabela, legivel, 'incluir', null, l, motivo);
          resumo.incluidas++;
        } else if (vazia) {
          await cli.query(`DELETE FROM ${S}.${F.tabela} WHERE ${onde}`, valsChave);
          await logar(cli, sessao.login, F.tabela, legivel, 'excluir', atual, null, motivo);
          // linha excluída de propósito não volta na próxima sincronização com o Prevision
          if (F.tabela === 'avanco_fisico') await marcarExcluida(cli, 'prevision', `${k.obra_id}|${comp.slice(0, 7)}`, sessao.login);
          resumo.excluidas++;
        } else {
          const mudou = cols.filter(c => (atual[c] ?? null) !== (l[c] ?? null));
          if (!mudou.length) { resumo.iguais++; continue; }
          const sets = mudou.map((c, i) => `${c} = $${chaves.length + i + 1}`).join(', ');
          await cli.query(`UPDATE ${S}.${F.tabela} SET ${sets}, atualizado_por = $${chaves.length + mudou.length + 1}, atualizado_em = now() WHERE ${onde}`,
            [...valsChave, ...mudou.map(c => l[c]), sessao.login]);
          await logar(cli, sessao.login, F.tabela, legivel, 'alterar',
            Object.fromEntries(mudou.map(c => [c, atual[c]])), Object.fromEntries(mudou.map(c => [c, l[c]])), motivo);
          resumo.alteradas++;
        }
      }
      return resumo;
    });
  }

  // NCs do checklist: o conjunto de assuntos de uma obra no mês
  async function salvarNcs(compBruta, obraId, assuntoIds, sessao, motivo) {
    const comp = competencia(compBruta);
    const obra = numero(obraId, { inteiro: true });
    const novos = new Set((assuntoIds || []).map(a => numero(a, { inteiro: true })));
    return emTransacao(async cli => {
      await conferirTrava(cli, comp, sessao, motivo);
      const n = await nomes(cli);
      const atuais = new Set((await cli.query(`SELECT assunto_id FROM ${S}.conformidade_nc WHERE obra_id = $1 AND competencia = $2`, [obra, comp])).rows.map(r => r.assunto_id));
      let inc = 0, exc = 0;
      for (const a of novos) if (!atuais.has(a)) {
        await cli.query(`INSERT INTO ${S}.conformidade_nc (obra_id, competencia, assunto_id, criado_por) VALUES ($1,$2,$3,$4)`, [obra, comp, a, sessao.login]);
        await logar(cli, sessao.login, 'conformidade_nc', chaveLegivel(n, { obra_id: obra, competencia: comp, assunto_id: a }), 'incluir', null, { assunto: n.assunto[a] }, motivo);
        inc++;
      }
      for (const a of atuais) if (!novos.has(a)) {
        await cli.query(`DELETE FROM ${S}.conformidade_nc WHERE obra_id = $1 AND competencia = $2 AND assunto_id = $3`, [obra, comp, a]);
        await logar(cli, sessao.login, 'conformidade_nc', chaveLegivel(n, { obra_id: obra, competencia: comp, assunto_id: a }), 'excluir', { assunto: n.assunto[a] }, null, motivo);
        exc++;
      }
      return { incluidas: inc, excluidas: exc };
    });
  }

  // Ruído: lista de medições do mês (uma obra pode ter mais de uma).
  // linhas: [{id?, obra_id, ponto, horario, db, empresa, excluir?}]
  async function salvarRuido(compBruta, linhas, sessao, motivo) {
    const comp = competencia(compBruta);
    if (!Array.isArray(linhas)) throw erroUso('nada para gravar');
    return emTransacao(async cli => {
      await conferirTrava(cli, comp, sessao, motivo);
      const n = await nomes(cli);
      const resumo = { incluidas: 0, alteradas: 0, excluidas: 0, iguais: 0 };
      for (const b of linhas) {
        const id = numero(b.id, { inteiro: true });
        const l = {
          obra_id: numero(b.obra_id, { inteiro: true }),
          ponto: texto(20)(b.ponto) || 'P1',
          horario: hora(b.horario),
          db: numero(b.db, { min: 0.1, max: 199 }),
          empresa: texto(60)(b.empresa),
        };
        if (!l.obra_id) throw erroUso('medição sem obra');
        const atual = id ? (await cli.query(`SELECT obra_id, ponto, to_char(horario,'HH24:MI') AS horario, db, empresa FROM ${S}.medicao_ruido WHERE id = $1 AND competencia = $2 FOR UPDATE`, [id, comp])).rows[0] : null;
        if (id && !atual) throw erroUso('medição não encontrada: ' + id);
        const chave = chaveLegivel(n, { obra_id: l.obra_id, competencia: comp, id: id || undefined });
        if (b.excluir || l.db === null) {
          if (!atual) continue;
          await cli.query(`DELETE FROM ${S}.medicao_ruido WHERE id = $1`, [id]);
          await logar(cli, sessao.login, 'medicao_ruido', chave, 'excluir', atual, null, motivo);
          resumo.excluidas++;
        } else if (!atual) {
          const r = await cli.query(`INSERT INTO ${S}.medicao_ruido (obra_id, competencia, ponto, horario, db, empresa, criado_por, atualizado_por)
            VALUES ($1,$2,$3,$4,$5,$6,$7,$7) RETURNING id`, [l.obra_id, comp, l.ponto, l.horario, l.db, l.empresa, sessao.login]);
          chave.id = r.rows[0].id;
          await logar(cli, sessao.login, 'medicao_ruido', chave, 'incluir', null, l, motivo);
          resumo.incluidas++;
        } else {
          const mudou = Object.keys(l).filter(c => (atual[c] ?? null) !== (l[c] ?? null));
          if (!mudou.length) { resumo.iguais++; continue; }
          await cli.query(`UPDATE ${S}.medicao_ruido SET obra_id=$2, ponto=$3, horario=$4, db=$5, empresa=$6, atualizado_por=$7, atualizado_em=now() WHERE id=$1`,
            [id, l.obra_id, l.ponto, l.horario, l.db, l.empresa, sessao.login]);
          await logar(cli, sessao.login, 'medicao_ruido', chave, 'alterar',
            Object.fromEntries(mudou.map(c => [c, atual[c]])), Object.fromEntries(mudou.map(c => [c, l[c]])), motivo);
          resumo.alteradas++;
        }
      }
      return resumo;
    });
  }

  // ── Cadastros (admin) ────────────────────────────────────────────────────
  async function salvarCadastro(entidade, item, sessao) {
    const C = CADASTROS[entidade];
    if (!C) throw erroUso('cadastro desconhecido: ' + entidade, 404);
    const cols = Object.keys(C.campos);
    const l = {};
    for (const c of cols) l[c] = C.campos[c](item[c]);
    for (const c of C.obrigatorios) if (l[c] === null || l[c] === undefined) throw erroUso('campo obrigatório: ' + c);
    const id = numero(item.id, { inteiro: true });
    return emTransacao(async cli => {
      try {
        if (!id) {
          const r = await cli.query(`INSERT INTO ${S}.${C.tabela} (${cols.join(', ')}${C.tabela === 'assunto_nc' ? '' : ', atualizado_por'})
            VALUES (${cols.map((_, i) => '$' + (i + 1)).join(', ')}${C.tabela === 'assunto_nc' ? '' : ', $' + (cols.length + 1)}) RETURNING id`,
            C.tabela === 'assunto_nc' ? cols.map(c => l[c]) : [...cols.map(c => l[c]), sessao.login]);
          await logar(cli, sessao.login, C.tabela, { id: r.rows[0].id, nome: l.nome }, 'incluir', null, l);
          return { id: r.rows[0].id };
        }
        const atual = (await cli.query(`SELECT ${cols.join(', ')} FROM ${S}.${C.tabela} WHERE id = $1 FOR UPDATE`, [id])).rows[0];
        if (!atual) throw erroUso('registro não encontrado', 404);
        const mudou = cols.filter(c => (atual[c] ?? null) !== (l[c] ?? null));
        if (!mudou.length) return { id, iguais: true };
        const extra = C.tabela === 'assunto_nc' ? '' : `, atualizado_por = $${mudou.length + 2}, atualizado_em = now()`;
        await cli.query(`UPDATE ${S}.${C.tabela} SET ${mudou.map((c, i) => `${c} = $${i + 2}`).join(', ')}${extra} WHERE id = $1`,
          C.tabela === 'assunto_nc' ? [id, ...mudou.map(c => l[c])] : [id, ...mudou.map(c => l[c]), sessao.login]);
        await logar(cli, sessao.login, C.tabela, { id, nome: l.nome || atual.nome }, 'alterar',
          Object.fromEntries(mudou.map(c => [c, atual[c]])), Object.fromEntries(mudou.map(c => [c, l[c]])));
        return { id };
      } catch (e) {
        if (e.code === '23505') throw erroUso('já existe um cadastro com esse nome');
        throw e;
      }
    });
  }

  // meta de conformidade: {obra_id, vigente_desde: 'AAAA-MM', meta (0-1) | null = remover}
  async function salvarMeta(item, sessao) {
    const obra = numero(item.obra_id, { inteiro: true });
    const desde = competencia(item.vigente_desde);
    const meta = numero(item.meta, { min: 0, max: 1 });
    return emTransacao(async cli => {
      const n = await nomes(cli);
      const chave = chaveLegivel(n, { obra_id: obra, competencia: desde });
      const atual = (await cli.query(`SELECT meta FROM ${S}.meta_conformidade WHERE obra_id = $1 AND vigente_desde = $2`, [obra, desde])).rows[0];
      if (meta === null) {
        if (!atual) return {};
        await cli.query(`DELETE FROM ${S}.meta_conformidade WHERE obra_id = $1 AND vigente_desde = $2`, [obra, desde]);
        await logar(cli, sessao.login, 'meta_conformidade', chave, 'excluir', atual, null);
      } else if (!atual) {
        await cli.query(`INSERT INTO ${S}.meta_conformidade (obra_id, vigente_desde, meta, atualizado_por) VALUES ($1,$2,$3,$4)`, [obra, desde, meta, sessao.login]);
        await logar(cli, sessao.login, 'meta_conformidade', chave, 'incluir', null, { meta });
      } else if (atual.meta !== meta) {
        await cli.query(`UPDATE ${S}.meta_conformidade SET meta = $3, atualizado_por = $4, atualizado_em = now() WHERE obra_id = $1 AND vigente_desde = $2`, [obra, desde, meta, sessao.login]);
        await logar(cli, sessao.login, 'meta_conformidade', chave, 'alterar', atual, { meta });
      }
      return {};
    });
  }

  async function salvarParametro(chave, valor, sessao) {
    const v = numero(valor);
    if (v === null) throw erroUso('valor obrigatório');
    return emTransacao(async cli => {
      const atual = (await cli.query(`SELECT valor FROM ${S}.parametro WHERE chave = $1 FOR UPDATE`, [chave])).rows[0];
      if (!atual) throw erroUso('parâmetro desconhecido', 404);
      if (atual.valor === v) return {};
      await cli.query(`UPDATE ${S}.parametro SET valor = $2, atualizado_por = $3, atualizado_em = now() WHERE chave = $1`, [chave, v, sessao.login]);
      await logar(cli, sessao.login, 'parametro', { chave }, 'alterar', { valor: atual.valor }, { valor: v });
      return {};
    });
  }

  async function fecharMes(compBruta, fechar, sessao) {
    const comp = competencia(compBruta);
    if (fechar) {
      await q(`INSERT INTO ${S}.fechamento (competencia, fechado_por) VALUES ($1, $2) ON CONFLICT DO NOTHING`, [comp, sessao.login]);
    } else {
      await q(`DELETE FROM ${S}.fechamento WHERE competencia = $1`, [comp]);
    }
    await auditar(sessao.login, fechar ? 'mes_fechar' : 'mes_reabrir', { competencia: comp });
  }

  // Cadastros = o que só administrador mexe; todo o resto do log é Lançamentos.
  // origem: 'pessoas' (quem entrou no sistema) ou 'sistema' (sincronização com Salesforce e Prevision).
  const TABELAS_CADASTRO = ['obra', 'tipo_residuo', 'unidade_consumidora', 'assunto_nc', 'parametro'];
  async function historico({ tabela, login, competencia: comp, limite, grupo, origem }) {
    const cond = [], p = [];
    if (tabela) { p.push(tabela); cond.push(`tabela = $${p.length}`); }
    if (login) { p.push(login); cond.push(`login = $${p.length}`); }
    if (comp) { p.push(competencia(comp).slice(0, 7)); cond.push(`chave->>'competencia' = $${p.length}`); }
    if (grupo === 'cadastros' || grupo === 'lancamentos') {
      p.push(TABELAS_CADASTRO); cond.push(`${grupo === 'cadastros' ? '' : 'NOT '}(tabela = ANY($${p.length}))`);
    }
    if (origem === 'sistema') cond.push(`login = 'sistema'`);
    if (origem === 'pessoas') cond.push(`login <> 'sistema'`);
    const onde = cond.length ? 'WHERE ' + cond.join(' AND ') : '';
    // contagem do que foi pedido (todos os registros do filtro, não só os que cabem na tela)
    const pc = [...p, TABELAS_CADASTRO];
    const [itens, resumo] = await Promise.all([
      q(`SELECT id, ${quando('quando')} AS quando, login, tabela, chave, acao, antes, depois, motivo
           FROM ${S}.log_alteracao ${onde} ORDER BY id DESC LIMIT ${Math.min(Math.max(Number(limite) || 300, 1), 2000)}`, p),
      q(`SELECT CASE WHEN tabela = ANY($${pc.length}) THEN 'cadastros' ELSE 'lancamentos' END AS grupo, login, acao, count(*)::int AS n
           FROM ${S}.log_alteracao ${onde} GROUP BY 1, 2, 3`, pc),
    ]);
    return { itens: itens.rows, resumo: resumo.rows };
  }

  // ── Usuários e sessões (padrão do Projetos) ──────────────────────────────
  const hashToken = t => crypto.createHash('sha256').update(String(t)).digest('hex');
  async function sessaoGravar(token, s, validadeMs) {
    await q(`DELETE FROM ${S}.sessao WHERE expira_em < now()`);
    await q(`INSERT INTO ${S}.sessao (token_sha256, login, nome, email, criada_em, expira_em)
             VALUES ($1, $2, $3, $4, to_timestamp($5 / 1000.0), to_timestamp($5 / 1000.0) + $6 * interval '1 millisecond')`,
      [hashToken(token), s.login, s.nome || null, s.email || null, s.criadaEm, validadeMs]);
  }
  async function sessaoLer(token) {
    const { rows } = await q(`SELECT login, nome, email, (extract(epoch FROM criada_em) * 1000)::bigint AS criada
                                FROM ${S}.sessao WHERE token_sha256 = $1 AND expira_em > now()`, [hashToken(token)]);
    const r = rows[0];
    return r ? { login: r.login, nome: r.nome, email: r.email, criadaEm: Number(r.criada) } : null;
  }
  const sessaoApagar = token => q(`DELETE FROM ${S}.sessao WHERE token_sha256 = $1`, [hashToken(token)]);
  const sessoesApagarDe = login => q(`DELETE FROM ${S}.sessao WHERE login = $1`, [login]);

  async function acessoDe(u, inicial) {
    const ini = inicial || { situacao: 'pendente', nivel: 'leitor', por: null };
    const { rows } = await q(
      `INSERT INTO ${S}.usuario AS x (login, nome, email, situacao, nivel, decidido_por, decidido_em)
       VALUES ($1, $2, $3, $4, $5, $6, CASE WHEN $6::text IS NULL THEN NULL ELSE now() END)
       ON CONFLICT (login) DO UPDATE SET nome = coalesce(EXCLUDED.nome, x.nome), email = coalesce(EXCLUDED.email, x.email)
       RETURNING situacao, nivel, (xmax = 0) AS novo`,
      [u.login, u.nome || null, u.email || null, ini.situacao, ini.nivel, ini.por]);
    return rows[0];
  }
  const marcarAcesso = login => q(`UPDATE ${S}.usuario SET ultimo_acesso = now() WHERE login = $1`, [login]);
  async function situacaoDe(login) {
    return (await q(`SELECT situacao, nivel FROM ${S}.usuario WHERE login = $1`, [login])).rows[0] || { situacao: null };
  }
  async function usuarios() {
    return (await q(`SELECT login, nome, email, situacao, nivel, decidido_por, ${quando('pedido_em')} AS pedido_em,
                            ${quando('decidido_em')} AS decidido_em, ${quando('ultimo_acesso')} AS ultimo_acesso
                       FROM ${S}.usuario ORDER BY (situacao = 'pendente') DESC, nome NULLS LAST, login`)).rows;
  }
  // exclui o cadastro: se a pessoa entrar de novo com a conta da rede, volta como pendente
  async function excluirUsuario(login, por) {
    const { rowCount } = await q(`DELETE FROM ${S}.usuario WHERE login = $1`, [login]);
    if (!rowCount) throw erroUso('usuário não encontrado', 404);
    await auditar(por, 'usuario_excluir', { login });
  }
  async function decidirUsuario(login, { situacao, nivel }, por) {
    if (!['aprovado', 'rejeitado'].includes(situacao)) throw erroUso('situação inválida');
    if (!['leitor', 'editor', 'admin'].includes(nivel)) throw erroUso('nível inválido');
    const { rowCount } = await q(`UPDATE ${S}.usuario SET situacao = $2, nivel = $3, decidido_por = $4, decidido_em = now() WHERE login = $1`,
      [login, situacao, nivel, por]);
    if (!rowCount) throw erroUso('usuário não encontrado', 404);
    await auditar(por, 'usuario_decidir', { login, situacao, nivel });
  }

  async function saude() { await q('SELECT 1'); return true; }

  // ── Abas que são listas (Auditoria, Plano de Ação, Emergências, ...) ──────
  function filtroLista(L, f) {
    if (L.filtro === 'mes') { const c = competencia(f.competencia); return { onde: `competencia = $1`, p: [c], fixo: L.datado ? {} : { competencia: c }, comp: c }; }
    if (L.filtro === 'ano') {
      const a = numero(f.ano, { inteiro: true, min: 2000, max: 2100 });
      if (a === null) throw erroUso('informe o ano');
      return L.tabela === 'acoes_sga' ? { onde: 'ano = $1', p: [a], fixo: { ano: a }, ano: a }
        : { onde: 'extract(year FROM competencia) = $1', p: [a], fixo: {}, ano: a };
    }
    return { onde: 'true', p: [], fixo: {} };
  }
  async function listar(aba, f) {
    const L = LISTAS[aba];
    if (!L) throw erroUso('aba desconhecida: ' + aba, 404);
    const F = filtroLista(L, f);
    const extra = L.mesLinha ? ', extract(year FROM competencia)::int AS ano, extract(month FROM competencia)::int AS mes' : '';
    const origem = L.comOrigem ? ', origem' : '';
    const r = await q(`SELECT id, ${L.colunas.join(', ')}${extra}${origem}, atualizado_por, ${quando('atualizado_em')} AS atualizado_em
                         FROM ${S}.${L.tabela} WHERE ${F.onde} ORDER BY ${L.ordem}`, F.p);
    return r.rows;
  }
  async function salvarLista(aba, f, linhas, sessao, motivo) {
    const L = LISTAS[aba];
    if (!L) throw erroUso('aba desconhecida: ' + aba, 404);
    if (!Array.isArray(linhas) || !linhas.length) throw erroUso('nada para gravar');
    if (linhas.length > 2000) throw erroUso('linhas demais numa gravação');
    const F = filtroLista(L, f);
    const cols = Object.keys(L.campos);
    return emTransacao(async cli => {
      const n = await nomes(cli);
      const travas = new Set();
      const trava = async c => { const k = c || '-'; if (!travas.has(k)) { travas.add(k); await conferirTrava(cli, c, sessao, motivo); } };
      const resumo = { incluidas: 0, alteradas: 0, excluidas: 0, iguais: 0 };
      for (const b of linhas) {
        const id = numero(b.id, { inteiro: true });
        const atual = id ? (await cli.query(`SELECT ${[...L.colunas, ...(L.extraLeitura || [])].join(', ')} FROM ${S}.${L.tabela} WHERE id = $1 AND ${F.onde.replace(/\$1/g, '$2')} FOR UPDATE`, [id, ...F.p])).rows[0] : null;
        if (id && !atual) throw erroUso('linha não encontrada (recarregue a tela): ' + id);
        if (b.excluir) {
          if (!atual) continue;
          await trava(compDe(L, atual, F));
          await cli.query(`DELETE FROM ${S}.${L.tabela} WHERE id = $1`, [id]);
          const { sf_case_id: caso, ...visivel } = atual; // o vínculo com o Salesforce não é campo da tela
          await logar(cli, sessao.login, L.tabela, chaveLegivel(n, chaveLinha(L, atual, F, id)), 'excluir', visivel, null, motivo);
          if (caso) await marcarExcluida(cli, 'salesforce', caso, sessao.login);
          resumo.excluidas++;
          continue;
        }
        if (L.soExclusao) throw erroUso(L.soExclusao, 403);
        const l = { ...F.fixo };
        for (const c of cols) l[c] = L.campos[c](b[c]);
        if (L.preparar) L.preparar(l, F);
        for (const c of L.obrigatorios || []) if (l[c] === null || l[c] === undefined || l[c] === '') throw erroUso(`preencha ${L.rotulos[c] || c} em todas as linhas`);
        await trava(compDe(L, l, F));
        // o mesmo nome (setor, ou pessoa + conta) não pode aparecer em duas linhas do mesmo mês
        if (L.unico && l.competencia) {
          const cond = L.unico.map((c, i) => `lower(btrim(coalesce(${c}::text, ''))) = lower(btrim($${i + 3}::text))`).join(' AND ');
          const dup = await cli.query(`SELECT 1 FROM ${S}.${L.tabela} WHERE competencia = $1 AND id <> $2 AND ${cond} LIMIT 1`, [l.competencia, id || 0, ...L.unico.map(c => l[c] ?? '')]);
          if (dup.rowCount) throw erroUso(`"${l[L.unico[0]]}"${L.unico.length > 1 && l[L.unico[1]] ? ' (' + l[L.unico[1]] + ')' : ''} já tem uma linha em ${String(l.competencia).slice(5, 7)}/${String(l.competencia).slice(0, 4)}. Some as páginas na mesma linha ou use outro nome.`);
        }
        const gravar = L.colunas.filter(c => c in l);
        if (!atual) {
          const r = await cli.query(`INSERT INTO ${S}.${L.tabela} (${gravar.join(', ')}, criado_por, atualizado_por)
            VALUES (${gravar.map((_, i) => '$' + (i + 1)).join(', ')}, $${gravar.length + 1}, $${gravar.length + 1}) RETURNING id`, [...gravar.map(c => l[c]), sessao.login]);
          await logar(cli, sessao.login, L.tabela, chaveLegivel(n, chaveLinha(L, l, F, r.rows[0].id)), 'incluir', null, l, motivo);
          resumo.incluidas++;
        } else {
          const mudou = gravar.filter(c => !iguais(atual[c], l[c]));
          if (!mudou.length) { resumo.iguais++; continue; }
          await cli.query(`UPDATE ${S}.${L.tabela} SET ${mudou.map((c, i) => `${c} = $${i + 2}`).join(', ')}, atualizado_por = $${mudou.length + 2}, atualizado_em = now() WHERE id = $1`,
            [id, ...mudou.map(c => l[c]), sessao.login]);
          await logar(cli, sessao.login, L.tabela, chaveLegivel(n, chaveLinha(L, l, F, id)), 'alterar',
            Object.fromEntries(mudou.map(c => [c, atual[c]])), Object.fromEntries(mudou.map(c => [c, l[c]])), motivo);
          resumo.alteradas++;
        }
      }
      return resumo;
    });
  }

  // ── Relatório de impressão (arquivo exportado) → Impressões, por setor, por pessoa ──
  // rel = saída de carga/ler_relatorio_impressao.py. Sem `aplicar` só mostra o que entraria.
  // Com `aplicar`, o mês é SUBSTITUÍDO pelo relatório (subir de novo no fim do mês troca o parcial).
  // "Impressões" é uma linha por mês: só Total de pág e % Colorido vêm do relatório; Pág Dupla e as
  // economias (que o relatório não tem) ficam como estão.
  async function relatorioImpressao(rel, { aplicar, arquivo }, sessao, motivo) {
    const comp = competencia(rel.competencia);
    const soma = (a, c) => a.reduce((s, x) => s + (x[c] || 0), 0);
    const alvo = rel.tipo === 'conta' ? 'impressoes_setor' : 'impressoes_pessoa';
    const linhas = rel.tipo === 'conta' ? rel.setores : rel.pessoas;
    const pctCor = rel.total && rel.total.paginas ? Math.round(rel.total.cor / rel.total.paginas * 1e6) / 1e6 : null;
    const hoje = async cli => {
      const t = await cli.query(`SELECT count(*)::int AS n, coalesce(sum(paginas), 0)::int AS paginas, string_agg(DISTINCT origem, ', ') AS origens
                                   FROM ${S}.${alvo} WHERE competencia = $1`, [comp]);
      const m = rel.tipo === 'conta' ? (await cli.query(`SELECT total_pag, pct_colorido, origem FROM ${S}.impressoes WHERE competencia = $1`, [comp])).rows[0] || null : null;
      return { tabela: t.rows[0], mensal: m, fechado: await mesFechado(cli, comp) };
    };
    if (!aplicar) {
      const cli = await pool.connect();
      try {
        return { competencia: rel.competencia, tipo: rel.tipo, periodo: rel.periodo, parcial: rel.parcial, total: rel.total, pct_colorido: pctCor,
          avisos: rel.avisos, ignorados: rel.ignorados, linhas, hoje: await hoje(cli) };
      } finally { cli.release(); }
    }
    return emTransacao(async cli => {
      await conferirTrava(cli, comp, sessao, motivo);
      const por = `Importado do relatório ${arquivo || ''} (${rel.periodo.de} a ${rel.periodo.ate})`.replace('  ', ' ');
      const chave = { competencia: comp.slice(0, 7) };
      const antes = (await hoje(cli)).tabela;
      if (antes.n) {
        await cli.query(`DELETE FROM ${S}.${alvo} WHERE competencia = $1`, [comp]);
        await logar(cli, sessao.login, alvo, chave, 'excluir', { linhas: antes.n, paginas: antes.paginas }, null, motivo || 'Substituído pelo relatório importado');
      }
      const cols = rel.tipo === 'conta' ? ['nome'] : ['usuario', 'conta_atual'];
      for (const x of linhas) {
        const vals = [comp, ...cols.map(c => x[c] ?? null), x.pb, x.cor, x.paginas, x.pct];
        await cli.query(`INSERT INTO ${S}.${alvo} (competencia, ${cols.join(', ')}, pb, cor, paginas, pct, origem, criado_por, atualizado_por)
                         VALUES (${vals.map((_, i) => '$' + (i + 1)).join(', ')}, 'relatorio', $${vals.length + 1}, $${vals.length + 1})`, [...vals, sessao.login]);
      }
      await logar(cli, sessao.login, alvo, chave, 'incluir', null, { linhas: linhas.length, paginas: soma(linhas, 'paginas'), arquivo: arquivo || null, periodo: rel.periodo }, motivo || por);
      const resumo = { [alvo]: { removidas: antes.n, incluidas: linhas.length } };
      if (rel.tipo === 'conta') {
        const atual = (await cli.query(`SELECT total_pag, pct_colorido FROM ${S}.impressoes WHERE competencia = $1 FOR UPDATE`, [comp])).rows[0];
        const novo = { total_pag: rel.total.paginas, pct_colorido: pctCor };
        if (!atual) {
          await cli.query(`INSERT INTO ${S}.impressoes (competencia, total_pag, pct_colorido, origem, criado_por, atualizado_por) VALUES ($1, $2, $3, 'relatorio', $4, $4)`,
            [comp, novo.total_pag, novo.pct_colorido, sessao.login]);
          await logar(cli, sessao.login, 'impressoes', chave, 'incluir', null, novo, motivo || por);
          resumo.impressoes = { incluida: true };
        } else {
          const mudou = Object.keys(novo).filter(c => !iguais(atual[c], novo[c]));
          if (mudou.length) {
            await cli.query(`UPDATE ${S}.impressoes SET total_pag = $2, pct_colorido = $3, origem = 'relatorio', atualizado_por = $4, atualizado_em = now() WHERE competencia = $1`,
              [comp, novo.total_pag, novo.pct_colorido, sessao.login]);
            await logar(cli, sessao.login, 'impressoes', chave, 'alterar', Object.fromEntries(mudou.map(c => [c, atual[c]])), Object.fromEntries(mudou.map(c => [c, novo[c]])), motivo || por);
          }
          resumo.impressoes = { alterada: mudou.length > 0 };
        }
      }
      return resumo;
    });
  }

  // ── Salesforce: chamados de vizinhos lidos do DW corporativo ─────────────
  // Só leitura (sessão read-only). Credencial em DWC_PG* no .env. Vizinho =
  // Tipo_de_Cliente__c 'Comunidade'; a obra vem da conta (AccountId), porque
  // Empreendimento__c chega vazio. silver guarda uma linha por carga (LOAD_TS):
  // fica a mais recente de cada chamado.
  let dw = null;
  function dwPool() {
    if (!env.DWC_PGHOST || !env.DWC_PGUSER) throw erroUso('DW não configurado: preencha DWC_PGHOST, DWC_PGPORT, DWC_PGDATABASE, DWC_PGUSER e DWC_PGPASSWORD no .env', 503);
    if (!dw) {
      dw = new Pool({
        host: env.DWC_PGHOST, port: Number(env.DWC_PGPORT || 5432), database: env.DWC_PGDATABASE || 'dw_fg',
        user: env.DWC_PGUSER, password: env.DWC_PGPASSWORD || undefined, max: 2, connectionTimeoutMillis: 10000,
        options: '-c default_transaction_read_only=on -c statement_timeout=30000',
      });
      dw.on('error', e => console.error('[dw] ' + e.message));
    }
    return dw;
  }
  const sfChamados = (onde = 'true') => `
    WITH c AS (SELECT DISTINCT ON ("Id") * FROM silver.salesforce_case ORDER BY "Id", "LOAD_TS" DESC NULLS LAST),
         a AS (SELECT DISTINCT ON ("Id") "Id", "Name" FROM silver.salesforce_account ORDER BY "Id", "LOAD_TS" DESC NULLS LAST)
    SELECT c."Id" AS case_id, ltrim(c."CaseNumber", '0') AS numero,
           to_char(c."CreatedDate"::timestamptz AT TIME ZONE 'America/Sao_Paulo', 'YYYY-MM-DD') AS data,
           c."Tipo__c" AS tipo, c."Motivo__c" AS motivo, c."Status" AS status, c."Subject" AS assunto,
           c."AccountId" AS conta_id, a."Name" AS conta
      FROM c LEFT JOIN a ON a."Id" = c."AccountId"
     WHERE c."Tipo_de_Cliente__c" = 'Comunidade' AND lower(coalesce(c."IsDeleted"::text, 'false')) NOT IN ('true', '1')
       AND ${onde}
     ORDER BY c."CreatedDate"::timestamptz DESC`;

  // obra cujo nome aparece no nome da conta ou no assunto ("Vizinho Garden Park")
  function sugerirObra(obras, ...textos) {
    const t = textos.filter(Boolean).join(' ').toLowerCase();
    const achadas = obras.filter(o => o.ativa && t.includes(o.nome.toLowerCase()));
    achadas.sort((x, y) => y.nome.length - x.nome.length);
    return achadas.length ? achadas[0].id : null;
  }

  async function salesforcePrevia() {
    const pool = dwPool();
    const [sf, carga, contas, recl, obras] = await Promise.all([
      pool.query(sfChamados()),
      pool.query(`SELECT to_char(max("LOAD_TS"), 'DD/MM/YYYY HH24:MI') AS em FROM silver.salesforce_case`),
      q(`SELECT account_id, obra_id FROM ${S}.salesforce_conta`),
      q(`SELECT id, sf_case_id, cod_chamado, obra_id, to_char(data, 'YYYY-MM-DD') AS data, ocorrencia, categoria FROM ${S}.reclamacao`),
      q(`SELECT id, nome, ativa FROM ${S}.obra`),
    ]);
    const conta = new Map(contas.rows.map(c => [c.account_id, c.obra_id]));
    const porCase = new Map(recl.rows.filter(r => r.sf_case_id).map(r => [r.sf_case_id, r]));
    const porCod = new Map(recl.rows.filter(r => !r.sf_case_id && r.cod_chamado).map(r => [String(r.cod_chamado).trim().replace(/^0+/, ''), r]));
    const chamados = sf.rows.map(c => {
      const ja = porCase.get(c.case_id), lista = ja ? null : porCod.get(c.numero);
      const lembrada = conta.get(c.conta_id);
      const obra_id = ja ? ja.obra_id : lista ? lista.obra_id : lembrada ?? sugerirObra(obras.rows, c.conta, c.assunto);
      // linha da planilha que tem o mesmo Cod. Chamado (ainda sem vínculo ou já ligada): só vale se mês e obra baterem; senão, decisão em Integrações.
      // Já ligada: "usar o Salesforce" corrige só mês/obra e mantém ocorrência e categoria da planilha (o assunto do chamado é só um título).
      let diferenca = null;
      const base = ja || lista;
      if (base) {
        const campos = [];
        if (String(base.data || '').slice(0, 7) !== c.data.slice(0, 7)) campos.push('mês');
        if (lembrada && lembrada !== base.obra_id) campos.push('obra');
        if (campos.length) diferenca = { campos, registro_id: base.id,
          planilha: { data: base.data, obra_id: base.obra_id, ocorrencia: base.ocorrencia },
          sistema: { data: c.data, obra_id: lembrada ?? base.obra_id, ocorrencia: ja ? base.ocorrencia : c.assunto, categoria: ja ? base.categoria : c.tipo } };
      }
      return { ...c, obra_id, obra_origem: ja || lista ? 'lista' : lembrada ? 'conta' : obra_id ? 'sugerida' : null,
        situacao: ja ? 'importado' : lista ? 'na_lista' : 'novo', reclamacao_id: (ja || lista || {}).id || null, diferenca };
    });
    return { carga: carga.rows[0].em, chamados };
  }

  // itens: [{case_id, obra_id}]. O conteúdo do chamado é relido do DW (não vem da tela).
  async function salesforceImportar(itens, sessao, motivo) {
    if (!['editor', 'admin'].includes(sessao.nivel)) throw erroUso('seu acesso é só de consulta', 403);
    if (!Array.isArray(itens) || !itens.length) throw erroUso('nenhum chamado selecionado');
    if (itens.length > 500) throw erroUso('chamados demais numa importação');
    const escolha = new Map();
    for (const it of itens) {
      const obra = numero(it.obra_id, { inteiro: true, min: 1 });
      if (!it.case_id || typeof it.case_id !== 'string') throw erroUso('chamado inválido');
      if (!obra) throw erroUso('escolha a obra de todos os chamados selecionados');
      escolha.set(it.case_id, obra);
    }
    const sf = (await dwPool().query(sfChamados('c."Id" = ANY($1)'), [[...escolha.keys()]])).rows;
    if (sf.length !== escolha.size) throw erroUso('chamado não encontrado no DW (ou não é da Comunidade): recarregue a busca');
    return emTransacao(async cli => {
      const n = await nomes(cli);
      const travas = new Set();
      const resumo = { incluidas: 0, vinculadas: 0, iguais: 0 };
      for (const c of sf) {
        const obra_id = escolha.get(c.case_id);
        if (c.conta_id) {
          await cli.query(`INSERT INTO ${S}.salesforce_conta (account_id, nome, obra_id, atualizado_por) VALUES ($1, $2, $3, $4)
            ON CONFLICT (account_id) DO UPDATE SET nome = EXCLUDED.nome, obra_id = EXCLUDED.obra_id, atualizado_por = EXCLUDED.atualizado_por, atualizado_em = now()
            WHERE ${S}.salesforce_conta.obra_id IS DISTINCT FROM EXCLUDED.obra_id`, [c.conta_id, c.conta, obra_id, sessao.login]);
        }
        if ((await cli.query(`SELECT 1 FROM ${S}.reclamacao WHERE sf_case_id = $1`, [c.case_id])).rowCount) { resumo.iguais++; continue; }
        const comp = c.data.slice(0, 7) + '-01';
        if (!travas.has(comp)) { travas.add(comp); await conferirTrava(cli, comp, sessao, motivo); }
        const lista = (await cli.query(`SELECT id, obra_id FROM ${S}.reclamacao WHERE sf_case_id IS NULL AND ltrim(trim(cod_chamado), '0') = $1 ORDER BY id LIMIT 1 FOR UPDATE`, [c.numero])).rows[0];
        if (lista) {
          await cli.query(`UPDATE ${S}.reclamacao SET sf_case_id = $2, atualizado_por = $3, atualizado_em = now() WHERE id = $1`, [lista.id, c.case_id, sessao.login]);
          await logar(cli, sessao.login, 'reclamacao', chaveLegivel(n, { id: lista.id, obra_id: lista.obra_id, competencia: comp }), 'alterar', { sf_case_id: null }, { sf_case_id: c.case_id }, motivo);
          resumo.vinculadas++;
          continue;
        }
        const l = { obra_id, canal: 'Relacionamento', data: c.data, fonte: 'Salesforce', cod_chamado: c.numero, ocorrencia: c.assunto, categoria: c.tipo, sf_case_id: c.case_id };
        const cols = Object.keys(l);
        const r = await cli.query(`INSERT INTO ${S}.reclamacao (${cols.join(', ')}, origem, criado_por, atualizado_por)
          VALUES (${cols.map((_, i) => '$' + (i + 1)).join(', ')}, 'salesforce', $${cols.length + 1}, $${cols.length + 1}) RETURNING id`, [...cols.map(k => l[k]), sessao.login]);
        await logar(cli, sessao.login, 'reclamacao', chaveLegivel(n, { id: r.rows[0].id, obra_id, competencia: comp }), 'incluir', null, l, motivo);
        resumo.incluidas++;
      }
      return resumo;
    });
  }

  // ── Sincronização automática: Salesforce e Prevision (via DW) ─────────────
  // Roda 3x ao dia (server.js) e sob demanda. Regra comum: só entra o que tem obra
  // conhecida e mês aberto; o resto vira pendência para uma pessoa resolver.
  const SISTEMA = { login: 'sistema', nome: 'Sincronização automática', nivel: 'admin' };

  // Reclamações: o mesmo critério da tela (Reclamação, ainda fora da lista), mas só com
  // obra já conhecida (conta lembrada ou linha da lista). Conta nova / mês fechado = pendência.
  async function salesforceAutomatico() {
    const { chamados } = await salesforcePrevia();
    const fechados = new Set((await q(`SELECT to_char(competencia, 'YYYY-MM-DD') AS c FROM ${S}.fechamento`)).rows.map(r => r.c));
    const excluidas = new Set((await q(`SELECT chave FROM ${S}.integracao_excluida WHERE fonte = 'salesforce'`)).rows.map(r => r.chave));
    const auto = [], pendentes = [], achadas = [];
    let ignoradas = 0;
    for (const c of chamados) {
      if (excluidas.has(c.case_id)) { if (c.situacao !== 'importado') ignoradas++; continue; } // alguém excluiu de propósito
      if (c.situacao === 'importado') {
        if (c.diferenca) achadas.push({ chave: c.case_id, titulo: `Chamado ${c.numero} · ${c.conta || c.assunto || ''}`.trim(), planilha: c.diferenca.planilha, sistema: c.diferenca.sistema, registro_id: c.diferenca.registro_id });
        continue;
      }
      if (c.situacao !== 'na_lista' && !/^reclama/i.test(c.tipo || '')) continue; // dúvidas/solicitações não entram
      if (c.diferenca) {
        achadas.push({ chave: c.case_id, titulo: `Chamado ${c.numero} · ${c.conta || c.assunto || ''}`.trim(), planilha: c.diferenca.planilha, sistema: c.diferenca.sistema, registro_id: c.diferenca.registro_id });
        continue;
      }
      // obra vem da conta lembrada, da linha da lista ou do nome da obra no nome da conta/assunto (sem confirmação manual)
      const conhecida = !!c.obra_id;
      const fechado = fechados.has(c.data.slice(0, 7) + '-01');
      if (conhecida && !fechado) auto.push({ case_id: c.case_id, obra_id: c.obra_id });
      else pendentes.push({ numero: c.numero, data: c.data, conta_id: c.conta_id, conta: c.conta, assunto: c.assunto, categoria: c.tipo, status: c.status, sugestao: null,
        motivo: conhecida ? 'mês fechado' : 'conta sem obra' });
    }
    const resumo = auto.length ? await salesforceImportar(auto, SISTEMA, 'Importação automática (Salesforce)') : { incluidas: 0, vinculadas: 0, iguais: 0 };
    const diferencas = await registrarDiferencas('salesforce', achadas);
    return { ...resumo, ignoradas, diferencas, pendentes: pendentes.slice(0, 200), total_pendentes: pendentes.length };
  }

  // Avanço físico: Prevision, curva MONETÁRIA, foto mais recente: realizado no fim do mês
  // menos o do mês anterior (reproduz a planilha). Só meses completos, últimos 12.
  // Mês já lançado à mão que difere do Prevision NÃO é sobrescrito: vira divergência.
  async function previsionAutomatico() {
    const pool = dwPool();
    const lista = (await pool.query(`SELECT DISTINCT ON ("ProjectId") "ProjectId" AS id, "ProjectName" AS nome, "Phase" AS fase
      FROM silver.prevision_lista_de_projetos ORDER BY "ProjectId", "Load_Timestamp" DESC`)).rows;
    const [maps, obras, fech, excl] = await Promise.all([
      q(`SELECT project_id, obra_id, ignorar FROM ${S}.prevision_projeto`),
      q(`SELECT id, nome, ativa FROM ${S}.obra`),
      q(`SELECT to_char(competencia, 'YYYY-MM-DD') AS c FROM ${S}.fechamento`),
      q(`SELECT chave FROM ${S}.integracao_excluida WHERE fonte = 'prevision'`),
    ]);
    const mapa = new Map(maps.rows.map(m => [m.project_id, m]));
    const fechados = new Set(fech.rows.map(r => r.c));
    const excluidas = new Set(excl.rows.map(r => r.chave));
    const semObra = [], alvo = new Map(); // project_id -> {obra_id, nome}
    for (const p of lista) {
      const m = mapa.get(p.id);
      if (m && m.ignorar) continue;
      if (m && m.obra_id) alvo.set(p.id, { obra_id: m.obra_id, nome: p.nome });
      else if (p.fase === 'construction') semObra.push({ project_id: p.id, nome: p.nome, sugestao: sugerirObra(obras.rows, p.nome) });
    }
    const hoje = new Date(Date.now() - 3 * 3600 * 1000); // data em Brasília
    const fim = (a, m) => new Date(Date.UTC(a, m + 1, 0)).toISOString().slice(0, 10); // último dia do mês m (0-11)
    const meses = []; // 12 meses completos, do mais antigo
    for (let k = 12; k >= 1; k--) { const d = new Date(Date.UTC(hoje.getUTCFullYear(), hoje.getUTCMonth() - k, 1)); meses.push({ comp: d.toISOString().slice(0, 10), fim: fim(d.getUTCFullYear(), d.getUTCMonth()) }); }
    const dm = new Date(Date.UTC(hoje.getUTCFullYear(), hoje.getUTCMonth() - 13, 1));
    const datas = [fim(dm.getUTCFullYear(), dm.getUTCMonth()), ...meses.map(m => m.fim)];
    // divergentes = valor fora de 0–100% no Prevision (corrigir lá); diferencas = planilha × Prevision (decididas em Integrações)
    const out = { incluidas: 0, atualizadas: 0, iguais: 0, fechadas: 0, excluidas: 0, divergentes: [], diferencas: 0, semObra, projetos: alvo.size };
    const achadas = [];
    if (!alvo.size) return out;
    const cv = (await pool.query(`SELECT c."ProjectId" AS pid, c."Date"::date::text AS dia, c."Realized" AS r
      FROM silver.prevision_curvas_monetaria c
      JOIN (SELECT "ProjectId", max("Load_Timestamp") AS lt FROM silver.prevision_curvas_monetaria WHERE "ProjectId" = ANY($1) GROUP BY 1) u
        ON u."ProjectId" = c."ProjectId" AND u.lt = c."Load_Timestamp"
     WHERE c."Perspective" = 'monetary' AND c."Date"::date = ANY($2::date[])`, [[...alvo.keys()], datas])).rows;
    const real = new Map(cv.map(r => [r.pid + '|' + r.dia, Number(r.r)]));
    const ids = [...new Set([...alvo.values()].map(a => a.obra_id))];
    const exist = (await q(`SELECT obra_id, to_char(competencia, 'YYYY-MM-DD') AS comp, avanco_pct, origem FROM ${S}.avanco_fisico
      WHERE obra_id = ANY($1) AND competencia >= $2`, [ids, meses[0].comp])).rows;
    const jaTem = new Map(exist.map(e => [e.obra_id + '|' + e.comp, e]));
    await emTransacao(async cli => {
      const n = await nomes(cli);
      for (const [pid, a] of alvo) {
        for (let i = 0; i < meses.length; i++) {
          const atual = real.get(pid + '|' + meses[i].fim), ant = real.get(pid + '|' + datas[i]);
          if (atual == null || ant == null || ant >= 0.9999 || (atual === 0 && ant === 0)) continue; // sem curva, obra ainda não começou ou já terminada
          const v = Math.round((atual - ant) * 1e6) / 1e6, comp = meses[i].comp, ex = jaTem.get(a.obra_id + '|' + comp);
          if (excluidas.has(a.obra_id + '|' + comp.slice(0, 7))) { out.excluidas++; continue; } // alguém excluiu de propósito
          if (v < 0 || v > 1) { out.divergentes.push({ obra: n.obra[a.obra_id], comp: comp.slice(0, 7), sistema: ex ? Number(ex.avanco_pct) : null, prevision: v, nota: 'valor fora de 0–100%' }); continue; }
          if (ex && ex.origem !== 'prevision') {
            if (Math.abs(Number(ex.avanco_pct) - v) <= 0.0006) {
              // o valor da planilha confere com o Prevision: passa a constar como vindo dele (o valor não muda)
              await cli.query(`UPDATE ${S}.avanco_fisico SET origem = 'prevision' WHERE obra_id = $1 AND competencia = $2`, [a.obra_id, comp]);
              out.iguais++;
            }
            else achadas.push({ chave: a.obra_id + '|' + comp.slice(0, 7), titulo: `${n.obra[a.obra_id]} · ${comp.slice(5, 7)}/${comp.slice(0, 4)}`,
              planilha: { avanco_pct: Number(ex.avanco_pct) }, sistema: { avanco_pct: v } });
            continue;
          }
          if (ex && Math.abs(Number(ex.avanco_pct) - v) < 0.0000005) { out.iguais++; continue; }
          if (fechados.has(comp)) { out.fechadas++; continue; }
          const chave = chaveLegivel(n, { obra_id: a.obra_id, competencia: comp });
          if (ex) {
            await cli.query(`UPDATE ${S}.avanco_fisico SET avanco_pct = $3, atualizado_por = $4, atualizado_em = now() WHERE obra_id = $1 AND competencia = $2`, [a.obra_id, comp, v, SISTEMA.login]);
            await logar(cli, SISTEMA.login, 'avanco_fisico', chave, 'alterar', { avanco_pct: Number(ex.avanco_pct) }, { avanco_pct: v }, 'Sincronização automática (Prevision)');
            out.atualizadas++;
          } else {
            await cli.query(`INSERT INTO ${S}.avanco_fisico (obra_id, competencia, avanco_pct, origem, criado_por, atualizado_por) VALUES ($1, $2, $3, 'prevision', $4, $4)`, [a.obra_id, comp, v, SISTEMA.login]);
            await logar(cli, SISTEMA.login, 'avanco_fisico', chave, 'incluir', null, { avanco_pct: v }, 'Sincronização automática (Prevision)');
            out.incluidas++;
          }
        }
      }
    });
    out.divergentes = out.divergentes.slice(0, 100);
    out.diferencas = await registrarDiferencas('prevision', achadas);
    return out;
  }

  // Diferenças planilha × sistema: uma linha por (fonte, chave). A decisão fica guardada e só volta a
  // ser pedida se o valor do sistema mudar depois dela. Pendente que deixou de existir (valores
  // ficaram iguais, ou a linha foi excluída) sai da lista. Devolve quantas estão pendentes.
  const estavel = o => JSON.stringify(o, Object.keys(o || {}).sort());
  async function registrarDiferencas(fonte, achadas) {
    const atuais = (await q(`SELECT id, chave, situacao, sistema FROM ${S}.integracao_diferenca WHERE fonte = $1`, [fonte])).rows;
    const porChave = new Map(atuais.map(d => [d.chave, d]));
    for (const a of achadas) {
      const ex = porChave.get(a.chave);
      if (!ex) {
        await q(`INSERT INTO ${S}.integracao_diferenca (fonte, chave, titulo, planilha, sistema, registro_id) VALUES ($1,$2,$3,$4,$5,$6)`,
          [fonte, a.chave, a.titulo, JSON.stringify(a.planilha), JSON.stringify(a.sistema), a.registro_id || null]);
      } else if (estavel(ex.sistema) !== estavel(a.sistema)) {
        await q(`UPDATE ${S}.integracao_diferenca SET titulo = $2, planilha = $3, sistema = $4, registro_id = $5, situacao = 'pendente', decidido_por = NULL, decidido_em = NULL WHERE id = $1`,
          [ex.id, a.titulo, JSON.stringify(a.planilha), JSON.stringify(a.sistema), a.registro_id || null]);
      } else if (ex.situacao === 'pendente') {
        await q(`UPDATE ${S}.integracao_diferenca SET titulo = $2, planilha = $3, registro_id = $4 WHERE id = $1`, [ex.id, a.titulo, JSON.stringify(a.planilha), a.registro_id || null]);
      }
    }
    const vivas = new Set(achadas.map(a => a.chave));
    const sumiram = atuais.filter(d => d.situacao === 'pendente' && !vivas.has(d.chave)).map(d => d.id);
    if (sumiram.length) await q(`DELETE FROM ${S}.integracao_diferenca WHERE id = ANY($1)`, [sumiram]);
    return (await q(`SELECT count(*)::int AS n FROM ${S}.integracao_diferenca WHERE fonte = $1 AND situacao = 'pendente'`, [fonte])).rows[0].n;
  }

  let sincRodando = false;
  async function sincronizar(origem, slot) {
    if (sincRodando) return null;
    sincRodando = true;
    const out = { mudou: false };
    try {
      for (const [fonte, fn] of [['salesforce', salesforceAutomatico], ['prevision', previsionAutomatico]]) {
        const id = (await q(`INSERT INTO ${S}.sincronizacao (fonte, origem, slot) VALUES ($1, $2, $3) RETURNING id`, [fonte, origem, slot || null])).rows[0].id;
        try {
          const r = await fn();
          await q(`UPDATE ${S}.sincronizacao SET status = 'ok', concluido_em = now(), resumo = $2 WHERE id = $1`, [id, JSON.stringify(r)]);
          out[fonte] = r;
          if (r.incluidas || r.vinculadas || r.atualizadas) out.mudou = true;
        } catch (e) {
          console.error(`[sincronizacao] ${fonte}: ${e.message}`);
          await q(`UPDATE ${S}.sincronizacao SET status = 'erro', concluido_em = now(), erro = $2 WHERE id = $1`, [id, String(e.message).slice(0, 500)]);
          out[fonte] = { erro: e.message };
        }
      }
    } finally { sincRodando = false; }
    return out;
  }
  async function sincronizacaoFeita(slot) {
    return (await q(`SELECT 1 FROM ${S}.sincronizacao WHERE origem = 'agendada' AND slot = $1 LIMIT 1`, [slot])).rowCount > 0;
  }
  async function integracoes() {
    const u = await q(`SELECT DISTINCT ON (fonte) fonte, origem, status, erro, resumo, ${quando('iniciado_em')} AS quando
      FROM ${S}.sincronizacao ORDER BY fonte, iniciado_em DESC`);
    const [dif, prevision, salesforce, excl] = await Promise.all([
      q(`SELECT id, fonte, titulo, planilha, sistema, situacao, decidido_por, ${quando('decidido_em')} AS decidido_em
           FROM ${S}.integracao_diferenca ORDER BY (situacao = 'pendente') DESC, fonte, titulo`),
      q(`SELECT a.obra_id, o.nome AS obra, to_char(a.competencia, 'YYYY-MM') AS comp, a.avanco_pct, ${quando('a.atualizado_em')} AS atualizado_em
           FROM ${S}.avanco_fisico a JOIN ${S}.obra o ON o.id = a.obra_id
          WHERE a.origem = 'prevision' ORDER BY a.competencia DESC, o.nome LIMIT 600`),
      q(`SELECT r.id, to_char(r.data, 'YYYY-MM-DD') AS data, o.nome AS obra, r.cod_chamado, r.ocorrencia, r.categoria, r.status_acoes
           FROM ${S}.reclamacao r JOIN ${S}.obra o ON o.id = r.obra_id
          WHERE r.origem = 'salesforce' OR r.sf_case_id IS NOT NULL ORDER BY r.data DESC, r.id DESC LIMIT 600`),
      q(`SELECT fonte, count(*)::int AS n FROM ${S}.integracao_excluida GROUP BY fonte`),
    ]);
    return { rodando: sincRodando, fontes: Object.fromEntries(u.rows.map(r => [r.fonte, r])),
      diferencas: dif.rows, doPrevision: prevision.rows, doSalesforce: salesforce.rows,
      excluidas: Object.fromEntries(excl.rows.map(r => [r.fonte, r.n])) };
  }

  // Decisão sobre uma diferença planilha × sistema: manter o que veio da planilha ou usar o do sistema.
  async function decidirDiferenca(idBruto, decisao, sessao) {
    if (!['editor', 'admin'].includes(sessao.nivel)) throw erroUso('seu acesso é só de consulta', 403);
    if (!['planilha', 'sistema'].includes(decisao)) throw erroUso('decisão inválida');
    const id = numero(idBruto, { inteiro: true, min: 1 });
    return emTransacao(async cli => {
      const d = (await cli.query(`SELECT * FROM ${S}.integracao_diferenca WHERE id = $1 FOR UPDATE`, [id])).rows[0];
      if (!d) throw erroUso('diferença não encontrada: recarregue a tela', 404);
      if (d.situacao !== 'pendente') throw erroUso('esta diferença já foi decidida: recarregue a tela');
      const n = await nomes(cli);
      const fonte = d.fonte === 'prevision' ? 'Prevision' : 'Salesforce';
      const motivo = `Diferença planilha × ${fonte}: ${decisao === 'sistema' ? 'usado o valor do ' + fonte : 'mantido o valor da planilha'}`;
      if (d.fonte === 'prevision') {
        const [obraId, comp7] = d.chave.split('|'), obra = Number(obraId), comp = comp7 + '-01';
        await conferirTrava(cli, comp, sessao, motivo);
        const chave = chaveLegivel(n, { obra_id: obra, competencia: comp });
        if (decisao === 'sistema') {
          const v = d.sistema.avanco_pct;
          const atual = (await cli.query(`SELECT avanco_pct FROM ${S}.avanco_fisico WHERE obra_id = $1 AND competencia = $2 FOR UPDATE`, [obra, comp])).rows[0];
          if (atual) {
            await cli.query(`UPDATE ${S}.avanco_fisico SET avanco_pct = $3, origem = 'prevision', atualizado_por = $4, atualizado_em = now() WHERE obra_id = $1 AND competencia = $2`, [obra, comp, v, sessao.login]);
            await logar(cli, sessao.login, 'avanco_fisico', chave, 'alterar', { avanco_pct: Number(atual.avanco_pct) }, { avanco_pct: v }, motivo);
          } else {
            await cli.query(`INSERT INTO ${S}.avanco_fisico (obra_id, competencia, avanco_pct, origem, criado_por, atualizado_por) VALUES ($1, $2, $3, 'prevision', $4, $4)`, [obra, comp, v, sessao.login]);
            await logar(cli, sessao.login, 'avanco_fisico', chave, 'incluir', null, { avanco_pct: v }, motivo);
          }
        } else {
          await logar(cli, sessao.login, 'integracao_diferenca', chave, 'alterar', null, { decisao: 'manteve a planilha', avanco_pct: d.planilha.avanco_pct }, motivo);
        }
      } else {
        const r = (await cli.query(`SELECT id, obra_id, data FROM ${S}.reclamacao WHERE id = $1 FOR UPDATE`, [d.registro_id])).rows[0];
        if (!r) throw erroUso('a linha da planilha já foi excluída: recarregue a tela', 409);
        await conferirTrava(cli, String(r.data).slice(0, 7) + '-01', sessao, motivo);
        const chave = chaveLegivel(n, { id: r.id, obra_id: r.obra_id, competencia: String(r.data).slice(0, 7) + '-01' });
        if (decisao === 'sistema') {
          const s = d.sistema;
          await conferirTrava(cli, s.data.slice(0, 7) + '-01', sessao, motivo);
          await cli.query(`UPDATE ${S}.reclamacao SET data = $2, obra_id = $3, ocorrencia = $4, categoria = $5, fonte = 'Salesforce', sf_case_id = $6, origem = 'salesforce', atualizado_por = $7, atualizado_em = now() WHERE id = $1`,
            [r.id, s.data, s.obra_id, s.ocorrencia, s.categoria, d.chave, sessao.login]);
          await logar(cli, sessao.login, 'reclamacao', chave, 'alterar', d.planilha, s, motivo);
        } else {
          await cli.query(`UPDATE ${S}.reclamacao SET sf_case_id = $2, atualizado_por = $3, atualizado_em = now() WHERE id = $1`, [r.id, d.chave, sessao.login]);
          await logar(cli, sessao.login, 'integracao_diferenca', chave, 'alterar', null, { decisao: 'manteve a planilha' }, motivo);
        }
      }
      await cli.query(`UPDATE ${S}.integracao_diferenca SET situacao = $2, decidido_por = $3, decidido_em = now() WHERE id = $1`, [id, decisao, sessao.login]);
      return { ok: true };
    });
  }

  // conta do Salesforce -> obra (os próximos chamados dela entram sozinhos)
  async function salesforceContaMapear(it, sessao) {
    if (!['editor', 'admin'].includes(sessao.nivel)) throw erroUso('seu acesso é só de consulta', 403);
    if (!it || typeof it.account_id !== 'string' || !it.account_id) throw erroUso('conta inválida');
    const obra = numero(it.obra_id, { inteiro: true, min: 1 });
    if (!obra) throw erroUso('escolha a obra');
    await q(`INSERT INTO ${S}.salesforce_conta (account_id, nome, obra_id, atualizado_por) VALUES ($1, $2, $3, $4)
      ON CONFLICT (account_id) DO UPDATE SET nome = EXCLUDED.nome, obra_id = EXCLUDED.obra_id, atualizado_por = EXCLUDED.atualizado_por, atualizado_em = now()`,
      [it.account_id, String(it.nome || '').slice(0, 200), obra, sessao.login]);
    await auditar(sessao.login, 'salesforce_conta_ligar', { account_id: it.account_id, obra_id: obra });
  }
  async function previsionMapear(it, sessao) {
    if (!['editor', 'admin'].includes(sessao.nivel)) throw erroUso('seu acesso é só de consulta', 403);
    if (!it || typeof it.project_id !== 'string' || !it.project_id) throw erroUso('projeto inválido');
    const obra = it.ignorar ? null : numero(it.obra_id, { inteiro: true, min: 1 });
    if (!it.ignorar && !obra) throw erroUso('escolha a obra');
    await q(`INSERT INTO ${S}.prevision_projeto (project_id, nome, obra_id, ignorar, atualizado_por) VALUES ($1, $2, $3, $4, $5)
      ON CONFLICT (project_id) DO UPDATE SET obra_id = EXCLUDED.obra_id, ignorar = EXCLUDED.ignorar, atualizado_por = EXCLUDED.atualizado_por, atualizado_em = now()`,
      [it.project_id, String(it.nome || '').slice(0, 200), obra, !!it.ignorar, sessao.login]);
    await auditar(sessao.login, 'prevision_mapear', { project_id: it.project_id, obra_id: obra, ignorar: !!it.ignorar });
  }

  // aba de consulta trazida como está na planilha (ex.: Histórico de resíduos)
  async function abaCongelada(aba) {
    const r = await q(`SELECT linha, valores FROM ${S}.aba_congelada WHERE aba = $1 ORDER BY linha`, [aba]);
    return r.rows;
  }

  // ── Recarga da planilha (só homologação) ─────────────────────────────────
  // quanto se perde: tudo que não veio da carga (lançado ou alterado no sistema)
  async function feitosNoSistema() {
    const r = await q(`SELECT
      (SELECT count(*) FROM ${S}.log_alteracao)::int AS alteracoes,
      (SELECT count(DISTINCT login) FROM ${S}.log_alteracao)::int AS pessoas,
      (SELECT count(*) FROM ${S}.fechamento)::int AS meses_fechados`);
    return r.rows[0];
  }
  // carga_inicial.sql já traz BEGIN/TRUNCATE/.../COMMIT: usuários e sessões ficam
  async function recarregar(sqlCarga, login) {
    const antes = await feitosNoSistema();
    await q(sqlCarga);
    // os ids das linhas mudam na carga: decisões e exclusões guardadas deixam de valer
    await q(`TRUNCATE ${S}.integracao_diferenca, ${S}.integracao_excluida`);
    await auditar(login, 'recarregar_planilha', antes);
    return antes;
  }

  return {
    feitosNoSistema, recarregar, listar, salvarLista, relatorioImpressao, abaCongelada, salesforcePrevia, salesforceImportar,
    sincronizar, sincronizacaoFeita, integracoes, previsionMapear, decidirDiferenca, salesforceContaMapear,
    cadastros, mes, panorama, salvarGrade, salvarNcs, salvarRuido,
    salvarCadastro, salvarMeta, salvarParametro, fecharMes, historico, auditar,
    sessaoGravar, sessaoLer, sessaoApagar, sessoesApagarDe,
    acessoDe, marcarAcesso, situacaoDe, usuarios, decidirUsuario, excluirUsuario, saude,
    nome: `${env.PGUSER}@${env.PGHOST}:${env.PGPORT || 5432}/${env.PGDATABASE} (schema ${S})`,
  };
}

module.exports = { criarDados, erroUso };
