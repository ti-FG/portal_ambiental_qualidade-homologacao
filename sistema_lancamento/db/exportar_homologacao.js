// exportar_homologacao.js — gera db/carga_homologacao.sql a partir do banco da HML
//
//   node db\exportar_homologacao.js db\carga_homologacao.sql
//
// Lê (só leitura) o banco da homologação (PG* do www\.env da HML, ou ENV_FILE) e grava um
// SQL de INSERTs com os mesmos dados do sistema, SEM o que foi feito para teste e SEM os
// logs. Não grava credenciais no arquivo. Aplicar depois de db/schema.sql, em schema vazio.
'use strict';
const fs = require('fs');

let Client;
try { ({ Client } = require('pg')); }
catch (e) { ({ Client } = require('C:/hml/apps/ambiental_qualidade/www/node_modules/pg')); }

const ENV_FILE = process.env.ENV_FILE || 'C:/hml/apps/ambiental_qualidade/www/.env';
const env = {};
fs.readFileSync(ENV_FILE, 'utf8').split(/\r?\n/).forEach(l => {
  const i = l.indexOf('=');
  if (i > 0 && !l.startsWith('#')) env[l.slice(0, i).trim()] = l.slice(i + 1).trim();
});

const S = 'portal_ambiental_qualidade';
const SAIDA = process.argv[2];
if (!SAIDA) { console.error('uso: node exportar_homologacao.js <arquivo_de_saida.sql>'); process.exit(2); }

// logs e acessos do sistema: não são dados
const EXCLUIR_TABELAS = new Set(['auditoria', 'log_alteracao', 'sessao', 'sincronizacao', 'usuario']);
// cadastros criados só para testar (nenhum lançamento estava ligado a eles)
const TESTE = { obra: [15], unidade_consumidora: [26], assunto_nc: [18], tipo_residuo: [17] };
// FKs que apontam para os cadastros de teste: [tabela, coluna, cadastro]
const FK_TESTE = [
  ['conformidade_nc', 'assunto_id', 'assunto_nc'],
  ['residuo', 'tipo_residuo_id', 'tipo_residuo'],
  ['consumo', 'uc_id', 'unidade_consumidora'],
  ['unidade_consumidora', 'agrupa_em_id', 'unidade_consumidora'],
];

const lit = v => (v === null ? 'NULL' : "'" + String(v).replace(/'/g, "''") + "'");

(async () => {
  // tipos devolvidos como texto: o servidor converte de volta no INSERT
  const c = new Client({
    host: env.PGHOST, port: +env.PGPORT, database: env.PGDATABASE, user: env.PGUSER, password: env.PGPASSWORD,
    types: { getTypeParser: () => v => v },
  });
  await c.connect();
  await c.query('SET default_transaction_read_only = on');
  const q = async (s, p) => (await c.query(s, p)).rows;

  const tabs = (await q(`select table_name t from information_schema.tables where table_schema=$1 and table_type='BASE TABLE' order by 1`, [S]))
    .map(r => r.t).filter(t => !EXCLUIR_TABELAS.has(t));
  const cols = {}, pk = {};
  for (const t of tabs) {
    cols[t] = (await q(`select column_name c from information_schema.columns where table_schema=$1 and table_name=$2 and is_generated='NEVER' order by ordinal_position`, [S, t])).map(r => r.c);
    pk[t] = (await q(`select a.attname c from pg_index i join pg_attribute a on a.attrelid=i.indrelid and a.attnum=any(i.indkey) where i.indrelid='${S}.${t}'::regclass and i.indisprimary`)).map(r => r.c);
  }

  // ordem de inserção pelas FKs
  const fks = await q(`select conrelid::regclass::text a, confrelid::regclass::text b from pg_constraint where contype='f' and connamespace='${S}'::regnamespace`);
  const nome = x => x.replace(S + '.', '').replace(/"/g, '');
  const dep = {}; tabs.forEach(t => { dep[t] = new Set(); });
  fks.forEach(f => { const a = nome(f.a), b = nome(f.b); if (a !== b && dep[a] && tabs.includes(b)) dep[a].add(b); });
  const ordem = [], feito = new Set();
  while (ordem.length < tabs.length) {
    let prog = false;
    for (const t of tabs) if (!feito.has(t) && [...dep[t]].every(d => feito.has(d))) { ordem.push(t); feito.add(t); prog = true; }
    if (!prog) throw new Error('ciclo de FK entre: ' + tabs.filter(t => !feito.has(t)));
  }

  const out = [], resumo = [], seqs = [];
  for (const t of ordem) {
    let rows = await q(`select ${cols[t].map(x => `"${x}"`).join(',')} from ${S}."${t}" order by ${pk[t].length ? pk[t].map(x => `"${x}"`).join(',') : '1'}`);
    const total = rows.length;

    if (TESTE[t]) rows = rows.filter(r => !TESTE[t].includes(+r.id));
    for (const [tb, col, ref] of FK_TESTE) if (tb === t) rows = rows.filter(r => r[col] === null || !TESTE[ref].includes(+r[col]));
    if (cols[t].includes('obra_id')) rows = rows.filter(r => !TESTE.obra.includes(+r.obra_id));

    // edição de teste em linha real: GESSO / Blue View / set-2026 voltou ao valor da planilha (0)
    let corrigidas = 0;
    if (t === 'residuo') {
      const bv = (await q(`select id from ${S}.obra where nome='Blue View'`))[0].id;
      const gs = (await q(`select id from ${S}.tipo_residuo where nome='GESSO'`))[0].id;
      rows.forEach(r => {
        if (r.obra_id === bv && r.tipo_residuo_id === gs && r.competencia.startsWith('2026-09-01') && +r.volume_m3 === 1) { r.volume_m3 = '0'; corrigidas++; }
      });
    }

    // unidade_consumidora: a que agrupa vem antes da agrupada
    if (t === 'unidade_consumidora') {
      const ids = new Set(), pend = [...rows], ord = [];
      while (pend.length) {
        const i = pend.findIndex(r => r.agrupa_em_id === null || ids.has(r.agrupa_em_id));
        if (i < 0) throw new Error('agrupa_em_id em ciclo');
        const [r] = pend.splice(i, 1); ids.add(r.id); ord.push(r);
      }
      rows = ord;
    }

    resumo.push({ t, total, saem: total - rows.length, entram: rows.length, corrigidas });
    if (!rows.length) continue;

    out.push(`-- ${t}: ${rows.length} linhas`);
    const colunas = cols[t].map(x => `"${x}"`).join(', ');
    for (let i = 0; i < rows.length; i += 200) {
      const lote = rows.slice(i, i + 200).map(r => '(' + cols[t].map(x => lit(r[x])).join(', ') + ')').join(',\n  ');
      // schema.sql já semeia 'parametro'
      const conflito = t === 'parametro' ? ' ON CONFLICT (chave) DO UPDATE SET valor=EXCLUDED.valor, descricao=EXCLUDED.descricao' : '';
      out.push(`INSERT INTO ${t} (${colunas}) VALUES\n  ${lote}${conflito};`);
    }
    if (cols[t].includes('id') && (await q(`select 1 from pg_class where relname='${t}_id_seq' and relnamespace='${S}'::regnamespace`)).length) seqs.push(t);
    out.push('');
  }
  await c.end();

  const cab = [
    `-- carga_homologacao.sql — dados do sistema de homologação (8095) em ${new Date().toISOString().slice(0, 10)}, SEM o que foi feito para teste.`,
    `-- Gerado por db/exportar_homologacao.js a partir do banco da HML (só leitura). NÃO contém credenciais.`,
    `-- Aplicar DEPOIS de db/schema.sql, em schema VAZIO. Em schema com dados (ex.: o do DW) o guard abaixo aborta.`,
    `-- Fora desta carga: usuario, sessao, auditoria (acessos), log_alteracao e sincronizacao (logs do sistema).`,
    `-- Fora por serem de teste: obra "Teste", unidade_consumidora "Teste", assunto_nc "Teste", tipo_residuo "teste"`,
    `--   (nenhuma linha de lançamento estava ligada a elas); GESSO / Blue View / set-2026 (1 m³ de teste) volta a 0, valor da planilha.`,
    `-- Linhas por tabela (total na HML / fora / entram):`,
    ...resumo.map(r => `--   ${r.t.padEnd(24)} ${String(r.total).padStart(5)} / ${String(r.saem).padStart(3)} / ${String(r.entram).padStart(5)}${r.corrigidas ? `   (${r.corrigidas} corrigida)` : ''}`),
    '', `SET search_path = ${S};`, 'BEGIN;', '',
    `DO $$ BEGIN IF EXISTS (SELECT 1 FROM obra) THEN RAISE EXCEPTION 'carga_homologacao.sql: o schema já tem obras; esta carga é só para schema vazio'; END IF; END $$;`, '',
  ];
  const fim = ['-- sequências dos ids', ...seqs.map(t => `SELECT setval(pg_get_serial_sequence('${t}','id'), COALESCE((SELECT MAX(id) FROM ${t}),0)+1, false);`), '', 'COMMIT;', ''];
  fs.writeFileSync(SAIDA, [...cab, ...out, ...fim].join('\n'), 'utf8');

  console.table(resumo.filter(r => r.saem || r.corrigidas));
  console.log('tabelas:', ordem.length, ' arquivo:', SAIDA);
})().catch(e => { console.error('ERRO', e.message); process.exit(1); });
