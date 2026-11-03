// Gera o SQL que leva o schema do DW ao estado da HML (sem teste). Só leitura nos dois bancos.
const { Client } = require('pg'); const fs = require('fs');
const env = {};
fs.readFileSync('C:/hml/apps/ambiental_qualidade/www/.env', 'utf8').split(/\r?\n/).forEach(l => {
  const i = l.indexOf('='); if (i > 0 && !l.startsWith('#')) env[l.slice(0, i).trim()] = l.slice(i + 1).trim();
});
const S = 'portal_ambiental_qualidade';
const raw = { getTypeParser: () => v => v };
const lit = v => (v === null ? 'NULL' : "'" + String(v).replace(/'/g, "''") + "'");
const AUDIT = new Set(['criado_em', 'criado_por', 'atualizado_em', 'atualizado_por', 'id']);
// tabelas a levar: a chave natural usada para saber se a linha já existe no DW é a linha inteira (sem id/auditoria)
const INSERIR = ['impressoes', 'impressoes_setor', 'impressoes_pessoa', 'reclamacao', 'salesforce_conta'];
(async () => {
  const h = new Client({ host: env.PGHOST, port: +env.PGPORT, database: env.PGDATABASE, user: env.PGUSER, password: 'x', types: raw });
  const d = new Client({ host: env.DWC_PGHOST, port: +env.DWC_PGPORT, database: env.DWC_PGDATABASE, user: env.DWC_PGUSER, password: env.DWC_PGPASSWORD, types: raw, connectionTimeoutMillis: 15000 });
  await h.connect(); await d.connect();
  await h.query('SET default_transaction_read_only = on'); await d.query('SET default_transaction_read_only = on');
  const obraH = Object.fromEntries((await h.query(`select id, nome from ${S}.obra`)).rows.map(r => [r.id, r.nome]));
  const out = [], resumo = [];
  const colunas = async (cli, t) => (await cli.query(`select column_name c from information_schema.columns where table_schema=$1 and table_name=$2 and is_generated='NEVER' order by ordinal_position`, [S, t])).rows.map(r => r.c);

  // ---- avanço físico: valores do Prevision (HML) e meses novos
  {
    const sel = `select a.obra_id, a.competencia, a.avanco_pct, a.origem, a.criado_por, a.criado_em, a.atualizado_por, a.atualizado_em from ${S}.avanco_fisico a where a.obra_id<>15`;
    const H = (await h.query(sel)).rows;
    const obraD = Object.fromEntries((await d.query(`select id, nome from ${S}.obra`)).rows.map(r => [r.nome, r.id]));
    const D = (await d.query(`select obra_id, competencia, avanco_pct from ${S}.avanco_fisico`)).rows;
    const dk = new Map(D.map(r => [r.obra_id + '|' + r.competencia, r]));
    const upd = [], ins = [];
    for (const r of H) {
      const nome = obraH[r.obra_id]; const idD = obraD[nome];
      if (!idD) throw new Error('obra da HML não existe no DW: ' + nome);
      const o = dk.get(idD + '|' + r.competencia);
      if (!o) ins.push({ ...r, nome }); else if (+o.avanco_pct !== +r.avanco_pct) upd.push({ ...r, nome });
    }
    out.push(`-- avanco_fisico: ${upd.length} meses com o valor do Prevision (HML) e ${ins.length} meses novos`);
    for (const r of upd) out.push(`UPDATE avanco_fisico SET avanco_pct=${lit(r.avanco_pct)}, origem=${lit(r.origem)}, atualizado_por=${lit(r.atualizado_por)}, atualizado_em=${lit(r.atualizado_em)} WHERE obra_id=(SELECT id FROM obra WHERE nome=${lit(r.nome)}) AND competencia=${lit(r.competencia)};`);
    for (const r of ins) out.push(`INSERT INTO avanco_fisico (obra_id, competencia, avanco_pct, origem, criado_por, criado_em, atualizado_por, atualizado_em) VALUES ((SELECT id FROM obra WHERE nome=${lit(r.nome)}), ${lit(r.competencia)}, ${lit(r.avanco_pct)}, ${lit(r.origem)}, ${lit(r.criado_por)}, ${lit(r.criado_em)}, ${lit(r.atualizado_por)}, ${lit(r.atualizado_em)});`);
    out.push('');
    resumo.push(['avanco_fisico', 'update ' + upd.length, 'insert ' + ins.length]);
  }

  // ---- tabelas em que a HML tem linhas que o DW não tem
  for (const t of INSERIR) {
    const colH = await colunas(h, t), colD = new Set(await colunas(d, t));
    const cols = colH.filter(c => colD.has(c));
    const cmp = cols.filter(c => !AUDIT.has(c));
    const temObra = cols.includes('obra_id');
    const obraD = Object.fromEntries((await d.query(`select id, nome from ${S}.obra`)).rows.map(r => [String(r.id), r.nome]));
    const chaveDW = r => cmp.map(c => c === 'obra_id' ? (r[c] === null ? '' : (obraD[r[c]] || '?')) : (r[c] ?? '')).join('\u0001');
    const chaveH = r => cmp.map(c => c === 'obra_id' ? (r[c] === null ? '' : (obraH[r[c]] || '?')) : (r[c] ?? '')).join('\u0001');
    const filtroH = temObra ? 'where obra_id is distinct from 15' : '';
    const H = (await h.query(`select ${cols.map(c => `"${c}"`).join(',')} from ${S}."${t}" ${filtroH}`)).rows;
    const D = (await d.query(`select ${cols.map(c => `"${c}"`).join(',')} from ${S}."${t}"`)).rows;
    const tem = new Map(); D.forEach(r => { const k = chaveDW(r); tem.set(k, (tem.get(k) || 0) + 1); });
    const faltam = [];
    for (const r of H) { const k = chaveH(r); const n = tem.get(k) || 0; if (n > 0) tem.set(k, n - 1); else faltam.push(r); }
    out.push(`-- ${t}: ${faltam.length} linhas que a HML tem e o DW não`);
    const ins = cols.filter(c => c !== 'id');
    for (const r of faltam) {
      const vals = ins.map(c => c === 'obra_id' ? (r[c] === null ? 'NULL' : `(SELECT id FROM obra WHERE nome=${lit(obraH[r[c]])})`) : lit(r[c]));
      out.push(`INSERT INTO ${t} (${ins.map(c => `"${c}"`).join(', ')}) VALUES (${vals.join(', ')});`);
    }
    out.push('');
    resumo.push([t, 'insert ' + faltam.length]);
  }
  await h.end(); await d.end();

  const cab = [
    `-- atualizar_dw_para_homologacao.sql — leva o schema ${S} do DW ao mesmo estado do sistema de homologação (8095) em ${new Date().toISOString().slice(0, 10)}, SEM o que foi teste.`,
    `-- Gerado por leitura dos dois bancos; só traz o que a HML tem e o DW não. NÃO apaga nada, não contém credenciais.`,
    `-- Rodar UMA vez, com o dono do schema. O guard aborta se já foi aplicado (existe relatório de impressão no DW).`,
    `-- Resumo: ${resumo.map(r => r.join(' ')).join(' | ')}`,
    '', `SET search_path = ${S};`, 'BEGIN;', '',
    `DO $$ BEGIN IF EXISTS (SELECT 1 FROM impressoes WHERE origem='relatorio') THEN RAISE EXCEPTION 'atualizar_dw_para_homologacao.sql: já aplicado (há relatório de impressão no DW)'; END IF; END $$;`, '',
  ];
  fs.writeFileSync(process.argv[2], [...cab, ...out, 'COMMIT;', ''].join('\n'), 'utf8');
  console.log(resumo.map(r => r.join(' ')).join('\n'));
})().catch(e => { console.log('ERRO', e.message); process.exit(1); });
