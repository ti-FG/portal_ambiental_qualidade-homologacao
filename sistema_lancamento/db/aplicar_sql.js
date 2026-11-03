// aplicar_sql.js — aplica um arquivo .sql no banco do www\.env (PG*), sem precisar do psql.
//
//   node db\aplicar_sql.js db\schema.sql
//
// Codigos de saida: 0 = aplicado; 3 = o proprio SQL abortou com "ja aplicado" (guard); 1 = erro.
'use strict';
const fs = require('fs');
const path = require('path');
const { carregarEnv } = require('../www/env');

carregarEnv(process.env.ENV_FILE || path.join(__dirname, '..', 'www', '.env'));
const { Client } = require(path.join(__dirname, '..', 'www', 'node_modules', 'pg'));

const arquivo = process.argv[2];
if (!arquivo) { console.error('uso: node db\\aplicar_sql.js <arquivo.sql>'); process.exit(1); }
const nome = path.basename(arquivo);

(async () => {
  let codigo = 0;
  const c = new Client({
    host: process.env.PGHOST, port: Number(process.env.PGPORT || 5432), database: process.env.PGDATABASE,
    user: process.env.PGUSER, password: process.env.PGPASSWORD || undefined, connectionTimeoutMillis: 15000,
  });
  try {
    await c.connect();
    await c.query(fs.readFileSync(arquivo, 'utf8'));   // o arquivo traz o proprio BEGIN/COMMIT
    console.log('ok: ' + nome);
  } catch (e) {
    if (/já aplicado/.test(e.message)) { console.log('ja aplicado, nada a fazer: ' + nome); codigo = 3; }
    else { console.error('ERRO em ' + nome + ': ' + e.message); codigo = 1; }
  } finally {
    await c.end().catch(() => {});   // fechar a conexao desfaz uma transacao que ficou aberta por erro
  }
  process.exit(codigo);
})();
