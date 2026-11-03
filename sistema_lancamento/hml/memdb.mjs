/**
 * Banco da HOMOLOGAÇÃO dos Lançamentos Ambientais: PostgreSQL embutido (PGlite).
 * Mesmo modelo do hml/memdb.mjs do FG Metas Engenharia, com uma diferença
 * de propósito: aqui o banco é PERSISTIDO em hml/dados. A homologação serve
 * para a Qualidade experimentar o sistema; o que elas lançarem não pode sumir
 * quando o serviço reiniciar.
 *
 * Na primeira subida (hml/dados vazio):
 *   1. cria o schema (db/schema.sql);
 *   2. carrega o histórico da planilha (db/carga_inicial.sql, gerado por
 *      carga/carga_planilha.py).
 * Nas seguintes: só reaplica db/schema.sql (idempotente: views e colunas novas).
 *
 * Nada toca o 10.1.1.152. Escuta só em 127.0.0.1:15439.
 *
 * Voltar ao estado da planilha: parar o serviço, apagar hml/dados, subir de novo.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { PGlite } from '@electric-sql/pglite';
import { PGLiteSocketServer } from '@electric-sql/pglite-socket';

const AQUI = path.dirname(fileURLToPath(import.meta.url));
const DB_DIR = path.join(AQUI, '..', 'db');
const DADOS = process.env.MEMDB_DADOS ?? path.join(AQUI, 'dados');
const HOST = '127.0.0.1';
const PORTA = Number(process.env.MEMDB_PORT ?? 15439); // 15433-15438/15440-15441 são de outros apps
const SCHEMA = 'portal_ambiental_qualidade';

const t0 = Date.now();
const log = (m) => console.log(`[memdb ${((Date.now() - t0) / 1000).toFixed(1)}s] ${m}`);

const novo = !fs.existsSync(path.join(DADOS, 'PG_VERSION'));
const db = new PGlite(DADOS);
await db.waitReady;

await db.exec(fs.readFileSync(path.join(DB_DIR, 'schema.sql'), 'utf8'));
log('schema aplicado');

if (novo) {
  await db.exec(fs.readFileSync(path.join(DB_DIR, 'carga_inicial.sql'), 'utf8'));
  const { rows } = await db.query(`SELECT
      (SELECT count(*) FROM ${SCHEMA}.obra)::int AS obras,
      (SELECT count(*) FROM ${SCHEMA}.residuo)::int AS residuos,
      (SELECT count(*) FROM ${SCHEMA}.conformidade)::int AS conformidade,
      (SELECT count(*) FROM ${SCHEMA}.consumo)::int AS consumo`);
  log(`banco novo: carga da planilha aplicada ${JSON.stringify(rows[0])}`);
} else {
  log(`banco existente em ${DADOS}`);
}

const server = new PGLiteSocketServer({ db, host: HOST, port: PORTA, maxConnections: 20 });
await server.start();
log(`pronto em postgresql://${HOST}:${PORTA}/postgres (schema ${SCHEMA})`);

// mesmo defeito do pglite-socket 0.2.11 tratado nos outros memdb: conexão que
// cai com erro não sai da contagem. A varredura devolve essas vagas.
setInterval(() => {
  let liberadas = 0;
  for (const h of server.handlers) {
    if (!h.active) { server.handlers.delete(h); liberadas++; }
  }
  if (liberadas > 0) log(`${liberadas} conexão(ões) mortas liberadas; ativas: ${server.handlers.size}`);
}, 30_000).unref();

const parar = async () => {
  await server.stop().catch(() => {});
  await db.close().catch(() => {});
  process.exit(0);
};
process.on('SIGINT', parar);
process.on('SIGTERM', parar);
