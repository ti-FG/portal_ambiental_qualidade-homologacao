// env.js — lê o .env (fora de public\) para process.env, sem dependência.
// Mesmo formato do carregarEnv do Painel de Metas: CHAVE=valor, # comenta,
// o valor pode ter '=' (senha). O que já está no ambiente vence o arquivo
// (assim o teste e o serviço podem sobrescrever uma chave sem editar o .env).
'use strict';

const fs = require('fs');

function carregarEnv(arquivo) {
  let raw;
  try { raw = fs.readFileSync(arquivo, 'utf8'); }
  catch (e) { console.error('[env] ' + arquivo + ' não pôde ser lido: ' + e.message); return; }
  raw.split(/\r?\n/).forEach(linha => {
    const l = linha.trim();
    if (!l || l.startsWith('#')) return;
    const i = l.indexOf('=');
    if (i < 0) return;
    const k = l.slice(0, i).trim();
    if (process.env[k] === undefined) process.env[k] = l.slice(i + 1).trim().replace(/^["']|["']$/g, '');
  });
}

module.exports = { carregarEnv };
