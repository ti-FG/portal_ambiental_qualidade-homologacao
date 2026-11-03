// =============================================================================
// server.js — Lançamentos Ambientais (FG · Qualidade), 21/10/2026
//
// Substitui a planilha "Controle Ambiental 2025_oficial_db.xlsx": a área lança
// aqui, o sistema grava no PostgreSQL (schema portal_ambiental_qualidade) e os
// painéis passam a ler do banco.
//
// Mesmo desenho do FG Engenharia e do Portfólio de Projetos: Node sem framework,
// login no AD (ldap.js), sessão no banco (só o hash do token), um único HTML em
// public\. Diferente do Metas, a API é por recurso (não sincroniza um estado
// inteiro):
//
//   GET  /api/cadastros                obras, tipos de resíduo, unidades, assuntos, metas, parâmetros
//   GET  /api/mes?competencia=AAAA-MM  tudo o que foi lançado no mês
//   GET  /api/panorama?meses=12        quantas obras têm cada lançamento, mês a mês
//   PUT  /api/lancamentos/:fonte       (editor) grade: avanco | conformidade | residuo | consumo
//   PUT  /api/ncs                      (editor) assuntos de NC de uma obra no mês
//   PUT  /api/ruido                    (editor) medições de ruído do mês
//   GET  /api/integracoes · POST /api/integracoes/rodar | prevision | salesforce-conta | diferenca
//                                      Salesforce e Prevision entram só pela sincronização; a pessoa decide diferenças e liga contas/projetos
//   PUT  /api/cadastros/:entidade      (admin) obra | tipo_residuo | unidade_consumidora | assunto_nc
//   PUT  /api/metas                    (admin) meta de conformidade com vigência
//   PUT  /api/parametros/:chave        (admin)
//   POST /api/fechamento               (admin) fecha ou reabre um mês
//   GET  /api/historico                log de alterações campo a campo
//   GET  /api/usuarios · POST /api/usuarios/decidir   (admin)
//   POST /api/auth/login | /api/auth/logout · GET /api/auth/sessao
//   GET  /api/versao · GET /api/saude (sem login, para monitoramento)
//
// Níveis: leitor (consulta), editor (lança em mês aberto), admin (cadastros,
// usuários, fechamento e correção de mês fechado com motivo). ADMIN_LOGINS são
// administradores fixos; as demais contas do AD pedem acesso na 1ª vez.
// =============================================================================
'use strict';

const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const zlib = require('zlib');
const { carregarEnv } = require('./env');

carregarEnv(path.join(__dirname, '.env'));

const ldap = require('./ldap');
const { criarDados } = require('./dados');

const PORT = Number(process.env.PORT) || 8095;
const ROOT = path.join(__dirname, 'public');
const AMBIENTE = (process.env.AMBIENTE || '').toLowerCase();
const VERSAO = require('./package.json').version;

const lista = v => String(v || '').split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
const ADMIN_LOGINS = lista(process.env.ADMIN_LOGINS || 'natalia.costa');

const dados = criarDados(process.env);

// ── Painel (dashboard) gerado do banco ───────────────────────────────────────
// painel/gerar_painel.py roda o gerador oficial do dashboard lendo deste banco
// e grava www/painel/painel.html. Uma geração por vez; gravação de lançamento
// agenda uma nova (com folga, para várias gravações seguidas virarem uma só).
const { execFile } = require('child_process');
const PYTHON = process.env.PYTHON || 'python';
const PAINEL_DIR = path.join(__dirname, 'painel');
const PAINEL_SCRIPT = path.join(__dirname, '..', 'painel', 'gerar_painel.py');
const CARGA_SCRIPT = path.join(__dirname, '..', 'carga', 'carga_planilha.py');
const RELATORIO_SCRIPT = path.join(__dirname, '..', 'carga', 'ler_relatorio_impressao.py');
const CONFERENCIA_SCRIPT = path.join(__dirname, '..', 'carga', 'conferencia_planilha.py');
const CARGA_SQL = path.join(__dirname, '..', 'db', 'carga_inicial.sql');
const os = require('os');
function rodarPython(script, args) {
  return new Promise((resolve, reject) => {
    execFile(PYTHON, [script, ...args], { cwd: path.dirname(script), timeout: 300000, windowsHide: true, maxBuffer: 16 * 1024 * 1024 }, (err, out, errOut) => {
      if (err) {
        console.error(`[python] ${path.basename(script)} falhou: ${err.message}\n${String(errOut || '').slice(-3000)}`);
        return reject(Object.assign(new Error(`${path.basename(script)} falhou (ver logs\\servidor.log)`), { status: 500 }));
      }
      resolve(String(out || ''));
    });
  });
}
const painel = { gerando: false, deNovo: false, timer: null, erro: null, ultimaDuracao: null };
function gerarPainel() {
  if (painel.gerando) { painel.deNovo = true; return; }
  painel.gerando = true; painel.erro = null;
  const t0 = Date.now();
  execFile(PYTHON, [PAINEL_SCRIPT], { cwd: path.dirname(PAINEL_SCRIPT), timeout: 180000, windowsHide: true, maxBuffer: 8 * 1024 * 1024 }, (err, out, errOut) => {
    painel.gerando = false;
    painel.ultimaDuracao = Math.round((Date.now() - t0) / 100) / 10;
    if (err) {
      painel.erro = 'a geração do painel falhou (ver logs\\servidor.log)';
      console.error('[painel] falhou: ' + (err.message || err) + '\n' + String(errOut || '').slice(-3000));
    } else {
      console.log('[painel] gerado em ' + painel.ultimaDuracao + 's');
    }
    if (painel.deNovo) { painel.deNovo = false; gerarPainel(); }
  });
}
function agendarPainel() {
  clearTimeout(painel.timer);
  painel.agendado = true;
  painel.timer = setTimeout(() => { painel.agendado = false; gerarPainel(); }, 5000);
}
function infoPainel() {
  let info = {};
  try { info = JSON.parse(fs.readFileSync(path.join(PAINEL_DIR, 'painel.json'), 'utf8')); } catch (e) {}
  return { ...info, existe: fs.existsSync(path.join(PAINEL_DIR, 'painel.html')), gerando: painel.gerando || painel.agendado, erro: painel.erro };
}

// ── Sessões: memória na frente, banco atrás ──────────────────────────────────
const sessoes = new Map(); // token -> {login, nome, email, nivel, criadaEm}
const VALIDADE_SESSAO = 12 * 60 * 60 * 1000;
async function sessaoDe(req) {
  const token = req.headers['x-sessao'];
  if (!token || typeof token !== 'string' || token.length > 200) return null;
  const agora = Date.now();
  for (const [t, s] of sessoes) if (agora - s.criadaEm > VALIDADE_SESSAO) sessoes.delete(t);
  let s = sessoes.get(token);
  if (!s) {
    s = await dados.sessaoLer(token);
    if (!s || agora - s.criadaEm > VALIDADE_SESSAO) return null;
    sessoes.set(token, s);
  }
  // o nível de agora (tela Usuários / .env), nunca o que estava gravado
  const nivel = await nivelDe(s.login);
  if (!nivel) { sessoes.delete(token); await dados.sessaoApagar(token); return null; }
  s.nivel = nivel;
  return s;
}
async function nivelDe(login) {
  if (ADMIN_LOGINS.includes(login)) return 'admin';
  const u = await dados.situacaoDe(login);
  return u.situacao === 'aprovado' ? u.nivel : null;
}
async function derrubarSessoes(login) {
  for (const [t, s] of sessoes) if (s.login === login) sessoes.delete(t);
  await dados.sessoesApagarDe(login);
}

// ── Tentativas de login: senha errada repetida trava a conta no AD ──────────
const LOGIN_MAX_FALHAS = Math.max(1, Number(process.env.LOGIN_MAX_FALHAS || 5));
const LOGIN_MAX_FALHAS_IP = Math.max(1, Number(process.env.LOGIN_MAX_FALHAS_IP || 20));
const LOGIN_JANELA = Math.max(1, Number(process.env.LOGIN_JANELA_MINUTOS || 15)) * 60 * 1000;
const falhas = new Map();
function falhasDe(chave) {
  const f = falhas.get(chave);
  if (f && Date.now() - f.desde > LOGIN_JANELA) { falhas.delete(chave); return null; }
  return f || null;
}
function bloqueado(login, ip) {
  const fl = falhasDe('l:' + login), fi = falhasDe('ip:' + ip);
  const fim = f => Math.ceil((f.desde + LOGIN_JANELA - Date.now()) / 60000);
  if (fl && fl.n >= LOGIN_MAX_FALHAS) return fim(fl);
  if (fi && fi.n >= LOGIN_MAX_FALHAS_IP) return fim(fi);
  return 0;
}
function contarFalha(login, ip) {
  for (const k of ['l:' + login, 'ip:' + ip]) {
    const f = falhasDe(k);
    if (f) f.n++; else falhas.set(k, { n: 1, desde: Date.now() });
  }
}

// ── HTTP ─────────────────────────────────────────────────────────────────────
const LIMITE_CORPO = 2 * 1024 * 1024;
function lerCorpo(req) {
  return new Promise((resolve, reject) => {
    const partes = []; let tamanho = 0;
    req.on('data', c => {
      tamanho += c.length;
      if (tamanho > LIMITE_CORPO) { reject(Object.assign(new Error('corpo grande demais'), { status: 413 })); req.destroy(); return; }
      partes.push(c);
    });
    req.on('end', () => {
      try { resolve(JSON.parse(Buffer.concat(partes).toString('utf8') || '{}')); }
      catch (e) { reject(Object.assign(new Error('JSON inválido'), { status: 400 })); }
    });
    req.on('error', reject);
  });
}
function enviar(res, status, cabecalhos, corpo) {
  const aceita = /\bgzip\b/.test((res.req && res.req.headers['accept-encoding']) || '');
  if (!aceita || corpo.length < 1024) { res.writeHead(status, cabecalhos); return res.end(corpo); }
  zlib.gzip(corpo, (err, gz) => {
    if (err) { res.writeHead(status, cabecalhos); return res.end(corpo); }
    res.writeHead(status, Object.assign({}, cabecalhos, { 'Content-Encoding': 'gzip', 'Vary': 'Accept-Encoding' }));
    res.end(gz);
  });
}
function json(res, status, obj) {
  enviar(res, status, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' }, Buffer.from(JSON.stringify(obj), 'utf8'));
}

async function rotaLogin(req, res) {
  const { login, senha } = await lerCorpo(req);
  const quem = ldap.normalizarLogin(login), ip = req.socket.remoteAddress || '';
  const minutos = bloqueado(quem, ip);
  if (minutos) return json(res, 429, { ok: false, motivo: 'muitas_tentativas', minutos });
  if (!ldap.ldapConfigurado) {
    console.error('[auth] LDAP_* em branco no .env: login indisponível');
    return json(res, 503, { ok: false, motivo: 'ad_nao_configurado' });
  }
  let r;
  try {
    r = await ldap.autenticar(login, senha);
  } catch (err) {
    console.error('[auth] AD indisponível: ' + (err.message || err));
    return json(res, 503, { ok: false, motivo: 'ad_indisponivel' });
  }
  if (!r.ok) {
    contarFalha(quem, ip);
    console.log('[auth] negado (' + r.motivo + '): ' + quem);
    return json(res, 401, { ok: false, motivo: 'credenciais_invalidas' });
  }
  falhas.delete('l:' + quem);
  const u = r.usuario;
  const fixo = ADMIN_LOGINS.includes(u.login);
  const acesso = await dados.acessoDe(u, fixo ? { situacao: 'aprovado', nivel: 'admin', por: 'configuração' } : null);
  if (!fixo) {
    if (acesso.situacao === 'pendente') {
      if (acesso.novo) { console.log('[auth] pedido de acesso: ' + u.login); dados.auditar(u.login, 'acesso_pedido', null); }
      return json(res, 403, { ok: false, motivo: 'aguardando' });
    }
    if (acesso.situacao === 'rejeitado') return json(res, 403, { ok: false, motivo: 'rejeitado' });
  }
  dados.marcarAcesso(u.login);
  const token = crypto.randomBytes(32).toString('hex');
  const sessao = { login: u.login, nome: u.nome, email: u.email, nivel: fixo ? 'admin' : acesso.nivel, criadaEm: Date.now() };
  sessoes.set(token, sessao);
  await dados.sessaoGravar(token, sessao, VALIDADE_SESSAO);
  console.log('[auth] ok: ' + u.login + ' (' + sessao.nivel + ')');
  dados.auditar(u.login, 'login', { nivel: sessao.nivel });
  return json(res, 200, { ok: true, token, usuario: publico(sessao) });
}
const publico = s => ({ login: s.login, nome: s.nome, nivel: s.nivel });

async function rotaSaude(res) {
  const r = { sistema: 'Lançamentos Ambientais', versao: VERSAO, ambiente: AMBIENTE || 'producao', ok: true, ad: ldap.ldapConfigurado };
  try { await dados.saude(); r.banco = 'ok'; } catch (e) { r.banco = 'falhou'; r.ok = false; console.error('[saude] ' + e.message); }
  return json(res, r.ok ? 200 : 503, r);
}

const TIPOS = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.png': 'image/png', '.svg': 'image/svg+xml', '.ico': 'image/x-icon', '.json': 'application/json' };
function estatico(req, res, url) {
  let rel = decodeURIComponent(url.pathname);
  if (rel === '/') rel = '/index.html';
  // nunca servir dotfiles nem sair de public\
  if (rel.split('/').some(p => p.startsWith('.'))) { res.writeHead(404); return res.end(); }
  const arq = path.normalize(path.join(ROOT, rel));
  if (!arq.startsWith(ROOT + path.sep)) { res.writeHead(403); return res.end(); }
  fs.readFile(arq, (err, buf) => {
    if (err) { res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' }); return res.end('não encontrado'); }
    let corpo = buf;
    const ext = path.extname(arq).toLowerCase();
    if (ext === '.html') {
      corpo = Buffer.from(buf.toString('utf8')
        .replace(/__AMBIENTE__/g, AMBIENTE)
        .replace(/__VERSAO__/g, VERSAO), 'utf8');
    }
    enviar(res, 200, { 'Content-Type': TIPOS[ext] || 'application/octet-stream', 'Cache-Control': ext === '.html' ? 'no-cache' : 'max-age=3600',
      'X-Content-Type-Options': 'nosniff', 'X-Frame-Options': 'SAMEORIGIN' }, corpo);
  });
}

async function rotear(req, res) {
  const url = new URL(req.url, 'http://localhost');
  const p = url.pathname, m = req.method;
  if (!p.startsWith('/api/') && !p.startsWith('/painel')) {
    if (m !== 'GET' && m !== 'HEAD') { res.writeHead(405); return res.end(); }
    return estatico(req, res, url);
  }
  if (p === '/api/saude') return rotaSaude(res);
  // o painel é público na rede, como o dashboard de hoje (8090)
  if (p === '/painel' || p === '/painel/') {
    return fs.readFile(path.join(PAINEL_DIR, 'painel.html'), (err, buf) => {
      if (err) { res.writeHead(503, { 'Content-Type': 'text/plain; charset=utf-8' }); return res.end('painel ainda não gerado'); }
      enviar(res, 200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-cache' }, buf);
    });
  }
  if (p === '/api/versao') return json(res, 200, { versao: VERSAO, ambiente: AMBIENTE || 'producao' });
  if (p === '/api/auth/login' && m === 'POST') return rotaLogin(req, res);

  const sessao = await sessaoDe(req);
  if (!sessao) return json(res, 401, { erro: 'sessão expirada: entre de novo' });
  const admin = sessao.nivel === 'admin';
  const soAdmin = () => json(res, 403, { erro: 'só administradores' });

  if (p === '/api/auth/sessao' && m === 'GET') return json(res, 200, { usuario: publico(sessao) });
  if (p === '/api/auth/logout' && m === 'POST') {
    const t = req.headers['x-sessao'];
    sessoes.delete(t); await dados.sessaoApagar(t);
    return json(res, 200, { ok: true });
  }
  if (p === '/api/cadastros' && m === 'GET') return json(res, 200, await dados.cadastros());
  if (p === '/api/mes' && m === 'GET') return json(res, 200, await dados.mes(url.searchParams.get('competencia')));
  if (p === '/api/panorama' && m === 'GET') return json(res, 200, await dados.panorama(url.searchParams.get('meses')));
  if (p === '/api/historico' && m === 'GET') {
    return json(res, 200, await dados.historico({
      tabela: url.searchParams.get('tabela'), login: url.searchParams.get('login'),
      competencia: url.searchParams.get('competencia'), limite: url.searchParams.get('limite'),
      grupo: url.searchParams.get('grupo'), origem: url.searchParams.get('origem'),
    }));
  }

  // Integrações: Salesforce e Prevision (via DW), 3x ao dia e sob demanda
  if (p === '/api/integracoes' && m === 'GET') return json(res, 200, { ...(await dados.integracoes()), horarios: HORARIOS_SINC });
  if (p === '/api/integracoes/rodar' && m === 'POST') {
    if (sessao.nivel !== 'admin') return json(res, 403, { erro: 'só administrador roda a sincronização na hora' });
    const r = await rodarSincronizacao('manual', null);
    return json(res, 200, r ? { ok: true } : { ok: false, erro: 'já tem uma sincronização em andamento' });
  }
  if (p === '/api/integracoes/prevision' && m === 'POST') {
    await dados.previsionMapear(await lerCorpo(req), sessao);
    const r = await rodarSincronizacao('manual', null);
    return json(res, 200, { ok: true, rodou: !!r });
  }
  if (p === '/api/integracoes/salesforce-conta' && m === 'POST') {
    await dados.salesforceContaMapear(await lerCorpo(req), sessao);
    const r = await rodarSincronizacao('manual', null);
    return json(res, 200, { ok: true, rodou: !!r });
  }
  // diferença planilha × sistema: manter a planilha ou usar o valor do Salesforce / Prevision
  if (p === '/api/integracoes/diferenca' && m === 'POST') {
    const b = await lerCorpo(req);
    await dados.decidirDiferenca(b.id, b.decisao, sessao);
    console.log(`[integracoes] diferença ${b.id}: ${b.decisao} por ${sessao.login}`);
    agendarPainel();
    return json(res, 200, { ok: true });
  }

  if (p === '/api/painel' && m === 'GET') return json(res, 200, infoPainel());
  if (p === '/api/painel/atualizar' && m === 'POST') {
    if (!['editor', 'admin'].includes(sessao.nivel)) return json(res, 403, { erro: 'seu acesso é só de consulta' });
    clearTimeout(painel.timer); painel.agendado = false;
    gerarPainel();
    return json(res, 202, infoPainel());
  }

  // Conferência planilha × sistema (.xlsx), para validar a migração linha a linha
  if (p === '/api/conferencia' && m === 'GET') {
    const arq = path.join(os.tmpdir(), `conferencia_${process.pid}_${Date.now()}.xlsx`);
    await rodarPython(CONFERENCIA_SCRIPT, [arq]);
    const buf = fs.readFileSync(arq); fs.unlink(arq, () => {});
    res.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="conferencia_planilha_${new Date().toISOString().slice(0, 10)}.xlsx"`, 'Cache-Control': 'no-store' });
    return res.end(buf);
  }

  // Recarregar da planilha: SÓ na homologação, só admin. Refaz a carga a partir
  // da planilha de agora e substitui cadastros e lançamentos (usuários ficam).
  if (p === '/api/recarga' && m === 'GET') {
    if (AMBIENTE !== 'homologacao' || !admin) return soAdmin();
    return json(res, 200, await dados.feitosNoSistema());
  }
  if (p === '/api/recarga' && m === 'POST') {
    if (AMBIENTE !== 'homologacao') return json(res, 403, { erro: 'recarga da planilha só existe na homologação' });
    if (!admin) return soAdmin();
    const b = await lerCorpo(req);
    const feitos = await dados.feitosNoSistema();
    if (feitos.alteracoes > 0 && b.confirmo !== 'APAGAR') {
      return json(res, 409, { erro: `há ${feitos.alteracoes} alteração(ões) feitas no sistema por ${feitos.pessoas} pessoa(s); a recarga apaga tudo isso`, feitos });
    }
    const saidaCarga = await rodarPython(CARGA_SCRIPT, []);
    const antes = await dados.recarregar(fs.readFileSync(CARGA_SQL, 'utf8'), sessao.login);
    console.log(`[recarga] por ${sessao.login}; perdidas ${JSON.stringify(antes)}`);
    gerarPainel();
    return json(res, 200, { ok: true, relatorio: fs.readFileSync(path.join(path.dirname(CARGA_SCRIPT), 'relatorio_carga.txt'), 'utf8'), saida: saidaCarga.slice(-200) });
  }

  // gravações de lançamento: o painel é regerado logo depois
  let r;
  if ((r = /^\/api\/lancamentos\/(\w+)$/.exec(p)) && m === 'PUT') {
    const b = await lerCorpo(req);
    const resumo = await dados.salvarGrade(r[1], b.competencia, b.linhas, sessao, b.motivo);
    console.log(`[lancamento] ${r[1]} ${b.competencia} por ${sessao.login}: ${JSON.stringify(resumo)}`);
    agendarPainel();
    return json(res, 200, { ok: true, resumo });
  }
  if (p === '/api/ncs' && m === 'PUT') {
    const b = await lerCorpo(req);
    const resumo = await dados.salvarNcs(b.competencia, b.obra_id, b.assuntos, sessao, b.motivo);
    agendarPainel();
    return json(res, 200, { ok: true, resumo });
  }
  if (p === '/api/congelada' && m === 'GET') return json(res, 200, await dados.abaCongelada(url.searchParams.get('aba')));
  // relatório de impressão (.xls/.xlsx exportado do sistema de impressão): lê, mostra a prévia e, com aplicar, grava
  if (p === '/api/impressoes/relatorio' && m === 'POST') {
    if (!['editor', 'admin'].includes(sessao.nivel)) return json(res, 403, { erro: 'seu acesso é só de consulta' });
    const b = await lerCorpo(req);
    const nome = path.basename(String(b.nome || 'relatorio'));
    if (!/\.(xls|xlsx)$/i.test(nome)) return json(res, 400, { erro: 'envie o arquivo do relatório como veio do sistema de impressão (.xls ou .xlsx)' });
    const buf = Buffer.from(String(b.conteudo || ''), 'base64');
    if (buf.length < 200) return json(res, 400, { erro: 'arquivo vazio' });
    const tmp = path.join(os.tmpdir(), 'impressao_' + crypto.randomBytes(8).toString('hex') + path.extname(nome).toLowerCase());
    fs.writeFileSync(tmp, buf);
    let rel;
    try {
      rel = await new Promise((resolve, reject) => {
        execFile(PYTHON, ['-I', RELATORIO_SCRIPT, tmp], { cwd: path.dirname(RELATORIO_SCRIPT), timeout: 60000, windowsHide: true, maxBuffer: 8 * 1024 * 1024, encoding: 'buffer' }, (err, out, errOut) => {
          if (err) {
            const msg = Buffer.from(errOut || '').toString('utf8').trim().split(/\r?\n/).pop();
            console.error(`[impressoes] leitura de ${nome} falhou: ${err.message}\n${Buffer.from(errOut || '').toString('utf8').slice(-1500)}`);
            return reject(Object.assign(new Error(msg || 'não consegui ler o arquivo'), { status: 400, uso: true }));
          }
          try { resolve(JSON.parse(Buffer.from(out).toString('utf8'))); } catch (e) { reject(Object.assign(new Error('leitura do relatório devolveu algo inesperado'), { status: 500 })); }
        });
      });
    } finally { fs.unlink(tmp, () => {}); }
    const esperado = b.tela === 'impressoes_pessoa' ? 'usuarios' : 'conta';
    if (rel.tipo !== esperado) {
      return json(res, 400, { erro: rel.tipo === 'usuarios'
        ? 'Este é o relatório por usuários. Suba-o na tela Impressões por pessoa.'
        : 'Este é o relatório por conta. Suba-o na tela Impressões ou Impressões por setor.' });
    }
    const r2 = await dados.relatorioImpressao(rel, { aplicar: !!b.aplicar, arquivo: nome }, sessao, b.motivo);
    if (b.aplicar) { console.log(`[impressoes] ${rel.tipo} ${rel.competencia} por ${sessao.login}: ${JSON.stringify(r2)}`); agendarPainel(); }
    return json(res, 200, b.aplicar ? { ok: true, resumo: r2, competencia: rel.competencia } : r2);
  }
  // abas que são listas (Auditoria, Plano de Ação, Emergências, Impressões, ...)
  if ((r = /^\/api\/lista\/(\w+)$/.exec(p)) && m === 'GET') {
    return json(res, 200, await dados.listar(r[1], { competencia: url.searchParams.get('competencia'), ano: url.searchParams.get('ano') }));
  }
  if ((r = /^\/api\/lista\/(\w+)$/.exec(p)) && m === 'PUT') {
    const b = await lerCorpo(req);
    const resumo = await dados.salvarLista(r[1], b, b.linhas, sessao, b.motivo);
    console.log(`[lista] ${r[1]} ${b.competencia || b.ano || ''} por ${sessao.login}: ${JSON.stringify(resumo)}`);
    agendarPainel();
    return json(res, 200, { ok: true, resumo });
  }
  if (p === '/api/ruido' && m === 'PUT') {
    const b = await lerCorpo(req);
    const resumo = await dados.salvarRuido(b.competencia, b.linhas, sessao, b.motivo);
    agendarPainel();
    return json(res, 200, { ok: true, resumo });
  }

  // daqui para baixo: só administradores
  if ((r = /^\/api\/cadastros\/(\w+)$/.exec(p)) && m === 'PUT') {
    if (!admin) return soAdmin();
    const r2 = await dados.salvarCadastro(r[1], await lerCorpo(req), sessao); agendarPainel(); return json(res, 200, { ok: true, ...r2 });
  }
  if (p === '/api/metas' && m === 'PUT') {
    if (!admin) return soAdmin();
    const r2 = await dados.salvarMeta(await lerCorpo(req), sessao); agendarPainel(); return json(res, 200, { ok: true, ...r2 });
  }
  if ((r = /^\/api\/parametros\/(\w+)$/.exec(p)) && m === 'PUT') {
    if (!admin) return soAdmin();
    const b = await lerCorpo(req);
    const r2 = await dados.salvarParametro(r[1], b.valor, sessao); agendarPainel(); return json(res, 200, { ok: true, ...r2 });
  }
  if (p === '/api/fechamento' && m === 'POST') {
    if (!admin) return soAdmin();
    const b = await lerCorpo(req);
    await dados.fecharMes(b.competencia, !!b.fechar, sessao);
    return json(res, 200, { ok: true });
  }
  if (p === '/api/usuarios' && m === 'GET') {
    if (!admin) return soAdmin();
    return json(res, 200, { usuarios: await dados.usuarios(), fixos: ADMIN_LOGINS });
  }
  if (p === '/api/usuarios/excluir' && m === 'POST') {
    if (!admin) return soAdmin();
    const b = await lerCorpo(req);
    const alvo = ldap.normalizarLogin(b.login);
    if (ADMIN_LOGINS.includes(alvo)) return json(res, 400, { erro: 'administrador fixo (.env): não se altera pela tela' });
    if (alvo === sessao.login) return json(res, 400, { erro: 'você não pode excluir o seu próprio acesso' });
    await dados.excluirUsuario(alvo, sessao.login);
    await derrubarSessoes(alvo);
    return json(res, 200, { ok: true });
  }
  if (p === '/api/usuarios/decidir' && m === 'POST') {
    if (!admin) return soAdmin();
    const b = await lerCorpo(req);
    const alvo = ldap.normalizarLogin(b.login);
    if (ADMIN_LOGINS.includes(alvo)) return json(res, 400, { erro: 'administrador fixo (.env): não se altera pela tela' });
    await dados.decidirUsuario(alvo, b, sessao.login);
    await derrubarSessoes(alvo);
    return json(res, 200, { ok: true });
  }
  return json(res, 404, { erro: 'rota desconhecida' });
}

// ── Sincronização agendada (Salesforce + Prevision) ─────────────────────────
// Às 09:00, 13:00 e 17:00 (Brasília). Se o servidor estava fora do ar no horário,
// a rodada acontece assim que ele volta (a mais recente já vencida do dia que ainda não rodou).
const HORARIOS_SINC = ['09:00', '13:00', '17:00'];
async function rodarSincronizacao(origem, slot) {
  const r = await dados.sincronizar(origem, slot);
  if (r && r.mudou) agendarPainel(); // novos dados: o painel é regerado em seguida
  if (r) console.log(`[sincronizacao] ${origem}${slot ? ' ' + slot : ''}: ${JSON.stringify({ salesforce: r.salesforce && (r.salesforce.erro || { incluidas: r.salesforce.incluidas, vinculadas: r.salesforce.vinculadas, pendentes: r.salesforce.total_pendentes }), prevision: r.prevision && (r.prevision.erro || { incluidas: r.prevision.incluidas, atualizadas: r.prevision.atualizadas, divergentes: r.prevision.divergentes.length }) })}`);
  return r;
}
async function tickSincronizacao() {
  try {
    const f = Object.fromEntries(new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Sao_Paulo', hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' })
      .formatToParts(new Date()).map(x => [x.type, x.value]));
    const dia = `${f.year}-${f.month}-${f.day}`, hhmm = `${f.hour}:${f.minute}`;
    const vencidos = HORARIOS_SINC.filter(h => h <= hhmm);
    if (!vencidos.length) return;
    const slot = `${dia} ${vencidos[vencidos.length - 1]}`;
    if (await dados.sincronizacaoFeita(slot)) return;
    await rodarSincronizacao('agendada', slot);
  } catch (e) { console.error('[sincronizacao] ' + e.message); }
}

const servidor = http.createServer((req, res) => {
  rotear(req, res).catch(err => {
    if (err.uso) return json(res, err.status || 400, { erro: err.message });
    if (err.status) return json(res, err.status, { erro: err.message });
    if (err.code === '23503') return json(res, 400, { erro: 'referência inválida (obra, tipo ou unidade inexistente)' });
    if (err.code === '23514') return json(res, 400, { erro: 'valor fora do permitido' });
    console.error('[erro] ' + req.method + ' ' + req.url + ': ' + (err.stack || err));
    if (!res.headersSent) json(res, 500, { erro: 'erro interno (ver logs\\servidor.log)' });
  });
});
servidor.listen(PORT, '0.0.0.0', () => {
  if (!fs.existsSync(path.join(PAINEL_DIR, 'painel.html'))) gerarPainel();
  if (process.env.SYNC_AUTO !== '0') { setTimeout(tickSincronizacao, 20000); setInterval(tickSincronizacao, 60000); }
  console.log(`[servidor] Lançamentos Ambientais ${VERSAO} em http://0.0.0.0:${PORT} · ${AMBIENTE || 'producao'} · banco ${dados.nome} · AD ${ldap.ldapConfigurado ? 'configurado' : 'NÃO configurado (LDAP_* em branco)'}`);
});
