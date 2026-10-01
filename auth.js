const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

const DATA_DIR = path.join(__dirname, 'data');
const USERS_FILE = path.join(DATA_DIR, 'users.json');
const CONFIG_FILE = path.join(DATA_DIR, 'config.json');

const SESSION_MS = 8 * 60 * 60 * 1000;
const PENDING_MS = 2 * 60 * 1000;
const MAX_FAILS = 5;
const LOCK_MS = 5 * 60 * 1000;

const sessions = new Map(); // sid -> { usuario, perfil, exp }
const pendings = new Map(); // token -> { usuario, exp, fails }
const fails = new Map(); // usuario|ip -> { count, until }

function hash(secret, salt = crypto.randomBytes(16).toString('hex')) {
  return { salt, hash: crypto.scryptSync(String(secret), salt, 64).toString('hex') };
}

function check(secret, rec) {
  const h = crypto.scryptSync(String(secret), rec.salt, 64);
  const ref = Buffer.from(rec.hash, 'hex');
  return h.length === ref.length && crypto.timingSafeEqual(h, ref);
}

const DUMMY = hash('x');

function readJson(file, fallback) {
  try { return JSON.parse(fs.readFileSync(file, 'utf8')); } catch (e) { return fallback; }
}

function writeJson(file, data) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  const tmp = file + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(data, null, 2));
  fs.renameSync(tmp, file);
}

const loadUsers = () => readJson(USERS_FILE, []);
const saveUsers = (u) => writeJson(USERS_FILE, u);
const loadConfig = () => readJson(CONFIG_FILE, null);

const same = (a, b) => String(a).toLowerCase() === String(b).toLowerCase();
const findUser = (usuario) => loadUsers().find((u) => same(u.usuario, usuario));
const publicUser = (u) => ({ usuario: u.usuario, perfil: u.perfil, ativo: u.ativo, criadoEm: u.criadoEm });

// Usado pelo script criar-admin.js: grava o administrador e a chave de admin (somente hashes).
function setupAdmin(usuario, senha, chave) {
  const users = loadUsers().filter((u) => !same(u.usuario, usuario));
  users.push({ usuario, ...hash(senha), perfil: 'admin', ativo: true, criadoEm: new Date().toISOString() });
  saveUsers(users);
  writeJson(CONFIG_FILE, { adminKey: hash(chave) });
}

function isConfigured() {
  return !!loadConfig() && loadUsers().some((u) => u.perfil === 'admin');
}

// ---- bloqueio por tentativas ----
function lockKey(usuario, ip) { return String(usuario).toLowerCase() + '|' + ip; }
function isLocked(usuario, ip) {
  const f = fails.get(lockKey(usuario, ip));
  return !!f && f.until > Date.now();
}
function recordFail(usuario, ip) {
  const k = lockKey(usuario, ip);
  const f = fails.get(k) || { count: 0, until: 0 };
  f.count += 1;
  if (f.count >= MAX_FAILS) { f.until = Date.now() + LOCK_MS; f.count = 0; }
  fails.set(k, f);
}
function clearFails(usuario, ip) { fails.delete(lockKey(usuario, ip)); }

// ---- login ----
function verifyPassword(usuario, senha) {
  const u = findUser(usuario);
  const ok = check(senha, u || DUMMY);
  return u && u.ativo && ok ? u : null;
}

function newToken() { return crypto.randomBytes(32).toString('hex'); }

function createSession(u) {
  const sid = newToken();
  sessions.set(sid, { usuario: u.usuario, perfil: u.perfil, exp: Date.now() + SESSION_MS });
  return sid;
}

function createPending(u) {
  const token = newToken();
  pendings.set(token, { usuario: u.usuario, exp: Date.now() + PENDING_MS, fails: 0 });
  return token;
}

function verifyAdminKey(token, chave, ip) {
  const p = pendings.get(token);
  if (!p || p.exp < Date.now()) { pendings.delete(token); return { erro: 'expirado' }; }
  if (isLocked(p.usuario, ip)) return { erro: 'bloqueado' };
  const cfg = loadConfig();
  if (!cfg || !check(chave, cfg.adminKey)) {
    p.fails += 1;
    recordFail(p.usuario, ip);
    if (p.fails >= MAX_FAILS) pendings.delete(token);
    return { erro: 'invalida' };
  }
  pendings.delete(token);
  clearFails(p.usuario, ip);
  const u = findUser(p.usuario);
  if (!u || !u.ativo) return { erro: 'invalida' };
  return { sid: createSession(u) };
}

// ---- sessão / cookies ----
function parseCookies(req) {
  const out = {};
  (req.headers.cookie || '').split(';').forEach((p) => {
    const i = p.indexOf('=');
    if (i > 0) out[p.slice(0, i).trim()] = decodeURIComponent(p.slice(i + 1).trim());
  });
  return out;
}

function sessionOf(req) {
  const sid = parseCookies(req).sid;
  const s = sid && sessions.get(sid);
  if (!s) return null;
  if (s.exp < Date.now()) { sessions.delete(sid); return null; }
  const u = findUser(s.usuario);
  if (!u || !u.ativo) { sessions.delete(sid); return null; }
  s.exp = Date.now() + SESSION_MS;
  return { sid, usuario: u.usuario, perfil: u.perfil };
}

const setCookie = (res, sid) =>
  res.setHeader('Set-Cookie', `sid=${sid}; HttpOnly; SameSite=Strict; Path=/; Max-Age=${SESSION_MS / 1000}`);
const clearCookie = (res) => res.setHeader('Set-Cookie', 'sid=; HttpOnly; SameSite=Strict; Path=/; Max-Age=0');

function destroyUserSessions(usuario) {
  for (const [sid, s] of sessions) if (same(s.usuario, usuario)) sessions.delete(sid);
}

// ---- middlewares ----
function requireAuth(req, res, next) {
  const s = sessionOf(req);
  if (s) { req.user = s; return next(); }
  if (req.path.startsWith('/api/')) return res.status(401).json({ erro: 'Sessão expirada. Entre novamente.' });
  res.redirect('/login.html');
}

function requireAdmin(req, res, next) {
  requireAuth(req, res, () => {
    if (req.user.perfil === 'admin') return next();
    if (req.path.startsWith('/api/')) return res.status(403).json({ erro: 'Acesso restrito ao administrador.' });
    res.redirect('/');
  });
}

// ---- gestão de usuários (admin) ----
const USER_RE = /^[a-z0-9._-]{3,30}$/i;

function listUsers() { return loadUsers().map(publicUser); }

function createUser(usuario, senha, perfil) {
  if (!USER_RE.test(usuario || '')) return { erro: 'Usuário: 3 a 30 caracteres (letras, números, ponto, hífen ou sublinhado).' };
  if (String(senha || '').length < 6) return { erro: 'A senha precisa ter ao menos 6 caracteres.' };
  const users = loadUsers();
  if (users.some((u) => same(u.usuario, usuario))) return { erro: 'Já existe um usuário com esse nome.' };
  users.push({ usuario, ...hash(senha), perfil: perfil === 'admin' ? 'admin' : 'usuario', ativo: true, criadoEm: new Date().toISOString() });
  saveUsers(users);
  return { ok: true };
}

function updateUser(usuario, patch, atual) {
  const users = loadUsers();
  const u = users.find((x) => same(x.usuario, usuario));
  if (!u) return { erro: 'Usuário não encontrado.' };
  if (patch.senha != null) {
    if (String(patch.senha).length < 6) return { erro: 'A senha precisa ter ao menos 6 caracteres.' };
    Object.assign(u, hash(patch.senha));
  }
  if (patch.ativo != null) {
    if (same(u.usuario, atual) && !patch.ativo) return { erro: 'Você não pode desativar o próprio usuário.' };
    u.ativo = !!patch.ativo;
  }
  saveUsers(users);
  if (!u.ativo || patch.senha != null) destroyUserSessions(u.usuario);
  return { ok: true };
}

function deleteUser(usuario, atual) {
  if (same(usuario, atual)) return { erro: 'Você não pode excluir o próprio usuário.' };
  const users = loadUsers();
  const rest = users.filter((u) => !same(u.usuario, usuario));
  if (rest.length === users.length) return { erro: 'Usuário não encontrado.' };
  saveUsers(rest);
  destroyUserSessions(usuario);
  return { ok: true };
}

module.exports = {
  setupAdmin, isConfigured, isLocked, recordFail, clearFails, verifyPassword,
  createSession, createPending, verifyAdminKey, sessionOf, setCookie, clearCookie, parseCookies,
  sessions, requireAuth, requireAdmin, listUsers, createUser, updateUser, deleteUser,
};
