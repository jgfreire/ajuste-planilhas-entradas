const express = require('express');
const multer = require('multer');
const JSZip = require('jszip');
const path = require('path');
const { ajustarPlanilha } = require('./processor');
const auth = require('./auth');

const app = express();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 100 * 1024 * 1024, files: 500 },
});
const PORT = process.env.PORT || 3010;

app.use(express.json({ limit: '10kb' }));
app.use(express.static(path.join(__dirname, 'public'))); // somente a tela de login

// ---- páginas protegidas ----
app.get('/', auth.requireAuth, (req, res) => res.sendFile(path.join(__dirname, 'app', 'index.html')));
app.get('/admin', auth.requireAdmin, (req, res) => res.sendFile(path.join(__dirname, 'app', 'admin.html')));

// ---- login ----
const GENERICO = 'Usuário ou senha inválidos.';

app.post('/api/login', (req, res) => {
  const { usuario, senha } = req.body || {};
  if (!usuario || !senha) return res.status(400).json({ erro: GENERICO });
  if (!auth.isConfigured()) return res.status(503).json({ erro: 'Sistema sem administrador cadastrado. Rode criar-admin.js.' });
  if (auth.isLocked(usuario, req.ip)) return res.status(429).json({ erro: 'Muitas tentativas. Aguarde alguns minutos.' });
  const u = auth.verifyPassword(usuario, senha);
  if (!u) { auth.recordFail(usuario, req.ip); return res.status(401).json({ erro: GENERICO }); }
  if (u.perfil === 'admin') return res.json({ etapa: 2, pendente: auth.createPending(u) });
  auth.clearFails(usuario, req.ip);
  auth.setCookie(res, auth.createSession(u));
  res.json({ ok: true });
});

// Segunda etapa (só aparece para perfil administrativo): código de verificação.
app.post('/api/login/verificar', (req, res) => {
  const { pendente, codigo } = req.body || {};
  const r = auth.verifyAdminKey(pendente, String(codigo || ''), req.ip);
  if (r.erro === 'expirado') return res.status(401).json({ erro: 'Tempo esgotado. Faça login novamente.', reiniciar: true });
  if (r.erro === 'bloqueado') return res.status(429).json({ erro: 'Muitas tentativas. Aguarde alguns minutos.', reiniciar: true });
  if (r.erro) return res.status(401).json({ erro: 'Código inválido.' });
  auth.setCookie(res, r.sid);
  res.json({ ok: true });
});

app.post('/api/logout', (req, res) => {
  const sid = auth.parseCookies(req).sid;
  if (sid) auth.sessions.delete(sid);
  auth.clearCookie(res);
  res.json({ ok: true });
});

app.get('/api/me', auth.requireAuth, (req, res) => res.json({ usuario: req.user.usuario, perfil: req.user.perfil }));

// ---- gestão de usuários (somente administrador) ----
const resp = (res, r) => (r.erro ? res.status(400).json({ erro: r.erro }) : res.json(r));
app.get('/api/usuarios', auth.requireAdmin, (req, res) => res.json(auth.listUsers()));
app.post('/api/usuarios', auth.requireAdmin, (req, res) => {
  const { usuario, senha, perfil } = req.body || {};
  resp(res, auth.createUser(usuario, senha, perfil));
});
app.patch('/api/usuarios/:usuario', auth.requireAdmin, (req, res) => resp(res, auth.updateUser(req.params.usuario, req.body || {}, req.user.usuario)));
app.delete('/api/usuarios/:usuario', auth.requireAdmin, (req, res) => resp(res, auth.deleteUser(req.params.usuario, req.user.usuario)));


const utf8 = (s) => Buffer.from(s, 'latin1').toString('utf8');

// Processa em memória: nada é gravado em disco no servidor.
// Recebe uma ou várias planilhas (com o caminho relativo de cada uma) e devolve
// um .xlsx (1 arquivo) ou um .zip com a estrutura de pastas preservada.
app.post('/api/ajustar', auth.requireAuth, upload.array('planilhas'), async (req, res) => {
  const files = req.files || [];
  if (!files.length) return res.status(400).json({ erro: 'Envie ao menos um arquivo .xlsx.' });

  let caminhos = [];
  try { caminhos = JSON.parse(req.body.caminhos || '[]'); } catch (e) { /* usa o nome do arquivo */ }
  const pasta = String(req.body.pasta || '').trim();

  const resultados = [];
  const zip = new JSZip();
  let unico = null;

  for (let i = 0; i < files.length; i++) {
    const caminho = (caminhos[i] || utf8(files[i].originalname)).replace(/\\/g, '/');
    const nome = caminho.split('/').pop();
    if (nome.startsWith('~$') || !/\.xlsx$/i.test(nome)) {
      resultados.push({ caminho, status: 'ignorado', motivo: /^~\$/.test(nome) ? 'arquivo temporário do Excel' : 'não é .xlsx' });
      continue;
    }
    try {
      const { report, buffer } = await ajustarPlanilha(files[i].buffer);
      if (!buffer) {
        resultados.push({ caminho, status: 'sem-colunas', motivo: report.erro });
        continue;
      }
      const saida = caminho.replace(/\.xlsx$/i, '') + '_ajustada.xlsx';
      zip.file(saida, buffer);
      unico = { nome: saida.split('/').pop(), buffer };
      resultados.push({ caminho, status: 'ok', saida, report });
    } catch (e) {
      resultados.push({ caminho, status: 'erro', motivo: 'Não foi possível ler a planilha: ' + e.message });
    }
  }

  const ok = resultados.filter((r) => r.status === 'ok').length;
  let arquivo = null, nomeSaida = null;
  if (ok === 1 && files.length === 1) {
    arquivo = unico.buffer.toString('base64');
    nomeSaida = unico.nome;
  } else if (ok > 0) {
    arquivo = (await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })).toString('base64');
    nomeSaida = (pasta || 'planilhas') + '_ajustadas.zip';
  }
  res.json({ resultados, arquivo, nomeSaida });
});

app.use((err, req, res, next) => {
  res.status(400).json({ erro: err.code === 'LIMIT_FILE_SIZE' ? 'Algum arquivo passa de 100 MB.' : err.message });
});

app.listen(PORT, () => console.log(`Ajuste de planilha de entradas em http://localhost:${PORT}`));
