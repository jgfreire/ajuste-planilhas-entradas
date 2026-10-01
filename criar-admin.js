// Uso: node criar-admin.js <usuario> <senha> <chave-de-admin>
// Grava somente hashes em data/users.json e data/config.json.
const auth = require('./auth');
const [usuario, senha, chave] = process.argv.slice(2);
if (!usuario || !senha || !chave) {
  console.log('Uso: node criar-admin.js <usuario> <senha> <chave-de-admin>');
  process.exit(1);
}
auth.setupAdmin(usuario, senha, chave);
console.log('Administrador "' + usuario + '" configurado.');
