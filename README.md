# Ajuste de Planilhas de Entradas

Sistema web (Node/Express) que ajusta planilhas `.xlsx` de entradas:

- remove as colunas **Chave NF**, **Tags (etiquetas)** e **Eventos**;
- adiciona a coluna **REVENDA / USO OU CONSUMO / IMOBILIZADO**, com lista suspensa;
- deixa a saída formatada (cabeçalho, filtro, painel congelado, bordas, larguras);
- aceita uma pasta ou várias planilhas e devolve um `.zip` com as planilhas `_ajustada.xlsx`.

Os arquivos são processados em memória; nada é gravado em disco.

## Como rodar

```bash
npm install
node criar-admin.js <usuario> <senha> <chave-de-admin>   # cria o administrador (só grava hashes em data/)
node server.js                                            # http://localhost:3010
```

No Windows, `iniciar.bat` instala as dependências (se preciso) e inicia o sistema.

## Login

- Usuário comum: usuário e senha.
- Administrador: após usuário e senha corretos, a tela pede um "Código de verificação" (a chave de admin).
- O administrador tem a página extra **Usuários** (`/admin`) para cadastrar, desativar, redefinir senha e excluir usuários.

`data/` (usuários e chave, somente hashes) não é versionada.
