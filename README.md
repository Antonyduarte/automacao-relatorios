# Automacao do relatorio TrueSafe

Automacao Node.js que acessa o TrueSafe Cloud, gera o relatorio de acessos,
baixa o arquivo e cria uma planilha Excel tratada.

## Instalar e executar

```bash
npm install
npx playwright install chrome
npm start
```

As credenciais sao lidas do arquivo `.env`:

```env
TRUESAFE_URL=https://www.truesafeweb.com.br
TRUESAFE_EMAIL=seu-email
TRUESAFE_SENHA=sua-senha
RELATORIOS_PASTA=C:\\Users\\opton\\Documents\\relatorios
```

## Estrutura

```text
src/
├── browser/
│   └── browser.js
├── excel/
│   └── excel.js
├── services/
│   └── automation.service.js
└── index.js
```

`browser.js` concentra login, filtros e download. `excel.js` le CSV/XLSX e
monta a planilha final. A remocao de duplicatas usa a combinacao `Usuario` e
`Documento`.
