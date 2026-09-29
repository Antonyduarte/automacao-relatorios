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

`browser.js` concentra login, filtros e download. `excel.js` processa o
relatorio, remove duplicatas comparando a linha completa, cria uma tabela real
e insere uma nova aba no arquivo principal. A aba recebe o nome da Data
Inicial no formato `DD-MM-YY`.

Variaveis adicionais:

```env
TIME_ZONE=America/Sao_Paulo
EXCEL_PRINCIPAL_PATH=C:\Users\opton\Documents\Automacao\saida\Relatorios-jantar.xlsx
DOWNLOAD_PATH=C:\Users\opton\Documents\Automacao\downloads
HEADLESS=false
```

As datas sao calculadas automaticamente no fuso de `TIME_ZONE`: Data Inicial
e ontem, Data Final e hoje. O arquivo principal precisa existir e ser um `.xlsx` valido. A
automacao interrompe sem alterar o arquivo quando a aba `DD-MM-YY` ja existe.
