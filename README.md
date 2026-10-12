# Automação de relatórios TrueSafe

Automação em Node.js que acessa o TrueSafe Cloud, gera relatórios de acessos,
baixa os arquivos e adiciona os dados tratados a uma planilha Excel.

## Requisitos

- Node.js 18 ou superior;
- acesso válido ao TrueSafe Cloud;
- Google Chrome instalado, quando a execução for feita com `HEADLESS=false`.

## Instalação

```bash
npm install
npx playwright install chrome
```

## Configuração

Crie um arquivo `.env` na raiz do projeto. Nunca publique esse arquivo nem
compartilhe suas credenciais.

```env
TRUESAFE_URL=https://www.truesafeweb.com.br
TRUESAFE_EMAIL=seu-email-de-login
TRUESAFE_SENHA=sua-senha

TIME_ZONE=America/Sao_Paulo
HORA_INICIAL=22:30
HORA_FINAL=02:30
BLOQUEIO=Catraca Refeitorio

EXCEL_JANTAR_PDZ_PATH=C:\caminho\para\Relatorios-jantar.xlsx
DOWNLOAD_PATH=C:\caminho\para\downloads
HEADLESS=false
```

O arquivo principal informado em `EXCEL_JANTAR_PDZ_PATH` precisa existir e ser
um arquivo `.xlsx` válido. O diretório definido em `DOWNLOAD_PATH` será criado
automaticamente, caso ainda não exista.

## Execução

```bash
npm start
```

Por padrão, a automação calcula o período no fuso de `TIME_ZONE`, usando o dia
anterior como data inicial e o dia atual como data final. Os horários são
controlados por `HORA_INICIAL` e `HORA_FINAL`.

Quando a planilha já possui abas de relatórios, o programa identifica a última
data processada e tenta gerar somente os dias pendentes. Cada relatório é salvo
em uma aba no formato `DD-MM-YY`. Se a aba já existir, ela não é sobrescrita.

Para executar sem abrir a janela do Chrome:

```env
HEADLESS=true
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

- `browser.js`: login, filtros, período e download do relatório;
- `excel.js`: leitura, limpeza, deduplicação e formatação dos dados;
- `automation.service.js`: configuração e fluxo principal da automação;
- `index.js`: ponto de entrada da aplicação.

## Arquivos ignorados

Credenciais, dependências e relatórios locais são mantidos fora do Git por meio
do `.gitignore`:

- `.env`;
- `node_modules/`;
- `downloads/`;
- `saida/`;
- `brutos/`.

Antes de publicar o projeto, confirme que nenhum relatório contendo nomes,
documentos ou outros dados pessoais foi adicionado ao repositório.

## Aviso de segurança e uso

Use a automação somente com autorização para acessar o sistema e os dados
gerados. As credenciais devem ser armazenadas apenas no `.env` ou em um gerenciador
seguro de segredos. Se uma senha for exposta em um commit, troque-a imediatamente.
