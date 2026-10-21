# Automação de relatórios TrueSafe

Automação em Node.js que acessa o TrueSafe Cloud, gera relatórios de acessos
para os períodos de jantar e almoço, baixa os arquivos exportados e adiciona
os dados tratados a planilhas Excel.

## O que a automação faz

- acessa o TrueSafe Cloud e realiza o login;
- filtra o bloqueio configurado;
- gera relatórios de jantar e almoço em uma única sessão;
- limpa os arquivos exportados, mantém as 11 colunas principais e remove
  duplicidades por usuário e documento;
- grava cada dia em uma aba `DD-MM-YY`, sem sobrescrever abas existentes;
- identifica os dias pendentes até ontem e processa cada perfil separadamente.

## Requisitos

- Node.js 18 ou superior;
- acesso autorizado ao TrueSafe Cloud;
- Google Chrome instalado quando `HEADLESS=false` (padrão).

## Instalação

```bash
npm install
npx playwright install chrome
```

## Configuração

Copie `.env.example` para `.env` e preencha os valores. O arquivo `.env` não
deve ser publicado nem compartilhado, pois contém credenciais.

```env
TRUESAFE_URL=https://www.truesafeweb.com.br
TRUESAFE_EMAIL=seu-email-de-login
TRUESAFE_SENHA=sua-senha

TIME_ZONE=America/Sao_Paulo

EXCEL_JANTAR_PDZ_PATH=C:\caminho\para\Relatorios-jantar.xlsx
EXCEL_ALMOCO_PDZ_PATH=C:\caminho\para\Relatorios-almoco.xlsx
DOWNLOAD_PATH=C:\caminho\para\downloads
HEADLESS=false
```

Os dois arquivos definidos em `EXCEL_JANTAR_PDZ_PATH` e
`EXCEL_ALMOCO_PDZ_PATH` precisam existir e ser arquivos `.xlsx` válidos. A pasta
de `DOWNLOAD_PATH` é criada automaticamente quando necessário.

### Horários e bloqueios

As variáveis abaixo são opcionais. Se não forem informadas, serão usados os
valores padrão indicados:

```env
HORA_INICIAL_JANTAR=22:30
HORA_FINAL_JANTAR=02:30
BLOQUEIO_JANTAR=Catraca Refeitorio

HORA_INICIAL_ALMOCO=10:30
HORA_FINAL_ALMOCO=14:30
BLOQUEIO_ALMOCO=Catraca Refeitorio
```

Também é possível definir `HORA_INICIAL`, `HORA_FINAL` e `BLOQUEIO` como
valores comuns. Eles funcionam como alternativa para o perfil de jantar e para
o bloqueio do almoço quando não houver uma configuração específica.

## Execução

```bash
npm start
```

O login é realizado uma vez. O período de jantar cruza a madrugada: a data
final é o dia seguinte à data inicial. O período de almoço usa a mesma data
para início e fim. Em ambos os casos, o dia atual não é processado; o limite é
sempre o dia anterior no fuso definido em `TIME_ZONE`.

Na primeira execução de um arquivo sem abas no formato `DD-MM-YY`, somente o
relatório de ontem é gerado. Nas execuções seguintes, a automação continua a
partir da maior data já registrada e processa eventuais dias pendentes.

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

- `browser.js`: login, filtros, períodos e download dos relatórios;
- `excel.js`: leitura, limpeza, deduplicação e formatação dos dados;
- `automation.service.js`: configuração dos perfis e fluxo da automação;
- `index.js`: ponto de entrada da aplicação.

## Arquivos locais e segurança

O `.gitignore` mantém fora do Git as credenciais, dependências e arquivos
gerados localmente: `.env`, `node_modules/`, `downloads/`, `saida/` e
`brutos/`.

    