# Automação de relatórios TrueSafe

Automação em Node.js que acessa o TrueSafe Cloud, gera relatórios de acessos,
baixa os arquivos e adiciona os dados tratados a planilhas Excel de jantar e
almoço.

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

EXCEL_JANTAR_PDZ_PATH=C:\caminho\para\Relatorios-jantar.xlsx
EXCEL_ALMOCO_PDZ_PATH=C:\caminho\para\Relatorios-almoco.xlsx
DOWNLOAD_PATH=C:\caminho\para\downloads
HEADLESS=false
```

Os arquivos informados em `EXCEL_JANTAR_PDZ_PATH` e
`EXCEL_ALMOCO_PDZ_PATH` precisam existir e ser arquivos `.xlsx` válidos. O
diretório definido em `DOWNLOAD_PATH` será criado automaticamente, caso ainda
não exista.

As configurações abaixo são opcionais e permitem substituir os valores padrão
de cada perfil:

```env
HORA_INICIAL_JANTAR=22:30
HORA_FINAL_JANTAR=02:30
BLOQUEIO_JANTAR=Catraca Refeitorio

HORA_INICIAL_ALMOCO=10:30
HORA_FINAL_ALMOCO=14:30
BLOQUEIO_ALMOCO=Catraca Refeitorio
```

Também é possível usar `HORA_INICIAL`, `HORA_FINAL` e `BLOQUEIO` como valores
comuns. Eles funcionam como alternativa para o perfil de jantar.

## Execução

```bash
npm start
```

O login é feito uma única vez e os relatórios são processados em sequência para
os perfis de jantar e almoço. O período é calculado no fuso de `TIME_ZONE`,
usando o dia anterior como data inicial.

O relatório de jantar considera o período entre `HORA_INICIAL_JANTAR` e
`HORA_FINAL_JANTAR` cruzando a madrugada. O relatório de almoço usa o mesmo dia
para as datas inicial e final.

Quando a planilha já possui abas de relatórios, o programa identifica a última
data processada e tenta gerar somente os dias pendentes, até o dia anterior.
Cada relatório é salvo em uma aba no formato `DD-MM-YY`. Se a aba já existir,
ela não é sobrescrita.

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
- `automation.service.js`: configuração dos perfis e fluxo principal da automação;
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
