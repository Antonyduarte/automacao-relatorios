# Automação — Relatório de Acessos (Catraca Refeitório) TrueSafe Cloud

Script Node.js que abre o Google Chrome, faz login no TrueSafe Cloud, gera o
relatório de acessos filtrado (ontem 22:30 até hoje 02:30, só "Catraca
Refeitorio"), baixa o arquivo e entrega um `.xlsx` limpo e formatado como
Tabela do Excel, com duplicadas removidas e linha de Totais.

## 1. Pré-requisitos

- Node.js 18 ou superior instalado.
- **Google Chrome instalado** na máquina (o script abre o Chrome de verdade,
  não o navegador interno do Playwright).

## 2. Instalação

```bash
cd truesafe-automation
npm install
npx playwright install chrome
```

O segundo comando garante que o Playwright sabe onde está o Chrome instalado.

## 3. Credenciais

As credenciais já estão preenchidas em `.env` (email e senha que você me
passou). Esse arquivo **não deve ser commitado/compartilhado** — o
`.gitignore` já está configurado para ignorá-lo.

```
TRUESAFE_URL=https://www.truesafeweb.com.br
TRUESAFE_EMAIL=recepcaoperdizes@bembrasil.ind.br
TRUESAFE_SENHA=98427114
```

Se a senha mudar, basta editar esse arquivo.

## 4. Rodando

```bash
npm start
```

O Chrome vai abrir visivelmente (`headless: false`) para você acompanhar. Ao
final:

- o arquivo baixado "cru" fica em `downloads/`;
- o arquivo final tratado (tabela + totais) fica em
  `saida/relatorio_catraca_refeitorio_AAAA-MM-DD.xlsx`.

## 5. O que o script faz, passo a passo

1. Abre `https://www.truesafeweb.com.br` e tenta logar.
2. Vai direto para `relatorio_acessos_cloud.php` (a tela "Acessos + Filtros").
3. Preenche `Data Inicial` = ontem e `Data Final` = hoje (campos `#data_inicio`
   e `#data_fim`, confirmados no HTML que você enviou).
4. Preenche `Horário Inicial` = 22:30 e `Horário Final` = 02:30 (`#hora_ini`,
   `#hora_fim`).
5. Vai na aba "Bloqueios", desmarca "Exibir Todos" e marca só
   "Catraca Refeitorio".
6. Clica em "Gerar Excel" (o formulário abre em nova aba com
   `target="_blank"` — o script já trata isso esperando o popup e o
   download).
7. Processa o arquivo baixado:
   - remove as 2 primeiras linhas;
   - limpa o conteúdo da coluna A da última linha;
   - monta uma Tabela do Excel começando em **A2**;
   - remove linhas duplicadas;
   - adiciona a linha de **Totais** (soma nas colunas numéricas, contagem na
     primeira coluna).

## 6. Pontos que podem precisar de ajuste fino

Eu confirmei os seletores de **login (parcialmente), campos de data/hora,
aba Bloqueios e botão Gerar Excel** direto no HTML que você mandou
(`index.txt` e `acessos_cloud.txt`). Só uma coisa eu não pude confirmar:

- **Tela de login**: os dois arquivos HTML que você enviou já são páginas
  internas (pós-login) — não tinham o formulário de login em si. A função
  `login()` tenta seletores comuns (`input[type="email"]`,
  `input[name="senha"]`, botão "Entrar", etc.). **Se o login falhar**, abra o
  site em uma aba anônima, clique com o botão direito no campo de e-mail e
  senha → "Inspecionar", e me mande esse trecho de HTML (igual você fez com
  as outras duas telas) que eu ajusto os seletores certinho.

- **Estrutura exata do arquivo baixado**: como não tenho o arquivo real
  gerado pelo "Gerar Excel" em mãos, a lógica de "tabela a partir de A2" foi
  implementada literalmente como você pediu (linha 1 fica em branco de
  propósito, cabeçalho na linha 2, dados a partir da linha 3). Rode uma vez
  e, se a linha 1 do resultado final não estiver como você espera (por
  exemplo, se sobrar algum texto de título ali), me avise o que apareceu que
  eu ajusto a função `processarPlanilha()`.

## 7. Estrutura dos arquivos

```
truesafe-automation/
├── automatizar_relatorio.js   # script principal
├── package.json
├── .env                       # credenciais (não versionar)
├── .gitignore
├── downloads/                 # arquivo baixado "cru" (criado ao rodar)
└── saida/                     # arquivo .xlsx final (criado ao rodar)
```
