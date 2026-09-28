/**
 * Automação do Relatório de Acessos (TrueSafe Cloud) - Catraca Refeitório
 * ------------------------------------------------------------------------
 * O que este script faz:
 *  1. Abre o Google Chrome de verdade (via Playwright, channel: "chrome");
 *  2. Acessa https://www.truesafeweb.com.br e faz login;
 *  3. Vai em "Relatórios" > "Acessos + Filtros" (relatorio_acessos_cloud.php);
 *  4. Preenche Data Inicial = ontem, Data Final = hoje;
 *  5. Preenche Horário Inicial = 22:30, Horário Final = 02:30;
 *  6. Na aba "Bloqueios": desmarca "Exibir Todos" e marca só "Catraca Refeitorio";
 *  7. Clica em "Gerar Excel" e salva o arquivo baixado;
 *  8. Abre o arquivo baixado, remove as 2 primeiras linhas, limpa a coluna A
 *     da última linha, monta uma Tabela do Excel a partir de A2, remove
 *     linhas duplicadas e insere a linha de Totais;
 *  9. Salva o resultado final em .xlsx pronto para uso.
 *
 * Seletores usados (confirmados a partir do HTML enviado):
 *   #data_inicio, #data_fim   -> <input type="date">  (formato exigido: YYYY-MM-DD)
 *   #hora_ini,   #hora_fim    -> <input type="time">  (formato: HH:mm)
 *   aba Bloqueios             -> a[href="#bloqueios"]
 *   "Exibir Todos" (Bloqueios)-> #exibir_bloqueios_all
 *   "Catraca Refeitorio"      -> input[name="bloqueios[]"][value="Catraca Refeitorio"]
 *   Botão "Gerar Excel"       -> button[name="excel"]
 *
 * ATENÇÃO - único ponto que eu NÃO consegui confirmar pelo HTML enviado:
 *   a tela de LOGIN em si. Os dois HTMLs que você mandou (index.txt e
 *   acessos_cloud.txt) já são páginas internas, pós-login. A função login()
 *   abaixo tenta alguns seletores comuns (email/senha) e, se não encontrar
 *   nada, tenta simplesmente prosseguir (assumindo sessão já ativa).
 *   Se o login falhar, me mande o HTML da tela de login que eu ajusto os
 *   seletores certinho, igual fiz com as outras duas telas.
 */

require('dotenv').config();
const path = require('path');
const fs = require('fs');
const { chromium } = require('playwright');
const ExcelJS = require('exceljs');
const { parse: parseCsvSync } = require('csv-parse/sync');

// ----------------------------------------------------------------------
// CONFIGURAÇÃO
// ----------------------------------------------------------------------
const BASE_URL = process.env.TRUESAFE_URL || 'https://www.truesafeweb.com.br';
const EMAIL = process.env.TRUESAFE_EMAIL;
const SENHA = process.env.TRUESAFE_SENHA;

const HORA_INICIAL = '22:30';
const HORA_FINAL = '02:30';
const BLOQUEIO_DESEJADO = 'Catraca Refeitorio'; // valor exato do value= no HTML (sem acento)

const DOWNLOAD_DIR = path.join(__dirname, 'downloads');
const OUTPUT_DIR = path.join(__dirname, 'saida');

// ----------------------------------------------------------------------
// HELPERS DE DATA
// ----------------------------------------------------------------------
function pad2(n) {
  return String(n).padStart(2, '0');
}

// input type="date" exige o valor no formato YYYY-MM-DD, independente do
// formato exibido (dd/mm/aaaa) na tela.
function toISODate(date) {
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
}

function getHojeEOntem() {
  const hoje = new Date();
  const ontem = new Date(hoje);
  ontem.setDate(hoje.getDate() - 1);
  return { hojeISO: toISODate(hoje), ontemISO: toISODate(ontem) };
}

// ----------------------------------------------------------------------
// PASSOS DA AUTOMAÇÃO NO NAVEGADOR
// ----------------------------------------------------------------------

async function login(page) {
  await page.goto(BASE_URL, { waitUntil: 'domcontentloaded' });

  // Se já caiu direto no Dashboard (sessão salva/cookie), não há o que fazer.
  const jaLogado = await page.locator('text=Dashboard').first().isVisible().catch(() => false);
  if (jaLogado) {
    console.log('Sessão já autenticada, pulando tela de login.');
    return;
  }

  console.log('Tentando localizar o formulário de login...');

  // Tentativas genéricas de seletor para o campo de usuário/e-mail.
  const camposUsuario = [
    'input[type="email"]',
    'input[name="usuario"]',
    'input[name="login"]',
    'input[name="usuario"]',
    '#email',
    '#login',
    '#usuario',
  ];
  const camposSenha = [
    'input[type="password"]',
    'input[name="senha"]',
    '#senha',
    '#password',
  ];

  let campoUsuario = null;
  for (const sel of camposUsuario) {
    const loc = page.locator(sel).first();
    if (await loc.count()) {
      campoUsuario = loc;
      break;
    }
  }

  let campoSenha = null;
  for (const sel of camposSenha) {
    const loc = page.locator(sel).first();
    if (await loc.count()) {
      campoSenha = loc;
      break;
    }
  }

  if (!campoUsuario || !campoSenha) {
    console.warn(
      'ATENÇÃO: não encontrei os campos de login com os seletores padrão. ' +
      'Me envie o HTML da tela de login para eu ajustar. Prosseguindo mesmo assim...'
    );
    return;
  }

  await campoUsuario.fill(EMAIL);
  await campoSenha.fill(SENHA);

  // Botão de submit: tenta por texto comum, senão usa o botão type=submit do form.
  const botaoEntrar = page.locator(
    'button:has-text("Entrar"), button:has-text("Login"), button:has-text("Acessar"), input[type="submit"]'
  ).first();

  await Promise.all([
    page.waitForNavigation({ waitUntil: 'domcontentloaded' }).catch(() => {}),
    botaoEntrar.click(),
  ]);

  console.log('Login enviado.');
}

async function abrirRelatorioAcessosFiltros(page) {
  // Vai direto para a página do relatório (confirmado no HTML: relatorio_acessos_cloud.php).
  // Isso evita depender do hover/clique no menu lateral, que pode variar.
  const url = new URL('relatorio_acessos_cloud.php', BASE_URL).toString();
  await page.goto(url, { waitUntil: 'domcontentloaded' });

  // Confirma que a tela certa carregou.
  await page.locator('#data_inicio').waitFor({ state: 'attached' });
  console.log('Tela "Acessos + Filtros" carregada.');
}

async function preencherPeriodoEHorario(page) {
  const { hojeISO, ontemISO } = getHojeEOntem();

  await page.fill('#data_inicio', ontemISO);
  await page.fill('#data_fim', hojeISO);
  await page.fill('#hora_ini', HORA_INICIAL);
  await page.fill('#hora_fim', HORA_FINAL);

  console.log(`Período definido: ${ontemISO} ${HORA_INICIAL} até ${hojeISO} ${HORA_FINAL}`);
}

async function configurarBloqueios(page) {
  // Vai para a aba "Bloqueios"
  await page.click('a[href="#bloqueios"]');
  await page.locator('#bloqueios').waitFor({ state: 'visible' });

  // Desmarca "Exibir Todos" (isso libera os checkboxes individuais, que
  // no HTML enviado vêm com disabled="" enquanto "Exibir Todos" está ligado).
  const exibirTodos = page.locator('#exibir_bloqueios_all');
  await exibirTodos.setChecked(false, { force: true });

  // Pequena folga para o JS da página remover o "disabled" dos checkboxes.
  await page.waitForTimeout(300);

  // Marca apenas "Catraca Refeitorio".
  const catracaRefeitorio = page.locator(
    `input[name="bloqueios[]"][value="${BLOQUEIO_DESEJADO}"]`
  );
  await catracaRefeitorio.waitFor({ state: 'attached' });
  await catracaRefeitorio.setChecked(true, { force: true });

  console.log('Aba "Bloqueios" configurada: apenas Catraca Refeitório marcada.');
}

async function gerarEBaixarExcel(context, page) {
  if (!fs.existsSync(DOWNLOAD_DIR)) fs.mkdirSync(DOWNLOAD_DIR, { recursive: true });

  // O formulário do relatório tem target="_blank", então o clique abre uma
  // nova aba (popup) e o download acontece nessa nova aba.
  const [popup] = await Promise.all([
    context.waitForEvent('popup'),
    page.click('button[name="excel"]'),
  ]);

  const download = await popup.waitForEvent('download');
  const sugerido = download.suggestedFilename();
  const destino = path.join(DOWNLOAD_DIR, sugerido);
  await download.saveAs(destino);

  console.log(`Arquivo baixado: ${destino}`);

  await popup.close().catch(() => {});
  return destino;
}

// ----------------------------------------------------------------------
// PROCESSAMENTO DA PLANILHA (CSV OU XLSX) COM EXCELJS
// ----------------------------------------------------------------------

/**
 * Lê o arquivo baixado (csv ou xlsx/xls) e devolve uma matriz simples de
 * strings/valores: linhas x colunas, exatamente como veio do relatório.
 */
async function lerArquivoComoMatriz(caminhoArquivo) {
  const ext = path.extname(caminhoArquivo).toLowerCase();

  if (ext === '.csv') {
    const conteudo = fs.readFileSync(caminhoArquivo, 'latin1'); // relatórios BR costumam vir em latin1/ANSI
    const registros = parseCsvSync(conteudo, {
      bom: true,
      relax_column_count: true,
      skip_empty_lines: false,
    });
    return registros;
  }

  // .xlsx / .xls
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(caminhoArquivo);
  const ws = wb.worksheets[0];
  const matriz = [];
  ws.eachRow({ includeEmpty: true }, (row) => {
    const valores = row.values.slice(1); // ExcelJS usa índice 1-based e values[0] é undefined
    matriz.push(valores.map((v) => (v == null ? '' : v)));
  });
  return matriz;
}

/**
 * Aplica as regras pedidas:
 *  - remove as 2 primeiras linhas;
 *  - limpa o conteúdo da coluna A da última linha;
 *  - remove linhas duplicadas (considerando o conjunto de dados, sem contar o cabeçalho);
 *  - monta um workbook novo com uma Tabela do Excel começando em A2
 *    (cabeçalho na linha 2, dados a partir da linha 3) e linha de Totais.
 *
 * OBS.: não tenho o arquivo real baixado para conferir a estrutura exata
 * das linhas de título/rodapé do relatório. Ajuste a constante
 * LINHA_DO_CABECALHO abaixo se, depois de rodar uma vez, você perceber
 * que o cabeçalho das colunas não caiu certinho na linha 2 do Excel final.
 */
async function processarPlanilha(caminhoOrigem, caminhoSaida) {
  let linhas = await lerArquivoComoMatriz(caminhoOrigem);

  // 1) Excluir as duas primeiras linhas
  linhas = linhas.slice(2);

  if (linhas.length === 0) {
    throw new Error('Depois de remover as 2 primeiras linhas não sobrou nenhuma linha. Confira o arquivo baixado.');
  }

  // 2) Excluir o conteúdo da última linha, coluna A (mantém a linha, só limpa a célula A)
  const ultimaLinha = linhas[linhas.length - 1];
  if (ultimaLinha.length > 0) {
    ultimaLinha[0] = '';
  }

  // A primeira linha remanescente vira o cabeçalho das colunas.
  const cabecalho = linhas[0].map((c) => (c === '' || c == null ? '(coluna)' : String(c)));
  const dados = linhas.slice(1);

  // 4) Remover duplicadas (compara a linha inteira)
  const vistos = new Set();
  const dadosSemDuplicadas = [];
  for (const linha of dados) {
    const chave = JSON.stringify(linha);
    if (!vistos.has(chave)) {
      vistos.add(chave);
      dadosSemDuplicadas.push(linha);
    }
  }

  // Detecta quais colunas são numéricas, para decidir a função da linha de Totais.
  const numColunas = cabecalho.length;
  const colunaEhNumerica = new Array(numColunas).fill(true);
  for (const linha of dadosSemDuplicadas) {
    for (let c = 0; c < numColunas; c++) {
      const valor = linha[c];
      const ehVazio = valor === '' || valor == null;
      const ehNumero = !ehVazio && !isNaN(Number(valor));
      if (!ehVazio && !ehNumero) colunaEhNumerica[c] = false;
    }
  }

  // 3) e 5) Montar a tabela do Excel a partir de A2 + linha de Totais
  const wbSaida = new ExcelJS.Workbook();
  const wsSaida = wbSaida.addWorksheet('Relatorio');

  // Linha 1 fica em branco de propósito, para a tabela começar em A2 (conforme pedido).
  const colunasTabela = cabecalho.map((nome, idx) => ({
    name: nome,
    filterButton: true,
    totalsRowFunction: idx === 0 ? 'count' : (colunaEhNumerica[idx] ? 'sum' : 'none'),
  }));

  wsSaida.addTable({
    name: 'TabelaRelatorio',
    ref: 'A2',
    headerRow: true,
    totalsRow: true,
    style: {
      theme: 'TableStyleMedium2',
      showRowStripes: true,
    },
    columns: colunasTabela,
    rows: dadosSemDuplicadas,
  });

  // Ajusta a largura das colunas de forma simples, baseada no conteúdo.
  cabecalho.forEach((_, idx) => {
    const col = wsSaida.getColumn(idx + 1);
    let maior = String(cabecalho[idx] || '').length;
    for (const linha of dadosSemDuplicadas) {
      const tam = String(linha[idx] ?? '').length;
      if (tam > maior) maior = tam;
    }
    col.width = Math.min(Math.max(maior + 2, 10), 40);
  });

  if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  await wbSaida.xlsx.writeFile(caminhoSaida);

  console.log(`Planilha final gerada em: ${caminhoSaida}`);
  console.log(`Linhas de dados (após remover duplicadas): ${dadosSemDuplicadas.length}`);
}

// ----------------------------------------------------------------------
// EXECUÇÃO PRINCIPAL
// ----------------------------------------------------------------------
(async () => {
  if (!EMAIL || !SENHA) {
    console.warn('Aviso: TRUESAFE_EMAIL/TRUESAFE_SENHA não definidos no .env. O login pode falhar.');
  }

  // channel: "chrome" abre o Google Chrome instalado na máquina (não o Chromium interno do Playwright).
  const browser = await chromium.launch({
    channel: 'chrome',
    headless: false,
  });
  const context = await browser.newContext({ acceptDownloads: true });
  const page = await context.newPage();

  try {
    await login(page);
    await abrirRelatorioAcessosFiltros(page);
    await preencherPeriodoEHorario(page);
    await configurarBloqueios(page);

    const arquivoBaixado = await gerarEBaixarExcel(context, page);

    const nomeSaida = `relatorio_catraca_refeitorio_${toISODate(new Date())}.xlsx`;
    const caminhoSaida = path.join(OUTPUT_DIR, nomeSaida);
    await processarPlanilha(arquivoBaixado, caminhoSaida);

    console.log('Automação concluída com sucesso!');
  } catch (erro) {
    console.error('Falha na automação:', erro);
    process.exitCode = 1;
  } finally {
    await browser.close();
  }
})();
