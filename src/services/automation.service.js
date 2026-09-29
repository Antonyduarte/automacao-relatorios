const path = require('path');
const { chromium } = require('playwright');
const browser = require('../browser/browser');
const { inserirRelatorioNoArquivoPrincipal } = require('../excel/excel');

const pastaRelatorios = process.env.RELATORIOS_PASTA || path.resolve(__dirname, '../../saida');
const config = {
  baseUrl: process.env.TRUESAFE_URL || 'https://www.truesafeweb.com.br',
  email: process.env.TRUESAFE_EMAIL,
  senha: process.env.TRUESAFE_SENHA,
  timeZone: process.env.TIME_ZONE || 'America/Sao_Paulo',
  horaInicial: process.env.HORA_INICIAL || '22:30',
  horaFinal: process.env.HORA_FINAL || '02:30',
  bloqueioDesejado: process.env.BLOQUEIO || 'Catraca Refeitorio',
  downloadDir: process.env.DOWNLOAD_PATH || path.join(pastaRelatorios, 'brutos'),
  arquivoPrincipal: process.env.EXCEL_PRINCIPAL_PATH || path.join(pastaRelatorios, 'Relatorios-jantar.xlsx'),
};

async function executarAutomacao() {
  let browserInstance;
  try {
    // A numeração dos logs ajuda a localizar em qual etapa uma execução falhou.
    console.log('[1/8] Iniciando navegador...');
    browserInstance = await chromium.launch({ channel: 'chrome', headless: process.env.HEADLESS === 'true' });
    const context = await browserInstance.newContext({ acceptDownloads: true });
    const page = await context.newPage();

    console.log('[2/8] Acessando sistema...');
    await browser.login(page, config);
    await browser.abrirRelatorioAcessosFiltros(page, config);
    console.log('[3/8] Preenchendo período...');
    const periodo = await browser.preencherPeriodoEHorario(page, config);
    await browser.configurarBloqueios(page, config);

    console.log('[4/8] Gerando relatório...');
    const arquivoBaixado = await browser.gerarEBaixarExcel(context, page, config);
    console.log('[5/8] Download concluído:');
    console.log(`      ${path.basename(arquivoBaixado)}`);

    const nomeAba = browser.nomeAbaDaDataInicial(periodo.dataInicial);
    console.log('[6/8] Processando relatório...');
    console.log('      Removendo última célula da coluna A e as duas primeiras linhas...');
    console.log('      Criando tabela, removendo duplicatas e adicionando contagem de registros...');
    console.log('[7/8] Abrindo arquivo principal...');
    console.log(`      Criando aba: ${nomeAba}`);
    const resultado = await inserirRelatorioNoArquivoPrincipal(config.arquivoPrincipal, nomeAba, arquivoBaixado);
    console.log('[8/8] Arquivo salvo com sucesso.');
    console.log(`      ${config.arquivoPrincipal}`);
    console.log(`      Registros originais: ${resultado.registrosOriginais}; após duplicatas: ${resultado.dados.length}`);
    console.log('Concluído.');
  } catch (erro) {
    console.error(`Falha na automação: ${erro.message}`);
    if (process.env.DEBUG === 'true') console.error(erro.stack);
    process.exitCode = 1;
  } finally {
    if (browserInstance) await browserInstance.close().catch(() => {});
  }
}

module.exports = { executarAutomacao, config };
