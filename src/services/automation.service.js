const path = require('path');
const { chromium } = require('playwright');
const browser = require('../browser/browser');
const { processarPlanilha } = require('../excel/excel');

const pastaRelatorios = process.env.RELATORIOS_PASTA || 'C:\\Users\\opton\\Documents\\relatórios';
const config = {
  baseUrl: process.env.TRUESAFE_URL || 'https://www.truesafeweb.com.br',
  email: process.env.TRUESAFE_EMAIL,
  senha: process.env.TRUESAFE_SENHA,
  horaInicial: '22:30',
  horaFinal: '02:30',
  bloqueioDesejado: 'Catraca Refeitorio',
  downloadDir: path.join(pastaRelatorios, 'brutos'),
  outputDir: pastaRelatorios,
};

async function executarAutomacao() {
  if (!config.email || !config.senha) {
    console.warn('Aviso: TRUESAFE_EMAIL/TRUESAFE_SENHA nao definidos no .env.');
  }

  const browserInstance = await chromium.launch({ channel: 'chrome', headless: false });
  const context = await browserInstance.newContext({ acceptDownloads: true });
  const page = await context.newPage();

  try {
    await browser.login(page, config);
    await browser.abrirRelatorioAcessosFiltros(page, config);
    await browser.preencherPeriodoEHorario(page, config);
    await browser.configurarBloqueios(page, config);

    const arquivoBaixado = await browser.gerarEBaixarExcel(context, page, config);
    const nomeSaida = `relatorio_catraca_refeitorio_${browser.toISODate(new Date())}.xlsx`;
    await processarPlanilha(arquivoBaixado, path.join(config.outputDir, nomeSaida), config.outputDir);
    console.log('Automacao concluida com sucesso!');
  } catch (erro) {
    console.error('Falha na automacao:', erro);
    process.exitCode = 1;
  } finally {
    await browserInstance.close();
  }
}

module.exports = { executarAutomacao, config };
