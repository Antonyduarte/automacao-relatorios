const fs = require('fs');
const { chromium } = require('playwright');
const ExcelJS = require('exceljs');
const browser = require('../browser/browser');
const { inserirRelatorioNoArquivoPrincipal } = require('../excel/excel');

if (!process.env.DOWNLOAD_PATH) {
  throw new Error('DOWNLOAD_PATH precisa estar definido no .env.');
}
if (!process.env.EXCEL_JANTAR_PDZ_PATH) {
  throw new Error('EXCEL_JANTAR_PDZ_PATH precisa estar definido no .env.');
}
if (!process.env.EXCEL_ALMOCO_PDZ_PATH) {
  throw new Error('EXCEL_ALMOCO_PDZ_PATH precisa estar definido no .env.');
}

// Configuração comum a todos os perfis (acesso ao site, fuso horário, pasta
// de downloads). Cada perfil abaixo herda isso e soma o que é próprio dele.
const configBase = {
  baseUrl: process.env.TRUESAFE_URL || 'https://www.truesafeweb.com.br',
  email: process.env.TRUESAFE_EMAIL,
  senha: process.env.TRUESAFE_SENHA,
  timeZone: process.env.TIME_ZONE || 'America/Sao_Paulo',
  downloadDir: process.env.DOWNLOAD_PATH,
};

// Um "perfil" descreve um relatório completo e independente: horário,
// bloqueio filtrado, arquivo de saída, título da tabela, e como a dataFinal
// é calculada a partir da dataInicial (jantar cruza a madrugada; almoço não).
const PERFIS = [
  {
    nome: 'Jantar',
    horaInicial: process.env.HORA_INICIAL_JANTAR || process.env.HORA_INICIAL || '22:30',
    horaFinal: process.env.HORA_FINAL_JANTAR || process.env.HORA_FINAL || '02:30',
    bloqueioDesejado: process.env.BLOQUEIO_JANTAR || process.env.BLOQUEIO || 'Catraca Refeitorio',
    arquivoPrincipal: process.env.EXCEL_JANTAR_PDZ_PATH,
    tituloTabela: 'Relatório Jantar',
    // O período cruza a madrugada: dataFinal é sempre o dia seguinte à dataInicial.
    montarPeriodo: browser.periodoParaDataInicial,
  },
  {
    nome: 'Almoço',
    horaInicial: process.env.HORA_INICIAL_ALMOCO || '10:30',
    horaFinal: process.env.HORA_FINAL_ALMOCO || '14:30',
    bloqueioDesejado: process.env.BLOQUEIO_ALMOCO || process.env.BLOQUEIO || 'Catraca Refeitorio',
    arquivoPrincipal: process.env.EXCEL_ALMOCO_PDZ_PATH,
    tituloTabela: 'Relatório Almoço',
    // O período NÃO cruza a madrugada: dataInicial e dataFinal são o mesmo dia.
    montarPeriodo: browser.periodoMesmoDia,
  },
];

// Nome de aba no formato DD-MM-YY (o mesmo que browser.nomeAbaDaDataInicial gera).
const PADRAO_NOME_ABA = /^(\d{2})-(\d{2})-(\d{2})$/;

// Lê o arquivo principal (se existir) e devolve, em ISO (YYYY-MM-DD), a
// dataInicial mais recente já usada em alguma aba. Retorna null se o
// arquivo ainda não existe ou não tem nenhuma aba nesse formato de data.
async function obterUltimaDataProcessada(caminhoPrincipal) {
  if (!fs.existsSync(caminhoPrincipal)) return null;

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(caminhoPrincipal);

  let maiorISO = null;
  for (const worksheet of workbook.worksheets) {
    const match = PADRAO_NOME_ABA.exec(worksheet.name);
    if (!match) continue;
    const [, dia, mes, ano2] = match;
    const iso = `20${ano2}-${mes}-${dia}`;
    if (maiorISO === null || iso > maiorISO) maiorISO = iso;
  }
  return maiorISO;
}

// Monta a lista (em ISO) de dataInicial que ainda faltam gerar, até ontem
// (a mesma regra de sempre: nunca gera um relatório para hoje — para o
// jantar, porque o período só fecha às horaFinal da madrugada seguinte; para
// o almoço, por consistência, já que o dia de hoje ainda não terminou).
function listarDatasFaltantes(ultimaISO, ontemISO) {
  if (ultimaISO === null) return [ontemISO]; // nenhuma aba ainda: comportamento igual ao de antes (só ontem)
  if (ultimaISO >= ontemISO) return []; // já está em dia

  const datas = [];
  let cursor = browser.adicionarDiasISO(ultimaISO, 1);
  while (cursor <= ontemISO) {
    datas.push(cursor);
    cursor = browser.adicionarDiasISO(cursor, 1);
  }
  return datas;
}

// Gera todos os relatórios pendentes de UM perfil, reaproveitando a página
// (já logada) que foi aberta em executarAutomacao. Abre a tela de filtros e
// configura a aba "Bloqueios" (desmarcar Exibir Todos, marcar só o bloqueio
// desejado) UMA ÚNICA VEZ aqui, fora do loop de datas: clicar em "Gerar
// Excel" não navega a página (o download acontece numa aba separada, fechada
// logo em seguida), então ela continua aberta no mesmo formulário entre uma
// data e outra — não há motivo para recarregar a tela nem reconfigurar os
// Bloqueios de novo, e recarregar era justamente o que resetava tudo para
// "Exibir Todos".
async function processarPerfil(context, page, perfil, ontemISO) {
  console.log(`\n=== Perfil: ${perfil.nome} ===`);
  console.log('Verificando última data processada em:', perfil.arquivoPrincipal);
  const ultimaISO = await obterUltimaDataProcessada(perfil.arquivoPrincipal);
  console.log(ultimaISO ? `Última data processada: ${ultimaISO}` : 'Nenhuma aba de data encontrada ainda.');

  const datasFaltantes = listarDatasFaltantes(ultimaISO, ontemISO);
  if (datasFaltantes.length === 0) {
    console.log('Nenhum relatório pendente. Já está em dia até', ontemISO);
    return;
  }
  console.log(`Relatórios pendentes (${datasFaltantes.length}): ${datasFaltantes.join(', ')}`);

  await browser.abrirRelatorioAcessosFiltros(page, perfil);
  await browser.configurarBloqueios(page, perfil);

  for (let i = 0; i < datasFaltantes.length; i += 1) {
    const dataInicialISO = datasFaltantes[i];
    const periodo = perfil.montarPeriodo(dataInicialISO);
    const prefixo = `  (${i + 1}/${datasFaltantes.length}) ${periodo.dataInicial}`;

    try {
      console.log(`${prefixo} - Preenchendo período...`);
      await browser.preencherPeriodoEHorario(page, perfil, periodo);

      console.log(`${prefixo} - Gerando e baixando excel...`);
      const arquivoBaixado = await browser.gerarEBaixarExcel(context, page, perfil);

      const nomeAba = browser.nomeAbaDaDataInicial(periodo.dataInicial);
      console.log(`${prefixo} - Processando e salvando na aba "${nomeAba}"...`);
      const resultado = await inserirRelatorioNoArquivoPrincipal(
        perfil.arquivoPrincipal,
        nomeAba,
        arquivoBaixado,
        perfil.tituloTabela
      );
      console.log(`${prefixo} - OK (registros originais: ${resultado.registrosOriginais}; após duplicatas: ${resultado.dados.length})`);
    } catch (erroDoDia) {
      // Um dia com problema não deve impedir os outros de serem gerados.
      console.error(`${prefixo} - FALHOU: ${erroDoDia.message}`);
      if (process.env.DEBUG === 'true') console.error(erroDoDia.stack);
    }
  }
}

async function executarAutomacao() {
  const { ontemISO } = browser.getHojeEOntem(configBase.timeZone);

  let browserInstance;
  try {
    console.log('[1/3] Iniciando navegador...');
    browserInstance = await chromium.launch({ channel: 'chrome', headless: process.env.HEADLESS === 'true' });
    const context = await browserInstance.newContext({ acceptDownloads: true });
    const page = await context.newPage();

    console.log('[2/3] Acessando sistema...');
    await browser.login(page, configBase);

    console.log('[3/3] Gerando relatórios pendentes de cada perfil...');
    // Login é feito uma única vez; cada perfil (Jantar, Almoço, ...) é
    // processado em sequência na mesma sessão/página.
    for (const perfil of PERFIS) {
      const perfilCompleto = { ...configBase, ...perfil };
      await processarPerfil(context, page, perfilCompleto, ontemISO);
    }

    console.log('\nConcluído.');
  } catch (erro) {
    console.error(`Falha na automação: ${erro.message}`);
    if (process.env.DEBUG === 'true') console.error(erro.stack);
    process.exitCode = 1;
  } finally {
    if (browserInstance) await browserInstance.close().catch(() => {});
  }
}

module.exports = {
  executarAutomacao,
  configBase,
  PERFIS,
  obterUltimaDataProcessada,
  listarDatasFaltantes,
};
