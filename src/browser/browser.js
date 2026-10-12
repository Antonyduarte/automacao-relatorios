const fs = require('fs');
const path = require('path');

function pad2(value) {
  return String(value).padStart(2, '0');
}

function toISODate(date) {
  return `${date.getUTCFullYear()}-${pad2(date.getUTCMonth() + 1)}-${pad2(date.getUTCDate())}`;
}

function getHojeEOntem(timeZone = 'America/Sao_Paulo') {
  // O relatório cruza a madrugada, então as datas precisam ser calculadas
  // no fuso do negócio, e não necessariamente no fuso configurado no PC.
  const partes = new Intl.DateTimeFormat('en-CA', {
    timeZone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).formatToParts(new Date()).reduce((acc, parte) => {
    if (parte.type !== 'literal') acc[parte.type] = Number(parte.value);
    return acc;
  }, {});
  const hoje = new Date(Date.UTC(partes.year, partes.month - 1, partes.day));
  const ontem = new Date(hoje);
  ontem.setUTCDate(ontem.getUTCDate() - 1);
  return { hojeISO: toISODate(hoje), ontemISO: toISODate(ontem) };
}

function validarDataBr(valor, nome) {
  const match = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(String(valor || '').trim());
  if (!match) throw new Error(`${nome} deve estar no formato DD/MM/YYYY.`);
  const data = new Date(Date.UTC(Number(match[3]), Number(match[2]) - 1, Number(match[1])));
  if (data.getUTCFullYear() !== Number(match[3]) || data.getUTCMonth() !== Number(match[2]) - 1 || data.getUTCDate() !== Number(match[1])) {
    throw new Error(`${nome} não é uma data válida: ${valor}`);
  }
  return valor;
}

function dataBrParaInputDate(valor) {
  const [dia, mes, ano] = validarDataBr(valor, 'Data').split('/');
  return `${ano}-${mes}-${dia}`;
}

function nomeAbaDaDataInicial(dataInicial) {
  // A mesma data usada no filtro do sistema define o nome da aba.
  const [dia, mes, ano] = validarDataBr(dataInicial, 'Data Inicial').split('/');
  return `${dia}-${mes}-${ano.slice(-2)}`;
}

function isoParaBr(dataISO) {
  return dataISO.split('-').reverse().join('/');
}

function adicionarDiasISO(dataISO, dias) {
  const [ano, mes, dia] = dataISO.split('-').map(Number);
  const data = new Date(Date.UTC(ano, mes - 1, dia));
  data.setUTCDate(data.getUTCDate() + dias);
  return toISODate(data);
}

// Monta o período de um relatório (dataInicial = dataInicialISO às
// horaInicial, dataFinal = dia seguinte às horaFinal), a mesma regra de
// sempre usada: dataFinal é sempre o dia imediatamente após dataInicial.
function periodoParaDataInicial(dataInicialISO) {
  return {
    dataInicial: isoParaBr(dataInicialISO),
    dataFinal: isoParaBr(adicionarDiasISO(dataInicialISO, 1)),
  };
}

async function login(page, config) {
  await page.goto(config.baseUrl, { waitUntil: 'domcontentloaded' });

  // Com um acesso já lembrado antes, o site pode passar por um loader
  // ("Concluindo o login. Aguarde...") e cair direto nesta tela de
  // confirmação, sem pedir e-mail/senha de novo. Trata isso aqui também,
  // antes de procurar o formulário de login.
  await tratarTelaLembrarAcesso(page);

  const jaLogado = await page.locator('text=Dashboard').first().isVisible().catch(() => false);
  if (jaLogado) {
    console.log('Sessao ja autenticada, pulando tela de login.');
    return;
  }

  // Há mais de uma versão da tela de login; estes seletores cobrem os campos
  // que já foram encontrados nas diferentes telas do sistema.
  const camposUsuario = [
    'input[type="email"]',
    'input[name="email"]',
    'input[name="login"]',
    'input[name="usuario"]',
    '#email',
    '#login',
    '#usuario',
  ];
  const camposSenha = ['input[type="password"]', 'input[name="senha"]', '#senha', '#password'];

  let campoUsuario = null;
  for (const seletor of camposUsuario) {
    const loc = page.locator(seletor).first();
    if (await loc.count()) {
      campoUsuario = loc;
      break;
    }
  }

  let campoSenha = null;
  for (const seletor of camposSenha) {
    const loc = page.locator(seletor).first();
    if (await loc.count()) {
      campoSenha = loc;
      break;
    }
  }

  if (!campoUsuario || !campoSenha) throw new Error('Não foi possível localizar os campos de login.');
  if (!config.email || !config.senha) throw new Error('TRUESAFE_EMAIL e TRUESAFE_SENHA precisam estar definidos no .env.');

  await campoUsuario.fill(config.email);
  await campoSenha.fill(config.senha);

  const botaoEntrar = page.locator(
    'button:has-text("Entrar"), button:has-text("Login"), button:has-text("Acessar"), input[type="submit"]'
  ).first();

  await Promise.all([
    page.waitForNavigation({ waitUntil: 'domcontentloaded' }).catch(() => {}),
    botaoEntrar.click(),
  ]);

  await tratarTelaLembrarAcesso(page);
}

// Depois do login, o TrueSafe às vezes pergunta "Quer que o TrueSafe lembre
// seu acesso?" (form action="manter_conectado.php", botão
// name="remember" value="yes"). Se aparecer, clica em "Sim"; se não
// aparecer (sessão já lembrada, por exemplo), segue em frente normalmente.
async function tratarTelaLembrarAcesso(page) {
  const botaoSim = page.locator('button[name="remember"][value="yes"]').first();
  const apareceu = await botaoSim.isVisible({ timeout: 5000 }).catch(() => false);
  if (!apareceu) return;

  console.log('Tela "Lembrar acesso" detectada, clicando em "Sim, lembrar meu acesso".');
  await Promise.all([
    page.waitForNavigation({ waitUntil: 'domcontentloaded' }).catch(() => {}),
    botaoSim.click(),
  ]);
}

async function abrirRelatorioAcessosFiltros(page, config) {
  const url = new URL('relatorio_acessos_cloud.php', config.baseUrl).toString();
  await page.goto(url, { waitUntil: 'domcontentloaded' });
  await page.locator('#data_inicio').waitFor({ state: 'attached' });
  console.log('Tela "Acessos + Filtros" carregada.');
}

async function preencherPeriodoEHorario(page, config, periodo) {
  // periodo = { dataInicial, dataFinal } em DD/MM/YYYY, já calculado por
  // quem chamou (antes: sempre ontem/hoje; agora: cada data pendente).
  //
  // Os campos de data/hora ficam na aba "Período" (href="#periodo"). Se a
  // aba "Bloqueios" (ou outra) estiver ativa no momento — o que acontece
  // depois que configurarBloqueios roda, já que ela só é configurada uma
  // vez no início —, esses campos ficam escondidos e um page.fill() nelas
  // trava esperando ficarem visíveis para sempre. Por isso voltamos pra
  // aba Período aqui, toda vez, antes de preencher.
  await page.click('a[href="#periodo"]');
  await page.locator('#periodo').waitFor({ state: 'visible' });

  await page.fill('#data_inicio', dataBrParaInputDate(periodo.dataInicial));
  await page.fill('#data_fim', dataBrParaInputDate(periodo.dataFinal));
  await page.fill('#hora_ini', config.horaInicial);
  await page.fill('#hora_fim', config.horaFinal);
  console.log(`Período definido: ${periodo.dataInicial} ${config.horaInicial} até ${periodo.dataFinal} ${config.horaFinal}`);
  return periodo;
}

async function definirCheckbox(page, selector, marcado) {
  const checkbox = page.locator(selector).first();
  await checkbox.waitFor({ state: 'attached' });

  // Usar check/uncheck (em vez de alterar .checked via evaluate) dispara o
  // mesmo fluxo de eventos de um clique real. O TrueSafe usa esse fluxo para
  // atualizar os bloqueios quando "Exibir todas" é desmarcada.
  if (marcado) {
    await checkbox.check();
  } else {
    await checkbox.uncheck();
  }

  await page.waitForFunction(
    ({ selector, marcado }) => document.querySelector(selector)?.checked === marcado,
    { selector, marcado }
  );
}

async function configurarBloqueios(page, config) {
  await page.click('a[href="#bloqueios"]');
  await page.locator('#bloqueios').waitFor({ state: 'visible' });

  // O TrueSafe oculta os inputs e usa o clique no próprio checkbox para
  // atualizar os bloqueios. locator.check()/uncheck() não funciona nesse
  // caso porque exige que o input esteja visível.
  const clicarCheckbox = async (selector, marcado) => {
    const checkbox = page.locator(selector).first();
    await checkbox.waitFor({ state: 'attached' });
    await checkbox.evaluate((elemento, deveMarcar) => {
      if (elemento.checked !== deveMarcar) elemento.click();
    }, marcado);
    await page.waitForFunction(
      ({ seletor, deveMarcar }) => document.querySelector(seletor)?.checked === deveMarcar,
      { seletor: selector, deveMarcar: marcado }
    );
  };

  await clicarCheckbox('#exibir_bloqueios_all', false);
  // A desmarcação de "Exibir todas" atualiza as opções individuais de forma
  // assíncrona; aguarde essa atualização antes de marcar a opção desejada.
  await page.waitForTimeout(300);
  await clicarCheckbox('input[name="bloqueios[]"][value="Catraca Refeitorio"]', true);
 
  await page.waitForFunction(() => {
    const todas = document.querySelector('#exibir_bloqueios_all');
    const refeitorio = document.querySelector('input[name="bloqueios[]"][value="Catraca Refeitorio"]');
    return todas?.checked === false && refeitorio?.checked === true;
  });

  console.log('Aba "Bloqueios" configurada.');
}

async function gerarEBaixarExcel(context, page, config) {
  if (!fs.existsSync(config.downloadDir)) fs.mkdirSync(config.downloadDir, { recursive: true });

  // O evento fica associado ao clique atual, evitando reutilizar um arquivo
  // antigo que já esteja na pasta de downloads.
  const [download] = await Promise.all([
    context.waitForEvent('download', { timeout: 60000 }),
    page.click('button[name="excel"]'),
  ]);

  const destino = path.join(config.downloadDir, download.suggestedFilename());
  await download.saveAs(destino);

  for (const pagina of context.pages()) {
    if (pagina !== page) await pagina.close().catch(() => {});
  }

  console.log(`Arquivo baixado: ${destino}`);
  return destino;
}

module.exports = {
  dataBrParaInputDate,
  getHojeEOntem,
  adicionarDiasISO,
  periodoParaDataInicial,
  gerarEBaixarExcel,
  abrirRelatorioAcessosFiltros,
  configurarBloqueios,
  preencherPeriodoEHorario,
  login,
  nomeAbaDaDataInicial,
  toISODate,
  validarDataBr,
};
