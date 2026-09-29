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

async function login(page, config) {
  await page.goto(config.baseUrl, { waitUntil: 'domcontentloaded' });

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
}

async function abrirRelatorioAcessosFiltros(page, config) {
  const url = new URL('relatorio_acessos_cloud.php', config.baseUrl).toString();
  await page.goto(url, { waitUntil: 'domcontentloaded' });
  await page.locator('#data_inicio').waitFor({ state: 'attached' });
  console.log('Tela "Acessos + Filtros" carregada.');
}

async function preencherPeriodoEHorario(page, config) {
  // O período é recalculado a cada execução como ontem/hoje.
  const { hojeISO, ontemISO } = getHojeEOntem(config.timeZone);
  const periodo = {
    dataInicial: ontemISO.split('-').reverse().join('/'),
    dataFinal: hojeISO.split('-').reverse().join('/'),
  };
  await page.fill('#data_inicio', dataBrParaInputDate(periodo.dataInicial));
  await page.fill('#data_fim', dataBrParaInputDate(periodo.dataFinal));
  await page.fill('#hora_ini', config.horaInicial);
  await page.fill('#hora_fim', config.horaFinal);
  console.log(`Período definido: ${periodo.dataInicial} ${config.horaInicial} até ${periodo.dataFinal} ${config.horaFinal}`);
  return periodo;
}

async function definirCheckbox(page, selector, marcado) {
  await page.waitForSelector(selector, { state: 'attached' });
  await page.evaluate(
    ({ selector, marcado }) => {
      const el = document.querySelector(selector);
      if (!el) throw new Error(`Checkbox nao encontrado: ${selector}`);
      if (el.checked !== marcado) {
        el.checked = marcado;
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
      }
    },
    { selector, marcado }
  );
}

async function configurarBloqueios(page, config) {
  await page.click('a[href="#bloqueios"]');
  await page.locator('#bloqueios').waitFor({ state: 'visible' });
  await definirCheckbox(page, '#exibir_bloqueios_all', false);
  await page.waitForTimeout(300);
  await definirCheckbox(page, `input[name="bloqueios[]"][value="${config.bloqueioDesejado}"]`, true);
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
  gerarEBaixarExcel,
  abrirRelatorioAcessosFiltros,
  configurarBloqueios,
  preencherPeriodoEHorario,
  login,
  nomeAbaDaDataInicial,
  toISODate,
  validarDataBr,
};
