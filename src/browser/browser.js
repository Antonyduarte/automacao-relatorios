const fs = require('fs');
const path = require('path');

function pad2(value) {
  return String(value).padStart(2, '0');
}

function toISODate(date) {
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
}

function getHojeEOntem() {
  const hoje = new Date();
  const ontem = new Date(hoje);
  ontem.setDate(hoje.getDate() - 1);
  return { hojeISO: toISODate(hoje), ontemISO: toISODate(ontem) };
}

async function login(page, config) {
  await page.goto(config.baseUrl, { waitUntil: 'domcontentloaded' });

  const jaLogado = await page.locator('text=Dashboard').first().isVisible().catch(() => false);
  if (jaLogado) {
    console.log('Sessao ja autenticada, pulando tela de login.');
    return;
  }

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

  if (!campoUsuario || !campoSenha) {
    console.warn('Nao encontrei os campos de login. Prosseguindo mesmo assim...');
    return;
  }

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
  const { hojeISO, ontemISO } = getHojeEOntem();
  await page.fill('#data_inicio', ontemISO);
  await page.fill('#data_fim', hojeISO);
  await page.fill('#hora_ini', config.horaInicial);
  await page.fill('#hora_fim', config.horaFinal);
  console.log(`Periodo definido: ${ontemISO} ${config.horaInicial} ate ${hojeISO} ${config.horaFinal}`);
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
  gerarEBaixarExcel,
  abrirRelatorioAcessosFiltros,
  configurarBloqueios,
  preencherPeriodoEHorario,
  login,
  toISODate,
};
