const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const { parse: parseCsv } = require('csv-parse/sync');

// A tabela final usa somente as colunas A:K; os campos extras do export ficam
// fora do resultado, incluindo o conteúdo das colunas L até Q.
const CABECALHOS_TABELA = [
  'Data', 'Hora', 'Usuario', 'Documento', 'Evento', 'Bloqueio',
  'Credencial', 'Categoria', 'Cod Unid', 'Unidade', 'Empresa/Ref.',
];

function normalizarValor(valor) {
  if (valor instanceof Date) return valor.toISOString();
  if (valor == null) return '';
  return String(valor).trim();
}

function dividirLinha(linha) {
  // Alguns exports chegam como XLSX, mas com a linha inteira na coluna A
  // separada por ponto e vírgula. XLSX normal continua sendo tratado igual.
  if (linha.length === 1 && typeof linha[0] === 'string' && linha[0].includes(';')) {
    return parseCsv(linha[0], { delimiter: ';', relax_column_count: true })[0];
  }
  return linha;
}

async function lerMatriz(caminho) {
  if (!fs.existsSync(caminho)) throw new Error(`Arquivo de relatório não encontrado: ${caminho}`);
  if (fs.statSync(caminho).size === 0) throw new Error(`Arquivo de relatório vazio: ${caminho}`);

  if (path.extname(caminho).toLowerCase() === '.csv') {
    const texto = fs.readFileSync(caminho, 'utf8');
    return parseCsv(texto, { bom: true, delimiter: ';', relax_column_count: true, skip_empty_lines: false });
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(caminho);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('O relatório não possui nenhuma planilha.');

  const matriz = [];
  worksheet.eachRow({ includeEmpty: true }, row => {
    matriz.push(dividirLinha(row.values.slice(1).map(normalizarValor)));
  });
  return matriz;
}

function limparMatrizNaOrdemSolicitada(matriz) {
  const linhas = matriz.map(linha => [...linha]);
  let ultimaLinhaColunaA = -1;
  for (let i = linhas.length - 1; i >= 0; i -= 1) {
    if (normalizarValor(linhas[i][0]) !== '') { ultimaLinhaColunaA = i; break; }
  }
  if (ultimaLinhaColunaA < 0) throw new Error('O relatório não possui dados na coluna A.');

  // Primeiro remove somente o conteúdo da última célula preenchida da coluna A.
  linhas[ultimaLinhaColunaA][0] = '';
  // Depois remove as duas primeiras linhas inteiras.
  const restantes = linhas.slice(2);
  if (restantes.length < 3) throw new Error('O relatório não possui título, cabeçalho e dados suficientes após remover as duas primeiras linhas.');
  return restantes;
}

function removerLinhasVazias(linhas) {
  return linhas.filter(linha => linha.some(valor => normalizarValor(valor) !== ''));
}

function deduplicar(linhas) {
  // Considera duplicata quando Usuario (índice 2) e Documento (índice 3) se
  // repetem juntos, independente do que muda nas outras colunas (Data, Hora, etc.).
  const vistos = new Set();
  return linhas.filter(linha => {
    const chave = JSON.stringify([normalizarValor(linha[2]), normalizarValor(linha[3])]);
    if (vistos.has(chave)) return false;
    vistos.add(chave);
    return true;
  });
}

function converterDataParaExcel(valor) {
  const texto = normalizarValor(valor);
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(texto);
  if (!match) return valor;
  return new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
}

function prepararRelatorio(caminho) {
  return lerMatriz(caminho).then(matriz => {
    // A ordem é intencional: remoções cabeçalho e só depois a remoção de duplicatas
    const aposRemocoes = limparMatrizNaOrdemSolicitada(matriz);
    const linhas = removerLinhasVazias(aposRemocoes);
    const cabecalho = linhas[0].map(normalizarValor);
    if (!cabecalho.length || cabecalho.every(valor => !valor)) throw new Error('Cabeçalho do relatório não identificado.');

    const quantidadeColunas = CABECALHOS_TABELA.length;
    cabecalho.splice(0, cabecalho.length, ...CABECALHOS_TABELA);
    // linhas[0] = título/unidade (ex.: "Bem Brasil Alimentos - Perdizes"), já
    // usado acima só para validar que existia conteúdo; linhas[1] = cabeçalho
    // real do export (Data, Hora, Usuário...). Os dados começam em linhas[2].
    const dados = linhas.slice(2).map(linha => {
      const copia = Array.from({ length: quantidadeColunas }, (_, i) => normalizarValor(linha[i]));
      copia[0] = converterDataParaExcel(copia[0]);
      return copia;
    });
    const dadosSemDuplicatas = deduplicar(dados);
    if (!dadosSemDuplicatas.length) throw new Error('O relatório não possui registros após o cabeçalho.');

    return { cabecalho, dados: dadosSemDuplicatas, registrosOriginais: dados.length };
  });
}

function nomeTabelaSeguro(nome) {
  return `Tabela_${nome.replace(/[^A-Za-z0-9_]/g, '_')}`;
}

function adicionarRelatorioNaAba(workbook, nomeAba, relatorio, tituloTabela = 'Relatório Jantar') {
  if (workbook.getWorksheet(nomeAba)) {
    throw new Error(`Já existe uma aba chamada ${nomeAba}. O relatório dessa data já pode ter sido processado.`);
  }
  // A checagem acima impede que uma execução repetida substitua a aba do dia.
  const worksheet = workbook.addWorksheet(nomeAba);
  // A tabela recebe uma única linha de cabeçalho, com os nomes oficiais.
  const nomesUnicos = [...CABECALHOS_TABELA];

  const nomesTabelasExistentes = new Set();
  workbook.worksheets.forEach(ws => Object.keys(ws.tables).forEach(nome => nomesTabelasExistentes.add(nome)));
  let nomeTabela = nomeTabelaSeguro(nomeAba);
  let indiceTabela = 2;
  while (nomesTabelasExistentes.has(nomeTabela)) nomeTabela = `${nomeTabelaSeguro(nomeAba)}_${indiceTabela++}`;

  // Linha de título acima da tabela (linha 1), fora da tabela em si.
  // O texto vem de quem chamou (ex.: "Relatório Jantar", "Relatório Almoço").
  worksheet.getCell('A1').value = tituloTabela;
  worksheet.mergeCells('A1:K1');
  worksheet.getCell('A1').alignment = {
    horizontal: 'center',
    vertical: 'middle'
  }
  worksheet.getCell('A1').font = { bold: true, size: 20 };

  // É uma tabela estruturada do Excel, não apenas uma formatação visual
  // A coluna inicial conta registros e identificadores não são somados
  worksheet.addTable({
    name: nomeTabela,
    ref: 'A2',
    headerRow: true,
    totalsRow: true,
    style: { theme: 'TableStyleMedium2', showRowStripes: true },
    columns: nomesUnicos.map((nome, i) => ({
      name: nome,
      filterButton: true,
      // O ExcelJS sempre escreve o texto de totalsRowLabel na 1ª coluna da
      // linha de totais e NUNCA calcula uma fórmula nela (limitação da
      // própria biblioteca, independente do totalsRowFunction que a gente
      // passar). Por isso o rótulo "Total" fica na coluna A (Data) e a
      // quantidade de registros (fórmula SUBTOTAL de verdade, calculada
      // pelo Excel) fica na coluna B (Hora), ao lado.
      ...(i === 0 ? { totalsRowLabel: 'Total' } : {}),
      totalsRowFunction: i === 1 ? 'count' : 'none',
    })),
    rows: relatorio.dados,
  });

  worksheet.views = [{ state: 'frozen', ySplit: 1 }];

  // Coluna A (Data) alinhada totalmente à esquerda. Por padrão o Excel
  // alinha datas/números à direita; aqui forçamos a esquerda em cada célula
  // de dado e também na célula da linha de totais.
  const primeiraLinhaDeDados = 3; // linha 1 = título, linha 2 = cabeçalho da tabela (ref: 'A2')
  const ultimaLinhaDeDados = primeiraLinhaDeDados + relatorio.dados.length - 1;
  const linhaDeTotais = ultimaLinhaDeDados + 1;
  for (let linha = primeiraLinhaDeDados; linha <= linhaDeTotais; linha += 1) {
    worksheet.getCell(linha, 1).alignment = { horizontal: 'left' };
  }

  worksheet.columns.forEach((col, i) => {
    const maior = Math.max(nomesUnicos[i].length, ...relatorio.dados.map(linha => String(linha[i] ?? '').length));
    col.width = Math.min(Math.max(maior + 2, 10), 40);
  });
  return worksheet;
}

async function inserirRelatorioNoArquivoPrincipal(caminhoPrincipal, nomeAba, caminhoRelatorio, tituloTabela) {
  if (!fs.existsSync(caminhoPrincipal)) throw new Error(`Arquivo principal não encontrado: ${caminhoPrincipal}`);
  if (fs.statSync(caminhoPrincipal).size === 0) throw new Error(`O arquivo principal está vazio ou não é um Excel válido: ${caminhoPrincipal}`);

  // O arquivo existente é carregado e alterado em memória, preservando as
  // abas que já estavam nele
  const workbook = new ExcelJS.Workbook();
  try { await workbook.xlsx.readFile(caminhoPrincipal); }
  catch (erro) { throw new Error(`Não foi possível abrir o arquivo principal. Verifique se é um .xlsx válido e se não está bloqueado pelo Excel. ${erro.message}`); }
  const relatorio = await prepararRelatorio(caminhoRelatorio);
  adicionarRelatorioNaAba(workbook, nomeAba, relatorio, tituloTabela);
  await workbook.xlsx.writeFile(caminhoPrincipal);
  return relatorio;
}

module.exports = { prepararRelatorio, inserirRelatorioNoArquivoPrincipal, lerMatriz };