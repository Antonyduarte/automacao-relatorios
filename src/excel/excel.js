const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const { parse: parseCsvSync } = require('csv-parse/sync');

const CABECALHO_FINAL = [
  'Data', 'Hora', 'Usuario', 'Documento', 'Evento', 'Bloqueio',
  'Credencial', 'Categoria', 'Cod unid.', 'Unidade', 'Empresa/Ref.',
];
const NUM_COLUNAS_FINAIS = CABECALHO_FINAL.length;

async function lerArquivoComoMatriz(caminhoArquivo) {
  const ext = path.extname(caminhoArquivo).toLowerCase();

  if (ext === '.csv') {
    const buffer = fs.readFileSync(caminhoArquivo);
    let conteudo;
    try {
      conteudo = new TextDecoder('utf-8', { fatal: true }).decode(buffer);
    } catch {
      conteudo = buffer.toString('latin1');
    }

    const primeiraLinha = conteudo.split(/\r?\n/)[0] || '';
    const delimitador = (primeiraLinha.match(/;/g) || []).length >=
      (primeiraLinha.match(/,/g) || []).length ? ';' : ',';

    return parseCsvSync(conteudo, {
      bom: true,
      delimiter: delimitador,
      relax_column_count: true,
      relax_quotes: true,
      skip_empty_lines: false,
    });
  }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(caminhoArquivo);
  const matriz = [];
  wb.worksheets[0].eachRow({ includeEmpty: true }, (row) => {
    matriz.push(row.values.slice(1).map((value) => (value == null ? '' : value)));
  });
  return matriz;
}

async function processarPlanilha(caminhoOrigem, caminhoSaida, outputDir) {
  let linhas = (await lerArquivoComoMatriz(caminhoOrigem)).slice(2);
  if (linhas.length < 2) {
    throw new Error('Depois de remover as 2 primeiras linhas, nao sobrou titulo+cabeçalho+dados suficientes.');
  }

  const cortarColunas = (linha) => linha.slice(0, NUM_COLUNAS_FINAIS);
  const linhaTitulo = cortarColunas(linhas[0]);
  const dados = linhas.slice(2).map(cortarColunas);
  if (dados.length === 0) throw new Error('Nao sobrou nenhuma linha de dados depois do processamento.');

  dados[dados.length - 1][0] = '';

  const vistos = new Set();
  const dadosSemDuplicadas = dados.filter((linha) => {
    const chave = JSON.stringify([linha[2], linha[3]]);
    if (vistos.has(chave)) return false;
    vistos.add(chave);
    return true;
  });

  for (const linha of dadosSemDuplicadas) {
    const data = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(linha[0]).trim());
    if (data) linha[0] = new Date(Date.UTC(+data[1], +data[2] - 1, +data[3]));
    if (linha[8] !== '' && linha[8] != null && !isNaN(Number(linha[8]))) linha[8] = Number(linha[8]);
  }

  const wbSaida = new ExcelJS.Workbook();
  const wsSaida = wbSaida.addWorksheet('Relatorio');
  linhaTitulo.forEach((valor, idx) => {
    if (valor !== '' && valor != null) wsSaida.getCell(1, idx + 1).value = valor;
  });

  wsSaida.addTable({
    name: 'TabelaRelatorio',
    ref: 'A2',
    headerRow: true,
    totalsRow: true,
    style: { theme: 'TableStyleMedium2', showRowStripes: true },
    columns: CABECALHO_FINAL.map((nome, idx) => ({
      name: nome,
      filterButton: true,
      ...(idx === 0 ? { totalsRowLabel: 'Total' } : {}),
      totalsRowFunction: idx === 2 ? 'count' : 'none',
    })),
    rows: dadosSemDuplicadas,
  });

  wsSaida.getColumn(1).numFmt = 'dd/mm/yyyy';
  CABECALHO_FINAL.forEach((nome, idx) => {
    const coluna = wsSaida.getColumn(idx + 1);
    let maior = nome.length;
    for (const linha of dadosSemDuplicadas) {
      const tamanho = linha[idx] instanceof Date ? 10 : String(linha[idx] ?? '').length;
      if (tamanho > maior) maior = tamanho;
    }
    coluna.width = Math.min(Math.max(maior + 2, 10), 40);
  });

  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });
  await wbSaida.xlsx.writeFile(caminhoSaida);
  console.log(`Planilha final gerada em: ${caminhoSaida}`);
  console.log(`Linhas de dados (apos remover duplicadas): ${dadosSemDuplicadas.length}`);
}

module.exports = { processarPlanilha };
