require('dotenv').config()
const fs = require('fs');
const path = require('path');
const util = require('util');

const caminhoLog = path.resolve(__dirname, process.env.LOGS);

function formatarMensagem(argumentos) {
    return util.format(...argumentos);
}

function gravar(nivel, argumentos) {
    const dataHora = new Date().toISOString();
    const mensagem = `[${dataHora}] [${nivel}] ${formatarMensagem(argumentos)}\n`;

    try {
        fs.appendFileSync(caminhoLog, mensagem, 'utf8');
    } catch (erro) {
        // O log não deve interromper a automação caso o arquivo não possa ser gravado.
        process.stderr.write(`[LOGGER] Não foi possível gravar o log: ${erro.message}\n`);
    }
}

function iniciar() {
    const consoleLogOriginal = console.log.bind(console);
    const consoleWarnOriginal = console.warn.bind(console);
    const consoleErrorOriginal = console.error.bind(console);

    console.log = (...argumentos) => {
        gravar('INFO', argumentos);
        consoleLogOriginal(...argumentos);
    };

    console.warn = (...argumentos) => {
        gravar('WARN', argumentos);
        consoleWarnOriginal(...argumentos);
    };

    console.error = (...argumentos) => {
        gravar('ERROR', argumentos);
        consoleErrorOriginal(...argumentos);
    };

    gravar('INFO', ['===== Nova execução =====']);
}

module.exports = { iniciar, caminhoLog };