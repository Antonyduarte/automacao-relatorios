require('dotenv').config();
require('./logs/logs').iniciar();
const { executarAutomacao } = require('./services/automation.service');

executarAutomacao();