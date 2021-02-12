const faturamento = require('./services/faturamento');
const forClients = require('./services/forClients');
const insurencedb = require('./services/insurance-db-service');
const processable = require('./json/processable.json');
const unprocessable = require('./json/unprocessable.json');

const createExcel = () => {
  faturamento(processable)
  forClients(processable)
  insurencedb(unprocessable)
}

createExcel();