const config = require('config')

exports.getFileToProcess = function () {
  const fileJson = require('../json/unimedFinalNaoProcessados.json')
  return fileJson
}

const insuranceData = getFileToProcess()

function getPositionValorField(header) {
  const idxField = header.indexOf('Valor')
  return idxField < 0 ? header.length - 1 : idxField
}

exports.isEmpty = function () {
  insuranceData.length === 0;
}

exports.getDataIdCompany = function () {
  return ({
    companyGroupId: insuranceData[0].GrupoEmpresa.id,
    groupCompanyName: insuranceData[0].GrupoEmpresa.name,
    taxInstallment: insuranceData[0].Parcela
  });
}

exports.getCompanies = function () {
  return insuranceData[0].Empresas;
}

exports.getHeders = function () {
  return Object.keys(insuranceData[0].Empresas[0].Pessoas[0]);
}

exports.addPIACColumn = function (header) {
  if (!header.includes('PIAC')) {
    header.splice(getPositionValorField(header), 0, 'PIAC')
  }
  return header.indexOf('PIAC')
}

exports.hasPIACColumn = function (header) {
  header.includes('PIAC')
}
