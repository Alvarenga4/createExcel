const {
  createWorkBook,
  newWorkSheet,
  newCell,
  writeExcelFile
} = require('./services/excel-service.js')

const { getTargetPathFile } = require('./util')

const insuranceDbService = require('./services/insurance-db-service');

insuranceDbService.getFileToProcess()

if (insuranceDbService.isEmpty()) {
  console.log('Não existem dados para processar.')
  process.exit(0)
}

const wb = createWorkBook()
const ws = newWorkSheet(wb, 'ClientesNaoEncontrados')

newCell(ws, {
  line: 1,
  column: 1,
  value: 'Clientes não encontrados durante o processamento'
})

const { companyGroupId, groupCompanyName, taxInstallment } = insuranceDbService.getDataIdCompany()

newCell(ws, {
  line: 3,
  column: 1,
  value: `${companyGroupId} - ${groupCompanyName}`
})

newCell(ws, {
  line: 3,
  column: 3,
  value: `Mês/ano: ${taxInstallment}`
})

const companies = insuranceDbService.getCompanies()
const xlsHeader = insuranceDbService.getHeders()
const hasPiacColumn = insuranceDbService.hasPIACColumn(xlsHeader)

const fieldPiacPosition = insuranceDbService.addPIACColumn(xlsHeader)

for (let i = 0, limit = xlsHeader.length; i < limit; i++) {
  newCell(ws, {
    line: 5,
    column: 1 + i,
    value: xlsHeader[i]
  })
}

let actualLine = 5
let actualColumn = 0

companies.forEach(company => {
  company.people.forEach(person => {
    actualLine++
    Object.values(person).forEach(fieldValue => {
      actualColumn++
      newCell(ws, {
        line: actualLine,
        column: actualColumn,
        value: fieldValue
      })

      if (!hasPiacColumn && actualColumn === fieldPiacPosition) {
        actualColumn++
        newCell(ws, {
          line: actualLine,
          column: actualColumn,
          value: 0,
          type: 'number'
        })
      }
    })
    actualColumn = 0
  })
})

writeExcelFile(wb, './')
