const xl = require('excel4node')


const getNumber = (x) => {
  return Number(x.replace(/\./g, "").replace(',', '.'))
}

const createExcelForUnprocessables = unprocessables => {
  const wb = new xl.Workbook()

  const ws = headers(wb, 'Contratos_Não_Encontrados')

  const styleForPrice = wb.createStyle({
    numberFormat: 'R$ #,##0.00; ($#,##0.00)',
    alignment: {
      horizontal: 'left',
    },
  });
  let currentRow = 2

  unprocessables.forEach(contract => {

    contract.Pessoas.forEach((person, index) => {
      if (!person.Valor) console.log(person)

      ws.cell(currentRow, 1).string(person.Contrato)
      ws.cell(currentRow, 2).string(person.Código)
      ws.cell(currentRow, 3).string(person.Beneficiário)
      ws.cell(currentRow, 4).string(person.Matrícula)
      ws.cell(currentRow, 5).string(person.CPF)
      ws.cell(currentRow, 6).string(person.Plano)
      ws.cell(currentRow, 7).string(person.Tipo)
      ws.cell(currentRow, 8).string(person.Idade)
      ws.cell(currentRow, 9).string(person.Dependência)
      ws.cell(currentRow, 10).string(person['Data Limite'])
      ws.cell(currentRow, 11).string(person['Data Inclusão'])
      ws.cell(currentRow, 12).string(person['Data Exclusão'])
      ws.cell(currentRow, 13).string(person.Rubrica)
      ws.cell(currentRow, 14).number(person.PIAC ? person.PIAC : 0)
      ws.cell(currentRow, 15).number(getNumber(person.Valor)).style(styleForPrice)
      ws.cell(currentRow, 16).string(person.LCAT)
      ws.cell(currentRow, 17).string(person.Lotação)
      ws.cell(currentRow, 18).string(person['D/A'])
      ws.cell(currentRow, 19).string(person['Data Nascimento'])

      currentRow++
    })
  })

  wb.write(`Contratos_Não_Encontrados.xlsx`)
}

function headers(wb, SheetName) {
  const ws = wb.addWorksheet(SheetName)
  ws.cell(1, 1).string('Contrato')
  ws.cell(1, 2).string('Código')
  ws.cell(1, 3).string('Beneficiário')
  ws.cell(1, 4).string('Matrícula')
  ws.cell(1, 5).string('CPF')
  ws.cell(1, 6).string('Plano')
  ws.cell(1, 7).string('Tipo')
  ws.cell(1, 8).string('Idade')
  ws.cell(1, 9).string('Dependência')
  ws.cell(1, 10).string('Data Limite')
  ws.cell(1, 11).string('Data Inclusão')
  ws.cell(1, 12).string('Data Exclusão')
  ws.cell(1, 13).string('Rubrica')
  ws.cell(1, 14).string('PIAC')
  ws.cell(1, 15).string('Valor')
  ws.cell(1, 16).string('LCAT')
  ws.cell(1, 17).string('Lotação')
  ws.cell(1, 18).string('D/A')
  ws.cell(1, 19).string('Data Nascimento')
  return ws
}
module.exports = (unprocessables) => createExcelForUnprocessables(unprocessables)
