const xl = require('excel4node')
const companyGroups = require('../json/file.json');

const createExcelForCompanies = companyGroups => {
  companyGroups.forEach(group => {
    const groupCompanyName = group.GrupoEmpresa.name
    const competencia = group.Competencia
    const taxInstallment = group.Reajuste
    const installment = group.Parcela
    const wb = new xl.Workbook()

    const styleForPrice = wb.createStyle({
      numberFormat: 'R$ #,##0.00; ($#,##0.00)',
      alignment: {
        horizontal: 'left',
      },
    });

    const ws = headers(wb, groupCompanyName)

    ws.cell(1, 1).string(groupCompanyName)
    ws.cell(2, 1).string(`Prêmio - ${competencia}`)

    let currentRow = 5
    let totalRowsOfPerson = 0
    let totalValue = 0
    group.Empresas.forEach(company => {
      company.Pessoas.forEach(person => {
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
        ws.cell(currentRow, 15).number(person.ValorFinal).style(styleForPrice)
        ws.cell(currentRow, 16).string(person.LCAT)
        ws.cell(currentRow, 17).string(person.Lotação)
        ws.cell(currentRow, 18).string(person['D/A'])
        ws.cell(currentRow, 19).string(person['Data Nascimento'])

        totalValue += person.ValorFinal
        currentRow++
        totalRowsOfPerson++
      })
    })
    totalValue += taxInstallment

    ws.cell(totalRowsOfPerson + 5, 13).string(`Reajuste Dezembro ${installment}/12`)
    ws.cell(totalRowsOfPerson + 5, 15).number(taxInstallment).style(styleForPrice)
    ws.cell(totalRowsOfPerson + 6, 15).number(totalValue).style(styleForPrice)
    wb.write(`${groupCompanyName} - Demonstrativo Analitico de Faturamento_${competencia.replace('/', '_')}.xlsx`)
  })
}

function headers(wb, groupCompanyName) {
  const ws = wb.addWorksheet(groupCompanyName)
  ws.cell(4, 1).string('Contrato')
  ws.cell(4, 2).string('Código')
  ws.cell(4, 3).string('Beneficiário')
  ws.cell(4, 4).string('Matrícula')
  ws.cell(4, 5).string('CPF')
  ws.cell(4, 6).string('Plano')
  ws.cell(4, 7).string('Tipo')
  ws.cell(4, 8).string('Idade')
  ws.cell(4, 9).string('Dependência')
  ws.cell(4, 10).string('Data Limite')
  ws.cell(4, 11).string('Data Inclusão')
  ws.cell(4, 12).string('Data Exclusão')
  ws.cell(4, 13).string('Rubrica')
  ws.cell(4, 14).string('PIAC')
  ws.cell(4, 15).string('Valor')
  ws.cell(4, 16).string('LCAT')
  ws.cell(4, 17).string('Lotação')
  ws.cell(4, 18).string('D/A')
  ws.cell(4, 19).string('Data Nascimento')
  return ws
}
module.exports = (companyGroups) => createExcelForCompanies(companyGroups)
