const xl = require('excel4node')

const numberFormat = 'R$ #,##0.00; ($#,##0.00)'

const border = {
  left: {
    style: 'thin',
    color: '#444444'
  },
  right: {
    style: 'thin',
    color: '#444444'
  },
  top: {
    style: 'thin',
    color: '#444444'
  },
  bottom: {
    style: 'thin',
    color: '#444444'
  },
  diagonal: {
    style: 'thin',
    color: '#444444'
  },
}

const alignment = {
  horizontal: 'center',
  vertical: 'center',
  wrapText: true,
  shrinkToFit: true
}

const fill = {
  type: 'pattern',
  patternType: 'solid',
  bgColor: '#c6c8cc',
  fgColor: '#c6c8cc'
}

const font = {
  size: 10,
  family: 'Roman',
  name: 'Arial',
}

header = (ws, styleType, line, title) => {
  ws.cell(line, 1, line, 6, true).string(title).style(styleType)
}


headerLine = (ws, styleType, line) => {
  ws.cell(line, 1).string('Valor (obrigatório)').style(styleType)
  ws.cell(line, 2).string('Operadora (Sem Over)').style(styleType)
  ws.cell(line, 3).string('Forma de pagamento (Boleto, Cartão) (obrigatório)').style(styleType)
  ws.cell(line, 4).string('Nome Grupo Empresarial').style(styleType)
  ws.cell(line, 5).string('Descrição (Opcional)').style(styleType)
  ws.cell(line, 6).string('Parcelas (caso seja parcelado, senão, deixe em branco)').style(styleType)
}

resize = (ws) => {
  ws.column(1).setWidth(25)
  ws.column(2).setWidth(15)
  ws.column(3).setWidth(15)
  ws.column(4).setWidth(15)
  ws.column(5).setWidth(70)
  ws.column(6).setWidth(48)
  ws.row(2).setHeight(50)
  ws.row(7).setHeight(50)
  ws.row(8).setHeight(50)
}

const getNumber = (x) => {
  //console.log(x)
  return Number(x.replace(/\./g, "").replace(',', '.'))
}


generateBillingXls = (data) => {

  const wb = new xl.Workbook()
  const ws = wb.addWorksheet('Aba')


  var styleFill = wb.createStyle({ border: border, alignment: alignment, fill: fill, font: { ...font, bold: true } });

  var style = wb.createStyle({ border: border, alignment: alignment, font: font });

  var styleNumber = wb.createStyle({ border: border, alignment: alignment, font: font, numberFormat: numberFormat });

  resize(ws)

  header(ws, styleFill, 1, 'Planilha de Faturamento')
  headerLine(ws, wb.createStyle({ border: border, alignment: alignment, font: font, fill: fill }), 2)

  let c = 2

  data.forEach(groupCompany => {
    c++
    ws.cell(c, 1).number(groupCompany.Empresas
      .map(x => x.ValorFinal)
      .reduce((ac, cv) => ac + cv) + data[0].Reajuste)
      .style(styleNumber)
    ws.cell(c, 2).number(groupCompany.Empresas
      .map(empresa => (getNumber(empresa.ValorTotal)))
      .reduce((ac, cv) => ac + cv))
      .style(styleNumber)
    ws.cell(c, 3).string('Boleto').style(style)
    ws.cell(c, 4).string(groupCompany.GrupoEmpresa.name).style(style)
    //pende conversa com o diogão
    ws.cell(c, 5).string(`Cobrança Plano de Saúde Vitta - Competência ${groupCompany.Competencia} e Reajuste Dezembro`).style(style)
    ws.cell(c, 6).string(`${groupCompany.Parcela}/12`)
      .style(style)

  })


  const fileName = `Cobrancas_${data[0].Competencia.replace('/', '_')}.xlsx`

  wb.write(fileName);
}


module.exports = (data) => generateBillingXls(data)


