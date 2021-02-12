const xl4node = require('excel4node')

function createWorkBook () {
  return new xl4node.Workbook()
}

function newWorkSheet (workBook, name) {
  return workBook.addWorksheet(name)
}

function newCell (workSheet, cellInfo) {
  const ws = workSheet.cell(cellInfo.line, cellInfo.column)
  switch (cellInfo.type) {
    case 'number':
      ws.number(cellInfo.value)
      break
    case 'formula':
      ws.formula(cellInfo.value)
      break
    case 'bool':
      ws.bool(cellInfo.value)
      break
    default:
      ws.string(cellInfo.value)
      break
  }

  if (cellInfo.style) {
    ws.style(cellInfo.style)
  }
}

function writeExcelFile (workBook, fileName) {
  return workBook.write(fileName)
}

module.exports = {
  createWorkBook,
  newWorkSheet,
  newCell,
  writeExcelFile
}
