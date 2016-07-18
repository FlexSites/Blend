
const xlsx = require('xlsx-style')

// get column header names
module.exports = function getColumnNames (sheet) {
  var range = xlsx.utils.decode_range(sheet['!ref'])
  var columns = []
  var inc = 1

  for (var c = 0; c <= range.e.c; c++) {
    var addr = xlsx.utils.encode_cell({ r: 0, c: c })
    var cell = sheet[addr]
    if (!cell) {
      inc++
      cell = { v: 'Extra_' + inc }
    }

    var column = cell.v.toString().replace(/[^a-zA-Z]/g, '').toUpperCase().trim()

    if (column === 'DATEFROM') column = 'BEGIN'
    else if (column === 'DOSFROM') column = 'BEGIN'
    else if (column === 'DATETO') column = 'END'
    else if (column === 'DOSTO') column = 'END'

    if (columns.indexOf(column) > -1 && column !== '') column += '2'
    columns.push(column)
  }

  return columns
}
