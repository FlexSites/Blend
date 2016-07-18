
const xlsx = require('xlsx-style')

// create sheet from array
module.exports = function sheetFromArray (data) {
  const _worksheet = {}
  const range = { s: { c: 1000000, r: 1000000 }, e: { c: 0, r: 0 } }

  // cell indexes for alignment
  const align = { headers: [2, 3, 13, 15], others: [8, 9, 10, 11, 12, 16, 17, 18] }

  for (var R = 0; R !== data.length; ++R) {
    var rowColor = data[R][21]

    for (var C = 0; C !== 21; ++C) {
      if (range.s.r > R) range.s.r = R
      if (range.s.c > C) range.s.c = C
      if (range.e.r < R) range.e.r = R
      if (range.e.c < C) range.e.c = C
      var cell = { v: data[R][C] }

      if (cell.v == null) continue
      var cellRef = xlsx.utils.encode_cell({ c: C, r: R })

      if (typeof cell.v === 'number') cell.t = 'n'
      else if (typeof cell.v === 'boolean') cell.t = 'b'
      else if (cell.v instanceof Date) {
        cell.t = 'n'; cell.z = xlsx.SSF._table[14]
        cell.v = datenum(cell.v)
      } else {
        cell.t = 's'
      }

      cell.s = { 'font': { 'name': 'Verdana', 'sz': 10 } }

      if (rowColor) cell.s.fill = { patternType: 'solid', fgColor: { rgb: rowColor } }

      // cell alignment
      if (R === 0) { // set aligment, font weight and font size for columns in header row
        cell.s.font.bold = true

        if (align.headers.indexOf(C) > -1) cell.s.alignment = { 'horizontal': 'left' }
        else cell.s.alignment = { 'horizontal': 'center' }
      } else { // set alignment for all other columns
        if (align.others.indexOf(C) > -1) cell.s.alignment = { 'horizontal': 'center' }
      }

      // format currency cells
      if (((C > 3 && C < 8) || C === 14 || C === 19) && R !== 0) {
        cell.t = 'n'
        cell.z = '$#,#0.00'
        cell.numFmt = '$0,000.00'
      }

      _worksheet[cellRef] = cell
    }
  }

  if (range.s.c < 1000000) _worksheet['!ref'] = xlsx.utils.encode_range(range)
  return _worksheet
}
