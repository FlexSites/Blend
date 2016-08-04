var excel = require('excel-stream')
var fs = require('fs')

fs.createReadStream(__dirname + '/docs/fuckup.xlsx')
  .pipe(excel({
  	sheet: 'Open'
  }))  // same as excel({sheetIndex: 0})
  .on('data', console.log)