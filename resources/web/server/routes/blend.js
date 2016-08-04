'use strict'

// const view = require('./view')

// const BrowserWindow = remote.require('browser-window')
// const BrowserWindow = require('electron').remote

const path = require('path')
const Bluebird = require('bluebird')
const URL = require('url')

const SAVE_FILEPATH = path.join(__dirname, 'streamed-workbook.xlsx')

const fs = require('fs-extra')
const superagent = require('superagent')
const xlsx = require('xlsx-style')
const moment = require('moment')
const async = require('async')
const config = require('config')
const url = require('url')
const Workbook = require('../lib/Workbook')

const EXCLUDEWORKBOOKS = config.get('excluded.workbooks')
const EXCLUDESHEETS = config.get('excluded.worksheets')

// make sure docs folder extists and is empty
// fs.emptyDirSync(path.join(__dirname, 'docs'))

module.exports = () => {
  return (req, res, next) => {
    const urlParts = url.parse(req.url, true)
    blend(req.user.googleClient, JSON.parse(urlParts.query.options), JSON.parse(urlParts.query.years), urlParts.query.txpoc, urlParts.query.color, (err, blendFile) => {
      if (err) return next(err)
      res.download(blendFile, 'blend-file.xlsx', (err) => {
        console.info('sendfile worked', err, blendFile)
      })
    })
  }
}

function blend (client, selectedOptions, selectedYears, sepTXPOC, sepColor, finished) {
  console.log(client)
  async.waterfall([
  //   function(next) {
  //   //   // ////////////////////////
  //   //   // download excel files //
  //   //   // ////////////////////////

  //     return client.list()
  //       .tap(files => console.log("FILES", files))
  //       .then(files => 
  //         Bluebird.mapSeries(files, file => {
  //           console.log('attempt', file.id)
  //           return client.head(file.id)
  //             .then((res) => {
  //               let name = res.name
  //               let url = res.exportLinks['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']
  //               if (url && filterByArray(name, EXCLUDEWORKBOOKS)) {
  //                 let pathname = path.join(__dirname, 'docs', name + '.xlsx')
  //                 console.log('downloading', url, 'to', pathname)
  //                 // return client.get(file.id, pathname)

  //                 const mystream = fs.createWriteStream(pathname)
  //                 const query = URL.parse(url).query
  //                 superagent
  //                   .get(url)
  //                   .set('Authorization', `Bearer ${client.accessToken}`)
  //                   .pipe(mystream)

  //                 console.log('RETURN', [pathname, name, selectedYears])
  //                 return Bluebird.fromCallback(cb => mystream.on('finish', cb))
  //                   .return([pathname, name, selectedYears])
  //               }
  //               return false
  //             })
  //         })
  //       )
  //       .asCallback(next)
  //   //   // next(null, glob.sync(__dirname + '/docs/**.xlsx')
  //   //   //   .map(pathname => {
  //   //   //     console.log('pathname', pathname)
  //   //   //     return [pathname, path.basename(pathname, '.xlsx'), selectedYears]
  //   //   //   })
  //   //   // )
  //   // },
  //   },
    function (next) {
      const workbook = new Workbook(client, {
        selectedOptions,
        selectedYears,
        sepTXPOC,
        sepColor
      })

      workbook.merge('0B040HUA3b4cqeTFyYnEyWW5YZkE')
        .then(() => console.log('GOT PATHS!!!'))
        .catch(ex => console.error(ex))
    }

  ], function (err, sheets) {
    if (err) console.error(err)
    else console.log('all done')

    // let workbook = new Workbook()

    // Object.keys(sheets).forEach(function (sheet) {
    //   workbook.SheetNames.push(sheet)

    //   let _array = sheets[sheet]
    //   _array.sort(function (a, b) {
    //     if (a && b) {
    //       return a[3].localeCompare(b[3])
    //     }
    //   })
    //   _array.unshift(_headers)

    //   let _sheet = sheetFromArray(_array)
    //   _sheet['!cols'] = wscols

    //   workbook.Sheets[sheet] = _sheet
    // })

    // let wopts = { tabSelected: false }

    // xlsx.writeFile(workbook, filepath, wopts)
    finished(null, SAVE_FILEPATH)

    // $('#status').html('Report Complete')
    // $('.progress').hide()
    // $('#new').show()
    // $('#quit').show()
  })
}

let _headers = ['REPORT', 'INV #', 'INS.', 'NAME', 'PAID', 'BILLED', 'ALLOWED', 'PT REPSONS.', 'DATE FROM', 'DATE TO', 'SENT', 'RECEIVED', 'DATE PAID', 'CHECK/CL #', 'BULK AMOUNT', 'ADDITIONS NOTES', 'F/U DATE', 'LOC', 'UNITS', 'BALANCE', 'TC']

let wscols = [{ wch: 10 }, { wch: 12 }, { wch: 20 }, { wch: 20 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 16 }, { wch: 12 }, { wch: 12 }, { wch: 18 }, { wch: 12 }, { wch: 30 }, { wch: 28 }, { wch: 10 }, { wch: 10 }, { wch: 12 }, { wch: 20 }]

// create sheet from array
let sheetFromArray = function (data) {
  let _worksheet = {}
  let range = { s: { c: 1000000, r: 1000000 }, e: { c: 0, r: 0 } }

  // cell indexes for alignment
  let align = { headers: [2, 3, 13, 15], others: [8, 9, 10, 11, 12, 16, 17, 18] }

  for (let R = 0; R != data.length; ++R) {
    let rowColor = data[R][21]

    for (let C = 0; C != 21; ++C) {
      if (range.s.r > R) range.s.r = R
      if (range.s.c > C) range.s.c = C
      if (range.e.r < R) range.e.r = R
      if (range.e.c < C) range.e.c = C
      let cell = { v: data[R][C]}

      if (cell.v == null) continue
      let cell_ref = xlsx.utils.encode_cell({ c: C, r: R })

      if (typeof cell.v === 'number') cell.t = 'n'
      else if (typeof cell.v === 'boolean') cell.t = 'b'
      else if (cell.v instanceof Date) {
        cell.t = 'n'; cell.z = xlsx.SSF._table[14]
        cell.v = datenum(cell.v)
      }
      else cell.t = 's'

      cell.s = { 'font': { 'name': 'Verdana', 'sz': 10 } }

      if (rowColor) cell.s.fill = { patternType: 'solid', fgColor: { rgb: rowColor } }

      // cell alignment
      if (R == 0) { // set aligment, font weight and font size for columns in header row

        cell.s.font.bold = true

        if (align.headers.indexOf(C) > -1)
          cell.s.alignment = { 'horizontal': 'left' }
        else
          cell.s.alignment = { 'horizontal': 'center' }
      } else { // set alignment for all other columns
        if (align.others.indexOf(C) > -1) cell.s.alignment = { 'horizontal': 'center' }
      }

      // format currency cells
      if (((C > 3 && C < 8) || C == 14 || C == 19) && R !== 0) {
        cell.t = 'n'
        cell.z = '$#,#0.00'
        cell.numFmt = '$0,000.00'
      }

      _worksheet[cell_ref] = cell
    }
  }

  if (range.s.c < 1000000) _worksheet['!ref'] = xlsx.utils.encode_range(range)
  return _worksheet
}

let getRange = function (dates) {
  let _dates = dates
    // remove all whitespace
    .map(function (d) { return d.replace(/\s/g, '') })
    // remove everthing except numbers, /, -
    .map(function (d) { return d.replace(/[^0-9\/-]/g, '') })
    // remove empty values
    .filter(function (d) { return d !== '' })
    // split at '-'
    .map(function (d) { return d.split('-') })
    // flatten array
    .reduce(function (a, b) { return a.concat(b) }, [])
    // trim trailing dashes (helps when filtering out duplicates / empty values)
    .map(function (d) { return d.replace(/\-+$/, '') })
    // trim trailing forward slashes
    .map(function (d) { return d.replace(/\/+$/, '') })
    // trim leading zeroes (helps when filtering out duplicates)
    .map(function (d) { return d.replace(/^0+/, '') })
    // filter out duplicates (return unique values)
    .filter(function (v, i, a) { return i == a.indexOf(v) })

  // find our year for fixes below
  let year
  _dates.forEach(function (date) {
    let parts = date.split('/')
    if (parts.length == 3) year = parts.pop()
  })

  _dates.forEach(function (date, d) {
    let slashes = (date.match(/\//g) || []).length

    if (slashes != 2) {
      let newDate = date + '/' + year
      _dates[d] = newDate
    }
  })

  // if we have only one date, duplicate it so we have from/to
  if (_dates.length == 1) _dates.push(_dates[0])

  // _dates.map(function(d){ return moment(new Date(d)).format('M/D/YY') })
  _dates = _dates.map(function (d) { return moment(new Date(d)).format('l') })

  return _dates
}

function readXLSX (url, title, selectedYears) {
  console.log('reading file!!!', title)
  let wb = xlsx.readFile(url, { cellStyles: true })
  if (wb) {
    wb.title = title

    wb.SheetNames = wb.SheetNames.filter(function (s) {
      return filterByArray(s, EXCLUDESHEETS)
    })

    wb.SheetNames = wb.SheetNames.filter(function (s) {
      return !filterByArray(s, selectedYears)
    })

    let parsedSheets = []
    wb.SheetNames.forEach(function (sn) {
      if (wb.Sheets[sn]) {
        parsedSheets[sn] = wb.Sheets[sn]
      }
    })

    wb.Sheets = parsedSheets
    return wb
  }
  return wb
}

// return unique array
let uniqueArray = function (a) {
  let uniq = a.filter(function (item, pos) {
    return a.indexOf(item) == pos
  })

  return uniq
}

// get column header names
let getColumnNames = function (sheet) {
  let range = xlsx.utils.decode_range(sheet['!ref'])
  let columns = []
  let inc = 1

  for (let c = 0; c <= range.e.c; c++) {
    let addr = xlsx.utils.encode_cell({ r: 0, c: c })
    let cell = sheet[addr]
    if (!cell) {
      inc++
      cell = { v: 'Extra_' + inc }
    }

    let column = cell.v.toString().replace(/[^a-zA-Z]/g, '').toUpperCase().trim()

    if (column === 'DATEFROM') column = 'BEGIN'
    else if (column === 'DOSFROM') column = 'BEGIN'
    else if (column === 'DATETO') column = 'END'
    else if (column === 'DOSTO') column = 'END'

    if (columns.indexOf(column) > -1 && column !== '') column += '2'
    columns.push(column)
  }

  return columns
}

// convert currency to number
let currencyNumber = function (amount) {
  let f = null
  if (amount && typeof amount === 'string') {
    f = parseFloat(amount.replace(/[^0-9\.-]+/g, ''))
  } else {
    f = amount
  }
  return f
}

// get color of row, use column F as reference
let getRowColor = function (sheet, row) {
  let cell = 'F' + row

  if (sheet[cell] && sheet[cell].s) {
    let s = sheet[cell].s

    if (s.fill) {
      // check for background color
      let bgColor = s.fill.bgColor.rgb
      if (bgColor) return bgColor.substring(2)

      // check for foreground color
      let fgColor = s.fill.fgColor.rgb
      if (fgColor) return fgColor.substring(2)
    }
  }

  // return null
  return 'FFFFFF'
}

let filterByArray = function (name, array) {
  let notfound = true
  array.forEach(function (exclude) {
    if (name.toLowerCase().indexOf(exclude) > -1) notfound = false
  })
  return notfound
}

let getReq = function (query, option) {
  return Bluebird.fromCallback(cb => client.children.list({
    auth: option.auth,
    q: query,
    folderId: option.id,
    forever: true,
    gzip: true
  }, cb))
}
