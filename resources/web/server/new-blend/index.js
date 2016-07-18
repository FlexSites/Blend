'use strict'

String.prototype.toTitleCase = function () {
  let str = this.replace(/\b[\w]+\b/g, function (txt) {
    return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase()
  })
  return str
}

// const view = require('./view')

// const BrowserWindow = remote.require('browser-window')
// const BrowserWindow = require('electron').remote

let google = require('googleapis')
const drive = google.drive('v2')
const JWT = google.auth.JWT
const request = require('request')
const path = require('path')
const Bluebird = require('bluebird')
const Excel = require('exceljs')

const SERVICE_ACCOUNT_EMAIL = '780463185858-924al4jpqutjbvcbq6bp1lk91t4gmt34@developer.gserviceaccount.com'
const CLIENT_ID = '780463185858-924al4jpqutjbvcbq6bp1lk91t4gmt34.apps.googleusercontent.com'
const SERVICE_ACCOUNT_KEY_FILE = path.resolve(__dirname, '..', 'elevated-numbers.pem')
const SCOPE = ['https://www.googleapis.com/auth/drive']

let jwt = new JWT(
  SERVICE_ACCOUNT_EMAIL,
  SERVICE_ACCOUNT_KEY_FILE,
  null,
  SCOPE)

const fs = require('fs-extra')
const xlsx = require('xlsx-style')
const XLSXStream = require('xlsx-stream')
const moment = require('moment')
const tojson = xlsx.utils.sheet_to_json
const async = require('async')

let EXCLUDEWORKBOOKS = ['laira', 'blank', 'request', 'alpine recovery lodge', 'olympus drug and alcohol', 'renaissance outpatient bountiful', 'renaissance ranch- orem', 'renaissance ranch- ut outpatient', ' old']
let EXCLUDESHEETS = ['fax', 'copy', 'appeal', 'laira', 'checks', 'responses', 'ineligible']

// make sure docs folder extists and is empty
// fs.emptyDirSync(path.join(__dirname, 'docs'))

module.exports = function (selectedOptions, selectedYears, sepTXPOC, sepColor, finished) {
  async.waterfall([
    function (next) {
      // /////////////////////////////////////////////////
      // find excel files in selected facility folders //
      // /////////////////////////////////////////////////

      // iterate over selected options
      Bluebird.reduce(
        selectedOptions,
        (all, option) =>
          getReq("mimeType = 'application/vnd.google-apps.spreadsheet'", option)
            .then(json => {
              if (json && json.items) {
                return all.concat(json.items)
              }
              return all
            }),
        []
      ).then(files => next(null, files))
      .catch(next)
    },

    function (files, next) {
      console.log('download excel files', files.length)
      // ////////////////////////
      // download excel files //
      // ////////////////////////
      let progress = 0
      let excel = []

      async.each(files, function (file, done) {
        console.log('async series')
        let count = 0
        let max = 6

        async.whilst(
          function () {
            if (progress === files.length) {
              count = max
            }

            return count < max
          },
          function (attempt) {
            count++
            console.log('attempt', file.id)
            drive.files.get({
              auth: jwt,
              fileId: file.id,
              forever: true,
              gzip: true
            }, function (err, res) {
              if (err) {
                console.log(err)
                attempt()
              } else {
                let title = res.title
                let url = res.exportLinks['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']
                if (url && filterByArray(title, EXCLUDEWORKBOOKS)) {
                  request.get({
                    url: url,
                    encoding: null,
                    'qs': { 'access_token': jwt.credentials.access_token },
                    forever: true,
                    gzip: true
                  }, (err, res) => {
                    if (count > 1) console.log('retry #%d: %s', count, title)
                    progress++
                    let pathname = path.join(__dirname, 'docs', title + '.xlsx')

                    fs.writeFileSync(pathname, res.body)

                    try {
                      excel.push([pathname, title, selectedYears])
                      count = max
                    } catch (e) {
                      console.log('error parsing excel')
                    }

                    attempt()
                  })
                } else {
                  progress++

                  count = max
                  attempt()
                }
              }
            })
          },
          (err) => {
            if (progress === files.length) {
              // view.resetProgress()
              // view.updateProgress(1)
              // $('#status').html('Parsing Excel files&hellip;')

              next(null, excel)
            }

            done()
          }
        ) // end whilst
      }) // end each
    },

    function (paths, next) {
      next(null, paths.map(args => readXLSX(...args)))
    },

    function (workbooks, next) {
      // /////////////////////////
      // parse excel workbooks //
      // /////////////////////////

      let _sheets

      if (sepTXPOC && sepColor) { // separate by TX/POC and row color
        _sheets = [
          'Open',
          'Open POC',
          // white, yellow
          'Write Off',
          'Write Off POC',
          // magenta
          'Payment Member',
          'Payment Member POC',
          // blue
          'Payment Facility',
          'Payment Facility POC',
          // green
          'Orange',
          'Orange POC',
          // orange
          'Confirmed Paid',
          'Confirmed Paid POC',
          // brown
          'Denied',
          'Denied POC' // red
        ]
      } else if (sepTXPOC && !sepColor) { // separate by TX/POC
        _sheets = [
          'Main',
          'Main POC'
        ]
      } else if (!sepTXPOC && sepColor) { // separate by row color
        _sheets = [
          'Open', 
          'Write Off', 
          'Payment Member', 
          'Payment Facility', 
          'Orange', 
          'Confirmed Paid', 
          'Denied' 
        ]

      } else { // do NOT separate rows
        _sheets = [ 'Main' ]
      }

      if (!workbooks.length) next('No workbooks found')

      var options = {
        filename: './streamed-workbook.xlsx'
      }
      var crazything = new Excel.stream.xlsx.WorkbookWriter(options)
      _sheets = _sheets.map((name) => crazything.addWorksheet(name))

      workbooks.forEach(function (wb) {
        let sheets = wb.Sheets
        let sheetNames = wb.SheetNames

        console.log('workbook parse')

        sheetNames.forEach(function (sheetName) {
          let sheet = sheets[sheetName]

          let headers = getColumnNames(sheet)

          let rows = tojson(sheet, { header: headers, range: 1 })

          console.log('sheetname', sheetName)

          rows.forEach(function (row, i) {
            let ins = row.INS
            let name = row.NAME
            if (name) name = name.trim()

            if (name && ins) {
              let start, end
              // test dates first, if not within selected years, skip
              let allDates = [row.DOS, row.BEGIN, row.END].filter(Boolean)
              let dosDates = getRange(allDates)

              if (dosDates) {
                start = dosDates[0]
                end = dosDates[1]
              }

              if (start) {
                let rowYear = new Date(start).getFullYear()
                if (selectedYears.indexOf(rowYear) < 0) return
              }

              let facility = wb.title
                .toLowerCase()
                .replace('claim', '')
                .replace('cl followup', '')
                .replace('followup', '')
                .replace('follow-up', '')
                .replace('follow up', '')
                .replace(/\d+/g, '')
                .trim()
                .toTitleCase()

              let billed = currencyNumber(row.BILLED) || 0
              let paid = currencyNumber(row.PAID) || 0
              let allowed = currencyNumber(row.ALLOWED) || 0
              let response = currencyNumber(row.PTRESPONS) || 0
              let bulkamount = currencyNumber(row.BULKAMOUNT) || 0
              let balance = (billed * 100 - paid * 100) / 100 || 0

              // TODO if loc = MED, MM or FT, GOP
              let loc
              if (sheetName.indexOf('POC') > -1) {
                loc = 'POC'
              } else {
                loc = row.LOC
                if (!loc) {
                  Object.keys(row).forEach(function (key) {
                    if (key.indexOf('LEVEL') > -1) {
                      loc = row[key]
                    }
                  })
                }
                if (!loc) loc = ''
              }

              // TODO should invalid units be '' or 0?
              let units = parseInt(row.UNITS)
              if (isNaN(units)) units = ''

              let notes = ''
              let check = ''
              Object.keys(row).forEach(function (key) {
                if (key.indexOf('NOTE') > -1) notes = row[key]
                if (key.indexOf('CHECK') > -1) check = row[key]
              })

              let fudate = moment(new Date(row.FUDATE)).format('l')
              if (fudate === 'Invalid date') fudate = ''

              let invoice = row.INV
              let sent = moment(new Date(row.SENT)).format('l')
              let received = moment(new Date(row.RECEIVED)).format('l')
              let paydate = moment(new Date(row.DATEPAID)).format('l')

              // TODO why are these invalid (missing)?
              if (sent === 'Invalid date') sent = ''
              if (received === 'Invalid date') received = ''
              if (paydate === 'Invalid date') paydate = ''

              // TODO update insurance names?
              // TODO verify rows with null color have in fact no color!
              let color = getRowColor(sheet, i + 2)

              let _color
              if (blue.indexOf(color) > -1) _color = colors.blue
              else if (red.indexOf(color) > -1) _color = colors.red
              else if (brown.indexOf(color) > -1) _color = colors.brown
              else if (magenta.indexOf(color) > -1) _color = colors.magenta
              else if (green.indexOf(color) > -1) _color = colors.green
              else if (orange.indexOf(color) > -1) _color = colors.orange

              // TODO add facility to array
              let _row = ['', invoice, ins, name, paid, billed, allowed, response, start, end, sent, received, paydate, check, bulkamount, notes, fudate, loc, units, balance, facility, _color]

              // if separate by TX/POC, check for POC in loc
              let add = sepTXPOC ? (loc === 'POC' ? ' POC' : '') : ''

              // if separate by color, else push all to 'Main' sheet
              if (sepColor) {
                if (white.indexOf(color) > -1) {
                  crazything.getWorksheet('Open' + add).addRow(_row).commit()
                } else if (blue.indexOf(color) > -1) {
                  crazything.getWorksheet('Payment Member' + add).addRow(_row).commit()
                } else if (red.indexOf(color) > -1) {
                  crazything.getWorksheet('Denied' + add).addRow(_row).commit()
                } else if (brown.indexOf(color) > -1) {
                  crazything.getWorksheet('Confirmed Paid' + add).addRow(_row).commit()
                } else if (magenta.indexOf(color) > -1) {
                  crazything.getWorksheet('Write Off' + add).addRow(_row).commit()
                } else if (green.indexOf(color) > -1) {
                  crazything.getWorksheet('Payment Facility' + add).addRow(_row).commit()
                } else if (orange.indexOf(color) > -1) {
                  crazything.getWorksheet('Orange' + add).addRow(_row).commit()
                }
              } else {
                crazything.getWorksheet('Main' + add).addRow(_row).commit()
              }
            } // end if name + insurance
          }) // end rows for each
        }) // end sheet for each
      }) // end workbooks for each

      // crazything.eachSheet(function (worksheet, key) {
      //   worksheet
      //     .commit()
      //     .then(stuff => console.log('key', key))
      //     .catch(ex => console.error(key, ex))
      // })

      setTimeout(function () {
        crazything.commit()
          .then(results => {
            console.log('stream written', results)
            return results
          })
          .catch(ex => {
            console.log('horrible thing', ex)
          })
      }, 1000)
      next(null, _sheets)
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

    let today = moment().format('M-D-YYYY h-mm-ssa')
    let filename = 'BLEND ' + today + '.xlsx'
    let filepath = path.join(__dirname, filename)
    // let wopts = { tabSelected: false }

    // xlsx.writeFile(workbook, filepath, wopts)
    finished(null, filepath)

    // $('#status').html('Report Complete')
    // $('.progress').hide()
    // $('#new').show()
    // $('#quit').show()
  })
}

let colors = {
  blue: 'FF00B0F0',
  green: 'FF66FF33',
  brown: 'FF938953',
  red: 'FFFF0000',
  magenta: 'FFFF00FF',
  yellow: 'FFFFFF00',
  orange: 'FFFF9900'
}

// for comparing row colors
let white = ['FFFFFF', 'FFFF00'] // white, yellow
let blue = ['00B0F0', '24AEFF']
let green = ['00FF00', '66FF33', '6AA84F']
let brown = ['877852', '938950', '938953', '938954', '938955', '948A54', '988D55']
let red = ['C00000', 'FF0000']
let magenta = ['C27BA0', 'FF00FF', 'FF33CC']
let orange = ['FF9900']

let _headers = ['REPORT', 'INV #', 'INS.', 'NAME', 'PAID', 'BILLED', 'ALLOWED', 'PT REPSONS.', 'DATE FROM', 'DATE TO', 'SENT', 'RECEIVED', 'DATE PAID', 'CHECK/CL #', 'BULK AMOUNT', 'ADDITIONS NOTES', 'F/U DATE', 'LOC', 'UNITS', 'BALANCE', 'TC']

let wscols = [{ wch: 10 }, { wch: 12 }, { wch: 20 }, { wch: 20 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 16 }, { wch: 12 }, { wch: 12 }, { wch: 18 }, { wch: 12 }, { wch: 30 }, { wch: 28 }, { wch: 10 }, { wch: 10 }, { wch: 12 }, { wch: 20 }]

// excel workbook class
let Workbook = function () {
  if (!(this instanceof Workbook)) return new Workbook()

  this.SheetNames = []
  this.Sheets = {}
}

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
  return Bluebird.fromCallback(cb => drive.children.list({
    auth: jwt,
    q: query,
    folderId: option.id,
    forever: true,
    gzip: true
  }, cb))
}
