const google = require('googleapis')
const drive = google.drive('v2')
const JWT = google.auth.JWT
const request = require('request-promise')
const path = require('path')
const Bluebird = require('bluebird')

const fs = require('fs-extra')
const xlsx = require('xlsx-style')
const moment = require('moment')
const tojson = xlsx.utils.sheet_to_json
const async = require('async')

const sheetFromArray = require('./lib/sheet-from-array')
const getRange = require('./lib/get-range')
const getColumnNames = require('./lib/get-column-names')

// make sure docs folder exists but don't clear it out
fs.ensureDir(path.join(__dirname, 'docs'))

const stat = Bluebird.promisify(fs.stat, { context: fs })
const writeFile = Bluebird.promisify(fs.writeFile, { context: fs })

const {
  colors,
  HEADERS,
  EXCEL_MIMETYPE,
  EXCLUDEWORKBOOKS,
  EXCLUDESHEETS,
  white,
  blue,
  red,
  brown,
  magenta,
  green,
  orange
} = require('./constants')
const config = require('config')

const jwt = new JWT(
  config.get('google.drive.email'),
  config.get('google.drive.keyPath'),
  null,
  config.get('google.drive.scope')
)

// excel workbook class
class Workbook {
  constructor () {
    this.SheetNames = []
    this.Sheets = {}
  }
}

const wscols = [{ wch: 10 }, { wch: 12 }, { wch: 20 }, { wch: 20 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 16 }, { wch: 12 }, { wch: 12 }, { wch: 18 }, { wch: 12 }, { wch: 30 }, { wch: 28 }, { wch: 10 }, { wch: 10 }, { wch: 12 }, { wch: 20 }]

module.exports.getBlend = function (selectedOptions, selectedYears, sepTXPOC, sepColor, callback) {
  let _headers = HEADERS.concat([])
  selectedOptions = JSON.parse(selectedOptions)
  selectedYears = JSON.parse(selectedYears)

  let files = Bluebird.reduce(
    selectedOptions, 
    (all, option) => 
      listFiles(options.id)
        .then(files => {
          if (files.items) return all.concat(files.items)
          return all
        }), 
    []
  )
  .then(files => {
    console.log('FIlES', files.length)
    // ////////////////////////
    // download excel files //
    // ////////////////////////
    let excel = []
    Bluebird.all(
      files.map((file) => {
        return driveGet(file.id)
          .then(({ title, exportLinks }) => {
          // check if file already exists...
          const pathname = path.join(__dirname, 'docs', title + '.xlsx')
          console.log(`Checking if exists ${pathname}`)

          return stat(pathname)
            .catch((err) => {
              console.log(`"${pathname}" does not exist`)
              var url = res.exportLinks['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']
              if (url && filterByArray(title, EXCLUDEWORKBOOKS)) {
                return request.get({
                  url,
                  encoding: null,
                  'qs': { 'access_token': jwt.credentials.access_token },
                  forever: true,
                  gzip: true
                })
                .then(body => writeFile(pathname, body))
                .then(() => {
                  excel.push({
                    path: pathname,
                    title
                  })
                })
              }
            })
          })
        })
    )
    then(excel => {

    next(null, _sheets)
  },
  function (err, sheets) {
    if (err) console.log(err)
    else console.log('all done')

    var workbook = new Workbook()

    Object.keys(sheets).forEach(function (sheet) {
      workbook.SheetNames.push(sheet)

      var _array = sheets[sheet]
      _array.sort(function (a, b) {
        if (a && b) {
          return a[3].localeCompare(b[3])
        }
      })
      _array.unshift(_headers)

      var _sheet = sheetFromArray(_array)
      _sheet['!cols'] = wscols

      workbook.Sheets[sheet] = _sheet
    })

    var today = moment().format('M-D-YYYY h-mm-ssa')
    var filename = 'BLEND ' + today + '.xlsx'
    var filepath = path.join(__dirname, filename)
    var wopts = { tabSelected: false }

    xlsx.writeFile(workbook, filepath, wopts)

    return filepath
    // open(filepath)

  // $('#status').html('Report Complete')
  // $('.progress').hide()
  // $('#new').show()
  // $('#quit').show()
  }

  // ], function (err, sheets) {
  //     if (err) alert(err)
  //     else console.log('all done')

  //     var workbook = new Workbook()

  //     Object.keys(sheets).forEach(function (sheet) {

  //         workbook.SheetNames.push(sheet)

  //         var _array = sheets[sheet]
  //         _array.sort(function (a, b) {
  //             if (a && b) {
  //                 return a[3].localeCompare(b[3])
  //             }
  //         })
  //         _array.unshift(_headers)

  //         var _sheet = sheetFromArray(_array)
  //         _sheet['!cols'] = wscols

  //         workbook.Sheets[sheet] = _sheet

  //     })

  //     var today = moment().format('M-D-YYYY h-mm-ssa')
  //     var filename = 'BLEND ' + today + '.xlsx'
  //     var filepath = path.join(__dirname, filename)
  //     var wopts = { tabSelected: false }

  //     xlsx.writeFile(workbook, filepath, wopts)
  //     open(filepath)

//     $('#status').html('Report Complete')
//     $('.progress').hide()
//     $('#new').show()
//     $('#quit').show()
// })


const driveList = Bluebird.promisify(drive.children.list, { context: drive.children })
function listFiles (id) {
  return driveList({
    auth: jwt,
    q: EXCEL_MIMETYPE,
    folderId: id,
    forever: true,
    gzip: true
  })
}

const driveGet = Bluebird.promisify(drive.files.get, { context: drive.files })
function getFile(id) {
  return driveGet({
    auth: jwt,
    fileId: id,
    forever: true,
    gzip: true
  })
}

  // convert currency to number
function currencyNumber (amount) {
  var f = 0
  if (amount && typeof amount === 'string') {
    f = parseFloat(amount.replace(/[^0-9\.-]+/g, ''))
  } else {
    f = amount
  }
  return f
}

// get color of row, use column F as reference
function getRowColor (sheet, row) {
  var cell = 'F' + row

  if (sheet[cell] && sheet[cell].s) {
    var s = sheet[cell].s

    if (s.fill) {
      // check for background color
      var bgColor = s.fill.bgColor.rgb
      if (bgColor) return bgColor.substring(2)

      // check for foreground color
      var fgColor = s.fill.fgColor.rgb
      if (fgColor) return fgColor.substring(2)
    }
  }

  // return null
  return 'FFFFFF'
}

function filterByArray (name, array) {
  var notfound = true
  array.forEach(function (exclude) {
    if (name.toLowerCase().indexOf(exclude) > -1) notfound = false
  })
  return notfound
}
