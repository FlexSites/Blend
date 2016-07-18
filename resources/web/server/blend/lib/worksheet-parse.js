

module.exports = function parseWorksheet (excel) {
  // /////////////////////////
  // parse excel workbooks //
  // /////////////////////////
  console.log(excel)

  // var sepTXPOC = $('#sep-txpoc').is(':checked')
  // var sepColor = $('#sep-color').is(':checked')
  console.log('in sheets')

  var _sheets

  if (sepTXPOC && sepColor) { // separate by TX/POC and row color
    _sheets = {
      'Open': [], 'Open POC': [], // white, yellow
      'Write Off': [], 'Write Off POC': [], // magenta
      'Payment Member': [], 'Payment Member POC': [], // blue
      'Payment Facility': [], 'Payment Facility POC': [], // green
      'Orange': [], 'Orange POC': [], // orange
      'Confirmed Paid': [], 'Confirmed Paid POC': [], // brown
      'Denied': [], 'Denied POC': [] // red
    }
  } else if (sepTXPOC && !sepColor) { // separate by TX/POC
    _sheets = { 'Main': [], 'Main POC': [] }
  } else if (!sepTXPOC && sepColor) { // separate by row color
    _sheets = { 'Open': [], 'Write Off': [], 'Payment Member': [], 'Payment Facility': [], 'Orange': [], 'Confirmed Paid': [], 'Denied': [] }
  } else { // do NOT separate rows
    _sheets = { 'Main': [] }
  }

  if (!excel.length) next('No workbooks found', null)

  excel.forEach(function (workbook) {
    var wb = xlsx.readFile(workbook.path, { cellStyles: true })

    wb.title = workbook.title

    wb.SheetNames = wb.SheetNames.filter(function (s) {
      return filterByArray(s, EXCLUDESHEETS)
    })

    wb.SheetNames = wb.SheetNames.filter(function (s) {
      return !filterByArray(s, selectedYears)
    })

    var parsedSheets = []
    wb.SheetNames.forEach(function (sn) {
      if (wb.Sheets[sn]) {
        parsedSheets[sn] = wb.Sheets[sn]
      }
    })

    wb.Sheets = parsedSheets

    // filter out certain workbooks
    if (!EXCLUDEWORKBOOKS.indexOf(wb.title) > -1) {
      var sheets = wb.Sheets
      var sheetNames = wb.SheetNames

      // filter out sheets
      sheetNames = sheetNames.filter(function (s) {
        return (
        filterByArray(s, EXCLUDESHEETS)
        )
      })

      sheetNames.forEach(function (sheetName) {

        var sheet = sheets[sheetName]

        var headers = getColumnNames(sheet)

        var rows = tojson(sheet, { header: headers, range: 1 })

        rows.forEach(function (row, i) {
          var ins = row.INS
          var name = row.NAME
          if (name) name = name.trim()

          if (name && ins) {
            var start, end
            // test dates first, if not within selected years, skip
            var allDates = [row.DOS, row.BEGIN, row.END].filter(Boolean)
            var dosDates = getRange(allDates)

            if (dosDates) {
              start = dosDates[0]
              end = dosDates[1]
            }

            if (start) {
              var rowYear = new Date(start).getFullYear()
              if (selectedYears.indexOf(rowYear) < 0) return
            }

            var facility = wb.title
              .toLowerCase()
              .replace('claim', '')
              .replace('cl followup', '')
              .replace('followup', '')
              .replace('follow-up', '')
              .replace('follow up', '')
              .replace(/\d+/g, '')
              .trim()
              // .toTitleCase()

            // TODO if loc = MED, MM or FT, GOP
            var loc
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
            var units = parseInt(row.UNITS)
            if (isNaN(units)) units = ''

            var notes = ''
            var check = ''
            Object.keys(row).forEach(function (key) {
              if (key.indexOf('NOTE') > -1) notes = row[key]
              if (key.indexOf('CHECK') > -1) check = row[key]
            })

            // TODO update insurance names?
            // TODO verify rows with null color have in fact no color!
            var color = getRowColor(sheet, i + 2)

            var _color
            if (blue.indexOf(color) > -1) _color = colors.blue
            else if (red.indexOf(color) > -1) _color = colors.red
            else if (brown.indexOf(color) > -1) _color = colors.brown
            else if (magenta.indexOf(color) > -1) _color = colors.magenta
            else if (green.indexOf(color) > -1) _color = colors.green
            else if (orange.indexOf(color) > -1) _color = colors.orange

            var billed = currencyNumber(row.BILLED)
            var paid = currencyNumber(row.PAID)
            var balance = (billed * 100 - paid * 100) / 100 || 0

            // TODO add facility to array
            var _row = [
              '', 
              row.INV, 
              ins, 
              name, 
              paid, 
              billed, 
              currencyNumber(row.ALLOWED), 
              currencyNumber(row.PTRESPONS), 
              start, 
              end, 
              dateParse(row.SENT), 
              dateParse(row.RECEIVED), 
              dateParse(row.DATEPAID), 
              check, 
              currencyNumber(row.BULKAMOUNT), 
              notes, 
              dateParse(row.FUDATE), 
              loc, 
              units, 
              balance, 
              facility, 
              _color
            ].map(val => val === 'Invalid date' ? '' : val)

            // if separate by TX/POC, check for POC in loc
            var add = sepTXPOC ? (loc === 'POC' ? ' POC' : '') : ''

            // if separate by color, else push all to 'Main' sheet
            if (sepColor) {
              if (white.indexOf(color) > -1) {
                _sheets['Open' + add].push(_row)
              } else if (blue.indexOf(color) > -1) {
                _sheets['Payment Member' + add].push(_row)
              } else if (red.indexOf(color) > -1) {
                _sheets['Denied' + add].push(_row)
              } else if (brown.indexOf(color) > -1) {
                _sheets['Confirmed Paid' + add].push(_row)
              } else if (magenta.indexOf(color) > -1) {
                _sheets['Write Off' + add].push(_row)
              } else if (green.indexOf(color) > -1) {
                _sheets['Payment Facility' + add].push(_row)
              } else if (orange.indexOf(color) > -1) {
                _sheets['Orange' + add].push(_row)
              }
            } else {
              _sheets['Main' + add].push(_row)
            }
          } // end if name + insurance

        }) // end rows for each

      }) // end sheet for each

    }
  }) // end workbooks for each


function dateParse(str) {
  let date = moment(new Date(str)).format('l')
  if (date === 'Invalid date') return ''
  return date
}