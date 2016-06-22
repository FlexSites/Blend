"use strict";


String.prototype.toTitleCase = function () {
    var str = this.replace(/\b[\w]+\b/g, function (txt) {
        return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
    });
    return str;
}


var view;
require(['view'], function (file) {
    //foo is now loaded.
    view = file;
});


var google,
    drive,
    JWT,
    request,
    path;

require(['googleapis', 'request', 'path'], function(googleapis, _request_, _path_) {
    google = googleapis;
    drive = google.drive('v2');
    JWT = google.auth.JWT;
    request = _request_;
    path = _path_;
});

var SERVICE_ACCOUNT_EMAIL = '780463185858-924al4jpqutjbvcbq6bp1lk91t4gmt34@developer.gserviceaccount.com',
    CLIENT_ID = '780463185858-924al4jpqutjbvcbq6bp1lk91t4gmt34.apps.googleusercontent.com',
    SERVICE_ACCOUNT_KEY_FILE = path.join(__dirname, 'elevated-numbers.pem'),
    SCOPE = ['https://www.googleapis.com/auth/drive'];

var jwt = new JWT(
    SERVICE_ACCOUNT_EMAIL,
    SERVICE_ACCOUNT_KEY_FILE,
    null,
    SCOPE);

var fs = require('fs-extra'),
    xlsx = require('xlsx-style'),
    moment = require('moment'),
    tojson = xlsx.utils.sheet_to_json,
    async = require('async'),
    glob = require('glob');

var EXCLUDEWORKBOOKS = ['laira', 'blank', 'request', 'alpine recovery lodge', 'olympus drug and alcohol', 'renaissance outpatient bountiful', 'renaissance ranch- orem', 'renaissance ranch- ut outpatient', ' old'];
var EXCLUDESHEETS = ['fax', 'copy', 'appeal', 'laira', 'checks', 'responses', 'ineligible'];

var selectedYears = []
//////////////////////////////////
// process files on click blend //
//////////////////////////////////
$('#blend').on('click', function () {
    if (!$(this).hasClass('disabled')) {
        selectedYears = [].slice.call($('#years option:selected').map(function (y) {
            return parseInt($(this).val())
        }))
        process()
    }
})


////////////////
// new report //
////////////////
$('#new').on('click', function () {
    view.resetForNew();
})

view.buildYears();

///////////////////
// async testing //
///////////////////
async.waterfall([
    function (next) {
        ////////////////////////////////
        // google drive authorization //
        ////////////////////////////////
        jwt.authorize(function (err, tokens) {
            if (err) next(err)

            jwt.credentials = tokens

            next()
        })
    },

    function (next) {
        ////////////////////////////////
        // get list of client folders //
        ////////////////////////////////
        var req = drive.files.list({
            auth: jwt,
            folderId: '0B_kSXk5v54QYeFRucElOQWlpdG8',
            q: "mimeType = 'application/vnd.google-apps.folder' and trashed = false",
            orderBy: 'title',
            forever: true,
            gzip: true,
            timeout: 1,
            maxResults: 1000
        }, function (err, response, body) {
            if (err) console.log(err)
        })


        req.on('response', function (res) {
            var avgChunks = 250;
            var numChunks = 0;
            var data = ""

            res.on('data', function (chunk) {
                data += chunk
                numChunks += 1
                var percent = parseInt(numChunks * 100 / avgChunks)

                window.setTimeout(function () {
                    view.updateProgress(percent)
                }, 0)
            })

            res.on('end', function () {
                var json = JSON.parse(data)
                if (json && json.items) {

                    window.setTimeout(function () {
                        view.updateProgress(100)
                    }, 0)

                    var files = json.items

                    setTimeout(function () {
                        next(null, files)
                    }, 777)
                }
            })
        })
    },

    function (files, next) {
        ///////////////////////////////////////////////
        // filter folder list and create html select //
        ///////////////////////////////////////////////
        var select = document.createElement('select')
        select.id = 'facilities'
        select.name = 'facilities'
        select.multiple = 'multiple'

        // filter out certain workbooks
        files = files.filter(function (f) {
            return (
                filterByArray(f.title, EXCLUDEWORKBOOKS)
            )
        })

        if (files && files.length) {
            files.forEach(function (file) {

                if (file.parents && file.parents.length > 0 && file.parents[0].id === '0B_kSXk5v54QYeFRucElOQWlpdG8') {
                    var option = document.createElement('option')
                    option.value = file.id
                    option.text = file.title
                    select.appendChild(option)
                }
            })
        }

        $('.progress').hide()
        $('#status').html('Choose facilities&hellip;')
        $('#select').prepend(select)
        $('#select').show()

        $('#facilities').multiselect({
            maxHeight: 310,
            buttonWidth: '400px',
            enableFiltering: true,
            testing: 'something',
            includeSelectAllOption: true,
            enableCaseInsensitiveFiltering: true,

            selectAllName: 'all-facilities',
            selectAllText: $(':input[name="all-facilities"]').prop('checked') ? 'Deselect All' : 'Select All',


            onChange: function (option, checked) {
                var selectedOptions = $('#facilities option:selected')

                if (selectedOptions.length > 0)
                    $('#blend').removeClass('disabled')
                else
                    $('#blend').addClass('disabled')
            },

            buttonText: function (options, select) {
                if (options.length === 0) {
                    return 'Select Facilities';
                }
                else if (options.length > 4) {
                    return options.length + ' selected'
                }
                else {
                    var labels = [];
                    options.each(function () {
                        if ($(this).attr('label') !== undefined) {
                            labels.push($(this).attr('label'))
                        } else {
                            labels.push($(this).html())
                        }
                    })

                    return labels.join(', ') + ''
                }
            }
        })
    }

], function (err, res) {
    if (err) console.log('some error: ', err)
})


process = function () {

    async.waterfall([
        function (next) {
            ///////////////////////////////////////////////////
            // find excel files in selected facility folders //
            ///////////////////////////////////////////////////

            // hide select, reset/show progress, change status message
            $('#select').hide()
            view.resetProgress()
            view.updateProgress(1)
            $('.progress').show()
            $('#status').html('Locating Excel files&hellip;')

            // make sure docs folder extists and is empty
            fs.emptyDirSync(path.join(__dirname, 'docs'))

            // get value/title for all selected options
            var selectedOptions = [].slice.call($('#facilities option:selected').map(function (a, item) {
                return { id: item.value, title: item.innerHTML }
            }))

            // construct query for excel files only
            var query = "mimeType = 'application/vnd.google-apps.spreadsheet'"
            var progress = 0
            var excel = []

            // iterate over selected options
            selectedOptions.forEach(function (option, i, a) {

                var req = getReq(query, option)

                req.on('response', function (res) {
                    var data = ''

                    res.on('data', function (chunk) {
                        data += chunk
                    })

                    res.on('end', function () {

                        progress++
                        var percent = parseInt(progress * 100 / a.length)
                        view.updateProgress(percent)

                        var json = JSON.parse(data)

                        if (json && json.items) {
                            excel = excel.concat(json.items)
                        }


                        if (progress === a.length) {
                            //Refactor: This needs to be refactored and fixed. It does nothing.
                            view.resetProgress()
                            //updateProgress(100/a.length)
                            $('#status').html('Downloading Excel files&hellip;')

                            next(null, excel)
                        }

                    })
                })
            })
        },

        function (files, next) {
            //////////////////////////
            // download excel files //
            //////////////////////////
            var progress = 0
            var excel = []

            async.eachSeries(files, function (file, done) {

                var count = 0, max = 6

                async.whilst(
                    function () {

                        if (progress === files.length) {
                            count = max
                        }

                        return count < max
                    },
                    function (attempt) {
                        count++

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
                                var title = res.title
                                var url = res.exportLinks['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']
                                if (url && filterByArray(title, EXCLUDEWORKBOOKS)) {
                                    request.get({
                                        url: url,
                                        encoding: null,
                                        'qs': { 'access_token': jwt.credentials.access_token },
                                        forever: true,
                                        gzip: true
                                    }, function (err, res) {

                                        if (count > 1) console.log('retry #%d: %s', count, title)


                                        fs.writeFile(path.join(__dirname, 'docs', title + '.xlsx'), res.body, function (err) {

                                            if (err) {
                                                console.log(err)
                                                attempt()
                                            } else {

                                                progress++
                                                var percent = parseInt(progress * 100 / files.length)
                                                view.updateProgress(percent)

                                                try {
                                                    var wb = xlsx.read(res.body, { cellStyles: true })
                                                    if (wb) {
                                                        wb.title = title

                                                        wb.SheetNames = wb.SheetNames.filter(function (s) {
                                                            return filterByArray(s, EXCLUDESHEETS)
                                                        })
                                                        
                                                        wb.SheetNames = wb.SheetNames.filter(function(s){
                                                            return !filterByArray(s, selectedYears)
                                                        })

                                                        var parsedSheets = []
                                                        wb.SheetNames.forEach(function (sn) {
                                                            if (wb.Sheets[sn]) {
                                                                parsedSheets[sn] = wb.Sheets[sn]
                                                            }

                                                        })
                                                        
                                                        wb.Sheets = parsedSheets
                                                        

                                                        excel.push(wb)
                                                        count = max
                                                        wb = null
                                                    }
                                                } catch (e) {
                                                    console.log('error parsing excel')
                                                }
                                                
                                                attempt()
                                            }
                                        })
                                    }).on('response', function (res) {
                                        // Refactor: Should we be doing something with this response?
                                    })
                                } else {

                                    progress++
                                    var percent = parseInt(progress * 100 / files.length)
                                    view.updateProgress(percent)

                                    count = max
                                    attempt()
                                }
                            }
                        })
                    },
                    function (err) {

                        if (progress === files.length) {
                            view.resetProgress()
                            view.updateProgress(1)
                            $('#status').html('Parsing Excel files&hellip;')

                            next(null, excel)
                        }

                        done()
                    }
                )// end whilst

            })// end each
        },

        function (workbooks, next) {
            ///////////////////////////
            // parse excel workbooks //
            ///////////////////////////


            var sepTXPOC = $('#sep-txpoc').is(':checked')
            var sepColor = $('#sep-color').is(':checked')

            var _sheets

            if (sepTXPOC && sepColor) { // separate by TX/POC and row color
                _sheets = {
                    'Open': [], 'Open POC': [],             // white, yellow
                    'Write Off': [], 'Write Off POC': [],        // magenta
                    'Payment Member': [], 'Payment Member POC': [],   // blue
                    'Payment Facility': [], 'Payment Facility POC': [],  // green
                    'Orange': [], 'Orange POC': [], // orange
                    'Confirmed Paid': [], 'Confirmed Paid POC': [],   // brown
                    'Denied': [], 'Denied POC': []            // red
                }
            } else if (sepTXPOC && !sepColor) { // separate by TX/POC
                _sheets = { 'Main': [], 'Main POC': [] }
            } else if (!sepTXPOC && sepColor) { // separate by row color
                _sheets = { 'Open': [], 'Write Off': [], 'Payment Member': [], 'Payment Facility': [], 'Orange': [], 'Confirmed Paid': [], 'Denied': [] }
            } else { // do NOT separate rows
                _sheets = { 'Main': [] }
            }

            // filter out certain workbooks
            workbooks = workbooks.filter(function (w) {
                return (
                    filterByArray(w.title, EXCLUDEWORKBOOKS)
                )
            })

            if (!workbooks.length) next('No workbooks found')

            workbooks.forEach(function (wb) {
                var sheets = wb.Sheets
                var sheetNames = wb.SheetNames

                // filter out sheets
                sheetNames = sheetNames.filter(function (s) {
                    return (
                        filterByArray(s, EXCLUDESHEETS)
                    )
                })

                sheetNames.forEach(function (sheetName) {

                    var year = sheetName.replace(/[^0-9]/g, '')

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
                                .toTitleCase()


                            var billed = currencyNumber(row.BILLED) || 0
                            var paid = currencyNumber(row.PAID) || 0
                            var allowed = currencyNumber(row.ALLOWED) || 0
                            var response = currencyNumber(row.PTRESPONS) || 0
                            var bulkamount = currencyNumber(row.BULKAMOUNT) || 0
                            var balance = (billed * 100 - paid * 100) / 100 || 0

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

                            var fudate = moment(new Date(row.FUDATE)).format('l')
                            if (fudate === 'Invalid date') fudate = ''

                            var invoice = row.INV
                            var sent = moment(new Date(row.SENT)).format('l')
                            var received = moment(new Date(row.RECEIVED)).format('l')
                            var paydate = moment(new Date(row.DATEPAID)).format('l')

                            // TODO why are these invalid (missing)?
                            if (sent === 'Invalid date') sent = ''
                            if (received === 'Invalid date') received = ''
                            if (paydate === 'Invalid date') paydate = ''

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

                            // TODO add facility to array
                            var _row = ['', invoice, ins, name, paid, billed, allowed, response, start, end, sent, received, paydate, check, bulkamount, notes, fudate, loc, units, balance, facility, _color]

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

            }) // end workbooks for each


            next(null, _sheets)

        }

    ], function (err, sheets) {
        if (err) alert(err)
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

        $('#status').html('Report Complete')
        $('.progress').hide()
        $('#new').show()
        $('#quit').show()
    })
}

var colors = {
    blue: 'FF00B0F0',
    green: 'FF66FF33',
    brown: 'FF938953',
    red: 'FFFF0000',
    magenta: 'FFFF00FF',
    yellow: 'FFFFFF00',
    orange: 'FFFF9900'
}

// for comparing row colors
var white = ['FFFFFF', 'FFFF00'] // white, yellow
var blue = ['00B0F0', '24AEFF']
var green = ['00FF00', '66FF33', '6AA84F']
var brown = ['877852', '938950', '938953', '938954', '938955', '948A54', '988D55']
var red = ['C00000', 'FF0000']
var magenta = ['C27BA0', 'FF00FF', 'FF33CC']
var orange = ['FF9900']

var _headers = ['REPORT', 'INV #', 'INS.', 'NAME', 'PAID', 'BILLED', 'ALLOWED', 'PT REPSONS.', 'DATE FROM', 'DATE TO', 'SENT', 'RECEIVED', 'DATE PAID', 'CHECK/CL #', 'BULK AMOUNT', 'ADDITIONS NOTES', 'F/U DATE', 'LOC', 'UNITS', 'BALANCE', 'TC']

var wscols = [{ wch: 10 }, { wch: 12 }, { wch: 20 }, { wch: 20 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 16 }, { wch: 12 }, { wch: 12 }, { wch: 18 }, { wch: 12 }, { wch: 30 }, { wch: 28 }, { wch: 10 }, { wch: 10 }, { wch: 12 }, { wch: 20 }]


// excel workbook class
var Workbook = function () {
    if (!(this instanceof Workbook)) return new Workbook()

    this.SheetNames = []
    this.Sheets = {}
}

// create sheet from array
var sheetFromArray = function (data) {
    var _worksheet = {};
    var range = { s: { c: 1000000, r: 1000000 }, e: { c: 0, r: 0 } };

    // cell indexes for alignment
    var align = { headers: [2, 3, 13, 15], others: [8, 9, 10, 11, 12, 16, 17, 18] };

    for (var R = 0; R != data.length; ++R) {
        var rowColor = data[R][21]

        for (var C = 0; C != 21; ++C) {
            if (range.s.r > R) range.s.r = R;
            if (range.s.c > C) range.s.c = C;
            if (range.e.r < R) range.e.r = R;
            if (range.e.c < C) range.e.c = C;
            var cell = { v: data[R][C] };

            if (cell.v == null) continue;
            var cell_ref = xlsx.utils.encode_cell({ c: C, r: R });

            if (typeof cell.v === 'number') cell.t = 'n';
            else if (typeof cell.v === 'boolean') cell.t = 'b';
            else if (cell.v instanceof Date) {
                cell.t = 'n'; cell.z = xlsx.SSF._table[14];
                cell.v = datenum(cell.v);
            }
            else cell.t = 's';

            cell.s = { 'font': { 'name': 'Verdana', 'sz': 10 } }

            if (rowColor) cell.s.fill = { patternType: 'solid', fgColor: { rgb: rowColor } }

            // cell alignment
            if (R == 0) { // set aligment, font weight and font size for columns in header row

                cell.s.font.bold = true

                if (align.headers.indexOf(C) > -1)
                    cell.s.alignment = { 'horizontal': 'left' };
                else
                    cell.s.alignment = { 'horizontal': 'center' }

            } else { // set alignment for all other columns
                if (align.others.indexOf(C) > -1) cell.s.alignment = { 'horizontal': 'center' };
            }

            // format currency cells
            if (((C > 3 && C < 8) || C == 14 || C == 19) && R !== 0) {
                cell.t = 'n';
                cell.z = '$#,#0.00';
                cell.numFmt = '$0,000.00';
            }

            _worksheet[cell_ref] = cell;
        }

    }

    if (range.s.c < 1000000) _worksheet['!ref'] = xlsx.utils.encode_range(range);
    return _worksheet;
}

// Refactor: Are these ever used?
var prevBegin, prevEnd

var getRange = function (dates) {

    var _dates = dates
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
    var year
    _dates.forEach(function (date) {
        var parts = date.split('/')
        if (parts.length == 3) year = parts.pop()
    })

    _dates.forEach(function (date, d) {
        var slashes = (date.match(/\//g) || []).length

        if (slashes != 2) {
            var newDate = date + '/' + year
            _dates[d] = newDate
        }
    })

    // if we have only one date, duplicate it so we have from/to
    if (_dates.length == 1) _dates.push(_dates[0])

    // _dates.map(function(d){ return moment(new Date(d)).format('M/D/YY') })
    _dates = _dates.map(function (d) { return moment(new Date(d)).format('l') })

    return _dates
}

// return unique array
var uniqueArray = function (a) {
    var uniq = a.filter(function (item, pos) {
        return a.indexOf(item) == pos
    })

    return uniq
}

// get column header names
var getColumnNames = function (sheet) {

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

        if (column === "DATEFROM") column = "BEGIN"
        else if (column === "DOSFROM") column = "BEGIN"
        else if (column === "DATETO") column = "END"
        else if (column === "DOSTO") column = "END"

        if (columns.indexOf(column) > -1 && column !== '') column += '2'
        columns.push(column)
    }

    return columns;
}

// convert currency to number
var currencyNumber = function (amount) {
    var f = null
    if (amount && typeof amount === 'string') {
        f = parseFloat(amount.replace(/[^0-9\.-]+/g, ""))
    } else {
        f = amount
    }
    return f
}

// get color of row, use column F as reference
var getRowColor = function (sheet, row) {
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

var filterByArray = function (name, array) {
    var notfound = true
    array.forEach(function (exclude) {
        if (name.toLowerCase().indexOf(exclude) > -1) notfound = false
    })
    return notfound
}

var getReq = function (query, option) {
    var req = drive.children.list({
        auth: jwt,
        q: query,
        folderId: option.id,
        forever: true,
        gzip: true
    }, function (err, res) {
        setTimeout(function () {
            if (err) {
                console.log('find excel err', err)
                getReq(query, option)
            }
        }, 333)
    })

    return req
}