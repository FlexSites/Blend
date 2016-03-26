"use strict";

String.prototype.toTitleCase = function(){
    var str = this.replace(/\b[\w]+\b/g, function(txt){
        return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
    });
    return str;
}

var remote  = require('remote'),
    app     = remote.require('app');

var BrowserWindow = remote.require('browser-window')

var google  = require('googleapis'),
    drive   = google.drive('v2'),
    JWT     = google.auth.JWT,
    request = require('request'),
    path    = require('path');

var SERVICE_ACCOUNT_EMAIL    = '780463185858-924al4jpqutjbvcbq6bp1lk91t4gmt34@developer.gserviceaccount.com',
    CLIENT_ID                = '780463185858-924al4jpqutjbvcbq6bp1lk91t4gmt34.apps.googleusercontent.com',
    SERVICE_ACCOUNT_KEY_FILE = path.join( __dirname, 'elevated-numbers.pem' ),
    SCOPE                    = ['https://www.googleapis.com/auth/drive'];

var jwt = new JWT(
    SERVICE_ACCOUNT_EMAIL,
    SERVICE_ACCOUNT_KEY_FILE,
    null,
    SCOPE);

var fs     = require('fs-extra'),
    xlsx   = require('xlsx-style'),
    moment = require('moment'),
    tojson = xlsx.utils.sheet_to_json,
    open   = require('open'),
    async  = require('async'),
    glob   = require('glob');


//////////////////////////////////
// process files on click blend //
//////////////////////////////////
$('#blend').on('click', function(){
    if (!$(this).hasClass('disabled')) process()
})

/////////////////
// reload page //
/////////////////
$('#reload').on('click', function(){
    var focused = BrowserWindow.getFocusedWindow()
    if (focused) focused.reloadIgnoringCache()
})

//////////////
// quit app //
//////////////
$('#quit').on('click', function(){
    app.quit()
})

////////////////
// new report //
////////////////
$('#new').on('click', function(){
    $('.multiselect-search').val('')
    $('#facilities').multiselect('deselectAll', false)
    $('#facilities').multiselect('updateButtonText')
    $('li').show().removeClass('filter-hidden')

    $('#years').multiselect('selectAll', false)
    $('#years').multiselect('updateButtonText')

    $('#quit').hide()
    $('#new').hide()

    $('#status').html('Choose facilities&hellip;')

    $('#blend').addClass('disabled')

    $('#select').show()
    $('#blend').show()
})




///////////////////
// async testing //
///////////////////
async.waterfall([
    function(next){
        ////////////////////////////////
        // google drive authorization //
        ////////////////////////////////
        jwt.authorize(function(err, tokens){
            if (err) next(err)

            jwt.credentials = tokens

            next()
        })
    },

    function(next){
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
        }, function(err, response, body){
            if (err) console.log(err)
            // if (response) console.log(response)
            // if (body) console.log(body.toJSON())
        })
        var _progress = 0
        req.on('response', function(res){
            var len = parseInt(res.headers['content-length'], 10)
            var data = ""

            res.on('data', function(chunk){
                // console.log('chunk: ', chunk)
                // if (typeof chunk !== 'undefined')
                data += chunk
                _progress += chunk.length
                var percent = parseInt(_progress * 100 / len)

                window.setTimeout(function(){
                    // console.log('percent: ', percent)
                    updateProgress(percent)
                }, 0)

                // updateProgress(percent)
            })

            res.on('end', function(){
                var json = JSON.parse(data)
                if (json && json.items){
                    var files = json.items

                    setTimeout(function(){
                        next(null, files)
                    }, 777)
                }
            })
        })
    },

    function(files, next){
        ///////////////////////////////////////////////
        // filter folder list and create html select //
        ///////////////////////////////////////////////
        var select = document.createElement('select')
        select.id = 'facilities'
        select.name = 'facilities'
        select.multiple = 'multiple'

        var years = document.createElement('select')
        years.id = 'years'
        years.name = 'years'
        years.multiple = 'multiple'

        var thisYear = new Date().getFullYear()

        for (var y = 2011; y <= thisYear; y++){
            var option = document.createElement('option')
            option.value = y
            option.text  = y
            years.appendChild(option)
        }

        $('#select').prepend(years)
        // $('#select').prepend(document.createElement('br'))

        $('#years').multiselect({
            includeSelectAllOption: true,
            buttonWidth: '160px',
            buttonText: function(options, select){
                var len = select[0].options.length

                if (options.length === 0){
                    return 'Select Years'
                } else if (options.length === len){
                    return 'All Years'
                } else if (options.length > 2){
                    return options.length + ' years selected'
                } else {
                    var labels = [];
                    options.each(function(){
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

        $('#years').multiselect('selectAll', false)
        $('#years').multiselect('updateButtonText')

        // move years over to make it and facilties into a combo button visually
        $('#years').next('.btn-group').css({ 'marginLeft': '-1px' })


        console.log('files', files)

        // filter out certain workbooks
        files = files.filter(function(f){
            var lower = f.title.toLowerCase()
            return (
                    lower.indexOf('laira') < 0 &&
                    lower.indexOf('blank') < 0 &&
                    lower.indexOf('request') < 0 &&
                    lower.indexOf('alpine recovery lodge') < 0 &&
                    lower.indexOf('olympus drug and alcohol') < 0 &&
                    lower.indexOf('renaissance outpatient bountiful') < 0 &&
                    lower.indexOf('renaissance ranch- orem') < 0 &&
                    lower.indexOf('renaissance ranch- ut outpatient') < 0
                    )
        })

        if (files && files.length){
            files.forEach(function(file){

                if (file.parents && file.parents.length > 0 && file.parents[0].id === '0B_kSXk5v54QYeFRucElOQWlpdG8'){
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

            // selectAllText: (1 == 1 ? 'true' : 'false'),
            // selectAllNumber: true,

            // onSelectAll: function(){
            //     console.log( $(':input[name="all-facilities"]').prop('checked') )
            // },

            onChange: function(option, checked){
                var selectedOptions = $('#facilities option:selected')

                if (selectedOptions.length > 0)
                    $('#blend').removeClass('disabled')
                else
                    $('#blend').addClass('disabled')
            },

            buttonText: function(options, select){
                // console.log(this)
                if (options.length === 0) {
                    return 'Select Facilities';
                }
                else if (options.length > 4){
                    return options.length + ' selected'
                }
                else {
                    var labels = [];
                    options.each(function(){
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

        // $(':input[name="all-facilities"]').on('click', function(){
        //     console.log('clicked')
        //     // var label = $(this).parent()
        //     // var checked = this.checked

        //     // var a = checked ? 'Deselect All' : 'Select All'
        //     // var b = checked ? 'Select All' : 'Deselect All'

        //     // console.log(a, b)

        //     // label.html(label.html().replace(a, b))
        // })
    }

], function(err, res){
    if (err) console.log('some error: ', err)
})


process = function(){

    async.waterfall([
        function(next){
            ///////////////////////////////////////////////////
            // find excel files in selected facility folders //
            ///////////////////////////////////////////////////

            // hide select, reset/show progress, change status message
            $('#select').hide()
            resetProgress()
            updateProgress(1)
            $('.progress').show()
            $('#status').html('Locating Excel files&hellip;')

            // make sure docs folder extists and is empty
            fs.emptyDirSync( path.join( __dirname, 'docs' ))

            // get value/title for all selected options
            var selectedOptions = [].slice.call($('#facilities option:selected').map(function(a, item){
                return { id: item.value, title: item.innerHTML }
            }))

            // construct query for excel files only
            var query = "mimeType = 'application/vnd.google-apps.spreadsheet'"
            var progress = 0
            var excel = []

            // iterate over selected options
            selectedOptions.forEach(function(option, i, a){
                // console.log('option: ', option.title)

                var req = drive.children.list({
                    auth: jwt,
                    q: query,
                    folderId: option.id,
                    // timeout: 10000,
                    forever: true,
                    gzip: true
                }, function(err, res){
                    if (err) console.log('find excel err', err)
                })

                req.on('response', function(res){
                    var data = ''

                    res.on('data', function(chunk){
                        data += chunk
                    })

                    res.on('end', function(){

                        progress++
                        var percent = parseInt(progress * 100 / a.length)
                        updateProgress(percent)

                        var json = JSON.parse(data)

                        if(json && json.items){
                            excel = excel.concat(json.items)
                        }

                        if (progress === a.length){
                            resetProgress()
                            updateProgress(100/files.length)
                            $('#status').html('Downloading Excel files&hellip;')

                            next(null, excel)
                        }

                    })
                })
            })
        },

        function(files, next){
            //////////////////////////
            // download excel files //
            //////////////////////////
            var progress = 0
            var excel = []

            async.eachSeries(files, function(file, done){

                var count = 0, max = 6

                async.whilst(
                    function(){
                        // console.log('progress: %d, count: %d\n\n', progress, count)

                        if (progress === files.length){
                            count = max
                            // resetProgress()
                            // updateProgress(1)
                            // $('#status').html('Parsing Excel files&hellip;')
                            // next(null, excel)
                        }

                        return count < max
                    },
                    function(attempt){
                        count++

                        drive.files.get({
                            auth: jwt,
                            fileId: file.id,
                            forever: true,
                            gzip: true
                        }, function(err, res){
                            if (err){
                                console.log(err)
                                attempt()
                            } else {
                                var title = res.title
                                // console.log(title)
                                var url   = res.exportLinks['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']

                                // if (count > 1) console.log('%c count: %d\t%s ', 'background:#222, color:#bada55', count, title)
                                // if (count > 1) console.log('retry #%d: %s', count, title)

                                if (url && title.toLowerCase().indexOf('requests') < 0){
                                    request.get({
                                        url: url,
                                        encoding: null,
                                        'qs': { 'access_token': jwt.credentials.access_token },
                                        forever: true,
                                        gzip: true
                                    }, function(err, res){

                                        if (count > 1) console.log('retry #%d: %s', count, title)


                                        fs.writeFile(path.join(__dirname, 'docs', title + '.xlsx'), res.body, function(err){

                                            if (err){
                                                console.log(err)
                                                attempt()
                                            } else {
                                                // console.log('file saved: ', title + '.xlsx')

                                                progress++
                                                var percent = parseInt(progress * 100 / files.length)
                                                updateProgress(percent)

                                                try {
                                                    var wb = xlsx.read(res.body, { cellStyles: true })
                                                    if (wb){
                                                        wb.title = title
                                                        excel.push(wb)
                                                        count = max
                                                    }
                                                } catch(e) {
                                                    console.log('error parsing excel')
                                                }

                                                attempt()
                                            }
                                        })
                                    }).on('response', function(res){
                                        // console.log(res)
                                    })
                                } else {
                                    // console.log('ignore file: ', title)

                                    progress++
                                    var percent = parseInt(progress * 100 / files.length)
                                    updateProgress(percent)

                                    count = max
                                    attempt()
                                }
                            }
                        })
                    },
                    function(err){

                        if (progress === files.length){
                            resetProgress()
                            updateProgress(1)
                            $('#status').html('Parsing Excel files&hellip;')

                            next(null, excel)
                        }

                        done()
                    }
                )// end whilst

            })// end each
        },

        function(workbooks, next){
            ///////////////////////////
            // parse excel workbooks //
            ///////////////////////////
            console.log('parse workbooks')

            var selectedYears = [].slice.call($('#years option:selected').map(function(y){
                return parseInt($(this).val())
            }))

            var sepTXPOC  = $('#sep-txpoc').is(':checked')
            var sepColor  = $('#sep-color').is(':checked')

            var _sheets

            if (sepTXPOC && sepColor){ // separate by TX/POC and row color
                _sheets = {
                    'Open': [],            'Open POC': [],             // white, yellow
                    'Write Off': [],       'Write Off POC': [],        // magenta
                    'Payment Member': [],  'Payment Member POC': [],   // blue
                    'Payment Facility':[], 'Payment Facility POC':[],  // green
                    'Confirmed Paid': [],  'Confirmed Paid POC': [],   // brown
                    'Denied': [],          'Denied POC': []            // red
                }
            } else if (sepTXPOC && !sepColor){ // separate by TX/POC
                _sheets = { 'Main': [], 'Main POC': [] }
            } else if (!sepTXPOC && sepColor){ // separate by row color
                _sheets = { 'Open': [], 'Write Off': [], 'Payment Member': [], 'Payment Facility':[], 'Confirmed Paid': [], 'Denied': [] }
            } else { // do NOT separate rows
                _sheets = { 'Main': [] }
            }







            // filter out certain workbooks
            workbooks = workbooks.filter(function(w){
                var workbookLower = w.title.toLowerCase()
                return (
                        workbookLower.indexOf('laira') < 0 &&
                        workbookLower.indexOf('old') < 0
                        )
            })


            if (!workbooks.length) next('No workbooks found')

                workbooks.forEach(function(wb){
                    var sheets = wb.Sheets
                    var sheetNames = wb.SheetNames

                    // filter out sheets
                    sheetNames = sheetNames.filter(function(s){
                        var sheetLower = s.toLowerCase()
                        return (
                                sheetLower.indexOf('fax') < 0 &&
                                sheetLower.indexOf('copy') < 0 &&
                                sheetLower.indexOf('appeal') < 0 &&
                                sheetLower.indexOf('laira') < 0 &&
                                sheetLower.indexOf('checks') < 0 &&
                                sheetLower.indexOf('responses') < 0 &&
                                sheetLower.indexOf('ineligible') < 0
                                )
                    })


                    sheetNames.forEach(function(sheetName){

                        var year = sheetName.replace(/[^0-9]/g, '')
                        // console.log('year: ', year)

                        var sheet = sheets[sheetName]

                        var headers = getColumnNames(sheet)
                        // console.log(headers)

                        var rows = tojson( sheet, { header: headers, range: 1 })
                        // console.log(rows)

                        rows.forEach(function(row, i){

                            var ins  = row.INS
                            var name = row.NAME
                            if (name) name = name.trim()

                            if (name && ins){
                                var start, end
                                // test dates first, if not within selected years, skip
                                var allDates = [row.DOS, row.BEGIN, row.END].filter(Boolean)
                                var dosDates = getRange( allDates )

                                if (dosDates){
                                    start = dosDates[0]
                                    end   = dosDates[1]
                                }

                                if (start){
                                    var rowYear = new Date(start).getFullYear()
                                    // console.log(rowYear, selectedYears)
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


                                var billed   = currencyNumber( row.BILLED ) || 0
                                var paid     = currencyNumber( row.PAID ) || 0
                                var allowed  = currencyNumber( row.ALLOWED ) || 0
                                var response = currencyNumber( row.PTRESPONS ) || 0
                                var balance  = (billed * 100 - paid * 100) / 100 || 0

                                // TODO if loc = MED, MM or FT, GOP
                                var loc
                                if (sheetName.indexOf('POC') > -1){
                                    loc = 'POC'
                                } else {
                                    loc = row.LOC
                                    if (!loc){
                                        Object.keys(row).forEach(function(key){
                                            if (key.indexOf('LEVEL') > -1){
                                                // console.log('LEVEL instead of LOC')
                                                loc = row[key]
                                            }
                                        })
                                    }
                                    if (!loc) loc = ''
                                }

                                // if (loc === '') console.log('no LOC: ', wb.title, sheetName, (i + 2))

                                // if(!loc) console.log(headers)

                                // TODO should invalid units be '' or 0?
                                var units    = parseInt( row.UNITS )
                                // if (units == 0) console.log(wb.title, sheetName, (i + 2))
                                if (isNaN(units)) units = ''

                                var notes    = ''
                                var check    = ''
                                Object.keys(row).forEach(function(key){
                                    if (key.indexOf('NOTE') > -1) notes = row[key]
                                    if (key.indexOf('CHECK') > -1) check = row[key]
                                })

                                var fudate   = moment( new Date( row.FUDATE )).format('M/D/YY')
                                if (fudate === 'Invalid date') fudate = ''

                                var invoice  = row.INV


                                console.log(invoice, fudate)

                                // var start, end

                                // // var allDates = [row.DOS, row.BEGIN, row.END].filter(function(a){ return a !== undefined })
                                // var allDates = [row.DOS, row.BEGIN, row.END].filter(Boolean)
                                // var dosDates = getRange( allDates )

                                // if (!dosDates.length) console.log('NO dates!')




                                // var dosDates = getRange( row.DOS, row.BEGIN, row.END )

                                // if (dosDates){
                                //     start = dosDates[0]
                                //     end   = dosDates[1]
                                // }

                                // console.log(row.DOS, row.BEGIN, row.END)
                                // console.log(allDates)
                                // console.log(dosDates)
                                // console.log('--------------------------')


                                // var start    = moment( new Date( row.DATEFROM || row.DOSFROM )).format('M/D/YY')
                                // var end      = moment( new Date( row.DATETO || row.DOSTO )).format('M/D/YY')

                                // // TODO ensure start before end, end before/on today, both valid
                                // if (start === 'Invalid date' || end === 'Invalid date'){
                                //     // if (row.DOS)
                                //     //     console.log(row.DOS)
                                //     // else if (row.DOSFROM)
                                //     //     console.log(row.DOSFROM)
                                //     // else
                                //     //     console.log(row)
                                // }
                                // if (start === 'Invalid date') start = ''
                                // if (end === 'Invalid date') end = ''

                                var sent     = moment( new Date( row.SENT )).format('M/D/YY')
                                var received = moment( new Date( row.RECEIVED )).format('M/D/YY')
                                // var paydate  = moment( new Date( row.PAID2 )).format('M/D/YY')
                                var paydate  = moment( new Date( row.DATEPAID )).format('M/D/YY')

                                // TODO why are these invalid (missing)?
                                if (sent === 'Invalid date') sent = ''
                                if (received === 'Invalid date') received = ''
                                if (paydate === 'Invalid date') paydate = ''

                                // TODO update insurance names?


                                // TODO verify rows with null color have in fact no color!
                                var color = getRowColor( sheet, i + 2 )

                                var _color
                                if (blue.indexOf(color) > -1) _color = colors.blue
                                else if (red.indexOf(color) > -1) _color = colors.red
                                else if (brown.indexOf(color) > -1) _color = colors.brown
                                else if (magenta.indexOf(color) > -1) _color = colors.magenta
                                else if (green.indexOf(color) > -1) _color = colors.green

                                    // console.log('color: ', _color)

                                // TODO add facility to array
                                var _row = [ '', invoice, ins, name, paid, billed, allowed, response, start, end, sent, received, paydate, check, notes, fudate, loc, units, balance, facility , _color ]

                                // if (_color) _row.push(_color)



                                // if separate by TX/POC, check for POC in loc
                                var add = sepTXPOC ? (loc === 'POC' ? ' POC' : '') : ''

                                // if separate by color, else push all to 'Main' sheet
                                if (sepColor){
                                    if (white.indexOf(color) > -1){
                                        _sheets[ 'Open' + add ].push(_row)
                                    } else if (blue.indexOf(color) > -1){
                                        _sheets[ 'Payment Member' + add ].push(_row)
                                    } else if (red.indexOf(color) > -1){
                                        _sheets[ 'Denied' + add ].push(_row)
                                    } else if (brown.indexOf(color) > -1){
                                       _sheets[ 'Confirmed Paid' + add ].push(_row)
                                    } else if (magenta.indexOf(color) > -1){
                                        _sheets[ 'Write Off' + add ].push(_row)
                                    } else if (green.indexOf(color) > -1){
                                        _sheets[ 'Payment Facility' + add ].push(_row)
                                    }
                                } else {
                                    _sheets[ 'Main' + add ].push(_row)
                                }



                            } // end if name + insurance

                        }) // end rows for each

                    }) // end sheet for each

                }) // end workbooks for each


            next(null, _sheets)

        }

    ], function(err, sheets){
        if (err) alert(err)
        else console.log('all done')

        var workbook = new Workbook()


        Object.keys(sheets).forEach(function(sheet){

            workbook.SheetNames.push(sheet)

            var _array = sheets[sheet]
            _array.sort(function(a,b){
                if (a && b){
                    return a[3].localeCompare(b[3])
                }
            })
            _array.unshift(_headers)

            var _sheet = sheetFromArray( _array )
            _sheet['!cols'] = wscols

            // console.log(JSON.stringify(_sheet))

            workbook.Sheets[sheet] = _sheet

        })


        var today = moment().format('M-D-YYYY h-mm-ssa')
        var filename = 'BLEND ' + today + '.xlsx'
        var filepath = path.join(__dirname, filename)

        xlsx.writeFile( workbook, filepath )
        open(filepath)

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
    yellow: 'FFFFFF00'
}

// for comparing row colors
var white     = ['FFFFFF', 'FFFF00'] // white, yellow
var blue      = ['00B0F0']
var green     = ['00FF00', '66FF33', '6AA84F']
var brown     = ['877852', '938950', '938953', '938954', '938955', '948A54', '988D55']
var red       = ['C00000', 'FF0000']
var magenta   = ['C27BA0', 'FF00FF', 'FF33CC']

// var _headers = ['Patient', 'Facility', 'LOC', 'Insurance', 'DOS From', 'DOS To', 'Billed', 'Allowed', 'Responsibility', 'Paid', 'Balance', 'Sent', 'Received', 'Payment Date', 'Check/Claim #', 'Notes', 'Units']

var _headers = ['REPORT', 'INV #', 'INS.', 'NAME', 'PAID', 'BILLED', 'ALLOWED', 'PT REPSONS.', 'DATE FROM', 'DATE TO', 'SENT', 'RECEIVED', 'DATE PAID', 'CHECK/CL #', 'ADDITIONS NOTES', 'F/U DATE', 'LOC', 'UNITS', 'BALANCE', 'TC']

var wscols = [ {wch:10}, {wch:12}, {wch:20}, {wch:20}, {wch:10}, {wch:12}, {wch:12}, {wch:14}, {wch:12}, {wch:12}, {wch:16}, {wch:12}, {wch:12}, {wch:18}, {wch:30}, {wch:28}, {wch:10}, {wch:10}, {wch:12}, {wch:20} ]


// excel workbook class
var Workbook = function(){
    if(!(this instanceof Workbook)) return new Workbook()

    this.SheetNames = []
    this.Sheets = {}
}

// create sheet from array
var sheetFromArray = function(data){
    var _worksheet = {};
    var range = { s: { c: 1000000, r: 1000000 }, e: { c: 0, r: 0 }};
    // var colCount = data[0].length;

    // cell indexes for alignment
    var align         = { headers: [2,3,13,14], others: [8,9,10,11,12,15,16,17] };
    // var align         = { headers: [0,3,14,15], others: [2,4,5,11,12,13,16] };
    // var alignWriteOff = { headers: [0,3,15,16], others: [1,2,4,5,12,13,14,17] };

    for (var R = 0; R != data.length; ++R){
        var rowColor = data[R][20]

        // for (var C = 0; C != data[R].length; ++C){
        for (var C = 0; C != 20; ++C){
            if(range.s.r > R) range.s.r = R;
            if(range.s.c > C) range.s.c = C;
            if(range.e.r < R) range.e.r = R;
            if(range.e.c < C) range.e.c = C;
            var cell = {v: data[R][C] };

            if(cell.v == null) continue;
            var cell_ref = xlsx.utils.encode_cell({c:C,r:R});

            if(typeof cell.v === 'number') cell.t = 'n';
            else if(typeof cell.v === 'boolean') cell.t = 'b';
            else if(cell.v instanceof Date) {
                cell.t = 'n'; cell.z = xlsx.SSF._table[14];
                cell.v = datenum(cell.v);
            }
            else cell.t = 's';

            cell.s = { 'font': { 'name': 'Verdana', 'sz' : 10 } }

            if (rowColor) cell.s.fill = { patternType: 'solid', fgColor: { rgb: rowColor }}

            // cell.s = { 'font': { 'name': 'Verdana', 'sz' : 10 } }

            // if (rowColor) cell.s.fill.bgColor.rgb = rowColor

            // cell alignment
            if (R == 0){ // set aligment, font weight and font size for columns in header row

                cell.s.font.bold = true

                if (align.headers.indexOf(C) > -1)
                    cell.s.alignment = { 'horizontal': 'left' };
                else
                    cell.s.alignment = { 'horizontal': 'center' }

            } else { // set alignment for all other columns
                if (align.others.indexOf(C) > -1) cell.s.alignment = { 'horizontal': 'center' };
            }

            // format currency cells
            if (((C > 3 && C < 8) || C == 18) && R !== 0){
                cell.t = 'n';
                cell.z = '$#,#0.00';
                cell.numFmt = '$0,000.00';
            }

            _worksheet[cell_ref] = cell;
        }

    }

    if(range.s.c < 1000000) _worksheet['!ref'] = xlsx.utils.encode_range(range);

    // console.log(JSON.stringify(xlsx.utils.decode_range(_worksheet['!ref'])))

    return _worksheet;
}

var prevBegin, prevEnd

var getRange = function(dates){

    var _dates = dates
                    // remove all whitespace
                    .map(function(d){ return d.replace(/\s/g, '')})
                    // remove everthing except numbers, /, -
                    .map(function(d){ return d.replace(/[^0-9\/-]/g,'') })
                    // remove empty values
                    .filter(function(d){ return d !== '' })
                    // split at '-'
                    .map(function(d){ return d.split('-') })
                    // flatten array
                    .reduce(function(a,b){ return a.concat(b) }, [])
                    // trim trailing dashes (helps when filtering out duplicates / empty values)
                    .map(function(d){ return d.replace(/\-+$/, '') })
                    // trim trailing forward slashes
                    .map(function(d){ return d.replace(/\/+$/, '') })
                    // trim leading zeroes (helps when filtering out duplicates)
                    .map(function(d){ return d.replace(/^0+/, '') })
                    // filter out duplicates (return unique values)
                    .filter(function(v,i,a){ return i == a.indexOf(v) })

    // find our year for fixes below
    var year
    _dates.forEach(function(date){
        var parts = date.split('/')
        if (parts.length == 3) year = parts.pop()
    })

    _dates.forEach(function(date, d){
        var slashes = (date.match(/\//g) || []).length

        if (slashes != 2){
            var newDate = date + '/' + year
            _dates[d] = newDate
        }
    })

    // if we have only one date, duplicate it so we have from/to
    if (_dates.length == 1) _dates.push(_dates[0])

    // _dates.map(function(d){ return moment(new Date(d)).format('M/D/YY') })
    _dates = _dates.map(function(d){ return moment(new Date(d)).format('M/D/YY') })

    return _dates
}



// // convert date ranges into separate properly formatted dates
// _getRange = function(dos, begin, end) {

//     if (dos) {

//         var dates = dos.split('-');

//         dates.sort();

//         if (dates.length === 1) {
//             begin = prevBegin = end = prevEnd = dates[0];
//         } else if (dates.length === 2) {
//             begin = prevBegin = dates[0];
//             end   = prevEnd   = dates[1];
//         } else if (dates.length === 0) {
//             begin = prevBegin;
//             end   = prevEnd;
//         }

//         // remove unwanted characters
//         dates = dates.map(function(v, i){
//             return dates[i] =  v.replace(/[^0-9\/]/g, "") // remove everything but numbers (0-9) and /
//                                    .replace(/\/0/g,"\/") // replace /0#  with /# for easier dup matching
//                                    .replace(/\b0+/g, "") // remove leading zeros
//                                    .replace(/\/\d{2}(\d{2})/, "/$1") // replace 4-digit year with 2-digit
//         });

//         var uniq = uniqueArray( dates )
//             .filter(function(v){ return v!=='' }) // remove empty strings from unique array;
//             .filter(function(v){ return v.split('/').length !== 1 }); // filter out terms without / in them, (e.g %, $, words, decimals)

//         for (var i=0; i<uniq.length;i++){
//             uniq[i] = moment( new Date( uniq[i] ) ).format('MM/DD/YY');
//         }

//         // determine the correct year
//         var year = '';

//         var l = uniq.length;
//         while(l--){
//             var split = uniq[l].split('/');
//             if (split.length === 3) {
//                                                             // console.log(dos)
//                 // console.log(dos, begin, end)
//                 year = split.pop();

//                                                             // console.log(year)
//                 break;
//             }
//         }

//         uniq = uniqueArray( uniq );
//         if (uniq.length === 1) uniq.push(uniq[0]);

//         uniq = uniq.map(function(v){
//             var split = v.split('/')
//                                                             // console.log(split, split.length)
//             if (split.length < 3){
//                 // console.log('missing year: ', split)
//                 // console.log(v + '/'+ year)
//                 // console.log('----------')
//                 return v + '/' + year
//             } else {
//                 // console.log('date is correct format: ', v)
//                 // console.log('----------')
//                 return v
//             }
//             // return v.split('/').length < 3 ? v + '/' + year : v;
//         });

//                                                         // console.log('-------------------')

//         if (!uniq.length) {
//             uniq = prevDates;
//         } else {
//             prevDates = uniq;
//         }

//         return uniq;

//     }

//     else if (begin || end) {
//         var beginDates = [],
//             endDates = [],
//             dates = [];

//         // split mult. dates in begin field
//         if (begin) {
//             beginDates = begin.split('-');
//             if (beginDates.length) dates = dates.concat( beginDates );
//         }

//         // split mult. dates in end field
//         if (end) {
//             endDates = end.split('-');
//             if (endDates.length) dates = dates.concat( endDates );
//         }

//         // remove unwanted characters
//         dates = dates.map(function(v, i){
//             return dates[i] =  v.replace(/[^0-9\/]/g, "") // remove everything but numbers (0-9) and /
//                                    .replace(/\/0/g,"\/") // replace /0#  with /# for easier dup matching
//                                    .replace(/\b0+/g, "") // remove leading zeros
//                                    .replace(/\/\d{2}(\d{2})/, "/$1") // replace 4-digit year with 2-digit
//         });


//         var uniq = uniqueArray( dates )
//             .filter(function(v){ return v!=='' }) // remove empty strings from unique array;
//             .filter(function(v){ return v.split('/').length !== 1 }); // filter out terms without / in them, (e.g %, $, words, decimals)

//         // determine the correct year
//         var year = '';

//         var l = uniq.length;
//         while(l--){
//             var split = uniq[l].split('/');
//             if (split.length === 3) {
//                 year = split.pop();
//                 break;
//             }
//         }

//         uniq = uniq.map(function(v){
//             return v.split('/').length < 3 ? v + '/' + year : v;
//         });

//         uniq = uniqueArray( uniq );

//         if (uniq.length === 1) uniq.push(uniq[0]);

//         for (var i=0; i<uniq.length;i++){
//             uniq[i] = moment( new Date( uniq[i] ) ).format('MM/DD/YY');
//         }

//         if (!uniq.length) {
//             uniq = prevDates;
//         } else {
//             prevDates = uniq;
//         }

//         return uniq;

//     }

//     else {
//         begin = moment( new Date( prevBegin ) ).format('MM/DD/YY');
//         end   = moment( new Date( prevEnd) ).format('MM/DD/YY');

//         var uniq = [ begin, end ];

//         return uniq;
//     }
// }

// return unique array
var uniqueArray = function(a) {
    var uniq = a.filter(function(item, pos) {
        return a.indexOf(item) == pos
    })

    return uniq
}

// get column header names
var getColumnNames = function(sheet){

    var range = xlsx.utils.decode_range(sheet[ '!ref' ])
    var columns = []
    var inc = 1

    for (var c = 0; c <= range.e.c; c++){
        var addr = xlsx.utils.encode_cell({ r: 0, c: c })
        var cell = sheet[ addr ]

        if (!cell){
            inc++
            cell = { v: 'Extra_' + inc }
        }
        // if (!cell) continue

        var column = cell.v.toString().replace(/[^a-zA-Z]/g, '').toUpperCase().trim()

        if (column === "DATEFROM") column = "BEGIN"
        else if (column === "DOSFROM") column = "BEGIN"
        else if (column === "DATETO") column = "END"
        else if (column === "DOSTO") column = "END"

        if (columns.indexOf(column) > -1 && column !== '') column += '2'
        // if (columns.indexOf(column) > -1) column += '2'
        columns.push( column )
    }

    return columns;
}

// convert currency to number
var currencyNumber = function(amount) {
    var f = null
    if(amount) f = parseFloat( amount.replace(/[^0-9\.-]+/g,"") )
    return f
}

// get color of row, use column F as reference
var getRowColor = function(sheet, row){
    var cell = 'F' + row

    if (sheet[cell] && sheet[cell].s){
        var s = sheet[cell].s

        if (s.fill){
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

var resetProgress = function(){
    // turn off progress bar transitions, reset to 0, restore transitions
    $('.progress-bar').addClass('notransition')
    updateProgress(0)
    $('.progress-bar').remove('notransition')
}

var updateProgress = function(val){
     $('.progress-bar').css('width', val + '%').attr('aria-valuenow', val)
}
