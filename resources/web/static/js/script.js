require(['jquery', 'view', 'async'], function ($, view, async) {


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






});

