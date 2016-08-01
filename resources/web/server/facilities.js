var google = require('googleapis'),
    drive = google.drive('v2'),
    JWT = google.auth.JWT,
    path = require('path'),
    get = require('object-path').get;

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
    open = require('open'),
    async = require('async'),
    glob = require('glob');

var EXCLUDEWORKBOOKS = ['laira', 'blank', 'request', 'alpine recovery lodge', 'olympus drug and alcohol', 'renaissance outpatient bountiful', 'renaissance ranch- orem', 'renaissance ranch- ut outpatient', ' old'];
var EXCLUDESHEETS = ['fax', 'copy', 'appeal', 'laira', 'checks', 'responses', 'ineligible'];

var selectedYears = []

module.exports.getFacilities = function(callback) {
	var facilities = [];
	async.waterfall([
	    function (next) {
	        ////////////////////////////////
	        // google drive authorization //
	        ////////////////////////////////
	        jwt.authorize(function (err, tokens) {
	            if (err) next(err)

	            jwt.credentials = tokens
	          	console.log('authorize', err, tokens);
	            next(null);
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
	            	console.log('drive');
	        })


	        req.on('response', function (res) {
	            var avgChunks = 250;
	            var numChunks = 0;
	            var data = ""

	            res.on('data', function (chunk) {
	                data += chunk
	                numChunks += 1
	                var percent = parseInt(numChunks * 100 / avgChunks)

	                // window.setTimeout(function () {
	                //     view.updateProgress(percent)
	                // }, 0)
	            })

	            res.on('end', function () {
	            	console.log('ending stream');
	                var json = JSON.parse(data)
	                if (json && json.items) {

	                    // window.setTimeout(function () {
	                    //     view.updateProgress(100)
	                    // }, 0)

	                    var files = json.items

	                    next(null,files);

	                    // setTimeout(function () {
	                    //     next(null, files)
	                    // }, 777)
	                }
	            })
	        })
	    },

	    function (files, next) {
        // filter out certain workbooks
        files = files
        	.filter((file) => {
        		return !~EXCLUDEWORKBOOKS.indexOf(file) && get(file, 'parents.0.id') === '0B_kSXk5v54QYeFRucElOQWlpdG8';
        	})
			.map((file) => ({
				value: file.id,
				text: file.title,
			}));

			next(null, files);
	    }

	], callback)
}

