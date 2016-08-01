var google = require('googleapis'),
    drive = google.drive('v2'),
    JWT = google.auth.JWT,
    request = require('request-promise'),
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

module.exports.generate = function(callback) {
	

}

