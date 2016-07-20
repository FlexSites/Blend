"use strict";
var express = require('express');
var open = require("open");
var path = require('path');
var url = require('url');
var mime = require('mime')
var fs = require('fs')


var app = express();

var facilities = require('./server/facilities');
var blend = require('./server/new-blend');

app.use(express.static('static'));
app.use(express.static('app'));
app.use(express.static('node_modules'))

app.get('/', function (req, res) {
	res.sendFile(path.join(__dirname + 'static/index.html'));
});

var router = express.Router();

router.get('/', function(req, res) {
	res.json({ message: 'api is here'});
});



router.route('/facilities')
	.get(function(req, res) {
		facilities.getFacilities(function(err, facilities){
			return res.json({'facilities': facilities})
		});
	});

router.route('/blend')
	.get(function(req, res, next) {
		var url_parts = url.parse(req.url, true);
		blend(JSON.parse(url_parts.query.options), JSON.parse(url_parts.query.years), url_parts.query.txpoc, url_parts.query.color, function(err, blendFile){
			if (err) return next(err);
			res.download(blendFile, 'blend-file.xlsx', (err) => {
				console.log('sendfile worked', err, blendFile)
			})
			// var filename = path.basename(blendFile)
			// var mimetype = mime.lookup(filename)
			// res.setHeader('Content-Disposition', 'attachment; filename=blend.xlsx')
			// res.setHeader('Content-Type', 'application/octet-stream')
			// console.log('sending stream', filename, mimetype, blendFile)
			// var filestream = fs.createReadStream(blendFile)
			// filestream.pipe(res)
		});
	});


app.use('/api', router);

app.use((err, req, res, next) => {
	console.log(err)
	res.send(err)
})





app.listen(3000, function () {
	console.log('Example app listening on port 3000!');
});

