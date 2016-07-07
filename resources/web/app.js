"use strict";
var express = require('express');
var open = require("open");
var path = require('path');
var url = require('url');


var app = express();

var facilities = require('./server/facilities');
var blend = require('./server/blend');

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
			console.log('what the crap!!!', err, facilities)
			return res.json({'facilities': facilities})
		});
	});

router.route('/blend')
	.get(function(req, res) {
		var url_parts = url.parse(req.url, true);
		blend.getBlend(url_parts.query.options, url_parts.query.years, url_parts.query.txpoc, url_parts.query.color, function(err, blendFile){
			return res.sendFile(blendFile);
		});
	});


app.use('/api', router);





app.listen(3000, function () {
	console.log('Example app listening on port 3000!');
});

