"use strict";
var express = require('express');
var open = require("open");
var path = require('path');


var app = express();

var facilities = require('./server/facilities');

app.use(express.static('static'));
app.use(express.static('app'));
app.use(express.static('node_modules'))

app.get('/', function (req, res) {
	res.sendFile(path.join(__dirname + 'static/index.html'));
});

var router = express.Router();

router.get('/', function(req, res) {
	res.json({ message: 'api is here'});
	console.log(facilities);
	console.log(facilities.getFacilities());
});



router.route('/facilities')
	.get(function(req, res) {

	});

app.use('/api', router);





app.listen(3000, function () {
	console.log('Example app listening on port 3000!');
	open("http://localhost:3000/");
});

