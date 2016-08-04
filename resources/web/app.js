"use strict";

var express = require('express');
var open = require("open");
var path = require('path');
var url = require('url');
var mime = require('mime')
var fs = require('fs')
var Drive = require('./server/drive-client')
// var cookieParser = require('cookie-parser')
// var session = require('connect-session')
var passport = require('./server/authentication')


var app = express();

app.set('view engine', 'pug');
app.set('views', __dirname + '/static');

var facilities = require('./server/facilities');
var blend = require('./server/new-blend');

function isAuthenticated(req, res, next) {
	if (!req.isAuthenticated()) return res.redirect('/login')
	else if (req.url === '/login') return res.redirect('/')
	next()
}

app.get('/login', (req, res, next) => {
	res.render('login.pug', {

	})
})


app.use(express.static('node_modules'));

app.use(passport())
app.get('/', 
	isAuthenticated, 
	(req, res) => res.render('index', { user: req.user.emails[0].value })
);

app.use(express.static('static'));

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
		blend(new Drive(req.user.accessToken), JSON.parse(url_parts.query.options), JSON.parse(url_parts.query.years), url_parts.query.txpoc, url_parts.query.color, function(err, blendFile){
			if (err) return next(err);
			res.download(blendFile, 'blend-file.xlsx', (err) => {
				console.info('sendfile worked', err, blendFile)
			})
			// var filename = path.basename(blendFile)
			// var mimetype = mime.lookup(filename)
			// res.setHeader('Content-Disposition', 'attachment; filename=blend.xlsx')
			// res.setHeader('Content-Type', 'application/octet-stream')
			// console.log('sending stream', filename, mimetype, blendFile)
			// var filestream = fs.createReadStream(blendFile)
			// filestream.pipe(res)
		})
	})

app.use('/api', isAuthenticated, router)

app.use((err, req, res, next) => {
	console.error(err)
	res.render('login.pug', {
		errorMessage: err.message,
	})
})





app.listen(process.env.PORT || 3000, function () {
	console.info('Blend listening on port 3000!');
});

