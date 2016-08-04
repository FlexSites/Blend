"use strict"

const express = require('express')
const open = require("open")
const path = require('path')
const url = require('url')
const mime = require('mime')
const fs = require('fs')
// const cookieParser = require('cookie-parser')
// const session = require('connect-session')
const passport = require('./server/authentication')


const app = express()

app.set('view engine', 'pug')
app.set('views', __dirname + '/static')

const facilities = require('./server/routes/facilities')
const blend = require('./server/routes/blend')

function isAuthenticated(req, res, next) {
	if (!req.isAuthenticated()) return res.redirect('/login')
	else if (req.url === '/login') return res.redirect('/')
	next()
}

app.get('/login', (req, res, next) => {
	res.render('login.pug', {

	})
})


app.use(express.static('node_modules'))

app.use(passport())
app.get('/', 
	isAuthenticated, 
	(req, res) => res.render('index', { user: req.user.emails[0].value })
)

app.use(express.static('static'))

var router = express.Router()

router.get('/', function(req, res) {
	res.json({ message: 'api is here'})
})

router.use('/facilities', facilities())
router.use('/blend', blend())

app.use('/api', isAuthenticated, router)

app.use((err, req, res, next) => {
	console.error(err)
	res.render('login.pug', {
		errorMessage: err.message,
	})
})





app.listen(process.env.PORT || 3000, function () {
	console.info('Blend listening on port 3000!')
})

