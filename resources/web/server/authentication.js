const passport = require('passport')
const GoogleStrategy = require('passport-google-oauth20').Strategy
const config = require('config')
const Router = require('express').Router
const cookieParser = require('cookie-parser')
const session = require('express-session')


module.exports = () => {
  passport.use(
    new GoogleStrategy({
      clientID: config.get('google.key'),
      clientSecret: config.get('google.secret'),
      callbackURL: config.get('google.callback')
    },
    function (accessToken, refreshToken, profile, cb) {
      console.log('got a user', Object.assign({ accessToken, refreshToken }, profile))
      return cb(null, Object.assign({ accessToken, refreshToken }, profile))
    })
  )

  let router = Router()

  passport.serializeUser(function (user, done) {
    console.log('serialize', user)
    done(null, user)
  })

  passport.deserializeUser(function (user, done) {
    console.log('deserialize', user)
    done(null, user)
  })

  router.use(cookieParser())
  router.use(session({ secret: 'what the what' }))
  router.use(passport.initialize())
  router.use(passport.session())
  router.use((req, res, next) => {
    console.log('after login', req.user)
    next()
  })
  router.get('/user', (req, res) => {
    console.log('called user', req.user)
    res.send(req.user)
  })

  router.get('/auth/google',
    passport.authenticate('google', { scope: ['https://www.googleapis.com/auth/plus.login'] }))

  router.get('/auth/google/callback',
    passport.authenticate('google', { failureRedirect: '/' }),
    function (req, res) {
      res.redirect('/')
    })

  return router
}
