const passport = require('passport')
const GoogleStrategy = require('passport-google-oauth20').Strategy
const config = require('config')
const Router = require('express').Router
const cookieParser = require('cookie-parser')
const session = require('express-session')
const Drive = require('./google-client')


module.exports = () => {
  passport.use(
    new GoogleStrategy({
      clientID: config.get('google.key'),
      clientSecret: config.get('google.secret'),
      callbackURL: config.get('google.callback')
    },
    function (accessToken, refreshToken, profile, cb) {
      console.log('profile', profile);
      if(profile._json.domain !== "elevatedbilling.com"){
        return cb(new Error("Invalid host domain"));
      }
      const googleClient = new Drive(accessToken)
      return cb(null, Object.assign({ googleClient, accessToken, refreshToken }, profile))
    })
  )

  var router = Router()

  passport.serializeUser(function (user, done) {
    done(null, user)
  })

  passport.deserializeUser(function (user, done) {
    user.googleClient = new Drive(user.accessToken)
    done(null, user)
  })

  router.use(cookieParser())
  router.use(session({ secret: 'what the what' }))
  router.use(passport.initialize())
  router.use(passport.session())
  router.use((req, res, next) => {
    next()
  })
  router.get('/user', (req, res) => {
    res.send(req.user)
  })

  router.get('/auth/google',
    passport.authenticate('google', {
      scope: [
        'https://www.googleapis.com/auth/plus.login',
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/userinfo.profile',
        'https://www.googleapis.com/auth/userinfo.email'
      ]
  }))

  router.get('/auth/google/callback',
    passport.authenticate('google', { failureRedirect: '/login' }),
    function (req, res) {
      res.redirect('/')
    }); 

  return router
}
