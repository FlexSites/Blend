const passport = require('passport')
const GoogleStrategy = require('passport-google-oauth20').Strategy
const config = require('config')

passport.use(new GoogleStrategy({
    clientID: config.get('google.key'),
    clientSecret: config.get('google.secret'),
    callbackURL: config.get('google.callback')
  },
  function(accessToken, refreshToken, profile, cb) {
    User.findOrCreate({ googleId: profile.id }, function (err, user) {
      console.log('got a user', err, user)
      return cb(err, user);
    });
  }
));

module.exports = passport
