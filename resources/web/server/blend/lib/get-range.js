
const moment = require('moment')

module.exports = function getRange (dates) {
  var _dates = dates
    .map((d) =>
    // remove all whitespace
    d.replace(/\s/g, '')
      // remove everthing except numbers, /, -
      .replace(/[^0-9\/-]/g, '')
  )
    // remove empty values
    .filter((d) => d !== '')
    // split at '-'
    .map((d) => d.split('-'))
    // flatten array
    .reduce((a, b) => a.concat(b), [])
    // trim trailing dashes (helps when filtering out duplicates / empty values)
    .map((d) => d.replace(/\-+$/, '')
      // trim trailing forward slashes
      .replace(/\/+$/, '')
      // trim leading zeroes (helps when filtering out duplicates)
      .replace(/^0+/, '')
  )
    // filter out duplicates (return unique values)
    .filter((v, i, a) => i === a.indexOf(v))

  // find our year for fixes below
  var year
  _dates.forEach(function (date) {
    var parts = date.split('/')
    if (parts.length === 3) year = parts.pop()
  })

  _dates.forEach(function (date, d) {
    var slashes = (date.match(/\//g) || []).length

    if (slashes !== 2) {
      var newDate = date + '/' + year
      _dates[d] = newDate
    }
  })

  // if we have only one date, duplicate it so we have from/to
  if (_dates.length === 1) _dates.push(_dates[0])

  // _dates.map(function(d){ return moment(new Date(d)).format('M/D/YY') })
  _dates = _dates.map(function (d) { return moment(new Date(d)).format('l') })

  return _dates
}
