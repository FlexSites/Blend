const Color = require('color')
const diff = require('color-diff')

let { PALETTE, WHITE } = require('../constants')
PALETTE = Object.keys(PALETTE).map(key => hexToRGB(PALETTE[key]))

module.exports = (hex) => {
  if (typeof hex !== 'string') {
    return WHITE
  }

  if (hex.length > 6) hex = hex.substr(0, 6)
  return rgbToHex(diff.closest(hexToRGB(hex), PALETTE))
}

function keyCase (obj, isUppercase) {
  return Object.keys(obj)
    .reduce((result, key) => {
      result[isUppercase ? key.toUpperCase() : key.toLowerCase()] = obj[key]
      return result
    }, {})
}

function hexToRGB (hex) {
  const color = Color(`#${hex}`)
  return keyCase(color.rgb(), true)
}

function rgbToHex (rgb) {
  const color = Color(keyCase(rgb))
  return color.hexString()
}
