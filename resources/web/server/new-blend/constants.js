module.exports = {
  PALETTE: {
    blue: '0000FF',
    green: '00FF00',
    brown: '8B4513',
    red: 'FF0000',
    magenta: 'FF00FF',
    yellow: 'FFFF00',
    orange: 'FF9900'
  },

  // for comparing row colors
  white: ['FFFFFF', 'FFFF00'],
  blue: ['00B0F0', '24AEFF'],
  green: ['00FF00', '66FF33', '6AA84F'],
  brown: ['877852', '938950', '938953', '938954', '938955', '948A54', '988D55'],
  red: ['C00000', 'FF0000'],
  magenta: ['C27BA0', 'FF00FF', 'FF33CC'],
  orange: ['FF9900'],

  HEADERS: ['REPORT', 'INV #', 'INS.', 'NAME', 'PAID', 'BILLED', 'ALLOWED', 'PT REPSONS.', 'DATE FROM', 'DATE TO', 'SENT', 'RECEIVED', 'DATE PAID', 'CHECK/CL #', 'BULK AMOUNT', 'ADDITIONS NOTES', 'F/U DATE', 'LOC', 'UNITS', 'BALANCE', 'TC'],

  EXCLUDEWORKBOOKS: ['laira', 'blank', 'request', 'alpine recovery lodge', 'olympus drug and alcohol', 'renaissance outpatient bountiful', 'renaissance ranch- orem', 'renaissance ranch- ut outpatient', ' old'],

  EXCLUDESHEETS: ['fax', 'copy', 'appeal', 'laira', 'checks', 'responses', 'ineligible'],

  EXCEL_MIMETYPE: "mimeType = 'application/vnd.google-apps.spreadsheet'"
}

