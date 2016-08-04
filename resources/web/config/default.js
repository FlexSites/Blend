const path = require('path')

module.exports = {
  google: {
    drive: {
      email: '780463185858-924al4jpqutjbvcbq6bp1lk91t4gmt34@developer.gserviceaccount.com',
      id: '780463185858-924al4jpqutjbvcbq6bp1lk91t4gmt34.apps.googleusercontent.com',
      keyPath: path.join(__dirname, 'elevated-numbers.pem'),
      scope: ['https://www.googleapis.com/auth/drive'],
      folderId: '0B_kSXk5v54QYeFRucElOQWlpdG8'
    }
  },
  excluded: {
    workbooks: ['laira', 'blank', 'request', 'alpine recovery lodge', 'olympus drug and alcohol', 'renaissance outpatient bountiful', 'renaissance ranch- orem', 'renaissance ranch- ut outpatient', ' old'],
    worksheets: ['fax', 'copy', 'appeal', 'laira', 'checks', 'responses', 'ineligible']
  },
  defaultSheets: [
    'Open',
    'Write Off',
    'Payment Member',
    'Payment Facility',
    'Orange',
    'Confirmed Paid',
    'Denied'
  ]
}
