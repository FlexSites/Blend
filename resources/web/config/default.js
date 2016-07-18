const path = require('path')

module.exports = {
  google: {
    drive: {
      email: '780463185858-924al4jpqutjbvcbq6bp1lk91t4gmt34@developer.gserviceaccount.com',
      id: '780463185858-924al4jpqutjbvcbq6bp1lk91t4gmt34.apps.googleusercontent.com',
      keyPath: path.join(__dirname, 'elevated-numbers.pem'),
      scope: ['https://www.googleapis.com/auth/drive']
    }
  }
}
