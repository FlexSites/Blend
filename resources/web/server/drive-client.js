'use strict'

const google = require('googleapis')
const drive = google.drive('v2')
const config = require('config')
const OAuth2 = google.auth.OAuth2
const assert = require('assert')
const Bluebird = require('bluebird')
const Retry = require('bluebird-retry')
const request = require('request-promise')
const fs = require('fs')

const CLIENT_ID = config.get('google.key')
const CLIENT_SECRET = config.get('google.secret')
const CLIENT_CALLBACK = config.get('google.callback')

module.exports = class Drive {
	constructor(access_token) {
		assert(access_token, 'Access token is required.')
		this.accessToken = access_token
		this.client = new OAuth2(CLIENT_ID, CLIENT_SECRET, CLIENT_CALLBACK)
		this.client.setCredentials({ access_token })
	}

	head(id) {
		return Bluebird.fromCallback(cb => drive.files.get({
          auth: this.client,
          fileId: id,
          forever: true,
          gzip: true
        }, cb))
	}

	get(fileId, pathname) {
		return drive.files.export({
			auth: this.client,
			fileId,
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
		})
		.on('end', function() {
		  console.log('Done');
		  cb()
		})
		.on('error', function(err) {
		  console.log('Error during download', err);
		})
		.pipe(fs.createWriteStream(pathname));
	}

	get_old(id) {
		var uri = `https://www.googleapis.com/drive/v2/files/${id}?alt=media`
        return request.get({
          uri,
          headers: {
        	Authorization: `Bearer ${this.accessToken}`
          },
        })
	}

	list(query, option) {
	  return Bluebird.fromCallback(cb => drive.children.list({
	    auth: this.client,
	    q: query,
	    folderId: option.id,
	    forever: true,
	    gzip: true
	  }, cb))
	}
}