'use strict'

const google = require('googleapis')
const drive = google.drive('v3')
const sheets = google.sheets('v4')
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
const FOLDER_ID = config.get('google.drive.folderId')

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

	get(spreadsheetId, pathname) {
		return Bluebird.fromCallback(cb => sheets.spreadsheets.get({
			auth: this.client,
			spreadsheetId,
		}, cb))
	}

	list() {
	  return Bluebird.fromCallback(cb => drive.files.list({
	    auth: this.client,
	    pageSize: 1000,
	    q: `'${FOLDER_ID}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false`,
     	orderBy: 'name',
	  }, cb))
	  .then(res => res.files)
	}
}