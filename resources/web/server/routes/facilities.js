'use strict'

const config = require('config')
const Router = require('express').Router
const EXCLUDEWORKBOOKS = config.get('excluded.workbooks')

module.exports = () => {
  const router = new Router()

  const listFacilities = (req, res, next) => {
    const client = req.user.googleClient
    client.list()
      .tap(files => console.log('got files?', files.length))
      .then(files => files
        .filter((file) => !~EXCLUDEWORKBOOKS.indexOf(file.name.toLowerCase()))
        .map((file) => ({ value: file.id, text: file.name }))
      )
      .then(facilities => ({ facilities }))
      .then(res.json.bind(res))
      .catch(next)
  }

  router.route('/')
      .get(listFacilities)

  return router
}
