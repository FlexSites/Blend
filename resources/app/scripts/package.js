var path = require('path')
var package = require('electron-packager')
var json = require('jsonfile')

var buildPath = path.join( __dirname, 'build.json')
var buildJson = json.readFileSync(buildPath)

var rootPath = path.resolve(__dirname, '..')

buildJson.dir = rootPath
buildJson.sourcedir = rootPath
buildJson.out = path.join( rootPath, 'releases' )

package(buildJson, function(err, appPath){
    if(err){
        return error(err)
    }
    
    console.log('Application successfully built at ', appPath)
})