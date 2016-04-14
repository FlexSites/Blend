"use strict";

var app = require('app')
var BrowserWindow = require('browser-window')
var Menu = require('menu')

var mainWindow = null

app.on('window-all-closed', function () {
    app.quit()
})

app.on('ready', function () {

    process.on('uncaughtException', function (err) {
        alert(err)
    })

    var name = app.getName()
    var template = [
        {
            label: 'BLEND',
            submenu: [
                {
                    label: 'About ' + name,
                    role: 'about'
                },

                {
                    type: 'separator'
                },

                {
                    label: 'Hide Electron',
                    accelerator: 'Command+H',
                    selector: 'hide:'
                },

                {
                    label: 'Hide Others',
                    accelerator: 'Command+Shift+H',
                    selector: 'hideOtherApplications:'
                },

                {
                    type: 'separator'
                },

                {
                    label: 'Reload',
                    accelerator: 'CmdOrCtrl+R',
                    click: function (item, focusedWindow) {
                        if (focusedWindow)
                            focusedWindow.reloadIgnoringCache();
                    }
                },

                {
                    label: 'Close',
                    accelerator: 'CmdOrCtrl+W',
                    role: 'close'
                },

                {
                    type: 'separator'
                },

                {
                    label: 'Toggle Developer Tools',
                    accelerator: (function () {
                        if (process.platform == 'darwin')
                            return 'Alt+Command+I';
                        else
                            return 'Ctrl+Shift+I';
                    })(),
                    click: function (item, focusedWindow) {
                        if (focusedWindow) focusedWindow.toggleDevTools();
                    }
                },

                {
                    type: 'separator'
                },

                {
                    label: 'Quit',
                    accelerator: 'Command+Q',
                    click: function () { app.quit() }
                }
            ]
        }
    ]

    var menu = Menu.buildFromTemplate(template)
    Menu.setApplicationMenu(menu)

    // mainWindow = new BrowserWindow({ width: 971, height: 1000, show: false, center: true, 'use-content-size': true, frame: false, resizable: true, 'accept-first-mouse': true })

    mainWindow = new BrowserWindow({
        width: 971,
        height: 600,
        frame: true,
        center: true,
        resizable: true,
        transparent: false,
        title: "Blend",
        'always-on-top': false,

        'use-content-size': false,
        'accept-first-mouse': true,
        show: false
    })
    mainWindow.loadURL('file://' + __dirname + '/index.html')

    mainWindow.show()

    // mainWindow.openDevTools()

})
