"use strict";
const { app, BrowserWindow } = require('electron')
const path = require("path");
const absPath = path.resolve()

const { main } = require("../funct/funct1.js")

main()


const createWindow = () => {
    const win = new BrowserWindow({
        width: 800,
        height: 600,
        icon: absPath + "\\assets\\logo-etergy.ico",
        webPreferences: {
            plugins: true,
            enableRemoteModule: true,
            backgroundThrottling: false,
            sandbox: false,
            preload: path.join(__dirname, "preload.js"),
        },
    })
    win.loadFile('main/index.html').then(() => win.webContents.send("absPath", absPath))
}

app.whenReady().then(() => {
    createWindow()
    app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow() })
    app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit() })
})

