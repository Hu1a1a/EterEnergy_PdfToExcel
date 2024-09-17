"use strict";
const { app, BrowserWindow } = require('electron')
const path = require("path");
const absPath = path.resolve()
const express = require('express');
const server = express();
const { main } = require('../funct/funct1');
const { main2 } = require('../funct/funct2');

server.get("/funct1", (req, res) => main().then(() => res.sendStatus(200)));
server.get("/funct2", (req, res) => main2().then(() => res.sendStatus(200)));
server.listen(4321);

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
    win.loadFile('main/index.html')
}

app.whenReady().then(() => {
    createWindow()
    app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow() })
    app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit() })
})

