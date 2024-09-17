"use strict";
const { contextBridge, shell } = require("electron");
const path = require("path");
contextBridge.exposeInMainWorld("electron", {
    absPath: () => path.resolve(),
    opePath: (file) => shell.openPath(path.resolve() + file)
});
