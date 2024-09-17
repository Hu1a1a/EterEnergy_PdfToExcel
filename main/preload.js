"use strict";
const { contextBridge, ipcRenderer } = require("electron");

const { main } = require("../funct/funct1.js")
const { main2 } = require("../funct/funct2.js")

contextBridge.exposeInMainWorld("electron", {
  Funct1: () => main(),
  Funct2: () => main2()
});
