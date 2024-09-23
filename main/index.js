const path = window.electron.absPath()

function goTo(file) {
    window.electron.opePath(file)
}
function funct1() {
    httpGet("funct1")
}
function funct2() {
    httpGet("funct2")
}
function httpGet(theUrl) {
    const xmlHttp = new XMLHttpRequest();
    xmlHttp.open("GET", "http://localhost:4321/" + theUrl, false);
    xmlHttp.send(null);
    return xmlHttp.responseText;
}