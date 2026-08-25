// serveFruitBasket.js -- serve the web fruit basket samples with Node.
//
// Some browsers refuse to load a page's neighbours over the file
// protocol, and the React sample needs its libraries fetched. Running a
// small server sidesteps both. Node ships with everything used here --
// no npm install, no packages, no build step -- so this works the
// moment Node itself is installed, which the EdSharp installer offers
// as a checkbox.
//
//     node serveFruitBasket.js
//     node serveFruitBasket.js -port 8080
//
// It prints the address, opens the default browser, and serves the
// files in its own folder until you press Control+C.
//
// Written in Camel Type: Hungarian prefixes, lower camel case, one
// require per line in alphabetical order.

"use strict";

const childProcess = require("child_process");
const fs = require("fs");
const http = require("http");
const path = require("path");

const dTypes = {
    ".css": "text/css; charset=utf-8",
    ".htm": "text/html; charset=utf-8",
    ".html": "text/html; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".md": "text/plain; charset=utf-8",
    ".txt": "text/plain; charset=utf-8"
};

// The port asked for on the command line, or 8000.
function readPort() {
    var iAt, iPort;
    iAt = process.argv.indexOf("-port");
    if (iAt < 0) return 8000;
    iPort = Number(process.argv[iAt + 1]);
    if (!iPort || iPort < 1 || iPort > 65535) return 8000;
    return iPort;
}

// The file a request asks for, kept inside this folder whatever the
// request says.
function resolveFile(sUrl) {
    var sPath, sSafe;
    sPath = decodeURIComponent(sUrl.split("?")[0]);
    if (sPath === "/" || sPath === "") sPath = "/index.htm";
    sSafe = path.normalize(path.join(__dirname, sPath));
    if (!sSafe.startsWith(__dirname)) return "";
    return sSafe;
}

// A plain page listing the samples, so the address opens somewhere
// useful rather than on a directory listing.
function indexPage() {
    var lsLines, lsNames;
    lsNames = fs.readdirSync(__dirname).filter(function (sName) {
        return sName.toLowerCase().endsWith(".htm") && sName.toLowerCase() !== "index.htm";
    }).sort();
    lsLines = ["<!DOCTYPE html>", "<html lang=\"en\">", "<head>", "<meta charset=\"utf-8\" />",
        "<title>Fruit Basket Samples</title>", "</head>", "<body>", "<h1>Fruit Basket Samples</h1>", "<ul>"];
    lsNames.forEach(function (sName) {
        lsLines.push("<li><a href=\"" + sName + "\">" + sName.replace(".htm", "") + "</a></li>");
    });
    lsLines.push("</ul>", "</body>", "</html>");
    return lsLines.join("\n");
}

function startServer() {
    var iPort, serverFruit, sAddress;
    iPort = readPort();
    serverFruit = http.createServer(function (oRequest, oResponse) {
        var binBody, sFile, sType;
        sFile = resolveFile(oRequest.url);
        if (sFile === "") {
            oResponse.writeHead(403, {"Content-Type": "text/plain; charset=utf-8"});
            oResponse.end("Outside the samples folder.");
            return;
        }
        if (!fs.existsSync(sFile) && (oRequest.url === "/" || oRequest.url === "")) {
            oResponse.writeHead(200, {"Content-Type": dTypes[".htm"]});
            oResponse.end(indexPage());
            return;
        }
        if (!fs.existsSync(sFile) || fs.statSync(sFile).isDirectory()) {
            oResponse.writeHead(404, {"Content-Type": "text/plain; charset=utf-8"});
            oResponse.end("No such file: " + oRequest.url);
            return;
        }
        sType = dTypes[path.extname(sFile).toLowerCase()] || "application/octet-stream";
        binBody = fs.readFileSync(sFile);
        oResponse.writeHead(200, {"Content-Type": sType});
        oResponse.end(binBody);
    });

    serverFruit.listen(iPort, function () {
        sAddress = "http://localhost:" + iPort + "/";
        console.log("Serving the fruit basket samples at " + sAddress);
        console.log("Press Control+C to stop.");
        try {
            if (process.platform === "win32") childProcess.exec("start \"\" \"" + sAddress + "\"");
            else if (process.platform === "darwin") childProcess.exec("open \"" + sAddress + "\"");
            else childProcess.exec("xdg-open \"" + sAddress + "\"");
        } catch (oError) {
            console.log("Open that address in your browser.");
        }
    });

    serverFruit.on("error", function (oError) {
        if (oError.code === "EADDRINUSE") console.log("Port " + iPort + " is busy. Try: node serveFruitBasket.js -port 8080");
        else console.log("The server could not start: " + oError.message);
        process.exit(1);
    });
}

startServer();
