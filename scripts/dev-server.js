"use strict";

var fs = require("fs");
var http = require("http");
var path = require("path");

var host = "127.0.0.1";
var port = parseInt(process.env.YOTEIHYOU_PORT || "8765", 10);
var root = path.resolve(__dirname, "..");
var rootPrefix = root.toLowerCase() + path.sep;
var contentTypes = {
    ".css": "text/css; charset=utf-8",
    ".csv": "text/csv; charset=utf-8",
    ".gif": "image/gif",
    ".htm": "text/html; charset=utf-8",
    ".html": "text/html; charset=utf-8",
    ".ico": "image/x-icon",
    ".jpeg": "image/jpeg",
    ".jpg": "image/jpeg",
    ".js": "text/javascript; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".png": "image/png",
    ".svg": "image/svg+xml"
};

if (isNaN(port) || port < 1 || port > 65535) {
    process.stderr.write("YOTEIHYOU_PORT must be a number from 1 to 65535.\n");
    process.exit(1);
}

function sendText(response, statusCode, text) {
    var body = Buffer.from(text, "utf8");
    response.writeHead(statusCode, {
        "Content-Type": "text/plain; charset=utf-8",
        "Content-Length": body.length,
        "Cache-Control": "no-store"
    });
    response.end(body);
}

function resolveRequestPath(requestUrl) {
    var pathname;
    var relativePath;
    var targetPath;
    try {
        pathname = decodeURIComponent(new URL(requestUrl, "http://" + host).pathname || "/");
    } catch (error) {
        return null;
    }
    relativePath = pathname === "/" ? "index.html" : pathname.replace(/^\/+/, "");
    targetPath = path.resolve(root, relativePath);
    if (targetPath !== root && targetPath.toLowerCase().indexOf(rootPrefix) !== 0) {
        return null;
    }
    return targetPath;
}

var server = http.createServer(function (request, response) {
    var targetPath = resolveRequestPath(request.url || "/");
    if (!targetPath) {
        sendText(response, 400, "Bad request");
        return;
    }
    fs.stat(targetPath, function (statError, stats) {
        if (statError || !stats.isFile()) {
            sendText(response, 404, "Not found");
            return;
        }
        fs.readFile(targetPath, function (readError, data) {
            if (readError) {
                sendText(response, 500, "Unable to read file");
                return;
            }
            response.writeHead(200, {
                "Content-Type": contentTypes[path.extname(targetPath).toLowerCase()] || "application/octet-stream",
                "Content-Length": data.length,
                "Cache-Control": "no-store"
            });
            if (request.method === "HEAD") {
                response.end();
            } else {
                response.end(data);
            }
        });
    });
});

server.on("error", function (error) {
    if (error.code === "EADDRINUSE") {
        process.stderr.write("Port " + port + " is already in use. Stop the existing debug session and press F5 again.\n");
    } else {
        process.stderr.write(error.message + "\n");
    }
    process.exit(1);
});

server.listen(port, host, function () {
    process.stdout.write("Development server ready at http://" + host + ":" + port + "\n");
});
