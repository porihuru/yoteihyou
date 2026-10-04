"use strict";
var fs = require("fs"), vm = require("vm"), assert = require("assert"), path = require("path");
var app = fs.readFileSync(path.join(__dirname, "../js/app.js"), "utf8");
var fn = app.slice(app.indexOf("    function syncFixedTimeAxis()"), app.indexOf("    function updateFixedHeader()"));
[0.5, 1, 1.25, 2].forEach(function (scale) {
    var edges = [0, 155.25, 201.625, 248.125, 294.5];
    var cells = edges.slice(1).map(function () { return {style: {}}; });
    var sourceCells = cells.map(function (_, i) { return {getBoundingClientRect: function () {
        return {left: (edges[i] - 83.5) * scale, top: 50 * scale,
            width: (edges[i + 1] - edges[i]) * scale, height: 30.5 * scale};
    }}; });
    var head = {getBoundingClientRect: function () { return {height: 30.5 * scale, top: 50 * scale}; },
        getElementsByTagName: function () { return sourceCells; }};
    var table = {getBoundingClientRect: function () { return {left: -83.5 * scale, width: 294.5 * scale}; },
        getElementsByTagName: function () { return [head]; }};
    var copy = {style: {}, getElementsByTagName: function () { return cells; }};
    var context = {fixedHeader: {axisOverlay: {firstChild: copy, style: {}}}, state: {viewMode: "monthly"},
        currentDisplayZoom: scale * 100, byId: function () { return {style: {}, getElementsByTagName: function () { return [table]; }}; }};
    vm.runInNewContext(fn + "syncFixedTimeAxis();", context);
    cells.forEach(function (cell, i) {
        assert.strictEqual(parseFloat(cell.style.left), edges[i]);
        assert.strictEqual(parseFloat(cell.style.width), edges[i + 1] - edges[i]);
        assert.strictEqual(parseFloat(cell.style.height), 30.5);
    });
    assert.strictEqual(copy.style.height, "30.5px");
});
var css = fs.readFileSync(path.join(__dirname, "../css/style.css"), "utf8");
assert.ok(/\.fixed-daily-axis th\s*\{[^}]*position: absolute/.test(css));
assert.ok(app.indexOf('fixedHeader.backdrop.style.display = "none"') >= 0);
assert.ok(app.indexOf('Math.ceil(toolbarBottom / scale)') >= 0);
assert.ok(app.indexOf('copy.firstChild.removeAttribute("id")') >= 0);
console.log("5 fixed-header tests passed (fractional columns, horizontal scroll, 50/100/125/200% zoom, opaque backdrop)");
