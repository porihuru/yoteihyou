"use strict";
var fs = require("fs"), path = require("path"), vm = require("vm"), assert = require("assert");
var app = fs.readFileSync(path.join(__dirname, "../js/app.js"), "utf8");
var fn = app.slice(app.indexOf("    function constrainMonthlyCaptions("), app.indexOf("    function refreshMonthlyCaptions("));
var count = 0;
[0.9, 1, 1.25].forEach(function (scale) {
    [140, 900].forEach(function (textWidth) {
        var tableWidth = 600, start = 550;
        var caption = {style: {}, scrollWidth: textWidth};
        function width() { return Math.min(parseFloat(caption.style.width) || 50, parseFloat(caption.style.maxWidth) || Infinity); }
        Object.defineProperty(caption, "offsetWidth", {get: width});
        caption.getBoundingClientRect = function () {
            var left = (start + (parseFloat(caption.style.left) || 0)) * scale;
            return {left: left, right: left + width() * scale, width: width() * scale};
        };
        var table = {querySelectorAll: function () { return [caption]; }, getBoundingClientRect: function () {
            return {left: 0, right: tableWidth * scale};
        }};
        var scope = vm.createContext({}); vm.runInContext(fn, scope);
        [600, 400, 800].forEach(function (size) {
            tableWidth = size; start = size - 50;
            scope.constrainMonthlyCaptions(table);
            var rect = caption.getBoundingClientRect();
            assert.ok(rect.right <= tableWidth * scale - 2 + 0.001);
            assert.ok(rect.left >= -0.001);
            scope.constrainMonthlyCaptions(table);
            assert.ok(Math.abs(caption.getBoundingClientRect().right - rect.right) < 0.001);
        });
        count += 1;
    });
});
assert.ok(app.indexOf("refreshMonthlyCaptions();") >= 0);
console.log(count + " caption-bound tests passed (90/100/125%, resize, long titles, repeated layout)");
