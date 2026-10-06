"use strict";
var fs = require("fs"), vm = require("vm"), assert = require("assert"), path = require("path");
var elements = {}, requests = [], editable = true, changed = 0;
["target", "status", "save", "reload", "inherit", "time", "line"].forEach(function (id) {
    elements["weekly-group-" + id] = {value: "", options: [], appendChild: function (o) { this.options.push(o); }, selectedIndex: 0};
});
var context = {window: {YoteihyouUtil: {getJson: function (xhr) { return xhr; }},
    YoteihyouSharePointDataSource: function () {
        this.siteUrl = "https://example.invalid";
        this.getApiUrl = this.getListPath = function (s) { return s || "list"; };
        this.getEntityType = this.getDigest = function (done) { done("test"); };
        this.request = function (method, url, headers, body, done, fail) {
            requests.push({method: method, url: url, headers: headers, body: body, done: done, fail: fail});
        };
    }}, document: {getElementById: function (id) { return elements[id]; },
        createElement: function () { return {appendChild: function () {}}; }, createTextNode: function (s) { return s; }}};
vm.runInNewContext(fs.readFileSync(path.join(__dirname, "../js/weekly-group-settings.js"), "utf8"), context);
var C = context.window.YoteihyouWeeklyGroups;
var defaults = {showTime: true, showMultiDayLine: false};
var settings = new C({}, function () { return defaults; }, function () { return []; }, function () { return editable; }, function () { changed += 1; });
settings.rows = C.validate([["A", "", false, true], ["A", "KY", true, false], ["B", "KY", false, false]]);
assert.strictEqual(settings.get("A", "").showMultiDayLine, true);
assert.strictEqual(settings.get("A", "KY").showTime, true);
assert.strictEqual(settings.get("B", "KY").showTime, false);
assert.strictEqual(settings.get("A", "other"), defaults);
assert.throws(function () { C.validate([["A", "", 1, false]]); });
assert.throws(function () { C.validate([["A", "", true, false], ["A", "", false, true]]); });
settings.initialize();
requests.pop().done({d: {results: [{ID: 4, __metadata: {etag: '"2"'}, LayoutJson: JSON.stringify(settings.rows)}]}});
var target = elements["weekly-group-target"], save = elements["weekly-group-save"];
target.value = JSON.stringify(["A", "KY"]); settings.show();
elements["weekly-group-time"].checked = false; elements["weekly-group-line"].checked = true;
editable = false; save.onclick(); assert.strictEqual(requests.length, 0);
editable = true; save.onclick(); var request = requests.pop();
assert.strictEqual(request.headers["IF-MATCH"], '"2"');
var payload = JSON.parse(request.body), rows = JSON.parse(payload.LayoutJson);
assert.strictEqual(payload.Title, "weekly-groups-v1"); assert.strictEqual(rows.length, 3);
assert.strictEqual(settings.get("A", "KY").showTime, true);
request.fail("HTTP 412"); assert.strictEqual(settings.get("A", "KY").showTime, true);
assert.strictEqual(settings.record, null);
elements["weekly-group-reload"].onclick(); requests.pop().done({d: {results: [{ID: 4, __metadata: {etag: '"3"'}, LayoutJson: JSON.stringify(rows)}]}});
assert.strictEqual(settings.get("A", "KY").showMultiDayLine, true);
target.value = JSON.stringify(["A", "KY"]); settings.show(); elements["weekly-group-inherit"].checked = true;
save.onclick(); request = requests.pop(); assert.strictEqual(JSON.parse(JSON.parse(request.body).LayoutJson).length, 2);
request.done(); requests.pop().done({d: {results: [{ID: 4, __metadata: {etag: '"4"'}, LayoutJson: JSON.parse(request.body).LayoutJson}]}});
assert.strictEqual(settings.get("A", "KY"), defaults);
assert.strictEqual(settings.get("A", "").showTime, false);
console.log("Weekly group settings passed: independent parent/child and duplicate names, defaults, validation, read-only, ETag conflict, reset and reload");
