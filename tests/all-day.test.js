"use strict";
var assert = require("assert"), fs = require("fs"), vm = require("vm"), path = require("path");
var root = path.resolve(__dirname, ".."), count = 0;
function test(name, fn) { fn(); count += 1; console.log("OK " + name); }
var nodes = {"all-day": {checked: true, disabled: false}};
["start", "end"].forEach(function (p) {
    ["date", "time", "time-label", "time-picker", "time-toggle", "time-options", "calendar"].forEach(function (s) {
        nodes[p + "-" + s] = {value: "", style: {}, setAttribute: function () {}};
    });
});
var ctx = {window: {location: {}}, document: {getElementById: function (id) { return nodes[id]; }}, Date: Date};
vm.createContext(ctx);
["util", "sharepoint-data-source", "date-time-editor"].forEach(function (f) {
    vm.runInContext(fs.readFileSync(path.join(root, "js", f + ".js"), "utf8"), ctx);
});
var w = ctx.window, util = w.YoteihyouUtil, editor = w.YoteihyouDateTimeEditor;
// Simulate the site's local-date normalization on REST writes (not a payload echo).
function storedRow(payload) {
    var row = JSON.parse(JSON.stringify(payload));
    if (row.fAllDayEvent) {
        row.EventDate = util.formatDateKey(new Date(payload.EventDate)) + "T00:00:00Z";
        row.EndDate = util.formatDateKey(new Date(payload.EndDate)) + "T23:59:00Z";
    }
    return row;
}
function source(flag) { return new w.YoteihyouSharePointDataSource({siteUrl: "https://test.invalid", listTitle: "calendar", fields: {
    id: "ID", title: "Title", startDate: "EventDate", endDate: "EndDate", allDay: flag || "fAllDayEvent",
    category: "Category", location: "Location", description: "Description"
}}); }
test("日本時間で写真の09:00から翌日08:59へのずれを防ぐ", function () {
    process.env.TZ = "Asia/Tokyo";
    var item = source().toItem({EventDate: "2026-11-15T00:00:00Z", EndDate: "2026-11-15T23:59:00Z", fAllDayEvent: true});
    editor.set("start", item.startDate); editor.set("end", item.endDate);
    assert.strictEqual(nodes["start-date"].value, "2026-11-15");
    assert.strictEqual(nodes["end-date"].value, "2026-11-15");
    assert.strictEqual(nodes["start-time"].value, "0000");
    assert.strictEqual(nodes["end-time"].value, "2359");
});
test("1日・複数日・月年またぎ・閏日を繰り返し保存しても日付を保つ", function () {
    ["Asia/Tokyo", "UTC", "America/Los_Angeles"].forEach(function (tz) {
        process.env.TZ = tz;
        [["2026-11-15", "2026-11-15"], ["2026-11-15", "2026-11-18"],
            ["2026-12-31", "2027-01-02"], ["2028-02-29", "2028-03-01"], ["2026-03-08", "2026-03-09"]].forEach(function (dates) {
            var s = source(), row = {EventDate: dates[0] + "T00:00:00Z", EndDate: dates[1] + "T23:59:00Z", fAllDayEvent: true}, i, item;
            for (i = 0; i < 5; i += 1) {
                item = s.toItem(row);
                assert.strictEqual(util.formatDateKey(item.startDate), dates[0]);
                assert.strictEqual(util.formatDateKey(item.endDate), dates[1]);
                editor.set("start", item.startDate); editor.set("end", item.endDate);
                item.startDate = editor.read("start", true); item.endDate = editor.read("end", true);
                row = storedRow(s.toPayload(item, "Event"));
                assert.strictEqual(row.EventDate, dates[0] + "T00:00:00Z");
                assert.strictEqual(row.EndDate, dates[1] + "T23:59:00Z");
            }
        });
    });
    process.env.TZ = "Asia/Tokyo";
});
test("終日は逆転した時刻入力を無視しチェック操作で日付を変えない", function () {
    nodes["start-date"].value = nodes["end-date"].value = "2026-11-15";
    nodes["start-time"].value = "0900"; nodes["end-time"].value = "0859";
    nodes["all-day"].checked = true; editor.sync();
    assert.strictEqual(nodes["start-time-picker"].style.display, "none");
    assert.strictEqual(nodes["end-time-label"].style.display, "none");
    assert.strictEqual(util.formatDateTime(editor.read("start", true)), "2026-11-15 00:00");
    assert.strictEqual(util.formatDateTime(editor.read("end", true)), "2026-11-15 23:59");
    nodes["all-day"].checked = false; editor.sync();
    assert.strictEqual(nodes["start-time-picker"].style.display, "");
    assert.strictEqual(nodes["start-time"].disabled, false);
    assert.strictEqual(nodes["end-date"].value, "2026-11-15");
    assert.ok(editor.read("end", false) < editor.read("start", false));
});
test("通常予定と旧独自AllDay列の時差換算を維持", function () {
    [source(), source("AllDay")].forEach(function (s, i) {
        var row = {EventDate: "2026-11-15T00:00:00Z", EndDate: "2026-11-15T01:00:00Z", AllDay: true, fAllDayEvent: false};
        var item = s.toItem(row), payload = s.toPayload(item, "Event");
        assert.strictEqual(item.startDate.getHours(), 9);
        assert.strictEqual(payload.EventDate, row.EventDate);
        assert.strictEqual(item.allDay, !!i);
    });
});
test("サーバーの終日フラグ絞込を使わず日付の余裕を持って取得", function () {
    var s = source(), url;
    s.request = function (method, target, headers, body, done) { url = decodeURIComponent(target); done({responseText: '{"d":{"results":[]}}'}); };
    s.load({startDate: new Date(2026, 10, 15), endDate: new Date(2026, 10, 16)}, function () {}, assert.fail);
    assert.strictEqual(url.indexOf("fAllDayEvent eq"), -1);
    assert.ok(url.indexOf("EventDate lt datetime'2026-11-16T15:00:00Z'") >= 0);
    assert.ok(url.indexOf("EndDate ge datetime'2026-11-13T15:00:00Z'") >= 0);
});
test("月初・月末の終日予定を残し前後月の予定を除く（ページング含む）", function () {
    var s = source(), calls = 0, result;
    function row(id, start, end, allDay) { return {ID: id, EventDate: start, EndDate: end, fAllDayEvent: allDay}; }
    s.request = function (m, u, h, b, done) {
        calls += 1;
        done({responseText: JSON.stringify({d: calls === 1 ? {results: [
            row(1, "2026-11-01T00:00:00Z", "2026-11-01T23:59:00Z", true),
            row(2, "2026-10-31T00:00:00Z", "2026-10-31T23:59:00Z", true)
        ], __next: "next-page"} : {results: [
            row(3, "2026-11-30T00:00:00Z", "2026-11-30T23:59:00Z", true),
            row(4, "2026-12-01T00:00:00Z", "2026-12-01T23:59:00Z", true),
            row(5, "2026-10-31T15:00:00Z", "2026-10-31T16:00:00Z", false),
            row(6, "2026-10-30T00:00:00Z", "2026-11-02T23:59:00Z", true)
        ]}})});
    };
    s.load({startDate: new Date(2026, 10, 1), endDate: new Date(2026, 11, 1)}, function (items) { result = items; }, assert.fail);
    assert.strictEqual(calls, 2);
    assert.strictEqual(result.map(function (x) { return x.id; }).join(","), "1,3,5,6");
});
test("終日へ変更して保存・再取得した予定が月間の対象日に残る", function () {
    var s = source(), saved, result;
    var item = s.toItem({ID: 7, EventDate: "2026-11-15T00:00:00Z", EndDate: "2026-11-15T01:00:00Z", fAllDayEvent: false});
    editor.set("start", item.startDate); editor.set("end", item.endDate);
    item.allDay = true; item.startDate = editor.read("start", true); item.endDate = editor.read("end", true);
    saved = storedRow(s.toPayload(item, "Event")); saved.ID = 7;
    s.request = function (m, u, h, b, done) { done({responseText: JSON.stringify({d: {results: [saved]}})}); };
    s.load({startDate: new Date(2026, 10, 1), endDate: new Date(2026, 11, 1)}, function (items) { result = items; }, assert.fail);
    assert.strictEqual(result.length, 1);
    var app = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var scope = vm.createContext({startOfDay: function (d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }});
    vm.runInContext(app.slice(app.indexOf("    function itemOccursOn("), app.indexOf("    function getLayoutScope(")), scope);
    assert.strictEqual(scope.itemOccursOn(result[0], new Date(2026, 10, 15)), true);
    assert.strictEqual(scope.itemOccursOn(result[0], new Date(2026, 10, 16)), false);
});
test("既に保存された複数日を推測で短縮しない", function () {
    var item = source().toItem({EventDate: "2026-11-15T00:00:00Z", EndDate: "2026-11-16T23:59:00Z", fAllDayEvent: true});
    assert.strictEqual(util.formatDateKey(item.endDate), "2026-11-16");
});
test("日本時間の1日終日を前日15時から当日14時59分のUTCで送信", function () {
    process.env.TZ = "Asia/Tokyo";
    var p = source().toPayload({startDate: new Date(2026, 10, 15), endDate: new Date(2026, 10, 15, 23, 59), allDay: true}, "Event");
    assert.strictEqual(p.EventDate, "2026-11-14T15:00:00Z");
    assert.strictEqual(p.EndDate, "2026-11-15T14:59:00Z");
});
test("終日をドラッグして保存先で変換・再取得しても1日と複数日の日数を維持", function () {
    var app = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var scope = vm.createContext({Date: Date, Math: Math, DAILY_SNAP_MINUTES: 15,
        cloneScheduleItem: function (x) { var y = {}; Object.keys(x).forEach(function (k) { y[k] = x[k]; }); return y; },
        splitPurpose: function () { return {targets: ["monthly"]}; }, joinPurpose: function () { return "GP1｜月間"; }});
    vm.runInContext(app.slice(app.indexOf("    function moveItemToTarget("), app.indexOf("    function pasteDailyItem(")), scope);
    ["Asia/Tokyo", "America/Los_Angeles"].forEach(function (tz) {
        process.env.TZ = tz;
        [0, 2].forEach(function (span) {
            var s = source(), item = {startDate: new Date(2026, 2, 7), endDate: new Date(2026, 2, 7 + span, 23, 59), allDay: true};
            [new Date(2026, 2, 8), new Date(2026, 11, 31), new Date(2027, 0, 4)].forEach(function (target) {
                item = scope.moveItemToTarget(item, {day: target, section: "GP1", team: "", minute: 540});
                item = s.toItem(storedRow(s.toPayload(item, "Event")));
                assert.strictEqual(util.formatDateKey(item.startDate), util.formatDateKey(target));
                assert.strictEqual(util.formatDateKey(item.endDate), util.formatDateKey(new Date(target.getFullYear(), target.getMonth(), target.getDate() + span)));
            });
        });
    });
    process.env.TZ = "Asia/Tokyo";
});
test("保存後の日付を実際にGETし不一致と読込失敗を警告、再POSTしない", function () {
    ["match", "mismatch", "error"].forEach(function (mode) {
        var s = source(), posts = 0, gets = 0, warning;
        var item = {id: 1, etag: '"1"', title: "試験", allDay: true,
            startDate: new Date(2026, 10, 15), endDate: new Date(2026, 10, 15, 23, 59)};
        s.getEntityType = function (done) { done("Event"); }; s.getDigest = function (done) { done("digest"); };
        s.request = function (method, url, headers, body, done, fail) {
            if (method === "POST") { posts += 1; done({responseText: ""}); return; }
            gets += 1;
            if (mode === "error") { fail("通信失敗"); return; }
            done({responseText: JSON.stringify({d: {EventDate: "2026-11-15T00:00:00Z",
                EndDate: mode === "match" ? "2026-11-15T23:59:00Z" : "2026-11-16T23:59:00Z"}})});
        };
        s.update(item, function (saved, items, message) { warning = message || ""; }, assert.fail);
        assert.strictEqual(posts, 1); assert.strictEqual(gets, 1);
        if (mode === "match") { assert.strictEqual(warning, ""); }
        else { assert.ok(warning.indexOf(mode === "error" ? "確認に失敗" : "2026-11-16") >= 0); }
    });
});
console.log(count + " all-day tests passed");
