"use strict";
var assert = require("assert"), fs = require("fs"), vm = require("vm"), path = require("path");
var root = path.resolve(__dirname, ".."), count = 0;
function test(name, fn) { fn(); count += 1; console.log("OK " + name); }
var nodes = {};
["holiday-input", "holiday-status", "holiday-save", "holiday-reload"].forEach(function (id) { nodes[id] = {}; });
var ctx = {window: {location: {}, confirm: function () { return true; }}, document: {getElementById: function (id) { return nodes[id]; }}};
vm.createContext(ctx);
["js/util.js", "js/sharepoint-data-source.js", "config/holidays.js", "js/holiday-settings.js"].forEach(function (p) {
    vm.runInContext(fs.readFileSync(path.join(root, p), "utf8"), ctx);
});
var H = ctx.window.YoteihyouHolidays, options = {siteUrl: "https://test.invalid"};
test("3年49件、2026国民の休日と2027振替休日、2028未確定日なし", function () {
    var rows = H.validate(ctx.window.YOTEIHYOU_HOLIDAYS), counts = {};
    rows.forEach(function (r) { var y = r[0].slice(0, 4); counts[y] = (counts[y] || 0) + 1; });
    assert.deepStrictEqual(counts, {2026: 18, 2027: 17, 2028: 14});
    var h = new H(options);
    assert.strictEqual(h.name(new Date(2026, 8, 22)), "国民の休日");
    assert.strictEqual(h.name(new Date(2027, 2, 22)), "振替休日");
    assert.strictEqual(h.name(new Date(2028, 2, 20)), "");
    assert.strictEqual(h.name(new Date(2028, 8, 22)), "");
    assert.strictEqual(h.name(new Date(2028, 6, 17)), "海の日");
});
test("休日は土曜日より赤を優先し通常土日は従来どおり", function () {
    var h = new H(options);
    assert.strictEqual(h.dayClass(new Date(2028, 0, 1)), "sunday");
    assert.strictEqual(h.dayClass(new Date(2026, 9, 3)), "saturday");
    assert.strictEqual(h.dayClass(new Date(2026, 9, 4)), "sunday");
    assert.strictEqual(h.dayClass(new Date(2026, 9, 5)), "");
});
test("追加修正削除、空一覧、重複と不正日付の検証", function () {
    assert.strictEqual(H.parse("2026-12-29 | 独自休日\n2026-01-01 | 元日")[0][0], "2026-01-01");
    assert.strictEqual(H.parse("").length, 0);
    ["2026-02-30 | 不正", "2026-01-01 | ", "2026-01-01 | A\n2026-01-01 | B"].forEach(function (s) { assert.throws(function () { H.parse(s); }); });
});
test("共有読込で未登録・空一覧・破損・重複を区別", function () {
    var s = new H.Source(options), rows = [], result, failed = 0;
    s.client.request = function (m, url, h, b, done) { assert.ok(decodeURIComponent(url).indexOf("Title eq 'holidays-v1'") >= 0); done({responseText: JSON.stringify({d: {results: rows}})}); };
    function done(x) { result = x; }
    s.load(done, assert.fail); assert.strictEqual(result.rows.length, 49);
    rows = [{ID: 1, LayoutJson: "[]", __metadata: {etag: '"2"'}}]; s.load(done, assert.fail); assert.strictEqual(result.rows.length, 0);
    rows[0].LayoutJson = "<div>[]</div>"; s.load(assert.fail, function () { failed += 1; });
    rows.push(rows[0]); s.load(assert.fail, function () { failed += 1; }); assert.strictEqual(failed, 2);
});
test("共有保存は専用項目とETagを使い無条件上書きしない", function () {
    var s = new H.Source(options), called = 0;
    s.client.getEntityType = function (done) { done("Layout"); }; s.client.getDigest = function (done) { done("digest"); };
    s.client.request = function (m, u, h, b, done) {
        assert.strictEqual(h["IF-MATCH"], '"2"'); assert.strictEqual(JSON.parse(b).Title, "holidays-v1"); called += 1; done();
    };
    s.save({id: 1, etag: '"2"'}, [], function () {}, assert.fail);
    s.save({id: 1}, [], assert.fail, function () {}); assert.strictEqual(called, 1);
});
test("設定画面は読取専用を守り、保存成功後だけ反映", function () {
    var edit = false, changed = 0, h = new H(options, function () { return edit; }, function () { changed += 1; });
    var rows = [], writes = 0;
    h.source.load = function (done) { done({id: 1, etag: '"1"', rows: rows}); };
    h.source.save = function (record, value, done) { writes += 1; rows = value; done(); };
    h.initialize(); nodes["holiday-input"].value = "2026-12-29 | 独自休日";
    nodes["holiday-save"].onclick(); assert.strictEqual(writes, 0);
    edit = true; nodes["holiday-save"].onclick(); assert.strictEqual(writes, 1);
    assert.strictEqual(h.name(new Date(2026, 11, 29)), "独自休日");
    assert.ok(changed >= 2);
    h.source.save = function (record, value, done, fail) { fail("競合"); };
    nodes["holiday-input"].value = ""; nodes["holiday-save"].onclick();
    assert.strictEqual(h.name(new Date(2026, 11, 29)), "独自休日");
    assert.ok(nodes["holiday-status"].textContent.indexOf("競合") >= 0);
});
console.log(count + " holiday tests passed");
