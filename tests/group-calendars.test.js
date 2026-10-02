"use strict";
var assert = require("assert"), fs = require("fs"), vm = require("vm"), path = require("path");
var root = path.resolve(__dirname, ".."), count = 0;
function test(name, action) { action(); count += 1; console.log("OK " + name); }
function environment() {
    var context = {window: {location: {pathname: "/sites/a/SiteAssets/app/index.html", protocol: "https:", host: "test.invalid"}}, Date: Date, JSON: JSON};
    vm.createContext(context);
    ["util", "sharepoint-data-source", "group-calendar-source", "sharepoint-settings-data-source"].forEach(function (name) {
        vm.runInContext(fs.readFileSync(path.join(root, "js", name + ".js"), "utf8"), context);
    });
    return context.window;
}
var A = "11111111-1111-1111-1111-111111111111", B = "22222222-2222-2222-2222-222222222222";
function fixture() {
    var w = environment(), source = new w.YoteihyouGroupCalendarSource({siteUrl: "https://test.invalid/sites/a"});
    source.configure({groups: [{name: "GP1", calendarListTitle: "A"}, {name: "GP2科", teams: [
        {name: "GS班", calendarListTitle: "B"}, {name: "未設定班"}]}]});
    source.metadata.load = function (guid, done) { done({key: "events:" + guid, data: {}}); };
    source.metadata.save = function (record, done) { done(); };
    source.entries.forEach(function (entry, i) {
        if (!entry.client) { return; }
        entry.client.request = function (method, url, headers, body, done) {
            done({responseText: JSON.stringify({d: {Id: i ? B : A, BaseTemplate: 106}})});
        };
        entry.client.load = function (range, done) { done([{id: 1, etag: '"1"', title: "予定", startDate: new Date(2026, 9, 2), endDate: new Date(2026, 9, 2)}]); };
    });
    return source;
}
function loaded(source) { var result; source.load({}, function (items, warning) { result = {items: items, warning: warning}; }, assert.fail); return result; }
test("別リストのID=1を区別し階層と設定順を保持", function () {
    var s = fixture(), r = loaded(s);
    assert.strictEqual(r.items.length, 2);
    assert.strictEqual(r.items[0].id, A + ":1"); assert.strictEqual(r.items[1].id, B + ":1");
    assert.strictEqual(r.items[1].purpose, "GP2科／GS班｜日々・週間・月間");
    assert.ok(r.warning.indexOf("未設定班：接続先未設定") >= 0);
});
test("一つの予定表の接続失敗でも他の予定表を返す", function () {
    var s = fixture(); s.entries[0].client.load = function (range, done, fail) { fail("権限なし"); };
    var r = loaded(s); assert.strictEqual(r.items.length, 1); assert.ok(r.warning.indexOf("GP1：権限なし") >= 0);
});
test("同じリストの二重割当を検知して予定を二重表示しない", function () {
    var s = fixture(); s.entries[1].client.request = s.entries[0].client.request;
    var r = loaded(s); assert.strictEqual(r.items.length, 1); assert.ok(r.warning.indexOf("複数のグループ") >= 0);
});
test("保存先を振り分け、標準予定表に色や目的を送らない", function () {
    var s = fixture(), r = loaded(s), calls = 0;
    s.entries[1].client.create = function (item, done) {
        calls += 1;
        var payload = this.toPayload(item, "SP.Data.EventsListItem");
        assert.ok(!("Purpose" in payload)); assert.ok(!("undefined" in payload));
        assert.ok(!("LineStyle" in payload)); assert.strictEqual(payload.fAllDayEvent, true);
        item.id = 7; done(item);
    };
    s.metadata.save = function (record, done) {
        assert.strictEqual(record.data[7].targets, "月間"); assert.strictEqual(record.data[7].lineColor, "red"); done();
    };
    var item = r.items[1]; item.id = ""; item.purpose = "GP2科／GS班｜月間"; item.allDay = true; item.lineColor = "red";
    s.create(item, function (saved) { assert.strictEqual(saved.id, B + ":7"); assert.strictEqual(saved.lineColor, "red"); }, assert.fail);
    assert.strictEqual(calls, 1);
});
test("更新と削除は元リストの数値IDとETagを使用", function () {
    var s = fixture(), item = loaded(s).items[1], called = [];
    ["update", "remove"].forEach(function (method) {
        s.entries[1].client[method] = function (plain, done) { called.push(method); assert.strictEqual(plain.id, 1); assert.strictEqual(plain.etag, '"1"'); done(plain); };
        s[method](item, function () {}, assert.fail);
    });
    assert.deepStrictEqual(called, ["update", "remove"]);
});
test("既存予定のグループ変更を誤ったリストへの更新にしない", function () {
    var s = fixture(), item = loaded(s).items[0], message = "";
    item.purpose = "GP2科／GS班｜日々";
    s.update(item, assert.fail, function (text) { message = text; }); assert.ok(message.indexOf("別の予定表") >= 0);
});
test("補助設定読込失敗時は予定本体を書き換えない", function () {
    var s = fixture(), item = loaded(s).items[0], failed = false;
    s.metadata.load = function (guid, done, fail) { fail("JSON不正"); };
    s.entries[0].client.update = assert.fail;
    s.update(item, assert.fail, function () { failed = true; }); assert.ok(failed);
});
test("予定保存後の補助設定競合は成功＋警告で返し二重登録を防ぐ", function () {
    var s = fixture(), item = loaded(s).items[0], warning = "";
    s.entries[0].client.create = function (plain, done) { plain.id = 8; done(plain); };
    s.metadata.save = function (record, done, fail) { fail("競合"); };
    s.create(item, function (saved, items, text) { warning = text; assert.strictEqual(saved.id, A + ":8"); }, assert.fail);
    assert.ok(warning.indexOf("変更は完了") >= 0);
});
test("古い期間の応答が新しい取得結果を上書きしない", function () {
    var s = fixture(), callbacks = [], oldFinished = false;
    s.entries[0].client.load = function (range, done) { callbacks.push(done); };
    s.load({}, function () { oldFinished = true; }, assert.fail);
    s.load({}, function () {}, assert.fail);
    callbacks[1]([{id: 9}]); callbacks[0]([{id: 4}]);
    assert.ok(!oldFinished); assert.ok(s.items[A + ":9"]); assert.ok(!s.items[A + ":4"]);
});
test("親グループと小グループの両方に接続先を設定できる", function () {
    var s = fixture(); s.configure({groups: [{name: "科", calendarListTitle: "科予定", teams: [{name: "班", calendarListTitle: "班予定"}]}]});
    assert.strictEqual(s.entries.length, 2); assert.strictEqual(s.entries[0].team, ""); assert.strictEqual(s.entries[1].team, "班");
});
test("標準予定表GETに独自列を含めず標準終日列を取得", function () {
    var s = fixture(), client = s.entries[0].client;
    client.request = function (method, url, headers, body, done) {
        url = decodeURIComponent(url); assert.ok(url.indexOf("fAllDayEvent") >= 0);
        ["Purpose", "LineStyle", "LineColor", "TextColor", ",,"].forEach(function (text) { assert.strictEqual(url.indexOf(text), -1); });
        done({responseText: '{"d":{"results":[]}}'});
    };
    environment().YoteihyouSharePointDataSource.prototype.load.call(client, {}, function () {}, assert.fail);
});
test("補助設定の空データ・HTML混入・重複を区別", function () {
    var s = new (environment().YoteihyouGroupCalendarSource)({siteUrl: "https://test.invalid"}), rows = [], errors = 0;
    s.metadata.client.request = function (m, u, h, b, done) { done({responseText: JSON.stringify({d: {results: rows}})}); };
    s.metadata.load(A, function (r) { assert.strictEqual(Object.keys(r.data).length, 0); }, assert.fail);
    rows = [{ID: 1, LayoutJson: "<div>{}</div>"}]; s.metadata.load(A, assert.fail, function () { errors += 1; });
    rows = [{ID: 1}, {ID: 2}]; s.metadata.load(A, assert.fail, function () { errors += 1; }); assert.strictEqual(errors, 2);
});
test("補助設定保存はETag付きで独立したLayoutJsonレコードへ書く", function () {
    var s = new (environment().YoteihyouGroupCalendarSource)({siteUrl: "https://test.invalid"}), client = s.metadata.client, called = false;
    client.getEntityType = function (done) { done("SP.Data.LayoutListItem"); }; client.getDigest = function (done) { done("digest"); };
    client.request = function (m, u, h, b, done) {
        assert.strictEqual(h["IF-MATCH"], '"3"'); assert.strictEqual(JSON.parse(b).Title, "events:" + A); called = true; done();
    };
    s.metadata.save({key: "events:" + A, id: 2, etag: '"3"', data: {}}, function () {}, assert.fail); assert.ok(called);
});
test("繰り返し予定を誤って毎日の予定として表示しない", function () {
    var s = fixture(); s.entries[0].client.load = function (range, done) { done([{id: 3, recurring: true}]); };
    var r = loaded(s); assert.strictEqual(r.items.length, 1); assert.ok(r.warning.indexOf("繰り返し予定") >= 0);
});
test("組織設定の接続先を親子で保持し標準列に保存", function () {
    var C = environment().YoteihyouOrganizationSettingsDataSource;
    var s = new C({siteUrl: "https://test.invalid", organizationSettings: {fields: {
        id: "ID", groupName: "Title", teamName: "TeamName", calendarSiteUrl: "CalendarSiteUrl", calendarListTitle: "CalendarListTitle",
        monthlyRows: "MonthlyRows", weeklyRows: "WeeklyRows", dailyRows: "DailyRows", sortOrder: "SortOrder", isActive: "IsActive"
    }}});
    var parent = s.toItem({ID: 1, Title: "科", CalendarListTitle: "科予定", CalendarSiteUrl: "https://test.invalid"});
    var child = s.toItem({ID: 2, Title: "科", TeamName: "班", CalendarListTitle: "班予定"});
    var config = s.toOrganizationConfig([parent, child], {});
    assert.strictEqual(config.groups[0].calendarListTitle, "科予定");
    assert.strictEqual(config.groups[0].teams[0].calendarListTitle, "班予定");
    assert.strictEqual(s.toPayload(parent, "type").CalendarSiteUrl, "https://test.invalid");
});
test("グループ未登録を予定なしとして成功表示しない", function () {
    var s = fixture(); s.configure({groups: []});
    var r = loaded(s); assert.ok(r.warning.indexOf("有効なグループがありません") >= 0);
});
console.log(count + " group-calendar tests passed");
