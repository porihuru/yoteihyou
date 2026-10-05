"use strict";

var assert = require("assert");
var fs = require("fs");
var path = require("path");
var vm = require("vm");
var root = path.resolve(__dirname, "..");
var passed = 0;

function test(name, action) {
    action();
    passed += 1;
    process.stdout.write("OK " + name + "\n");
}

function loadBrowserScripts(files) {
    var browserWindow = {
        location: {
            protocol: "https:",
            host: "example.invalid",
            pathname: "/sites/test/SiteAssets/yoteihyou/index.html"
        },
        _spPageContextInfo: null
    };
    var context = vm.createContext({
        window: browserWindow,
        console: console,
        Date: Date,
        JSON: JSON,
        XMLHttpRequest: function () {}
    });
    files.forEach(function (file) {
        var source = fs.readFileSync(path.join(root, file), "utf8");
        vm.runInContext(source, context, {filename: file});
    });
    return context;
}

var context = loadBrowserScripts([
    "js/util.js",
    "js/csv-data-source.js",
    "js/sharepoint-data-source.js",
    "js/history-data-source.js",
    "js/access-counter.js",
    "js/data-service.js",
    "js/sharepoint-settings-data-source.js"
]);
var util = context.window.YoteihyouUtil;
var CsvDataSource = context.window.YoteihyouCsvDataSource;
var SharePointDataSource = context.window.YoteihyouSharePointDataSource;
var HistoryDataSource = context.window.YoteihyouHistoryDataSource;
var AccessCounter = context.window.YoteihyouAccessCounter;
var DataService = context.window.YoteihyouDataService;
var OrganizationSettingsDataSource = context.window.YoteihyouOrganizationSettingsDataSource;
var options = {
    siteUrl: "https://example.invalid/sites/test",
    listTitle: "予定表",
    pageSize: 500,
    fields: {
        id: "ID",
        title: "Title",
        startDate: "EventDate",
        endDate: "EndDate",
        category: "Category",
        location: "Location",
        description: "Description",
        purpose: "Purpose"
    }
};
var organizationSettingsOptions = {
    siteUrl: "https://example.invalid/sites/test",
    organizationSettings: {
        enabled: true,
        listTitle: "予定表組織設定",
        fields: {
            id: "ID",
            groupName: "Title",
            teamName: "TeamName",
            monthlyRows: "MonthlyRows",
            weeklyRows: "WeeklyRows",
            dailyRows: "DailyRows",
            autoRows: "AutoRows",
            sortOrder: "SortOrder",
            isActive: "IsActive"
        }
    }
};

test("日時を解析できる", function () {
    var date = util.parseDate("2026-09-16 09:05");
    assert.ok(date);
    assert.strictEqual(util.formatDateTime(date), "2026-09-16 09:05");
});

test("実在しない日付を拒否する", function () {
    assert.strictEqual(util.parseDate("2026-02-30 09:00"), null);
});

test("範囲外の時分を拒否する", function () {
    assert.strictEqual(util.parseDate("2026-09-16 24:00"), null);
    assert.strictEqual(util.parseDate("2026-09-16 09:60"), null);
});

test("日時の後ろに余分な文字がある値を拒否する", function () {
    assert.strictEqual(util.parseDate("2026-09-16 09:00 invalid"), null);
});

test("週間表示は月曜日から日曜日までとする", function () {
    var appSource = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var scope = vm.createContext({Date: Date});
    var start = appSource.indexOf("    function startOfDay(");
    var end = appSource.indexOf("    function formatJapaneseDate(");
    var monday;

    vm.runInContext(appSource.slice(start, end), scope);
    monday = scope.startOfWeek(new Date(2026, 8, 23));
    assert.strictEqual(monday.getTime(), new Date(2026, 8, 21).getTime());
    monday = scope.startOfWeek(new Date(2026, 8, 27));
    assert.strictEqual(monday.getTime(), new Date(2026, 8, 21).getTime());
    monday = scope.startOfWeek(new Date(2026, 8, 28));
    assert.strictEqual(monday.getTime(), new Date(2026, 8, 28).getTime());
});

test("CSVの予定を画面上で追加・更新・削除できる", function () {
    var source = new CsvDataSource({url: "data/schedule.csv"});
    var items;
    var item = {
        id: "",
        title: "一時予定",
        startDate: new Date(2026, 8, 16, 9, 0),
        endDate: new Date(2026, 8, 16, 10, 0)
    };
    assert.strictEqual(source.readOnly, false);
    source.create(item, function (created, snapshot) {
        items = snapshot;
    }, function (message) {
        throw new Error(message);
    });
    assert.strictEqual(items.length, 1);
    assert.ok(item.id.indexOf("csv-local-") === 0);

    item.title = "変更後";
    source.update(item, function (updated, snapshot) {
        items = snapshot;
    }, function (message) {
        throw new Error(message);
    });
    assert.strictEqual(items[0].title, "変更後");

    source.remove(item, function (removed, snapshot) {
        items = snapshot;
    }, function (message) {
        throw new Error(message);
    });
    assert.strictEqual(items.length, 0);
});

test("SharePointの設定項目を組織設定へ変換できる", function () {
    var source = new OrganizationSettingsDataSource(organizationSettingsOptions);
    var defaultAuto = source.toItem({ID: 10, Title: "既存設定", MonthlyRows: 5, WeeklyRows: 5, DailyRows: 5});
    var fixedRows = source.toItem({ID: 11, Title: "固定設定", MonthlyRows: 5, WeeklyRows: 5, DailyRows: 5, AutoRows: false});
    var converted = source.toOrganizationConfig([
        {id: 3, groupName: "GP2", teamName: "", monthlyRows: 1, weeklyRows: 2, dailyRows: 3, autoRows: true, sortOrder: 20, isActive: true},
        {id: 1, groupName: "GP1科", teamName: "GS班", monthlyRows: 5, weeklyRows: 2, dailyRows: 2, autoRows: false, sortOrder: 10, isActive: true},
        {id: 2, groupName: "GP1科", teamName: "JN班", monthlyRows: 4, weeklyRows: 3, dailyRows: 2, autoRows: true, sortOrder: 11, isActive: true},
        {id: 4, groupName: "無効", teamName: "", monthlyRows: 1, weeklyRows: 1, dailyRows: 1, sortOrder: 30, isActive: false}
    ], {separator: "／", metadataSeparator: "｜", targetSeparator: "・"});
    assert.strictEqual(defaultAuto.autoRows, true);
    assert.strictEqual(fixedRows.autoRows, false);
    assert.strictEqual(converted.groups.length, 2);
    assert.strictEqual(converted.groups[0].name, "GP1科");
    assert.strictEqual(converted.groups[0].teams.length, 2);
    assert.strictEqual(converted.groups[0].teams[1].name, "JN班");
    assert.strictEqual(converted.groups[0].teams[0].autoRows, false);
    assert.strictEqual(converted.groups[0].teams[1].autoRows, true);
    assert.strictEqual(converted.groups[1].name, "GP2");
    assert.strictEqual(converted.groups[1].dailyRows, 3);
    assert.strictEqual(converted.groups[1].autoRows, true);
});

test("設定画面はパスワードsnkを入力したときだけ開く", function () {
    var script = fs.readFileSync(path.join(root, "js/settings-controller.js"), "utf8");
    var panel = {style: {display: "none"}};
    var input = null;
    var alerts = 0;
    var opened = 0;
    var controller;
    var scope = vm.createContext({
        SettingsController: function () {},
        window: {
            prompt: function () { return input; },
            alert: function () { alerts += 1; }
        },
        byId: function () { return panel; }
    });
    vm.runInContext(script.slice(script.indexOf("    SettingsController.prototype.open = function"),
        script.indexOf("    SettingsController.prototype.close = function")), scope);
    controller = new scope.SettingsController();
    controller.onOpen = function () { opened += 1; };
    controller.clearForm = function () {};
    controller.renderItems = function () {};
    controller.load = function () {};
    controller.open();
    assert.strictEqual(panel.style.display, "none");
    assert.strictEqual(alerts, 0);
    input = "wrong";
    controller.open();
    assert.strictEqual(panel.style.display, "none");
    assert.strictEqual(alerts, 1);
    input = "snk";
    controller.open();
    assert.strictEqual(panel.style.display, "block");
    assert.strictEqual(opened, 1);
});

test("組織設定の更新時にETagを送る", function () {
    var source = new OrganizationSettingsDataSource(organizationSettingsOptions);
    var sentHeaders;
    var sentPayload;
    var item = {
        id: 5,
        etag: "\"8\"",
        groupName: "GP1",
        teamName: "",
        monthlyRows: 5,
        weeklyRows: 4,
        dailyRows: 3,
        autoRows: true,
        sortOrder: 10,
        isActive: true
    };
    source.client.getEntityType = function (success) {
        success("SP.Data.SettingsListItem");
    };
    source.client.getDigest = function (success) {
        success("digest");
    };
    source.client.request = function (method, url, headers, body, success) {
        sentHeaders = headers;
        sentPayload = JSON.parse(body);
        success({responseText: ""});
    };
    source.update(item, function () {}, function (message) {
        throw new Error(message);
    });
    assert.strictEqual(sentHeaders["IF-MATCH"], "\"8\"");
    assert.strictEqual(sentPayload.AutoRows, true);
    assert.strictEqual(sentHeaders["X-HTTP-Method"], "MERGE");
    assert.strictEqual(sentPayload.Title, "GP1");
    assert.strictEqual(sentPayload.WeeklyRows, 4);
});

test("SharePointのETagを予定へ保持する", function () {
    var source = new SharePointDataSource(options);
    var item = source.toItem({
        ID: 10,
        Title: "会議",
        EventDate: "2026-09-16T00:00:00Z",
        EndDate: "2026-09-16T01:00:00Z",
        Category: "",
        Location: "",
        Description: "",
        Purpose: "GP1｜日々",
        __metadata: {etag: "\"3\""}
    });
    assert.strictEqual(item.etag, "\"3\"");
});

test("予定の線設定とSharePointの登録情報を読み書きできる", function () {
    var settings = JSON.parse(JSON.stringify(options));
    var source;
    var item;
    var payload;
    settings.fields.lineStyle = "LineStyle";
    settings.fields.lineColor = "LineColor";
    settings.fields.textColor = "TextColor";
    source = new SharePointDataSource(settings);
    item = source.toItem({
        ID: 12, Title: "確認", EventDate: "2026-09-16T00:00:00Z",
        EndDate: "2026-09-16T01:00:00Z", LineStyle: "dotted", LineColor: "green", TextColor: "brown",
        Author: {Title: "作成者A"}, Editor: {Title: "更新者B"},
        Created: "2026-09-15T01:00:00Z", Modified: "2026-09-16T02:00:00Z"
    });
    assert.strictEqual(item.lineStyle, "dotted");
    assert.strictEqual(item.lineColor, "green");
    assert.strictEqual(item.textColor, "brown");
    assert.strictEqual(item.createdBy, "作成者A");
    assert.strictEqual(item.modifiedBy, "更新者B");
    assert.ok(item.createdAt instanceof Date);
    assert.ok(item.modifiedAt instanceof Date);
    payload = source.toPayload(item, "Test");
    assert.strictEqual(payload.LineStyle, "dotted");
    assert.strictEqual(payload.LineColor, "green");
    assert.strictEqual(payload.TextColor, "brown");
    assert.strictEqual(payload.Author, undefined);
    assert.strictEqual(payload.Modified, undefined);
});

test("予定の文字色を線色とは独立して設定・表示・履歴保存できる", function () {
    var html = fs.readFileSync(path.join(root, "index.html"), "utf8");
    var app = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var css = fs.readFileSync(path.join(root, "css/style.css"), "utf8");
    var csv = fs.readFileSync(path.join(root, "js/csv-data-source.js"), "utf8");
    var history = new HistoryDataSource(options);
    var entries;
    assert.strictEqual(util.normalizeTextColor("RED"), "red");
    assert.strictEqual(util.normalizeTextColor("invalid"), "default");
    assert.ok(html.indexOf('id="text-color-buttons"') >= 0);
    assert.ok(html.indexOf('id="text-color"') >= 0);
    assert.ok(app.indexOf('setTextColorChoice(item.textColor)') >= 0);
    assert.ok(app.indexOf('textColor: util.normalizeTextColor(byId("text-color").value)') >= 0);
    assert.ok(app.indexOf('textColor: util.normalizeTextColor(item.textColor)') >= 0);
    assert.ok(csv.indexOf('textColor: util.normalizeTextColor(source.TextColor)') >= 0);
    assert.ok(/\.event-text-color-red\s*\{[^}]*color:\s*#b32929;/.test(css));
    assert.ok(/body\.dark-mode \.event-text-color-green\s*\{[^}]*color:\s*#77c88a;/.test(css));
    history.recordSample("create", null, {id: 5, title: "色付き", textColor: "brown"});
    entries = history.getSample();
    assert.strictEqual(entries[0].after.textColor, "brown");
});

test("線の列がないSharePointリストも標準線で読み込む", function () {
    var settings = JSON.parse(JSON.stringify(options));
    var source;
    var requests = 0;
    settings.fields.lineStyle = "LineStyle";
    settings.fields.lineColor = "LineColor";
    source = new SharePointDataSource(settings);
    source.request = function (method, url, headers, body, success, failure) {
        requests += 1;
        if (requests === 1) {
            assert.ok(decodeURIComponent(url).indexOf("LineStyle") >= 0);
            failure("列がありません", 400);
        } else {
            assert.strictEqual(decodeURIComponent(url).indexOf("LineStyle"), -1);
            success({responseText: JSON.stringify({d: {results: [{ID: 1, Title: "既存", EventDate: "2026-09-16T00:00:00Z"}]}})});
        }
    };
    source.load(null, function (items) {
        assert.strictEqual(items[0].lineStyle, "solid");
        assert.strictEqual(items[0].lineColor, "default");
    }, function (message) { throw new Error(message); });
    assert.strictEqual(requests, 2);
    assert.strictEqual(source.styleFieldsAvailable, false);
});

test("文字色列がないSharePointリストは標準色で読み込み色付き保存を案内する", function () {
    var settings = JSON.parse(JSON.stringify(options));
    var source;
    var requests = 0;
    var error = "";
    settings.fields.lineStyle = "LineStyle";
    settings.fields.lineColor = "LineColor";
    settings.fields.textColor = "TextColor";
    source = new SharePointDataSource(settings);
    source.request = function (method, url, headers, body, success, failure) {
        requests += 1;
        if (requests === 1) {
            assert.ok(decodeURIComponent(url).indexOf("TextColor") >= 0);
            failure("列がありません", 400);
        } else {
            assert.strictEqual(decodeURIComponent(url).indexOf("TextColor"), -1);
            assert.ok(decodeURIComponent(url).indexOf("LineColor") >= 0);
            success({responseText: JSON.stringify({d: {results: [{ID: 1, Title: "既存",
                EventDate: "2026-09-16T00:00:00Z", LineStyle: "dotted", LineColor: "green"}]}})});
        }
    };
    source.load(null, function (items) {
        assert.strictEqual(items[0].textColor, "default");
        assert.strictEqual(items[0].lineColor, "green");
    }, function (message) { throw new Error(message); });
    assert.strictEqual(requests, 2);
    assert.strictEqual(source.textColorFieldAvailable, false);
    assert.strictEqual(source.styleFieldsAvailable, true);
    source.create({title: "色付き", textColor: "red"}, function () {
        throw new Error("保存されました");
    }, function (message) { error = message; });
    assert.ok(error.indexOf("TextColor") >= 0);
});

test("線列がなく文字色列だけあるSharePointリストも読み込む", function () {
    var settings = JSON.parse(JSON.stringify(options));
    var source;
    var requests = 0;
    settings.fields.lineStyle = "LineStyle";
    settings.fields.lineColor = "LineColor";
    settings.fields.textColor = "TextColor";
    source = new SharePointDataSource(settings);
    source.request = function (method, url, headers, body, success, failure) {
        requests += 1;
        if (decodeURIComponent(url).indexOf("LineStyle") >= 0) {
            failure("線列がありません", 400);
        } else {
            assert.ok(decodeURIComponent(url).indexOf("TextColor") >= 0);
            success({responseText: JSON.stringify({d: {results: [{ID: 1, Title: "既存",
                EventDate: "2026-09-16T00:00:00Z", TextColor: "red"}]}})});
        }
    };
    source.load(null, function (items) {
        assert.strictEqual(items[0].lineColor, "default");
        assert.strictEqual(items[0].textColor, "red");
    }, function (message) { throw new Error(message); });
    assert.strictEqual(requests, 3);
    assert.strictEqual(source.styleFieldsAvailable, false);
    assert.strictEqual(source.textColorFieldAvailable, true);
});

test("CSV履歴にサンプルと画面内操作を表示する", function () {
    var service = new DataService({
        csv: {url: "data/schedule.csv"}, sharePoint: options, defaultDataSource: "csv"
    });
    var item = {title: "追加テスト", startDate: new Date(2026, 8, 24, 9, 0),
        endDate: new Date(2026, 8, 24, 10, 0), purpose: "GP1｜日々"};
    var history;
    var samples = service.history.getSample();
    var counts = {create: 0, update: 0, delete: 0};
    samples.forEach(function (entry) { counts[entry.action] += 1; });
    assert.strictEqual(samples.length, 30);
    assert.strictEqual(counts.create, 10);
    assert.strictEqual(counts.update, 10);
    assert.strictEqual(counts.delete, 10);
    assert.ok(samples.some(function (entry) { return entry.at.getMonth() === 8; }));
    assert.ok(samples.some(function (entry) { return entry.at.getMonth() === 9; }));
    service.history.load = function (success) { success([]); };
    service.setEditingEnabled(true);
    service.create(item, function () {}, function (message) { throw new Error(message); });
    service.loadHistory(function (entries) { history = entries; }, function (message) { throw new Error(message); });
    assert.strictEqual(history.length, 31);
    assert.strictEqual(history[0].action, "create");
    assert.strictEqual(history[0].after.title, "追加テスト");
    assert.strictEqual(history[1].action, "create");
    assert.strictEqual(history[2].action, "update");
    assert.strictEqual(history[3].action, "delete");
});

test("SharePoint履歴は変更前後を保存し古い履歴を300件に整理する", function () {
    var history = new HistoryDataSource(options);
    var requests = [];
    var before = {id: 5, title: "変更前", startDate: new Date(2026, 8, 24, 9, 0),
        endDate: new Date(2026, 8, 24, 10, 0), purpose: "GP1｜日々"};
    var after = {id: 5, title: "変更後", startDate: new Date(2026, 8, 24, 10, 0),
        endDate: new Date(2026, 8, 24, 11, 0), purpose: "GP2｜週間"};
    var completed = false;
    var ids = [];
    var i;
    for (i = 301; i >= 1; i -= 1) { ids.push({ID: i}); }
    history.api.getEntityType = function (success) { success("SP.Data.HistoryListItem"); };
    history.api.getDigest = function (success) { success("digest"); };
    history.api.request = function (method, url, headers, body, success) {
        requests.push({method: method, url: url, headers: headers, body: body});
        if (method === "GET") { success({responseText: JSON.stringify({d: {results: ids}})}); }
        else { success({responseText: ""}); }
    };
    history.record("update", before, after, function () { completed = true; }, function (message) { throw new Error(message); });
    assert.strictEqual(completed, true);
    assert.strictEqual(JSON.parse(requests[0].body).Action, "update");
    assert.strictEqual(JSON.parse(JSON.parse(requests[0].body).BeforeJson).title, "変更前");
    assert.strictEqual(JSON.parse(JSON.parse(requests[0].body).AfterJson).title, "変更後");
    assert.ok(requests[2].url.indexOf("items(1)") >= 0);
    assert.strictEqual(requests[2].headers["X-HTTP-Method"], "DELETE");
});

test("組織設定の履歴は同じSharePoint履歴リストに変更前後を記録する", function () {
    var history = new HistoryDataSource(organizationSettingsOptions);
    var requests = [];
    var before = {id: 12, groupName: "GP1科", teamName: "GS班", monthlyRows: 5,
        weeklyRows: 3, dailyRows: 2, autoRows: true, sortOrder: 0, isActive: true};
    var after = {id: 12, groupName: "GP1科", teamName: "GS班", monthlyRows: 4,
        weeklyRows: 3, dailyRows: 2, autoRows: false, sortOrder: 1, isActive: true};
    var saved = false;
    var payload;
    history.api.getEntityType = function (success) { success("SP.Data.HistoryListItem"); };
    history.api.getDigest = function (success) { success("digest"); };
    history.api.request = function (method, url, headers, body, success) {
        requests.push({method: method, body: body});
        success({responseText: method === "GET" ? JSON.stringify({d: {results: []}}) : ""});
    };
    history.recordSettings("update", before, after, function () { saved = true; }, function (message) {
        throw new Error(message);
    });
    assert.strictEqual(saved, true);
    payload = JSON.parse(requests[0].body);
    assert.strictEqual(payload.Action, "update");
    assert.strictEqual(payload.ScheduleList, "予定表組織設定");
    assert.strictEqual(payload.ScheduleItemId, "12");
    assert.strictEqual(payload.Title, "GP1科／GS班");
    assert.strictEqual(JSON.parse(payload.BeforeJson).monthlyRows, 5);
    assert.strictEqual(JSON.parse(payload.AfterJson).monthlyRows, 4);
    assert.strictEqual(JSON.parse(payload.AfterJson).autoRows, false);
    assert.strictEqual(JSON.parse(payload.AfterJson).etag, undefined);
    assert.strictEqual(history.toEntry({ID: 3, Action: "update", BeforeJson: payload.BeforeJson,
        AfterJson: payload.AfterJson, ScheduleList: payload.ScheduleList,
        Author: {Title: "操作者"}}).actor, "操作者");
    history.recordSampleSettings("update", before, after);
    assert.strictEqual(history.getSample()[0].scheduleList, "予定表組織設定");
});

test("CSV表示中もSharePointの組織設定履歴を表示し取得失敗時は画面内履歴へ戻る", function () {
    var service = new DataService({csv: {url: "data/schedule.csv"},
        sharePoint: organizationSettingsOptions, defaultDataSource: "csv"});
    var settingsEntry = {action: "update", after: {kind: "organizationSettings", title: "GP1"}};
    var entries;
    var warning;
    service.history.load = function (success) {
        success([settingsEntry, {action: "create", after: {title: "実際の予定"}}]);
    };
    service.loadHistory(function (items, message) { entries = items; warning = message; }, function () {});
    assert.strictEqual(entries.length, 31);
    assert.strictEqual(entries[0], settingsEntry);
    assert.strictEqual(warning, "");
    service.history.recordSettings = function (action, before, after, success) { success(); };
    service.recordSettingsChange("create", null, {id: 2, groupName: "GP2"}, function (message) {
        assert.strictEqual(message, "");
    });
    assert.strictEqual(service.history.getSample()[0].after.kind, "organizationSettings");
    service.history.load = function (success, failure) { failure("権限なし"); };
    service.loadHistory(function (items, message) { entries = items; warning = message; }, function () {});
    assert.strictEqual(entries[0].after.groupName, "GP2");
    assert.ok(warning.indexOf("権限なし") >= 0);
});

test("履歴詳細では組織設定の行数と有効状態を表示する", function () {
    var app = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var detail = {innerHTML: ""};
    var scope = vm.createContext({
        byId: function () { return detail; },
        util: {escapeHtml: function (value) { return String(value); }},
        config: {sharePoint: {listTitle: "予定表"}},
        historyActionLabel: function () { return "更新"; },
        historyDate: function () { return "2026-09-26 12:00"; }
    });
    vm.runInContext(app.slice(app.indexOf("    function renderHistoryDetail("),
        app.indexOf("    function renderHistory(")), scope);
    scope.renderHistoryDetail({action: "update", actor: "操作者", scheduleList: "予定表組織設定",
        before: {kind: "organizationSettings", id: "3", groupName: "GP1", monthlyRows: 5,
            weeklyRows: 3, dailyRows: 1, autoRows: true, sortOrder: 0, isActive: true},
        after: {kind: "organizationSettings", id: "3", groupName: "GP1", monthlyRows: 4,
            weeklyRows: 3, dailyRows: 1, autoRows: false, sortOrder: 1, isActive: false}});
    assert.ok(detail.innerHTML.indexOf("設定ID: 3") >= 0);
    assert.ok(detail.innerHTML.indexOf("月間行数") >= 0);
    assert.ok(detail.innerHTML.indexOf("<td>5</td><td>4</td>") >= 0);
    assert.ok(detail.innerHTML.indexOf("<td>はい</td><td>いいえ</td>") >= 0);
    assert.ok(detail.innerHTML.indexOf("<td>0</td><td>1</td>") >= 0);
});

test("組織設定の追加・更新・削除が成功した後だけ履歴を記録する", function () {
    var script = fs.readFileSync(path.join(root, "js/settings-controller.js"), "utf8");
    var records = [];
    var reloads = [];
    var item = {groupName: "GP1", teamName: "", monthlyRows: 5};
    var before = {id: 4, groupName: "GP1", teamName: "", monthlyRows: 3};
    var failSave = false;
    var source = {
        canConnect: function () { return true; },
        create: function (value, success) { success({id: 8, groupName: value.groupName}); },
        update: function (value, success, failure) {
            if (failSave) { failure("保存失敗"); }
            else { success(value); }
        },
        remove: function (value, success) { success(value); }
    };
    var scope = vm.createContext({
        SettingsController: function () {},
        window: {confirm: function () { return true; }},
        byId: function () { return {}; }
    });
    var controller;
    vm.runInContext(script.slice(script.indexOf("    SettingsController.prototype.save = function"),
        script.indexOf("    SettingsController.prototype.initialize = function")), scope);
    controller = new scope.SettingsController();
    controller.source = source;
    controller.canEdit = function () { return true; };
    controller.readForm = function () { return item; };
    controller.setBusy = function () {};
    controller.setStatus = function () {};
    controller.clearForm = function () {};
    controller.load = function (showMessage, warning) { reloads.push({showMessage: showMessage, warning: warning}); };
    controller.recordChange = function (action, previous, current, done) {
        records.push({action: action, before: previous, after: current});
        done(action === "delete" ? "履歴失敗" : "");
    };
    controller.save({preventDefault: function () {}});
    assert.strictEqual(records[0].action, "create");
    assert.strictEqual(records[0].before, null);
    assert.strictEqual(records[0].after.id, 8);
    item = {id: 4, groupName: "GP1", teamName: "", monthlyRows: 5};
    controller.editingItem = before;
    controller.save({preventDefault: function () {}});
    assert.strictEqual(records[1].action, "update");
    assert.strictEqual(records[1].before.monthlyRows, 3);
    assert.strictEqual(records[1].after.monthlyRows, 5);
    failSave = true;
    controller.save({preventDefault: function () {}});
    assert.strictEqual(records.length, 2);
    controller.remove();
    assert.strictEqual(records[2].action, "delete");
    assert.strictEqual(records[2].after, null);
    assert.strictEqual(reloads[2].warning, "履歴失敗");
});

test("予定変更後の履歴記録失敗は予定成功と警告を返す", function () {
    var service = new DataService({
        csv: {url: "data/schedule.csv"}, sharePoint: options, defaultDataSource: "sharepoint"
    });
    var result = "";
    service.setEditingEnabled(true);
    service.sources.sharepoint.create = function (item, success) { success(item); };
    service.history.record = function (action, before, after, success, failure) { failure("権限がありません"); };
    service.create({title: "記録テスト"}, function (item, items, warning) {
        result = warning;
    }, function (message) { throw new Error(message); });
    assert.ok(result.indexOf("予定の変更は完了") >= 0);
    assert.ok(result.indexOf("権限がありません") >= 0);
});

test("履歴のボタンと閲覧パネルを備える", function () {
    var html = fs.readFileSync(path.join(root, "index.html"), "utf8");
    var app = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    assert.ok(html.indexOf('id="open-history"') >= 0);
    assert.ok(html.indexOf('id="history-panel"') >= 0);
    assert.ok(html.indexOf('src="js/history-data-source.js"') >= 0);
    assert.ok(app.indexOf('util.addEvent(byId("open-history"), "click", openHistory)') >= 0);
});

test("予定表は起動時に読取専用で更新ボタンからのみ編集できる", function () {
    var service = new DataService({
        csv: {url: "data/schedule.csv"}, sharePoint: options, defaultDataSource: "csv"
    });
    var error = "";
    service.create({title: "不可"}, function () { throw new Error("登録されました"); }, function (message) { error = message; });
    assert.strictEqual(service.isReadOnly(), true);
    assert.ok(error.indexOf("予定表更新") >= 0);
    service.setEditingEnabled(true);
    assert.strictEqual(service.isReadOnly(), false);
    service.setEditingEnabled(false);
    assert.strictEqual(service.isReadOnly(), true);
});

test("共通アクセスカウンターは競合時に再取得して加算する", function () {
    var counter = new AccessCounter(options);
    var reads = 0;
    var writes = 0;
    var result = 0;
    var readUrls = [];
    counter.api.getEntityType = function (success) { success("SP.Data.CounterListItem"); };
    counter.api.getDigest = function (success) { success("digest"); };
    counter.api.request = function (method, url, headers, body, success, failure) {
        if (method === "GET") {
            assert.strictEqual(readUrls.indexOf(url), -1, "競合後はキャッシュ済みURLを再利用しない");
            readUrls.push(url);
            assert.strictEqual(headers["Cache-Control"], "no-cache");
            reads += 1;
            success({responseText: JSON.stringify({d: {results: [{
                ID: 1, VisitCount: reads === 1 ? 10 : 11,
                __metadata: {etag: reads === 1 ? '"1"' : '"2"'}
            }]}})});
        } else {
            writes += 1;
            assert.strictEqual(headers["IF-MATCH"], writes === 1 ? '"1"' : '"2"');
            assert.strictEqual(JSON.parse(body).VisitCount, writes === 1 ? 11 : 12);
            assert.strictEqual(JSON.parse(body).__metadata.type, "SP.Data.CounterListItem");
            if (writes === 1) { failure("競合", 412); }
            else { success({responseText: ""}); }
        }
    };
    counter.increment(function (value) { result = value; }, function (message) { throw new Error(message); });
    assert.strictEqual(result, 12);
    assert.strictEqual(reads, 2);
    assert.strictEqual(writes, 2);
});

test("アクセス数の更新権限エラーはHTTP状態と失敗箇所を保持する", function () {
    var counter = new AccessCounter(options), received;
    counter.api.getEntityType = function (success) { success("SP.Data.CounterListItem"); };
    counter.api.getDigest = function (success) { success("digest"); };
    counter.api.request = function (method, url, headers, body, success, failure) {
        if (method === "GET") {
            success({responseText: JSON.stringify({d: {results: [{ID: 1, VisitCount: 10, __metadata: {etag: '"1"'}}]}})});
        } else { failure("アクセスが拒否されました", 403); }
    };
    counter.increment(function () { throw new Error("成功扱いになりました"); }, function (message, status) {
        received = {message: message, status: status};
    });
    assert.strictEqual(received.status, 403);
    assert.ok(received.message.indexOf("アクセス数の更新") >= 0);
});

test("カウンターの初期項目不足と不正値は明示し既存値を上書きしない", function () {
    [[], [{ID: 1, VisitCount: "不正", __metadata: {etag: '"1"'}}],
        [{ID: 1, VisitCount: null, __metadata: {etag: '"1"'}}]].forEach(function (rows) {
        var counter = new AccessCounter(options), message = "";
        counter.api.request = function (method, url, headers, body, success) {
            assert.strictEqual(method, "GET");
            success({responseText: JSON.stringify({d: {results: rows}})});
        };
        counter.increment(function () { throw new Error("成功扱いになりました"); }, function (error) { message = error; });
        assert.ok(message.indexOf(rows.length ? "VisitCountが不正" : "初期項目がありません") >= 0);
    });
});

test("アクセス数の競合は5回で終了し無条件上書きしない", function () {
    var counter = new AccessCounter(options), writes = 0, status;
    counter.api.getEntityType = function (success) { success("SP.Data.CounterListItem"); };
    counter.api.getDigest = function (success) { success("digest"); };
    counter.api.request = function (method, url, headers, body, success, failure) {
        if (method === "GET") {
            success({responseText: JSON.stringify({d: {results: [{ID: 1, VisitCount: 10, __metadata: {etag: '"1"'}}]}})});
        } else {
            assert.strictEqual(headers["IF-MATCH"], '"1"');
            writes += 1;
            failure("競合", 412);
        }
    };
    counter.increment(function () { throw new Error("成功扱いになりました"); }, function (message, code) { status = code; });
    assert.strictEqual(writes, 5);
    assert.strictEqual(status, 412);
});

test("上部に更新切替と小さな共有カウンターを置き縦余白を半減する", function () {
    var html = fs.readFileSync(path.join(root, "index.html"), "utf8");
    var css = fs.readFileSync(path.join(root, "css/style.css"), "utf8");
    var app = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    assert.ok(html.indexOf('id="edit-mode-status"') >= 0);
    assert.ok(html.indexOf('id="toggle-schedule-edit"') >= 0);
    assert.ok(html.indexOf('id="access-counter"') >= 0);
    assert.ok(html.indexOf('>アクセス 確認中</span>') >= 0);
    assert.ok(app.indexOf('"アクセス 集計失敗"') >= 0);
    assert.ok(/\.app-header\s*\{[^}]*padding:\s*7\.5px 18px;/.test(css));
    assert.ok(app.indexOf("accessCounter.increment(function (count)") >= 0);
    assert.ok(app.indexOf("service.setEditingEnabled(false);") >= 0);
});

test("日々のスクロール時は上部操作欄と時刻行を固定する", function () {
    var app = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var css = fs.readFileSync(path.join(root, "css/style.css"), "utf8");
    assert.ok(app.indexOf("function initializeFixedHeader()") >= 0);
    assert.ok(app.indexOf("function refreshFixedTimeAxis()") >= 0);
    assert.ok(app.indexOf("function updateFixedHeader()") >= 0);
    assert.ok(app.indexOf('util.addEvent(window, "scroll", updateFixedHeader)') >= 0);
    assert.ok(app.indexOf("source.cloneNode(true)") >= 0);
    assert.ok(css.indexOf(".fixed-daily-axis") >= 0);
    assert.ok(css.indexOf("position: sticky") === -1);
});

test("予定表・操作欄・グループ見出しの前に余白を入れない", function () {
    var css = fs.readFileSync(path.join(root, "css/style.css"), "utf8");
    var app = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    assert.ok(/\.app-shell\s*\{[^}]*margin:\s*0 auto 40px;/.test(css));
    assert.ok(/\.toolbar\s*\{[^}]*margin-top:\s*0;/.test(css));
    assert.ok(/\.list-view\s*\{[^}]*margin-top:\s*0;/.test(css));
    assert.ok(/\.calendar-wrap\s*\{[^}]*margin-top:\s*0;/.test(css));
    assert.ok(app.indexOf("if (isNaN(fixedHeader.toolbarGap))") >= 0);
});

test("日々・週間・月間の左右余白を従来の5％にする", function () {
    var css = fs.readFileSync(path.join(root, "css/style.css"), "utf8");
    assert.ok(/\.app-shell\s*\{[^}]*width:\s*calc\(95% \+ 64px\);/.test(css));
    assert.ok(/\.app-shell\s*\{[^}]*max-width:\s*99\.8%;/.test(css));
    assert.ok(/body\.editor-open \.app-shell\s*\{[^}]*width:\s*73%;/.test(css));
});

test("日々の件名背景は文字部分だけにして空白の縦罫線を隠さない", function () {
    var css = fs.readFileSync(path.join(root, "css/style.css"), "utf8");
    var app = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var captionRule = css.match(/\.daily-event-caption\s*\{([^}]*)\}/);
    assert.ok(captionRule);
    assert.ok(!/background:/.test(captionRule[1]));
    assert.ok(/\.daily-event-caption-text\s*\{[^}]*background:\s*#fff;/.test(css));
    assert.ok(app.indexOf('captionText.className = "daily-event-caption-text"') >= 0);
    assert.ok(app.indexOf('findParentByClass(source, "daily-event-caption")') >= 0);
});

test("週間のグループ内横罫線と、設定未取得時の予定内の横線を表示しない", function () {
    var css = fs.readFileSync(path.join(root, "css/style.css"), "utf8");
    var app = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    assert.ok(app.indexOf('row.className += " weekly-group-row"') >= 0);
    assert.ok(app.indexOf('rowIndex === 0 ? " weekly-group-first"') >= 0);
    assert.ok(app.indexOf('rowIndex === rowCount - 1 ? " weekly-group-last"') >= 0);
    assert.ok(/\.weekly-schedule tbody tr\.weekly-group-row td\s*\{[^}]*border-top-width:\s*0;[^}]*border-bottom-width:\s*0;/.test(css));
    assert.ok(/\.weekly-schedule \.period-event \.daily-event-line\s*\{[^}]*display:\s*none;/.test(css));
});

test("週間予定表の見出しと予定を左寄せにする", function () {
    var css = fs.readFileSync(path.join(root, "css/style.css"), "utf8");
    assert.ok(/\.weekly-schedule thead th,\s*\.weekly-schedule tbody th,\s*\.weekly-schedule tbody td\s*\{[^}]*text-align:\s*left;/.test(css));
    assert.ok(/\.weekly-schedule\.organization-schedule \.event-item\.period-event\s*\{[^}]*text-align:\s*left;/.test(css));
});

test("月間は単日予定の線を離し日またぎ予定の線だけつなぐ", function () {
    var app = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var css = fs.readFileSync(path.join(root, "css/style.css"), "utf8");
    var scope = vm.createContext({
        compareItems: function (a, b) { return a.startDate - b.startDate; },
        estimateDailyCaptionPixels: function (item) { return item.id === 4 ? 210 : 80; }
    });
    var spanning = {id: 1, startDate: new Date(2026, 8, 1)};
    var firstDay = {id: 2, startDate: new Date(2026, 8, 1, 10)};
    var secondDay = {id: 3, startDate: new Date(2026, 8, 2)};
    var aligned;
    vm.runInContext(app.slice(app.indexOf("    function alignMonthlyItemsByLane("),
        app.indexOf("    function createHeaderCell(")), scope);
    aligned = scope.alignMonthlyItemsByLane([
        [spanning, firstDay], [spanning, secondDay], [spanning]
    ]);
    assert.strictEqual(aligned.requiredRows, 2);
    assert.strictEqual(aligned.itemsByDate[0][0], spanning);
    assert.strictEqual(aligned.itemsByDate[1][0], spanning);
    assert.strictEqual(aligned.itemsByDate[2][0], spanning);
    aligned = scope.alignMonthlyItemsByLane([
        [{id: 4, startDate: new Date(2026, 8, 1)}],
        [{id: 5, startDate: new Date(2026, 8, 2)}],
        []
    ], 110);
    assert.strictEqual(aligned.requiredRows, 2);
    assert.ok(!aligned.itemsByDate[1][0]);
    assert.strictEqual(aligned.itemsByDate[1][1].id, 5);
    var monthEnd = [];
    for (var day = 0; day < 30; day += 1) { monthEnd.push([]); }
    monthEnd[27].push({id: 4, startDate: new Date(2026, 8, 28)});
    monthEnd[29].push({id: 5, startDate: new Date(2026, 8, 30)});
    aligned = scope.alignMonthlyItemsByLane(monthEnd, 57);
    assert.strictEqual(aligned.requiredRows, 2);
    assert.ok(aligned.itemsByDate[27][0]);
    assert.ok(aligned.itemsByDate[29][1]);
    assert.ok(app.indexOf('viewMode !== "monthly" || j === 0 || !itemOccursOn(item, dates[j - 1])') >= 0);
    assert.ok(/\.monthly-schedule tbody tr\.monthly-group-row td\s*\{[^}]*border-top-width:\s*0;[^}]*border-bottom-width:\s*0;/.test(css));
    assert.ok(/\.monthly-schedule\s*\{[^}]*min-width:\s*1878px;/.test(css));
    assert.ok(/\.monthly-schedule \.period-event \.daily-event-line\s*\{[^}]*width:\s*calc\(100% - 8px\);[^}]*margin:\s*2px 4px 3px;/.test(css));
    assert.ok(/\.monthly-schedule \.period-continues-after \.daily-event-line\s*\{[^}]*width:\s*calc\(100% - 2px\);[^}]*margin-left:\s*4px;/.test(css));
    assert.ok(/\.monthly-schedule \.period-continues-before \.daily-event-line\s*\{[^}]*width:\s*calc\(100% - 2px\);[^}]*margin-left:\s*-2px;/.test(css));
    assert.ok(/\.monthly-schedule \.period-continues-before\.period-continues-after \.daily-event-line\s*\{[^}]*width:\s*calc\(100% \+ 4px\);/.test(css));
    assert.ok(/\.monthly-schedule \.period-event-caption\s*\{[^}]*white-space:\s*nowrap;/.test(css));
});

test("月間の件名は省略せず右端で左へ広げる", function () {
    var app = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var css = fs.readFileSync(path.join(root, "css/style.css"), "utf8");
    var row = {};
    function makeCaption(left, captionRow) {
        var caption = {style: {}, offsetWidth: 50, scrollWidth: 140,
            parentNode: {parentNode: {parentNode: captionRow}}};
        caption.getBoundingClientRect = function () {
            var actualLeft = left + (parseFloat(caption.style.left) || 0);
            var width = parseFloat(caption.style.width) || 50;
            return {left: actualLeft, right: actualLeft + width, width: width};
        };
        return caption;
    }
    var captions = [makeCaption(100, row), makeCaption(260, {})];
    var table = {
        querySelectorAll: function () { return captions; },
        getBoundingClientRect: function () { return {left: 0, right: 300}; },
        getElementsByTagName: function () { return [{}, {getBoundingClientRect: function () {
            return {left: 0};
        }}]; }
    };
    var scope = vm.createContext({});
    vm.runInContext(app.slice(app.indexOf("    function constrainMonthlyCaptions("),
        app.indexOf("    function renderOrganizationSchedule(")), scope);
    scope.constrainMonthlyCaptions(table);
    assert.strictEqual(captions[0].style.width, "142px");
    assert.strictEqual(captions[1].style.width, "142px");
    assert.strictEqual(captions[1].style.left, "-104px");
    assert.strictEqual(captions[1].style.textAlign, "right");
    assert.ok(/\.monthly-schedule \.period-event-caption\s*\{[^}]*overflow:\s*visible;[^}]*text-overflow:\s*clip;[^}]*white-space:\s*nowrap;/.test(css));
    assert.ok(app.indexOf('constrainMonthlyCaptions(head.parentNode)') >= 0);
    assert.ok(app.indexOf('constrainMonthlyCaptions(table)') >= 0);
});

test("月間の空白ドラッグで横移動し予定のドラッグは妨げない", function () {
    var app = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var view = {tagName: "DIV", className: "", scrollWidth: 1880,
        clientWidth: 1000, scrollLeft: 0};
    var blank = {tagName: "TD", className: "organization-schedule-cell", parentNode: view};
    var button = {tagName: "BUTTON", className: "event-item", parentNode: blank};
    var pan = {active: false, moved: false, suppressClick: false};
    var timer;
    var page = {top: 100, left: 0};
    var scope = vm.createContext({
        monthlyPan: pan,
        window: {
            pageYOffset: page.top,
            pageXOffset: page.left,
            scrollTo: function (left, top) {
                page.left = left;
                page.top = top;
                this.pageYOffset = top;
            },
            setTimeout: function (callback) { timer = callback; }
        },
        document: {documentElement: {scrollTop: 0}, body: {scrollTop: 0}},
        byId: function () { return view; },
        hasClass: function (element, name) { return element.className.indexOf(name) >= 0; },
        addClass: function (element, name) { element.className += " " + name; },
        removeClass: function (element, name) { element.className = element.className.replace(name, ""); },
        preventEvent: function (event) { event.prevented = true; return false; }
    });
    vm.runInContext(app.slice(app.indexOf("    function beginMonthlyPan("),
        app.indexOf("    function formatDailyTime(")), scope);
    scope.beginMonthlyPan({button: 0, clientX: 500, clientY: 200, target: button});
    assert.strictEqual(pan.active, false);
    scope.beginMonthlyPan({button: 0, clientX: 500, clientY: 200, target: blank});
    scope.moveMonthlyPan({clientX: 300, clientY: 200});
    assert.strictEqual(view.scrollLeft, 200);
    scope.endMonthlyPan();
    assert.strictEqual(pan.active, false);
    var click = {};
    scope.suppressMonthlyPanClick(click);
    assert.strictEqual(click.prevented, true);
    timer();
    assert.strictEqual(pan.suppressClick, false);
    view.scrollWidth = view.clientWidth;
    scope.beginMonthlyPan({button: 0, clientX: 500, clientY: 300, target: blank});
    scope.moveMonthlyPan({clientX: 500, clientY: 180});
    assert.strictEqual(page.top, 220);
    assert.strictEqual(view.scrollLeft, 200);
    scope.endMonthlyPan();
    view.scrollWidth = 1880;
    scope.beginMonthlyPan({button: 0, clientX: 500, clientY: 300, target: blank});
    scope.moveMonthlyPan({clientX: 450, clientY: 350});
    assert.strictEqual(view.scrollLeft, 250);
    assert.strictEqual(page.top, 170);
    scope.endMonthlyPan();
});

test("印刷プレビューの用紙・向き・縦横倍率・グループを表示モード別に保存する", function () {
    var app = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var html = fs.readFileSync(path.join(root, "index.html"), "utf8");
    var css = fs.readFileSync(path.join(root, "css/style.css"), "utf8");
    var values = {
        print_monthly_paper: "A4",
        print_monthly_orientation: "portrait",
        print_monthly_scaleX: "80",
        print_monthly_scaleY: "120",
        print_monthly_groups: '["GP1\\u001f"]'
    };
    var scope = vm.createContext({readDisplayPreference: function (name) { return values[name] || ""; }});
    var preferences;
    vm.runInContext(app.slice(app.indexOf("    function readPrintPreferences("),
        app.indexOf("    function createPrintSource(")), scope);
    preferences = scope.readPrintPreferences("monthly");
    assert.strictEqual(preferences.paper, "A4");
    assert.strictEqual(preferences.orientation, "portrait");
    assert.strictEqual(preferences.scaleX, 80);
    assert.strictEqual(preferences.scaleY, 120);
    assert.strictEqual(preferences.excluded[0], "GP1\u001f");
    assert.strictEqual(scope.readPrintPreferences("weekly").paper, "A4");
    assert.strictEqual(scope.readPrintPreferences("daily").orientation, "landscape");
    ["print-preview", "print-paper", "print-orientation", "print-scale-x", "print-scale-y",
        "print-groups", "execute-print"].forEach(function (id) {
        assert.ok(html.indexOf('id="' + id + '"') >= 0, id);
    });
    assert.ok(app.indexOf('deletePreferenceCookie("print_" + modes[i] + "_" + printNames[j])') >= 0);
    assert.ok(app.indexOf('row.setAttribute("data-print-block", "1")') >= 0);
    assert.ok(css.indexOf("#print-area,") >= 0);
    assert.ok(css.indexOf(".print-preview-controls { display: none !important; }") >= 0);
    assert.ok(app.indexOf("applyPrintDimensions(source, widthRatio, heightRatio)") >= 0);
    assert.ok(app.indexOf("var fontRatio = Math.min(widthRatio, heightRatio)") >= 0);
    assert.ok(app.indexOf("element.style.fontSize = (metrics[i].fontSize * fontRatio)") >= 0);
    assert.ok(app.indexOf("element.style.height = Math.max(1, metrics[i].height * heightRatio)") >= 0);
    assert.ok(app.indexOf('table.style.width = "100%"') >= 0);
    assert.ok(app.indexOf("content.style.transform =") < 0);
    assert.ok(app.indexOf("var leftMarginMm = Math.max(20, marginMm)") >= 0);
    assert.ok(app.indexOf('content.style.left = leftMarginMm + "mm"') >= 0);
    assert.ok(app.indexOf("availableWidth = (paperWidth - leftMarginMm - marginMm) * pixelsPerMm") >= 0);
});

test("週間の土日見出しを色分けし予定のない土日だけ半幅にする", function () {
    var css = fs.readFileSync(path.join(root, "css/style.css"), "utf8");
    var html = fs.readFileSync(path.join(root, "index.html"), "utf8");
    var app = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var columns = Array.from({length: 8}, function () { return {style: {}}; });
    var elements = {
        "weekly-columns": {getElementsByTagName: function () { return columns; }},
        "weekly-table": {offsetWidth: 1400, getBoundingClientRect: function () { return {width: this.offsetWidth}; }},
        "month-title": {}, "print-heading": {}, "weekly-head": {}, "weekly-body": {}
    };
    var weekendItems = {};
    var scope = vm.createContext({
        state: {displayDate: new Date(2026, 8, 24)},
        weeklyColumnWeights: [],
        weeklyDisplaySettings: {showTime: true, showMultiDayLine: false},
        currentDisplayZoom: 100,
        startOfWeek: function () { return new Date(2026, 8, 21); },
        byId: function (id) { return elements[id]; },
        getItemsForDay: function (date, viewMode) {
            assert.strictEqual(viewMode, "weekly");
            return weekendItems[date.getDay()] || [];
        },
        renderOrganizationSchedule: function () {}
    });
    function columnWidth(column) {
        assert.ok(/^[\d.]+px$/.test(column.style.width));
        return parseFloat(column.style.width);
    }
    vm.runInContext(app.slice(app.indexOf("    function updateWeeklyColumnWidths()"),
        app.indexOf("    function findItemById(")), scope);
    scope.renderWeekly();
    assert.ok(html.indexOf('id="weekly-columns"') >= 0);
    assert.strictEqual(columns[0].style.width, "155px");
    assert.ok(Math.abs(columnWidth(columns[6]) * 2 - columnWidth(columns[1])) < 0.001);
    assert.ok(Math.abs(columnWidth(columns[7]) * 2 - columnWidth(columns[1])) < 0.001);
    weekendItems[6] = [{}];
    scope.renderWeekly();
    assert.ok(Math.abs(columnWidth(columns[6]) - columnWidth(columns[1])) < 0.001);
    assert.ok(Math.abs(columnWidth(columns[7]) * 2 - columnWidth(columns[1])) < 0.001);
    elements["weekly-table"].offsetWidth = 1000;
    scope.updateWeeklyColumnWidths();
    assert.ok(Math.abs(columnWidth(columns[6]) - columnWidth(columns[1])) < 0.001);
    assert.ok(Math.abs(columnWidth(columns[7]) * 2 - columnWidth(columns[1])) < 0.001);
    assert.ok(/\.weekly-schedule thead th\.saturday\s*,[^{}]+\{[^}]*color:\s*#1f5794;/.test(css));
    assert.ok(/\.weekly-schedule thead th\.sunday\s*,[^{}]+\{[^}]*color:\s*#b32929;/.test(css));
    assert.ok(/body\.dark-mode \.weekly-schedule thead th\.saturday\s*,[^{}]+\{[^}]*color:\s*#8dbdff;/.test(css));
    assert.ok(/body\.dark-mode \.weekly-schedule thead th\.sunday\s*,[^{}]+\{[^}]*color:\s*#ff9696;/.test(css));
});

test("SharePoint読込を表示期間で絞り込む", function () {
    var source = new SharePointDataSource(options);
    var requestedUrl = "";
    var loaded = false;
    source.request = function (method, url, headers, body, success) {
        requestedUrl = decodeURIComponent(url);
        success({responseText: JSON.stringify({d: {results: []}})});
    };
    source.load({
        startDate: new Date(2026, 8, 1, 0, 0),
        endDate: new Date(2026, 9, 1, 0, 0)
    }, function () {
        loaded = true;
    }, function (message) {
        throw new Error(message);
    });
    assert.strictEqual(loaded, true);
    assert.ok(requestedUrl.indexOf("$filter=EventDate lt datetime'") >= 0);
    assert.ok(requestedUrl.indexOf("EndDate ge datetime'") >= 0);
});

test("更新時に読込時のETagを送る", function () {
    var source = new SharePointDataSource(options);
    var sentHeaders;
    var succeeded = false;
    var item = {
        id: 10,
        etag: "\"3\"",
        title: "会議",
        startDate: new Date(2026, 8, 16, 9, 0),
        endDate: new Date(2026, 8, 16, 10, 0),
        category: "",
        location: "",
        description: "",
        purpose: "GP1｜日々"
    };
    source.getEntityType = function (success) {
        success("SP.Data.EventListItem");
    };
    source.getDigest = function (success) {
        success("digest");
    };
    source.request = function (method, url, headers, body, success) {
        sentHeaders = headers;
        success({responseText: ""});
    };
    source.update(item, function () {
        succeeded = true;
    }, function (message) {
        throw new Error(message);
    });
    assert.strictEqual(succeeded, true);
    assert.strictEqual(sentHeaders["IF-MATCH"], "\"3\"");
    assert.strictEqual(sentHeaders["X-HTTP-Method"], "MERGE");
});

test("ETagがない予定の更新を拒否する", function () {
    var source = new SharePointDataSource(options);
    var failureMessage = "";
    source.update({id: 10}, function () {
        throw new Error("更新が実行されました");
    }, function (message) {
        failureMessage = message;
    });
    assert.ok(failureMessage.indexOf("再読込") >= 0);
});

test("日々表示が組織縦・時刻横の時間枠を使用する", function () {
    var appSource = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var html = fs.readFileSync(path.join(root, "index.html"), "utf8");
    assert.ok(appSource.indexOf("dailyViewConfig.slotMinutes") >= 0);
    assert.ok(appSource.indexOf("dailyViewConfig.startHour") >= 0);
    assert.ok(appSource.indexOf("dailyViewConfig.endHour") >= 0);
    assert.ok(appSource.indexOf("dailyViewConfig.standardStartHour") >= 0);
    assert.ok(appSource.indexOf("dailyViewConfig.standardEndHour") >= 0);
    assert.ok(appSource.indexOf('createHeaderCell("グループ"') >= 0);
    assert.ok(appSource.indexOf('getOrganizationBlocks("daily"') >= 0);
    assert.ok(appSource.indexOf("function createDailyEventBar") >= 0);
    assert.ok(appSource.indexOf('button.className += " daily-event-short"') >= 0);
    assert.ok(appSource.indexOf("minimumVisualPixels = estimateDailyCaptionPixels") >= 0);
    assert.ok(appSource.indexOf("displayRange: {start: visualStart, end: visualEnd}") >= 0);
    assert.ok(appSource.indexOf('line.className = "daily-event-line"') >= 0);
    assert.ok(appSource.indexOf('caption.className = "daily-event-caption"') >= 0);
    assert.ok(html.indexOf('id="daily-head"') >= 0);
    assert.ok(html.indexOf("daily-horizontal-schedule") >= 0);
});

test("日々表示は通常7時から18時で必要な時間帯だけを広げる", function () {
    var appSource = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var day = new Date(2026, 8, 23);
    var scope = vm.createContext({
        Math: Math,
        sameDate: function (left, right) {
            return left.getFullYear() === right.getFullYear() &&
                left.getMonth() === right.getMonth() && left.getDate() === right.getDate();
        }
    });
    var start = appSource.indexOf("    function isDailyRangeTrigger(");
    var end = appSource.indexOf("    function getDailyItemRange(");
    var range;

    vm.runInContext(appSource.slice(start, end), scope);
    range = scope.getDailyDisplayRange([], day, 360, 1320, 420, 1080, 60);
    assert.deepStrictEqual({start: range.start, end: range.end}, {start: 420, end: 1080});

    range = scope.getDailyDisplayRange([
        {startDate: new Date(2026, 8, 23, 6, 30), endDate: new Date(2026, 8, 23, 7, 15)},
        {startDate: new Date(2026, 8, 23, 17, 30), endDate: new Date(2026, 8, 23, 19, 15)}
    ], day, 360, 1320, 420, 1080, 60);
    assert.deepStrictEqual({start: range.start, end: range.end}, {start: 360, end: 1200});

    range = scope.getDailyDisplayRange([
        {startDate: new Date(2026, 8, 22, 6, 0), endDate: new Date(2026, 8, 24, 20, 0)},
        {allDay: true, startDate: new Date(2026, 8, 23, 0, 0), endDate: new Date(2026, 8, 23, 23, 59)}
    ], day, 360, 1320, 420, 1080, 60);
    assert.deepStrictEqual({start: range.start, end: range.end}, {start: 420, end: 1080});
});

test("大グループと小グループを1行へ最小化できる", function () {
    var appSource = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var css = fs.readFileSync(path.join(root, "css/style.css"), "utf8");
    var scope = vm.createContext({});
    var start = appSource.indexOf("    function getOrganizationSectionsFromBlocks(");
    var end = appSource.indexOf("    function isOrganizationSectionCollapsed(");
    var sections;

    vm.runInContext(appSource.slice(start, end), scope);
    sections = scope.getOrganizationSectionsFromBlocks([
        {section: "GP1", team: ""},
        {section: "GP1科", team: "GS班"},
        {section: "GP1科", team: "JN班"}
    ]);
    assert.strictEqual(sections.length, 2);
    assert.strictEqual(sections[0].hasTeams, false);
    assert.strictEqual(sections[1].hasTeams, true);
    assert.strictEqual(sections[1].blocks.length, 2);
    assert.ok(appSource.indexOf("organizationCollapseState") >= 0);
    assert.ok(appSource.indexOf("appendOrganizationSectionRow(body, section") >= 0);
    assert.ok(appSource.indexOf("rowCount = collapsed ? 1 : getRenderedRowCount(block, requiredRows)") >= 0);
    assert.ok(css.indexOf(".organization-collapse-button") >= 0);
    assert.ok(css.indexOf(".organization-collapsed-cell") >= 0);
});

test("行数の自動調整は不要な空き行を表示しない", function () {
    var appSource = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var html = fs.readFileSync(path.join(root, "index.html"), "utf8");
    var settingsSource = fs.readFileSync(path.join(root, "js/settings-controller.js"), "utf8");
    var scope = vm.createContext({Math: Math, currentFontSize: 14});
    var start = appSource.indexOf("    function getRenderedRowCount(");
    var end = appSource.indexOf("    function getOrganizationKey(");

    vm.runInContext(appSource.slice(start, end), scope);
    assert.strictEqual(scope.getRenderedRowCount({rowCount: 5, autoRows: true}, 1), 1);
    assert.strictEqual(scope.getRenderedRowCount({rowCount: 5, autoRows: true}, 0), 1);
    assert.strictEqual(scope.getRenderedRowCount({rowCount: 5, autoRows: false}, 1), 5);
    assert.strictEqual(scope.getRenderedRowCount({rowCount: 5, autoRows: true}, 6), 6);
    assert.ok(html.indexOf('id="setting-auto-rows" checked') >= 0);
    assert.ok(settingsSource.indexOf('autoRows: byId("setting-auto-rows").checked') >= 0);
});

test("短時間予定の横線は文字枠を広げても実時刻に一致する", function () {
    var source = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var scope = vm.createContext({
        document: {
            createElement: function () {
                return {style: {}, children: [], appendChild: function (child) { this.children.push(child); }};
            },
            createTextNode: function (value) { return value; }
        },
        util: util,
        formatDailyTime: function (value) { return String(value); },
        currentFontSize: 14,
        dailyInteraction: {selectedItemId: "", clipboard: null}
    });
    vm.runInContext(source.slice(source.indexOf("    function getDailyCaptionText("),
        source.indexOf("    function renderDaily(")), scope);
    assert.ok(scope.estimateDailyCaptionPixels({title: "非常に長い会議件名", location: "第一会議室"}) > 120);
    [
        {actual: {start: 420, end: 430}, display: {start: 384, end: 466}},
        {actual: {start: 360, end: 370}, display: {start: 360, end: 442}},
        {actual: {start: 1310, end: 1320}, display: {start: 1238, end: 1320}},
        {actual: {start: 540, end: 600}, display: {start: 529, end: 611}}
    ].forEach(function (range) {
        var button = scope.createDailyEventBar({title: "test", startDate: new Date(2026, 8, 22)},
            range.actual, range.display, 360, 1320, 0);
        var line = button.children[1];
        var buttonLeft = parseFloat(button.style.left);
        var buttonWidth = parseFloat(button.style.width);
        var start = buttonLeft + buttonWidth * parseFloat(line.style.left) / 100;
        var end = start + buttonWidth * parseFloat(line.style.width) / 100;
        assert.ok(Math.abs(start - (range.actual.start - 360) / 960 * 100) < 0.00001);
        assert.ok(Math.abs(end - (range.actual.end - 360) / 960 * 100) < 0.00001);
    });
    var colored = scope.createDailyEventBar({
        title: "点線", startDate: new Date(2026, 8, 22), lineStyle: "dotted", lineColor: "brown",
        textColor: "green"
    }, {start: 540, end: 600}, {start: 529, end: 611}, 360, 1320, 0);
    assert.ok(colored.children[1].className.indexOf("line-dotted") >= 0);
    assert.ok(colored.children[1].className.indexOf("line-color-brown") >= 0);
    assert.ok(colored.children[0].children[0].className.indexOf("event-text-color-green") >= 0);
    assert.ok(colored.children[0].children[1].className.indexOf("event-text-color-green") >= 0);
    assert.ok(colored.children[2].children[0].className.indexOf("event-text-color-green") >= 0);
});

test("日々予定は文字の高さを維持して上下余白を縮める", function () {
    var appSource = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var css = fs.readFileSync(path.join(root, "css/style.css"), "utf8");
    var scope = vm.createContext({Math: Math, currentFontSize: 14});
    var start = appSource.indexOf("    function getDailyLaneMetrics(");
    var end = appSource.indexOf("    function createDailyEventBar(");
    var metrics;

    vm.runInContext(appSource.slice(start, end), scope);
    metrics = scope.getDailyLaneMetrics();
    assert.strictEqual(metrics.contentHeight, 37);
    assert.strictEqual(metrics.bodyVerticalPadding, 1);
    assert.strictEqual(metrics.eventHeight, 39);
    assert.strictEqual(metrics.laneHeight, 41);
    assert.ok(appSource.indexOf("groupBorderMargin + laneIndex * laneMetrics.laneHeight") >= 0);
    assert.ok(appSource.indexOf("getDailyGroupLayout(rowCount)") >= 0);
    assert.ok(/\.daily-event-bar\s*\{[\s\S]*?height:\s*39px;[\s\S]*?padding:\s*1px 0;/.test(css));
});

test("日々予定の段数が増えてもグループ外側の余白を増やさない", function () {
    var appSource = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var scope = vm.createContext({Math: Math, currentFontSize: 14});
    var start = appSource.indexOf("    function getDailyLaneMetrics(");
    var end = appSource.indexOf("    function createDailyEventBar(");

    vm.runInContext(appSource.slice(start, end), scope);
    assert.strictEqual(scope.getDailyGroupLayout(1).scheduledAreaHeight, 39);
    assert.strictEqual(scope.getDailyGroupLayout(1).groupBorderMargin, 2);
    assert.strictEqual(scope.getDailyGroupLayout(1).timelineHeight, 43);
    assert.strictEqual(scope.getDailyGroupLayout(4).scheduledAreaHeight, 162);
    assert.strictEqual(scope.getDailyGroupLayout(4).groupBorderMargin, 2);
    assert.strictEqual(scope.getDailyGroupLayout(4).timelineHeight, 166);
});

test("分離した日時入力は直接時刻・終日・不正値を扱える", function () {
    var nodes = {"start-date": {value: "2026-09-22"}, "start-time": {value: "0815"},
        "end-date": {value: "2026-09-23"}, "end-time": {value: "1700"}};
    var scope = vm.createContext({window: {YoteihyouUtil: util}, document: {
        getElementById: function (id) { return nodes[id]; }
    }});
    vm.runInContext(fs.readFileSync(path.join(root, "js/date-time-editor.js"), "utf8"), scope);
    var editor = scope.window.YoteihyouDateTimeEditor;
    assert.strictEqual(util.formatDateTime(editor.read("start", false)), "2026-09-22 08:15");
    nodes["start-time"].value = "07:10";
    assert.strictEqual(util.formatDateTime(editor.read("start", false)), "2026-09-22 07:10");
    nodes["start-time"].value = "2460";
    assert.strictEqual(editor.read("start", false), null);
    assert.strictEqual(util.formatDateTime(editor.read("start", true)), "2026-09-22 00:00");
    assert.strictEqual(util.formatDateTime(editor.read("end", true)), "2026-09-23 23:59");
    nodes["start-date"].value = "2026-02-30";
    assert.strictEqual(editor.read("start", true), null);
});

test("SharePointの終日フラグを保存して読み戻せる", function () {
    var settings = JSON.parse(JSON.stringify(options));
    settings.fields.allDay = "AllDay";
    var source = new SharePointDataSource(settings);
    var payload = source.toPayload({title: "終日予定", startDate: new Date(2026, 8, 22), allDay: true}, "Test");
    assert.strictEqual(payload.AllDay, true);
    assert.strictEqual(source.toItem(payload).allDay, true);
    payload.AllDay = false;
    assert.strictEqual(source.toItem(payload).allDay, false);
});

test("日々予定の移動と前後時刻の変更を15分単位で反映できる", function () {
    var appSource = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var scope = vm.createContext({
        Date: Date,
        Math: Math,
        DAILY_SNAP_MINUTES: 15,
        dailyViewConfig: {defaultStartHour: 9},
        util: util,
        service: {getMode: function () { return "csv"; }},
        splitPurpose: function () { return {targets: ["daily", "weekly"]}; },
        joinPurpose: function (section, team, targets) { return section + "/" + team + "/" + targets.join(","); },
        startOfDay: function (date) { return new Date(date.getFullYear(), date.getMonth(), date.getDate()); }
    });
    var item = {
        id: 1,
        etag: "etag",
        title: "会議",
        startDate: new Date(2026, 8, 23, 9, 0),
        endDate: new Date(2026, 8, 23, 10, 0),
        purpose: "GP1",
        isActive: true
    };
    var target = {day: new Date(2026, 8, 24), minute: 10 * 60 + 30, section: "GP2", team: "A班"};
    var moved;
    var resized;
    vm.runInContext(appSource.slice(appSource.indexOf("    function cloneScheduleItem("),
        appSource.indexOf("    function createEventButton(")), scope);
    moved = scope.moveItemToTarget(item, target);
    assert.strictEqual(moved.startDate.getTime(), new Date(2026, 8, 24, 10, 30).getTime());
    assert.strictEqual(moved.endDate.getTime(), new Date(2026, 8, 24, 11, 30).getTime());
    assert.strictEqual(moved.purpose, "GP2/A班/daily,weekly");
    resized = scope.getResizedDailyItem(item, "start", {day: new Date(2026, 8, 23), minute: 8 * 60 + 15});
    assert.strictEqual(resized.startDate.getTime(), new Date(2026, 8, 23, 8, 15).getTime());
    assert.throws(function () {
        scope.getResizedDailyItem(item, "end", {day: new Date(2026, 8, 23), minute: 9 * 60});
    }, /15分以上後/);
    var periodTarget = scope.getPeriodTarget({
        _periodMeta: {day: new Date(2026, 8, 25), section: "GP3", team: "B班"}
    }, item);
    assert.strictEqual(periodTarget.minute, 9 * 60);
    assert.strictEqual(periodTarget.section, "GP3");
    assert.strictEqual(scope.moveItemToTarget(item, periodTarget).startDate.getTime(),
        new Date(2026, 8, 25, 9, 0).getTime());
    assert.strictEqual(scope.getResizedPeriodItem(item, "end", periodTarget).endDate.getTime(),
        new Date(2026, 8, 25, 10, 0).getTime());
});

test("週間・月間の予定も線種・色とドラッグ・貼り付け操作を使う", function () {
    var app = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var css = fs.readFileSync(path.join(root, "css/style.css"), "utf8");
    assert.ok(app.indexOf('button.className = "event-item period-event"') >= 0);
    assert.ok(app.indexOf('beginPeriodDrag(event || window.event, item, button, "move")') >= 0);
    assert.ok(app.indexOf('getPeriodTarget(cell, item || (dailyInteraction.clipboard && dailyInteraction.clipboard.item))') >= 0);
    assert.ok(app.indexOf('scheduleCell._periodMeta = {') >= 0);
    assert.ok(app.indexOf('util.addEvent(byId("monthly-view"), "scroll", updateFixedHeader)') >= 0);
    assert.ok(css.indexOf(".period-event .daily-event-line") >= 0);
    assert.ok(css.indexOf(".organization-schedule-cell.period-drop-target") >= 0);
});

test("週間・月間の予定に開始終了時刻と横線の設定を保持する", function () {
    var app = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var scope = vm.createContext({
        util: util,
        weeklyDisplaySettings: {showTime: true, showMultiDayLine: false},
        dailyInteraction: {selectedItemId: "", clipboard: null},
        sameDate: function (left, right) {
            return left.getFullYear() === right.getFullYear() &&
                left.getMonth() === right.getMonth() && left.getDate() === right.getDate();
        },
        formatDailyTime: function (minute) {
            return ("0" + Math.floor(minute / 60)).slice(-2) + ("0" + minute % 60).slice(-2);
        },
        document: {
            createElement: function () {
                return {children: [], appendChild: function (child) { this.children.push(child); }};
            },
            createTextNode: function (value) { return value; }
        }
    });
    var day = new Date(2026, 8, 24);
    var item = {id: 4, title: "会議", startDate: new Date(2026, 8, 24, 9, 30),
        endDate: new Date(2026, 8, 24, 10, 45), lineStyle: "dotted", lineColor: "red",
        textColor: "green", location: "第１会議室"};
    var button;
    vm.runInContext(app.slice(app.indexOf("    function createEventButton("),
        app.indexOf("    function getOrganizationGroups(")), scope);
    button = scope.createEventButton(item, day);
    assert.strictEqual(button.children[0].children[0], "0930–1045");
    assert.ok(button.children[1].className.indexOf("line-dotted") >= 0);
    assert.ok(button.children[1].className.indexOf("line-color-red") >= 0);
    assert.ok(button.children[0].className.indexOf("event-text-color-green") >= 0);
    assert.ok(button.children[2].className.indexOf("event-text-color-green") >= 0);
    assert.ok(button.children[2].children[0].indexOf("第１会議室") >= 0);
    assert.strictEqual(button.children.length, 3);
    item.endDate = new Date(2026, 8, 26, 10, 45);
    button = scope.createEventButton(item, new Date(2026, 8, 25), "monthly", false);
    assert.ok(button.className.indexOf("period-continues-before") >= 0);
    assert.ok(button.className.indexOf("period-continues-after") >= 0);
    assert.ok(button.children[0].className.indexOf("daily-event-line") >= 0);
    assert.strictEqual(button.children[1].children[0], "\u00a0");
    item.endDate = new Date(2026, 8, 24, 10, 45);
    item.allDay = true;
    button = scope.createEventButton(item, day);
    assert.strictEqual(button.children[0].className.indexOf("daily-event-line") >= 0, true);
    assert.strictEqual(button.children.length, 2);
    [true, false].forEach(function (showTime) { [true, false].forEach(function (showLine) {
        scope.weeklyDisplaySettings = {showTime: showTime, showMultiDayLine: showLine};
        [true, false].forEach(function (allDay) { [true, false].forEach(function (multiDay) {
            item.allDay = allDay;
            item.endDate = new Date(2026, 8, multiDay ? 26 : 24, 10, 45);
            button = scope.createEventButton(item, day, "weekly");
            assert.strictEqual(button.children.some(function (child) {
                return child.className.indexOf("period-event-times") >= 0;
            }), showTime && !allDay);
            assert.strictEqual(button.className.indexOf("period-multiday") >= 0, showLine && multiDay);
        }); });
        item.allDay = false;
        item.endDate = new Date(2026, 8, 26, 10, 45);
        button = scope.createEventButton(item, new Date(2026, 8, 25), "weekly");
        assert.strictEqual(button.className.indexOf("period-continues-before") >= 0, showLine);
        assert.strictEqual(button.className.indexOf("period-continues-after") >= 0, showLine);
        button = scope.createEventButton(item, day, "monthly");
        assert.strictEqual(button.children.length, 2);
        button = scope.createEventButton(item, day, "daily");
        assert.strictEqual(button.children.length, 3);
    }); });
});

test("日々予定にドラッグ操作とコピー操作のUIがある", function () {
    var appSource = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var css = fs.readFileSync(path.join(root, "css/style.css"), "utf8");
    var html = fs.readFileSync(path.join(root, "index.html"), "utf8");
    assert.ok(appSource.indexOf('beginDailyDrag(event || window.event, item, button, "move")') >= 0);
    assert.strictEqual(appSource.indexOf('beginDailyDrag(event || window.event, item, button, "start")'), -1);
    assert.strictEqual(appSource.indexOf('beginDailyDrag(event || window.event, item, button, "end")'), -1);
    assert.ok(appSource.indexOf("event.keyCode === 67") >= 0);
    assert.ok(appSource.indexOf("event.keyCode === 88") >= 0);
    assert.ok(appSource.indexOf("event.keyCode === 86") >= 0);
    assert.ok(html.indexOf('id="daily-context-menu"') >= 0);
    assert.ok(/\.daily-resize-handle\s*\{[\s\S]*?border:\s*0;[\s\S]*?background:\s*transparent;/.test(css));
    assert.ok(/\.daily-event-bar\.daily-event-selected \.daily-event-caption\s*\{[\s\S]*?outline:\s*0;/.test(css));
});

test("試験用CSVは平日に全グループを収録しGP1に複数日予定を含める", function () {
    var generator = require(path.join(root, "scripts/generate-sample-schedule.js"));
    var rows = generator.buildRows();
    var csvText = fs.readFileSync(path.join(root, "data/schedule.csv"), "utf8");
    var dailyCoverage = {};
    var groupCounts = {};
    var targetCounts = {daily: 0, weekly: 0, monthly: 0};
    var current = new Date(Date.UTC(2026, 8, 1));
    var last = new Date(Date.UTC(2026, 9, 31));
    var dateText;
    var group;
    var i;

    rows.forEach(function (row) {
        group = row.Purpose.split("｜")[0];
        groupCounts[group] = (groupCounts[group] || 0) + 1;
        dailyCoverage[row.EventDate.substring(0, 10) + "\u001f" + group] = true;
        assert.ok(row.Purpose.indexOf("日々") >= 0);
        targetCounts.daily += 1;
        if (row.Purpose.indexOf("週間") >= 0) { targetCounts.weekly += 1; }
        if (row.Purpose.indexOf("月間") >= 0) {
            targetCounts.monthly += 1;
            assert.ok(row.Purpose.indexOf("週間") >= 0);
        }
    });

    while (current.getTime() <= last.getTime()) {
        if (current.getUTCDay() === 0 || current.getUTCDay() === 6) {
            current = new Date(current.getTime() + 24 * 60 * 60 * 1000);
            continue;
        }
        dateText = current.getUTCFullYear() + "-" +
            (current.getUTCMonth() + 1 < 10 ? "0" : "") + (current.getUTCMonth() + 1) + "-" +
            (current.getUTCDate() < 10 ? "0" : "") + current.getUTCDate();
        for (i = 0; i < generator.groups.length; i += 1) {
            assert.strictEqual(dailyCoverage[dateText + "\u001f" + generator.groups[i]], true,
                dateText + " " + generator.groups[i]);
        }
        current = new Date(current.getTime() + 24 * 60 * 60 * 1000);
    }

    assert.strictEqual(rows.length, 1587);
    assert.deepStrictEqual(targetCounts, {daily: 1587, weekly: 794, monthly: 317});
    assert.strictEqual(csvText, generator.toCsv(rows));
    assert.strictEqual(groupCounts.GP1, 47);
    generator.busyGroups.forEach(function (busyGroup) {
        assert.strictEqual(groupCounts[busyGroup], 176);
    });
    assert.ok(rows.every(function (row) {
        var date = new Date(row.EventDate.substring(0, 10) + "T00:00:00Z");
        return date.getUTCDay() !== 0 && date.getUTCDay() !== 6;
    }));
    assert.ok(rows.some(function (row) {
        return row.Title === "GP1 日またぎ予定" && row.EventDate === "2026-09-08 22:00" &&
            row.EndDate === "2026-09-09 02:00";
    }));
    assert.ok(rows.some(function (row) {
        return row.Title === "GP1 月またぎ計画" && row.EventDate < "2026-10-01" && row.EndDate >= "2026-10-01";
    }));
    assert.ok(rows.some(function (row) {
        return row.Title === "GP1 週またぎ対応" && row.EventDate === "2026-10-09 13:00" && row.EndDate === "2026-10-13 12:00";
    }));
});

test("文字設定を検証して適用し、旧Cookieの文字設定を使用しない", function () {
    var source = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var fields = {"setting-font-family": {}, "setting-font-size": {},
        "monthly-view": {style: {}}, "weekly-view": {style: {}}, "daily-view": {style: {}}};
    var saved = {fontFamily: "mincho", fontSize: "20"};
    var scope = vm.createContext({
        document: {body: {style: {}}, documentElement: {style: {}}},
        currentFontSize: 14,
        byId: function (id) { return fields[id]; },
        storeDisplayPreference: function () { throw new Error("文字設定をCookieへ保存しました"); },
        readDisplayPreference: function (name) { return saved[name] || ""; },
        state: {},
        setDarkMode: function () {},
        setDisplayZoom: function () {},
        loadOrganizationCollapseState: function () {}
    });
    vm.runInContext(source.slice(source.indexOf("    function setFontPreferences("),
        source.indexOf("    function removeClass(")), scope);
    scope.setFontPreferences("mincho", "20", true);
    assert.strictEqual(scope.currentFontSize, 20);
    assert.strictEqual(scope.document.documentElement.style.fontSize, undefined);
    assert.strictEqual(scope.document.body.style.fontFamily, undefined);
    ["monthly-view", "weekly-view", "daily-view"].forEach(function (id) {
        assert.strictEqual(fields[id].style.fontSize, "20px");
        assert.ok(fields[id].style.fontFamily.indexOf("MS Mincho") >= 0);
    });
    scope.setFontPreferences("default", 14, false);
    scope.initializeDisplayPreferences();
    assert.strictEqual(fields["setting-font-family"].value, "default");
    assert.strictEqual(fields["setting-font-size"].value, "14");
    scope.setFontPreferences("__proto__", "999", false);
    assert.strictEqual(scope.currentFontSize, 14);
    assert.strictEqual(fields["setting-font-family"].value, "default");
});

test("固定見出しは拡大率と小数の列幅を保持し、スクロールバーによる幅変更に追従する", function () {
    var source = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    [70, 90, 100, 110, 150].forEach(function (zoom) {
        var scale = zoom / 100, width = 1000.375;
        var widths = [155.25, 281.125, 282.375, 280.625];
        var copies = widths.map(function () { return {style: {}}; });
        var cells = widths.map(function (value, index) { return {getBoundingClientRect: function () {
            var left = widths.slice(0, index).reduce(function (sum, w) { return sum + w; }, 0);
            return {left: left * scale, top: 0, height: 30 * scale, width: value * scale};
        }}; });
        var head = {getBoundingClientRect: function () { return {top: 0, height: 30 * scale}; }, getElementsByTagName: function () { return cells; }};
        var table = {getBoundingClientRect: function () { return {left: 0, width: width * scale}; },
            getElementsByTagName: function () { return [head]; }};
        var view = {style: {fontSize: "20px", fontFamily: "Meiryo"}, getElementsByTagName: function () { return [table]; }};
        var copy = {style: {}, getElementsByTagName: function () { return copies; }};
        var overlay = {style: {}, firstChild: copy};
        var scope = vm.createContext({fixedHeader: {axisOverlay: overlay}, currentDisplayZoom: zoom,
            getFixedCoordinateScale: function () { return scale; },
            state: {viewMode: "weekly"}, byId: function () { return view; }});
        vm.runInContext(source.slice(source.indexOf("    function syncFixedTimeAxis("),
            source.indexOf("    function updateFixedHeader(")), scope);
        scope.syncFixedTimeAxis();
        assert.ok(Math.abs(parseFloat(copy.style.width) - width) < 0.00001);
        copies.forEach(function (cell, index) {
            assert.ok(Math.abs(parseFloat(cell.style.width) - widths[index]) < 0.00001);
        });
        assert.strictEqual(overlay.style.fontSize, "20px");
        assert.strictEqual(overlay.style.fontFamily, "Meiryo");
        width -= 17;
        scope.syncFixedTimeAxis();
        assert.ok(Math.abs(parseFloat(copy.style.width) - width) < 0.00001);
    });
});

test("文字を大きくすると日々予定の行高も広がる", function () {
    var source = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var scope = vm.createContext({currentFontSize: 14});
    vm.runInContext(source.slice(source.indexOf("    function getDailyLaneMetrics("),
        source.indexOf("    function createDailyEventBar(")), scope);
    var normal = scope.getDailyGroupLayout(3).timelineHeight;
    scope.currentFontSize = 20;
    assert.ok(scope.getDailyGroupLayout(3).timelineHeight > normal);
    assert.ok(scope.getDailyLaneMetrics().laneHeight > scope.getDailyLaneMetrics().eventHeight);
});

test("共有文字設定をリストから取得しETag付きで保存する", function () {
    var browser = loadBrowserScripts(["js/util.js", "js/sharepoint-data-source.js", "js/display-settings-data-source.js"]);
    var source = new browser.window.YoteihyouDisplaySettingsDataSource({siteUrl: "https://example.invalid/sites/test"});
    var rows = [{ID: 8, FontFamily: "mincho", FontSize: 18, __metadata: {etag: '"3"'}}];
    var loaded;
    var error = "";
    var writes = 0;
    function failure(message) { error = message; }
    source.client.getEntityType = function (success) { success("SP.Data.DisplaySettingsListItem"); };
    source.client.getDigest = function (success) { success("digest"); };
    source.client.request = function (method, url, headers, body, success) {
        if (method === "GET") {
            assert.ok(decodeURIComponent(url).indexOf("Title eq 'default'") >= 0);
            success({responseText: JSON.stringify({d: {results: rows}})});
        } else {
            writes += 1;
            assert.strictEqual(headers["IF-MATCH"], '"3"');
            assert.strictEqual(headers["X-HTTP-Method"], "MERGE");
            assert.strictEqual(JSON.parse(body).FontFamily, "meiryo");
            assert.strictEqual(JSON.parse(body).FontSize, 20);
            success({});
        }
    };
    source.load(function (item) { loaded = item; }, failure);
    assert.strictEqual(loaded.family, "mincho");
    assert.strictEqual(loaded.size, 18);
    source.save(loaded, "meiryo", "20", function () {}, failure);
    assert.strictEqual(writes, 1);
    source.save(loaded, "invalid", 999, function () {}, failure);
    assert.ok(error);
    assert.strictEqual(writes, 1);
    loaded.etag = "";
    source.save(loaded, "meiryo", 20, function () {}, failure);
    assert.strictEqual(writes, 1);
    rows = [];
    error = "";
    source.load(function () { throw new Error("未準備リストを読込成功扱いにしました"); }, failure);
    assert.ok(error);
    source.client.request = function (method, url, headers, body, success, fail) { fail("競合"); };
    loaded.etag = '"3"';
    source.save(loaded, "meiryo", 20, function () { throw new Error("競合を保存成功扱いにしました"); }, failure);
    assert.strictEqual(error, "競合");
});

test("共有ログ記録設定を読み込み、対象列だけをETag付きで保存する", function () {
    var browser = loadBrowserScripts(["js/util.js", "js/sharepoint-data-source.js", "js/display-settings-data-source.js"]);
    var source = new browser.window.YoteihyouDisplaySettingsDataSource({siteUrl: "https://example.invalid/sites/test"});
    var enabled = false, version = 3, loaded, payload;
    function fail(message) { throw new Error(message); }
    source.client.getEntityType = function (done) { done("Settings"); };
    source.client.getDigest = function (done) { done("digest"); };
    source.client.request = function (method, url, headers, body, done) {
        if (method === "GET") {
            assert.ok(url.indexOf("$select=ID,SystemLogEnabled") >= 0);
            assert.ok(decodeURIComponent(url).indexOf("Title eq 'default'") >= 0);
            done({responseText: JSON.stringify({d: {results: [{ID: 8, SystemLogEnabled: enabled, __metadata: {etag: String(version)}}]}})});
        } else {
            assert.strictEqual(headers["IF-MATCH"], String(version));
            assert.strictEqual(headers["X-HTTP-Method"], "MERGE");
            assert.strictEqual(headers["X-RequestDigest"], "digest");
            payload = JSON.parse(body);
            assert.deepStrictEqual(Object.keys(payload).sort(), ["SystemLogEnabled", "__metadata"]);
            enabled = payload.SystemLogEnabled;
            version += 1;
            done({});
        }
    };
    source.loadLogging(function (item) { loaded = item; }, fail);
    assert.strictEqual(loaded.enabled, false);
    source.saveLogging(loaded, true, function () {}, fail);
    source.loadLogging(function (item) { loaded = item; }, fail);
    assert.strictEqual(loaded.enabled, true);
    assert.strictEqual(loaded.etag, "4");
    source.saveLogging(loaded, false, function () {}, fail);
    assert.strictEqual(enabled, false);
});

test("共有ログ設定の列不足・重複・未設定・権限不足・競合を成功扱いにしない", function () {
    var browser = loadBrowserScripts(["js/util.js", "js/sharepoint-data-source.js", "js/display-settings-data-source.js"]);
    var source = new browser.window.YoteihyouDisplaySettingsDataSource({siteUrl: "https://example.invalid/sites/test"});
    var failures = 0;
    function success() { throw new Error("不正な設定を成功扱いにしました"); }
    function fail() { failures += 1; }
    [[], [{ID: 8}], [{ID: 8, SystemLogEnabled: "false", __metadata: {etag: "1"}}],
        [{ID: 8, SystemLogEnabled: true}], [{ID: 8}, {ID: 9}]].forEach(function (rows) {
        source.client.request = function (method, url, headers, body, done) { done({responseText: JSON.stringify({d: {results: rows}})}); };
        source.loadLogging(success, fail);
    });
    source.client.getEntityType = function (done) { done("Settings"); };
    source.client.getDigest = function (done) { done("digest"); };
    [403, 412].forEach(function (status) {
        source.client.request = function (method, url, headers, body, done, failure) { failure("失敗", status); };
        source.saveLogging({id: 8, etag: "1"}, false, success, fail);
    });
    source.saveLogging({id: 8}, true, success, fail);
    source.client.siteUrl = "";
    source.loadLogging(success, fail);
    source.saveLogging({id: 8, etag: "1"}, true, success, fail);
    assert.strictEqual(failures, 10);
});

test("共有ログ記録設定は保存成功後だけ反映し、失敗・読取専用・二重操作では切り替えない", function () {
    var app = fs.readFileSync(path.join(root, "js/app.js"), "utf8"), buttons = {}, saveDone, saveFailed, loadDone, loadFailed;
    var writes = 0, readOnly = false, log = {enabled: false, setEnabled: function (value) { this.enabled = value; }};
    ["toggle-system-log", "reload-system-log-setting", "system-log-setting-status"].forEach(function (id) { buttons[id] = {}; });
    var scope = vm.createContext({sharedLogSettings: null, logSettingsBusy: false,
        window: {yoteihyouSystemLog: log}, byId: function (id) { return buttons[id]; },
        service: {isReadOnly: function () { return readOnly; }},
        displaySettingsSource: {
            loadLogging: function (done, fail) { loadDone = done; loadFailed = fail; },
            saveLogging: function (item, enabled, done, fail) { writes += 1; saveDone = done; saveFailed = fail; }
        }});
    vm.runInContext(app.slice(app.indexOf("    function updateLogSettingsControls("), app.indexOf("    function initializeSystemLog(")), scope);
    scope.loadLogSettings();
    assert.strictEqual(buttons["toggle-system-log"].disabled, true);
    loadDone({id: 8, etag: "1", enabled: false});
    readOnly = true;
    scope.toggleLogSettings();
    assert.strictEqual(writes, 0);
    readOnly = false;
    scope.toggleLogSettings();
    scope.toggleLogSettings();
    assert.strictEqual(writes, 1);
    assert.strictEqual(log.enabled, false);
    saveFailed("競合");
    assert.strictEqual(log.enabled, false);
    assert.strictEqual(buttons["toggle-system-log"].disabled, true);
    scope.loadLogSettings();
    loadDone({id: 8, etag: "2", enabled: false});
    scope.toggleLogSettings();
    saveDone();
    assert.strictEqual(log.enabled, true);
    loadDone({id: 8, etag: "3", enabled: true});
    assert.strictEqual(scope.sharedLogSettings.etag, "3");
    scope.loadLogSettings();
    loadFailed("接続失敗");
    assert.strictEqual(log.enabled, true);
    assert.strictEqual(buttons["toggle-system-log"].disabled, true);
});

var layoutContext = loadBrowserScripts(["js/util.js", "js/sharepoint-data-source.js", "js/schedule-layout.js"]);
var layout = layoutContext.window.YoteihyouScheduleLayout;

test("手動配置は時間順に戻らず、未指定予定を重ならない空きへ配置する", function () {
    var result = layout.pack([{id: "a", start: 0, end: 4}, {id: "b", start: 2, end: 5},
        {id: "c", start: 6, end: 8}, {id: "new", start: 1, end: 3}], [["b", "c"], ["a"]]);
    assert.strictEqual(result.lanes.b, 0);
    assert.strictEqual(result.lanes.c, 0);
    assert.strictEqual(result.lanes.a, 1);
    assert.strictEqual(result.lanes.new, 2);
});

test("時刻変更で同じ段が衝突したら後続の指定段も下げて重なりを防ぐ", function () {
    var result = layout.pack([{id: "a", start: 0, end: 5}, {id: "b", start: 2, end: 4},
        {id: "c", start: 8, end: 9}], [["a", "b"], ["c"]]);
    assert.strictEqual(result.lanes.a, 0);
    assert.strictEqual(result.lanes.b, 1);
    assert.strictEqual(result.lanes.c, 2);
    assert.strictEqual(result.count, 3);
});

test("削除済みIDと重複IDを無視し、空の指定段は詰める", function () {
    var result = layout.pack([{id: "a", start: 1, end: 2}], [["deleted"], ["a", "a"], ["a"]]);
    assert.strictEqual(result.count, 1);
    assert.strictEqual(result.lanes.a, 0);
});

test("午前の上下交換で午後の予定の段を変えず、余分な段も増やさない", function () {
    var entries = [{id: "a", start: 9, end: 10}, {id: "b", start: 9, end: 10},
        {id: "x", start: 15, end: 16}, {id: "y", start: 15, end: 16}];
    var original = {a: 0, b: 1, x: 0, y: 1};
    var moved = layout.move(entries, original, "b", 0);
    assert.deepStrictEqual(JSON.parse(JSON.stringify(moved.positions)), {a: 1, b: 0, x: 0, y: 1});
    assert.strictEqual(layout.pack(entries, moved).count, 2);
    assert.strictEqual(JSON.stringify(original), '{"a":0,"b":1,"x":0,"y":1}');
    assert.deepStrictEqual(JSON.parse(JSON.stringify(layout.pack(entries, moved).lanes)), JSON.parse(JSON.stringify(moved.positions)));
});

test("移動先が空いていればその段に同居し、ほかの段を詰め直さない", function () {
    var entries = [{id: "a", start: 9, end: 10}, {id: "b", start: 15, end: 16}, {id: "c", start: 9, end: 10}];
    var moved = layout.move(entries, {a: 0, b: 1, c: 2}, "b", 0);
    assert.strictEqual(JSON.stringify(moved.positions), '{"a":0,"b":0,"c":2}');
    assert.strictEqual(layout.pack(entries, moved).lanes.c, 2);
});

test("長い帯が元の段に戻らない場合も無関係な予定を動かさず空きへ置く", function () {
    var entries = [{id: "short", start: 3, end: 4}, {id: "wide", start: 2, end: 7},
        {id: "other", start: 5, end: 6}, {id: "far", start: 20, end: 21}];
    var moved = layout.move(entries, {short: 1, wide: 0, other: 1, far: 0}, "short", 0);
    assert.strictEqual(moved.positions.short, 0);
    assert.strictEqual(moved.positions.wide, 2);
    assert.strictEqual(moved.positions.other, 1);
    assert.strictEqual(moved.positions.far, 0);
    assert.strictEqual(layout.pack(entries, moved).count, 3);
});

test("v2は新規予定・日時変更・削除があっても関係のない指定段を維持する", function () {
    var entries = [{id: "a", start: 0, end: 8}, {id: "b", start: 2, end: 4},
        {id: "c", start: 9, end: 10}, {id: "new", start: 3, end: 5}];
    var packed = layout.pack(entries, {version: 2, positions: {a: 0, b: 0, c: 2, deleted: 1}});
    assert.strictEqual(packed.lanes.c, 2);
    assert.notStrictEqual(packed.lanes.a, packed.lanes.b);
    assert.notStrictEqual(packed.lanes.a, packed.lanes.new);
    assert.notStrictEqual(packed.lanes.b, packed.lanes.new);
    assert.strictEqual(packed.lanes.deleted, undefined);
});

test("月間の手動配置は複数日の帯を同じ段に保つ", function () {
    var source = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var a = {id: "a"}, b = {id: "b"};
    var scope = vm.createContext({layoutEngine: layout, compareItems: function () { return 0; },
        estimateDailyCaptionPixels: function () { return 30; }});
    vm.runInContext(source.slice(source.indexOf("    function alignMonthlyItemsByLane("),
        source.indexOf("    function ", source.indexOf("    function alignMonthlyItemsByLane(") + 10)), scope);
    var result = scope.alignMonthlyItemsByLane([[a], [a, b], [a]], 100, {version: 2, positions: {b: 0, a: 1}});
    assert.strictEqual(result.itemsByDate[0][1], a);
    assert.strictEqual(result.itemsByDate[1][1], a);
    assert.strictEqual(result.itemsByDate[2][1], a);
    assert.strictEqual(result.itemsByDate[1][0], b);
});

test("配置設定はETagを読み戻し競合時は成功扱いにしない", function () {
    var source = new layout.Source(options), loaded, request, error;
    source.client.request = function (method, url, headers, body, success) {
        success({responseText: JSON.stringify({d: {results: [{ID: 9, LayoutJson: '{"scope":[["a"]]}',
            __metadata: {etag: '"4"'}}]}})});
    };
    source.load("sharepoint:daily:2026-09-29", function (record) { loaded = record; }, function (message) { throw new Error(message); });
    assert.strictEqual(loaded.etag, '"4"');
    source.client.getEntityType = function (done) { done("SP.Data.LayoutListItem"); };
    source.client.getDigest = function (done) { done("digest"); };
    source.client.request = function (method, url, headers, body, success, failure) {
        request = {headers: headers, body: JSON.parse(body)};
        failure("競合", 412);
    };
    source.save(loaded, {scope: [["b"], ["a"]]}, function () { throw new Error("競合を無視"); }, function (message) { error = message; });
    assert.strictEqual(request.headers["IF-MATCH"], '"4"');
    assert.strictEqual(request.headers["X-HTTP-Method"], "MERGE");
    assert.strictEqual(request.body.LayoutJson, '{"scope":[["b"],["a"]]}');
    assert.strictEqual(error, "競合");
    assert.strictEqual(JSON.stringify(loaded.orders), '{"scope":[["a"]]}');
});

test("初回配置保存と解除は単一項目への書き込みで行う", function () {
    var source = new layout.Source(options), request, saved = false;
    source.client.getEntityType = function (done) { done("Layout"); };
    source.client.getDigest = function (done) { done("digest"); };
    source.client.request = function (method, url, headers, body, success) {
        request = {url: url, headers: headers, body: JSON.parse(body)}; success();
    };
    source.save({key: "daily:2026-09-29", orders: {}}, {}, function () { saved = true; }, function (message) { throw new Error(message); });
    assert.strictEqual(saved, true);
    assert.ok(/\/items$/.test(request.url));
    assert.strictEqual(request.headers["IF-MATCH"], undefined);
    assert.strictEqual(request.body.LayoutJson, "{}");
});

test("破損した配置・重複項目・ETagなしの保存を拒否する", function () {
    var source = new layout.Source(options), failures = 0;
    function fail() { failures += 1; }
    function unexpected() { throw new Error("不正データを受理"); }
    ["not json", "[]", '{"scope":1}', '{"scope":[[{}]]}'].forEach(function (json) {
        source.client.request = function (method, url, headers, body, success) {
            success({responseText: JSON.stringify({d: {results: [{ID: 1, LayoutJson: json, __metadata: {etag: '"1"'}}]}})});
        };
        source.load("key", unexpected, fail);
    });
    source.client.request = function (method, url, headers, body, success) {
        success({responseText: '{"d":{"results":[{},{}]}}'});
    };
    source.load("key", unexpected, fail);
    source.save({id: 1, key: "key"}, {}, unexpected, fail);
    assert.strictEqual(failures, 6);
});

test("表示・日付・グループの配置を分離し、無効なドロップと読取専用では保存しない", function () {
    var source = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var readonly = false, saved = 0;
    var scope = vm.createContext({
        util: util, layoutEngine: layout, layoutKey: "daily:2026-09-29", layoutRecord: {orders: {}},
        layoutDrag: null, layoutEditingEnabled: true, layoutSuppressClickUntil: 0, dailyInteraction: {}, selectedLayoutScope: "",
        splitPurpose: function (value) { return {section: value, team: ""}; },
        removeClass: function () {}, service: {isReadOnly: function () { return readonly; }},
        state: {viewMode: "daily"}, document: {querySelectorAll: function () { throw new Error("無効な操作でDOMを変更"); }}
    });
    vm.runInContext(source.slice(source.indexOf("    function getLayoutScope("), source.indexOf("    function compareItems(")), scope);
    scope.saveLayout = function () { saved += 1; };
    var day = new Date(2026, 8, 29), next = new Date(2026, 8, 30), item = {id: "1", purpose: "GP1"};
    assert.notStrictEqual(scope.getLayoutScope(item, day, "daily"), scope.getLayoutScope(item, next, "daily"));
    assert.notStrictEqual(scope.getLayoutScope(item, day, "daily"), scope.getLayoutScope({purpose: "GP2"}, day, "daily"));
    assert.strictEqual(scope.getLayoutScope(item, day, "monthly"), scope.getLayoutScope(item, next, "monthly"));
    scope.layoutDrag = {item: item, moved: true, target: null, key: scope.layoutKey};
    scope.finishLayoutDrag();
    assert.strictEqual(scope.layoutDrag, null);
    readonly = true;
    scope.layoutDrag = {item: item, moved: true, target: {}, key: scope.layoutKey};
    scope.finishLayoutDrag();
    assert.strictEqual(scope.layoutDrag, null);
    assert.strictEqual(saved, 0);
});

test("別期間へ移動した後に届いた共有配置の応答を表示しない", function () {
    var source = fs.readFileSync(path.join(root, "js/app.js"), "utf8"), callbacks = [], renders = 0;
    var date = new Date(2026, 8, 29);
    var scope = vm.createContext({util: util, layoutKey: "", layoutReady: false, layoutRecord: null,
        selectedLayoutScope: "", layoutRequestId: 0, layoutDrag: null, state: {viewMode: "daily"},
        service: {getMode: function () { return "sharepoint"; }},
        getCurrentLoadRange: function () { return {startDate: date}; },
        renderCurrentView: function () { renders += 1; },
        layoutSource: {load: function (key, success) { callbacks.push(success); }}
    });
    vm.runInContext(source.slice(source.indexOf("    function getLayoutScope("), source.indexOf("    function compareItems(")), scope);
    scope.ensureLayout();
    date = new Date(2026, 8, 30);
    scope.ensureLayout();
    callbacks[0]({orders: {old: []}});
    assert.strictEqual(scope.layoutRecord, null);
    assert.strictEqual(renders, 0);
    callbacks[1]({orders: {current: []}});
    assert.strictEqual(scope.layoutReady, true);
    assert.strictEqual(renders, 1);
});

test("挿入線は離れた予定の横位置によらずマウス付近に出し、拡大と横スクロールに追従する", function () {
    var source = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    ["daily", "weekly", "monthly"].forEach(function (mode) {
        [0.7, 1, 1.5].forEach(function (scale) {
            var host = {offsetWidth: 900, clientWidth: 896, clientLeft: 2, clientTop: 2,
                scrollLeft: 420, scrollTop: 30, children: [],
                getBoundingClientRect: function () { return {left: 20, top: 100, width: 900 * scale}; },
                appendChild: function (child) { this.children.push(child); child.parentNode = this; },
                removeChild: function (child) { this.children.splice(this.children.indexOf(child), 1); }};
            var scope = vm.createContext({state: {viewMode: mode}, layoutDrag: {},
                byId: function (id) { assert.strictEqual(id, mode + "-view"); return host; },
                addClass: function () {}, removeClass: function () {},
                document: {createElement: function () { return {style: {}, setAttribute: function () {}}; }}});
            vm.runInContext(source.slice(source.indexOf("    function clearLayoutDropIndicator("),
                source.indexOf("    function beginLayoutDrag(")), scope);
            var row = {left: 1000, right: 1200, top: 210, bottom: 250};
            scope.showLayoutDropIndicator(260, row, false);
            var line = host.children[0];
            var actualLeft = 20 + (parseFloat(line.style.left) - host.scrollLeft + host.clientLeft) * scale;
            var actualTop = 100 + (parseFloat(line.style.top) - host.scrollTop + host.clientTop) * scale;
            assert.ok(Math.abs(actualLeft + parseFloat(line.style.width) * scale / 2 - 260) < 0.01);
            assert.ok(Math.abs(actualTop - row.top) < 0.01);
            scope.clearLayoutDropIndicator();
            assert.strictEqual(host.children.length, 0);
            scope.showLayoutDropIndicator(260, row, true);
            line = host.children[0];
            actualTop = 100 + (parseFloat(line.style.top) - host.scrollTop + host.clientTop) * scale;
            assert.ok(Math.abs(actualTop - row.bottom) < 0.01);
            scope.cancelLayoutDrag();
            assert.strictEqual(host.children.length, 0);
            assert.strictEqual(scope.layoutDrag, null);
        });
    });
});

test("空白や段の隙間へのドラッグでも近い段を選び、表示した境界と保存先を一致させる", function () {
    var source = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var body = {}, day = new Date(2026, 8, 29), key = '["2026-09-29","GP1",""]';
    var cell = {parentNode: body, getAttribute: function () { return null; },
        _dailyMeta: {day: day, section: "GP1", team: ""}};
    function button(id, top) {
        return {getAttribute: function (name) { return name === "data-layout-id" ? id : name === "data-layout-row" ? (id === "target" ? "0" : "1") : key; },
            getBoundingClientRect: function () { return {left: 900, right: 1000, top: top, bottom: top + 39, height: 39}; }};
    }
    var target = button("target", 200), moved = button("moved", 241), indicator, saved;
    var scope = vm.createContext({util: util, layoutEngine: layout, state: {viewMode: "daily"},
        getLayoutDateScope: function (date) { return util.formatDateKey(date); },
        layoutDrag: {item: {id: "moved"}, scope: key, startX: 300, startY: 255, moved: false,
            entries: [{id: "target", start: 0, end: 1}, {id: "moved", start: 0, end: 1}], positions: {target: 0, moved: 1}},
        findParentByClass: function (element, className) { return element === cell && className === "daily-timeline-cell" ? cell : null; },
        document: {body: body, elementFromPoint: function () { return cell; }, querySelectorAll: function () { return [target, moved]; }},
        preventEvent: function () {}});
    vm.runInContext(source.slice(source.indexOf("    function clearLayoutDropIndicator("), source.indexOf("    function finishLayoutDrag(")), scope);
    scope.showLayoutDropIndicator = function (x, row, after) { indicator = {x: x, y: after ? row.bottom : row.top}; };
    scope.showLayoutPreview = function () {};
    scope.moveLayoutDrag({clientX: 300, clientY: 239.5});
    assert.strictEqual(scope.layoutDrag.target, target);
    assert.strictEqual(scope.layoutDrag.after, true);
    assert.deepStrictEqual(indicator, {x: 300, y: 239});
    saved = scope.layoutDrag.proposed;
    assert.strictEqual(JSON.stringify(saved.positions), '{"target":0,"moved":1}');
});

test("複数の予定と衝突しても直接衝突しない予定は固定し、再配置後に重なりを残さない", function () {
    var entries = [{id: "move", start: 2, end: 8}, {id: "a", start: 1, end: 4},
        {id: "b", start: 5, end: 9}, {id: "x", start: 0, end: 2}, {id: "y", start: 9, end: 10}];
    var result = layout.move(entries, {move: 1, a: 0, b: 0, x: 1, y: 0}, "move", 0);
    assert.strictEqual(result.positions.x, 1);
    assert.strictEqual(result.positions.y, 0);
    assert.strictEqual(result.positions.move, 0);
    entries.forEach(function (a, i) {
        entries.slice(i + 1).forEach(function (b) {
            assert.ok(result.positions[a.id] !== result.positions[b.id] || a.end <= b.start || b.end <= a.start);
        });
    });
    assert.deepStrictEqual(JSON.parse(JSON.stringify(layout.pack(entries, result).lanes)), JSON.parse(JSON.stringify(result.positions)));
});

test("共有設定は旧形式を読み、新形式を保存・再読込でき、不正な段番号は拒否する", function () {
    var source = new layout.Source(options), loaded, payload, failures = 0;
    function failed(message) { throw new Error(message); }
    source.client.request = function (method, url, headers, body, done) {
        done({responseText: JSON.stringify({d: {results: [{ID: 1, LayoutJson: '{"scope":[["a","x"],["b"]]}', __metadata: {etag: '"1"'}}]}})});
    };
    source.load("key", function (record) { loaded = record; }, failed);
    var entries = [{id: "a", start: 0, end: 1}, {id: "b", start: 0, end: 1}, {id: "x", start: 3, end: 4}];
    var old = layout.pack(entries, loaded.orders.scope);
    var updated = {scope: layout.move(entries, old.lanes, "b", 0)};
    source.client.getEntityType = function (done) { done("Layout"); };
    source.client.getDigest = function (done) { done("digest"); };
    source.client.request = function (method, url, headers, body, done) { payload = JSON.parse(body); done(); };
    source.save(loaded, updated, function () {}, failed);
    assert.strictEqual(JSON.parse(payload.LayoutJson).scope.version, 2);
    source.client.request = function (method, url, headers, body, done) {
        done({responseText: JSON.stringify({d: {results: [{ID: 1, LayoutJson: payload.LayoutJson, __metadata: {etag: '"2"'}}]}})});
    };
    source.load("key", function (record) { loaded = record; }, failed);
    assert.strictEqual(loaded.orders.scope.positions.x, 0);
    assert.strictEqual(loaded.orders.scope.positions.a, 1);
    assert.strictEqual(loaded.orders.scope.positions.b, 0);
    [-1, 0.5, "1", Infinity].forEach(function (row) {
        source.save(loaded, {scope: {version: 2, positions: {a: row}}}, function () { throw new Error("不正な段を保存"); }, function () { failures += 1; });
    });
    assert.strictEqual(failures, 4);
});

test("週間の指定段の途中が空でも、後の予定を上へ詰めない", function () {
    var source = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var a = {id: "a"}, b = {id: "b"};
    var scope = vm.createContext({layoutEngine: layout, weeklyDisplaySettings: {showMultiDayLine: false},
        getItemsForDay: function () { return [a, b]; },
        splitPurpose: function () { return {section: "G", team: ""}; },
        getLayoutRows: function () { return {version: 2, positions: {a: 0, b: 2}}; }});
    vm.runInContext(source.slice(source.indexOf("    function getItemsForOrganizationDay("), source.indexOf("    function alignMonthlyItemsByLane(")), scope);
    var result = scope.getItemsForOrganizationDay("G", "", new Date(), "weekly");
    assert.strictEqual(result.length, 3);
    assert.strictEqual(result[0], a);
    assert.strictEqual(result[1], undefined);
    assert.strictEqual(result[2], b);
});

test("上下入れ替えは更新モードで明示的にオンにし、更新終了と再開でオフを維持する", function () {
    var source = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var editable = false, cancelled = 0;
    var buttons = {
        "toggle-layout-edit": {disabled: true, setAttribute: function (name, value) { this[name] = value; }},
        "reset-layout": {}, "event-editor": {style: {display: "none"}}
    };
    var orders = {group: {version: 2, positions: {a: 1}}};
    var scope = vm.createContext({layoutEditingEnabled: false, layoutReady: true, layoutBusy: false,
        selectedLayoutScope: "group", layoutRecord: {orders: orders}, settingsController: null,
        byId: function (id) { return buttons[id]; },
        cancelLayoutDrag: function () { cancelled += 1; }, setMessage: function () {}, updateSourceControls: function () {},
        service: {isReadOnly: function () { return !editable; }, isEditingEnabled: function () { return editable; },
            setEditingEnabled: function (value) { editable = value; }, getMode: function () { return "csv"; }}});
    vm.runInContext(source.slice(source.indexOf("    function updateLayoutControls("), source.indexOf("    function saveLayout(")), scope);
    vm.runInContext(source.slice(source.indexOf("    function toggleScheduleEdit("), source.indexOf("    function sameDate(")), scope);
    scope.renderCurrentView = function () { scope.updateLayoutControls(); };
    scope.updateLayoutControls();
    assert.strictEqual(buttons["toggle-layout-edit"].disabled, true);
    scope.toggleLayoutEdit();
    assert.strictEqual(scope.layoutEditingEnabled, false);
    scope.toggleScheduleEdit();
    assert.strictEqual(buttons["toggle-layout-edit"].disabled, false);
    assert.strictEqual(scope.layoutEditingEnabled, false);
    assert.strictEqual(buttons["reset-layout"].disabled, true);
    scope.toggleLayoutEdit();
    assert.strictEqual(scope.layoutEditingEnabled, true);
    assert.strictEqual(buttons["toggle-layout-edit"]["aria-pressed"], "true");
    assert.strictEqual(buttons["reset-layout"].disabled, false);
    scope.toggleScheduleEdit();
    assert.strictEqual(scope.layoutEditingEnabled, false);
    assert.strictEqual(buttons["toggle-layout-edit"].disabled, true);
    scope.toggleScheduleEdit();
    assert.strictEqual(scope.layoutEditingEnabled, false);
    assert.strictEqual(scope.layoutRecord.orders, orders);
    assert.ok(cancelled >= 3);
});

test("上下入れ替え中は予定本体・文字・端をドラッグしても日時移動や編集を起動しない", function () {
    var source = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var moves = 0, edits = 0, reordered = 0;
    function element() { return {className: "event-item", children: [], setAttribute: function () {},
        appendChild: function (child) { this.children.push(child); }}; }
    var scope = vm.createContext({layoutEditingEnabled: true, layoutReady: true,
        service: {isReadOnly: function () { return false; }},
        util: {addEvent: function () {}}, getLayoutScope: function () { return "scope"; },
        addClass: function (node, name) { node.className += " " + name; },
        document: {createElement: element, createTextNode: function (text) { return text; }},
        beginLayoutDrag: function (event, button, item) { reordered += 1; assert.strictEqual(item.id, "a"); },
        selectDailyItem: function () {}, preventEvent: function () { return false; }});
    vm.runInContext(source.slice(source.indexOf("    function decorateLayoutButton("), source.indexOf("    function compareItems(")), scope);
    ["daily", "weekly", "monthly"].forEach(function (view) {
        var button = element();
        button.onmousedown = function () { moves += 1; };
        button.onclick = function () { edits += 1; };
        scope.decorateLayoutButton(button, {id: "a"}, new Date(), view, 0);
        [button, {className: "period-event-caption"}, {className: "daily-resize-start"}].forEach(function (target) {
            button.onmousedown({target: target});
            button.onclick({target: target});
        });
        assert.ok(button.className.indexOf("layout-editing-event") >= 0);
        assert.strictEqual(button.ondragstart({}), false);
    });
    assert.strictEqual(reordered, 9);
    assert.strictEqual(moves, 0);
    assert.strictEqual(edits, 0);
    scope.layoutEditingEnabled = false;
    var normal = element();
    normal.onmousedown = function () { moves += 1; };
    scope.decorateLayoutButton(normal, {id: "a"}, new Date(), "weekly", 0);
    normal.onmousedown({});
    assert.strictEqual(moves, 1);
});

test("週間表示設定の4通りを対象列だけにETag付きで保存し、不正値と競合を拒否する", function () {
    var browser = loadBrowserScripts(["js/util.js", "js/sharepoint-data-source.js", "js/display-settings-data-source.js"]);
    var source = new browser.window.YoteihyouDisplaySettingsDataSource({siteUrl: "https://example.invalid/sites/test"});
    var rows, loaded, writes = 0, failures = 0;
    function fail(message) { throw new Error(message); }
    source.client.getEntityType = function (done) { done("Settings"); };
    source.client.getDigest = function (done) { done("digest"); };
    source.client.request = function (method, url, headers, body, done) {
        if (method === "GET") {
            assert.ok(url.indexOf("$select=ID,WeeklyShowTime,WeeklyShowMultiDayLine") >= 0);
            assert.ok(decodeURIComponent(url).indexOf("Title eq 'default'") >= 0);
            done({responseText: JSON.stringify({d: {results: rows}})});
        } else {
            writes += 1;
            var payload = JSON.parse(body);
            assert.deepStrictEqual(Object.keys(payload).sort(), ["WeeklyShowMultiDayLine", "WeeklyShowTime", "__metadata"]);
            assert.strictEqual(headers["IF-MATCH"], "3");
            assert.strictEqual(headers["X-HTTP-Method"], "MERGE");
            assert.strictEqual(payload.WeeklyShowTime, loaded.showTime);
            assert.strictEqual(payload.WeeklyShowMultiDayLine, loaded.showMultiDayLine);
            done();
        }
    };
    [true, false].forEach(function (time) { [true, false].forEach(function (line) {
        rows = [{ID: 8, WeeklyShowTime: time, WeeklyShowMultiDayLine: line, __metadata: {etag: "3"}}];
        source.loadWeekly(function (item) { loaded = item; }, fail);
        assert.strictEqual(loaded.showTime, time);
        assert.strictEqual(loaded.showMultiDayLine, line);
        source.saveWeekly(loaded, time, line, function () {}, fail);
    }); });
    assert.strictEqual(writes, 4);
    function unexpected() { throw new Error("不正な設定を成功扱いにしました"); }
    function rejected() { failures += 1; }
    [[], [{ID: 8}], [{ID: 8}, {ID: 9}],
        [{ID: 8, WeeklyShowTime: "false", WeeklyShowMultiDayLine: true, __metadata: {etag: "3"}}],
        [{ID: 8, WeeklyShowTime: true, WeeklyShowMultiDayLine: false}]].forEach(function (value) {
        rows = value; source.loadWeekly(unexpected, rejected);
    });
    source.saveWeekly({id: 8}, true, true, unexpected, rejected);
    source.saveWeekly(loaded, "false", true, unexpected, rejected);
    [403, 412].forEach(function (status) {
        source.client.request = function (method, url, headers, body, done, error) { error(String(status)); };
        source.saveWeekly(loaded, false, true, unexpected, rejected);
    });
    assert.strictEqual(failures, 9);
    assert.strictEqual(writes, 4);
});

test("週間表示は保存成功後だけ反映し、読取専用・連打・失敗では現状を保つ", function () {
    var app = fs.readFileSync(path.join(root, "js/app.js"), "utf8"), controls = {}, loadDone, loadFail, saveDone, saveFail;
    var writes = 0, renders = 0, readOnly = false;
    ["setting-weekly-show-time", "setting-weekly-show-multi-day-line", "save-weekly-settings",
        "reload-weekly-settings", "weekly-settings-status"].forEach(function (id) { controls[id] = {}; });
    var scope = vm.createContext({sharedWeeklySettings: null, weeklySettingsBusy: false,
        weeklyDisplaySettings: {showTime: true, showMultiDayLine: false}, state: {viewMode: "weekly"},
        byId: function (id) { return controls[id]; }, cancelLayoutDrag: function () {},
        renderCurrentView: function () { renders += 1; }, updateFixedHeader: function () {},
        service: {isReadOnly: function () { return readOnly; }}, displaySettingsSource: {
            loadWeekly: function (done, fail) { loadDone = done; loadFail = fail; },
            saveWeekly: function (item, time, line, done, fail) { writes += 1; saveDone = done; saveFail = fail; }
        }});
    vm.runInContext(app.slice(app.indexOf("    function updateWeeklySettingsControls("),
        app.indexOf("    function initializeDisplayPreferences(")), scope);
    scope.loadWeeklySettings(); loadDone({id: 8, etag: "1", showTime: true, showMultiDayLine: false});
    controls["setting-weekly-show-time"].checked = false;
    controls["setting-weekly-show-multi-day-line"].checked = true;
    readOnly = true; scope.saveWeeklySettings(); assert.strictEqual(writes, 0);
    readOnly = false; scope.saveWeeklySettings(); scope.saveWeeklySettings(); assert.strictEqual(writes, 1);
    assert.strictEqual(scope.weeklyDisplaySettings.showTime, true);
    saveFail("412"); assert.strictEqual(scope.weeklyDisplaySettings.showMultiDayLine, false);
    assert.strictEqual(controls["save-weekly-settings"].disabled, true);
    scope.loadWeeklySettings(); loadDone({id: 8, etag: "2", showTime: true, showMultiDayLine: false});
    controls["setting-weekly-show-time"].checked = false;
    controls["setting-weekly-show-multi-day-line"].checked = true;
    scope.saveWeeklySettings(); saveDone();
    assert.strictEqual(scope.weeklyDisplaySettings.showTime, false);
    assert.strictEqual(scope.weeklyDisplaySettings.showMultiDayLine, true);
    loadFail("接続失敗"); assert.strictEqual(scope.weeklyDisplaySettings.showMultiDayLine, true);
    var before = renders;
    scope.state.viewMode = "monthly";
    scope.loadWeeklySettings(); loadDone({id: 8, etag: "3", showTime: true, showMultiDayLine: false});
    assert.strictEqual(renders, before);
});

test("週間の日またぎ予定は同じ行に揃い、重複を避けて週単位で配置を保存できる", function () {
    var app = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var a = {id: "a"}, b = {id: "b"}, c = {id: "c"};
    var scope = vm.createContext({layoutEngine: layout, compareItems: function (a, b) { return a.id.localeCompare(b.id); },
        weeklyDisplaySettings: {showMultiDayLine: true}, util: util,
        splitPurpose: function () { return {section: "GP1", team: ""}; }});
    vm.runInContext(app.slice(app.indexOf("    function alignWeeklyItemsByLane("), app.indexOf("    function createHeaderCell(")), scope);
    vm.runInContext(app.slice(app.indexOf("    function getLayoutScope("), app.indexOf("    function getLayoutRows(")), scope);
    var days = [[a, b], [b, c], [b]], result = scope.alignWeeklyItemsByLane(days, null);
    var lane = result.itemsByDate[0].indexOf(b);
    assert.strictEqual(result.itemsByDate[1][lane], b);
    assert.strictEqual(result.itemsByDate[2][lane], b);
    assert.strictEqual(result.requiredRows, 2);
    assert.strictEqual(result.spans.b.start, 0); assert.strictEqual(result.spans.b.end, 3);
    result = scope.alignWeeklyItemsByLane(days, {version: 2, positions: {b: 3}});
    assert.strictEqual(result.requiredRows, 4);
    result.itemsByDate.forEach(function (items) { assert.strictEqual(items[3], b); });
    var first = new Date(2026, 9, 5), next = new Date(2026, 9, 6);
    assert.strictEqual(scope.getLayoutScope(a, first, "weekly"), scope.getLayoutScope(a, next, "weekly"));
    scope.weeklyDisplaySettings.showMultiDayLine = false;
    assert.notStrictEqual(scope.getLayoutScope(a, first, "weekly"), scope.getLayoutScope(a, next, "weekly"));
});

function makeLogFixture(storage) {
    var ctx = loadBrowserScripts(["js/system-log.js"]), timers = [];
    ctx.window.setTimeout = function (fn) { timers.push(fn); return timers.length; };
    ctx.window.clearTimeout = function (id) { timers[id - 1] = null; };
    return {log: new ctx.window.YoteihyouSystemLog({storage: storage}), timers: timers, window: ctx.window};
}

test("システムログは300件を保持して再読込し、保存制限でも記録を続ける", function () {
    var value, storage = {getItem: function () { return value; }, setItem: function (key, data) { value = data; }};
    var fixture = makeLogFixture(storage), i;
    for (i = 0; i < 305; i += 1) { fixture.log.add("list", "読込", "success", String(i)); }
    assert.strictEqual(fixture.log.entries.length, 300);
    assert.strictEqual(fixture.log.entries[0].detail, "5");
    fixture.log.begin("unfinished", "読込");
    var reloaded = makeLogFixture(storage).log;
    assert.strictEqual(reloaded.entries[299].status, "warning");
    storage.setItem = function () { throw new Error("blocked"); };
    reloaded.add("list", "読込", "error", "失敗");
    assert.strictEqual(reloaded.persistent, false);
    assert.strictEqual(reloaded.entries[299].detail, "失敗");
});

test("ログ記録のオンオフは端末に保存せず、停止中も過去のログを保持して再開できる", function () {
    var saved = {}, storage = {getItem: function (key) { return saved[key]; },
        setItem: function (key, value) { saved[key] = value; }};
    var fixture = makeLogFixture(storage), log = fixture.log;
    assert.strictEqual(log.enabled, true);
    log.add("アプリ", "起動", "info", "開始");
    log.setEnabled(false);
    assert.strictEqual(saved["yoteihyou.systemLog.enabled"], undefined);
    log.add("JavaScript", "実行", "error", "停止中");
    log.begin("リスト", "読込")("success", "停止中");
    assert.strictEqual(log.entries.length, 1);
    assert.strictEqual(fixture.timers.length, 0);
    log = makeLogFixture(storage).log;
    saved["yoteihyou.systemLog.enabled"] = "false";
    assert.strictEqual(makeLogFixture(storage).log.enabled, true);
    assert.strictEqual(log.entries.length, 1);
    log.setEnabled(true);
    log.add("リスト", "読込", "success", "再開");
    assert.strictEqual(log.entries.length, 2);
    assert.strictEqual(makeLogFixture(storage).log.enabled, true);
});

test("ログ停止中も元の処理・コールバック・例外を維持し、開始済みの記録は完了する", function () {
    var fixture = makeLogFixture(), log = fixture.log, callback, received;
    var source = {load: function (done) { callback = done; return 7; }};
    log.observe(source, "load", "予定表", 0, 1);
    source.load(function (value) { received = value; });
    log.setEnabled(false);
    callback("完了");
    assert.strictEqual(received, "完了");
    assert.strictEqual(log.entries[0].status, "success");
    var done = function () {};
    assert.strictEqual(source.load(done), 7);
    assert.strictEqual(callback, done);
    assert.strictEqual(log.entries.length, 1);
    var failure = {load: function () { throw new Error("通信エラー"); }};
    log.observe(failure, "load", "予定表", 0, 1);
    assert.throws(function () { failure.load(); }, /通信エラー/);
    assert.strictEqual(log.entries.length, 1);
});

test("ブラウザー保存を利用できなくても取得済みのログ記録設定を適用できる", function () {
    var log = makeLogFixture({getItem: function () { throw new Error("blocked"); },
        setItem: function () { throw new Error("blocked"); }}).log;
    log.setEnabled(false);
    log.add("アプリ", "起動", "info", "停止中");
    assert.strictEqual(log.entries.length, 0);
    log.setEnabled(true);
    log.add("アプリ", "起動", "info", "再開");
    assert.strictEqual(log.entries.length, 1);
});

test("システムログの監視はコールバックと戻り値を維持し、処理結果を一度だけ記録する", function () {
    var fixture = makeLogFixture(), received;
    var source = {load: function (range, success) { assert.strictEqual(this, source); success([range]); return 7; }};
    fixture.log.observe(source, "load", "予定表", 1, 2);
    assert.strictEqual(source.load("range", function (items) { received = items; }), 7);
    assert.deepStrictEqual(received, ["range"]);
    assert.strictEqual(fixture.log.entries.length, 1);
    assert.strictEqual(fixture.log.entries[0].status, "success");
    var failureSource = {load: function (success, failure) { failure("権限不足", 403); }};
    fixture.log.observe(failureSource, "load", "設定", 0, 1);
    failureSource.load(null, function (message, status) { received = [message, status]; });
    assert.deepStrictEqual(received, ["権限不足", 403]);
    assert.strictEqual(fixture.log.entries[1].detail, "HTTP 403 / 権限不足");
});

test("接続確認はGETだけを使い、列不足・未設定・無効設定を区別する", function () {
    var fixture = makeLogFixture(), done = 0, requests = [];
    var client = {siteUrl: "/site", listTitle: "配置", getListPath: function () { return "list"; }, getApiUrl: function (path) { return path; },
        request: function (method, url, headers, body, success) { requests.push([method, url, body]); success({responseText: '{"d":{"results":[]}}'}); }};
    function complete() { done += 1; }
    fixture.log.probe(client, ["ID", "LayoutJson"], false, complete);
    assert.strictEqual(requests[0][0], "GET");
    assert.strictEqual(requests[0][2], null);
    assert.ok(requests[0][1].indexOf("LayoutJson") >= 0);
    assert.strictEqual(fixture.log.entries[0].status, "success");
    client.request = function (method, url, headers, body, success, failure) { failure("列がありません", 400); };
    fixture.log.probe(client, ["LayoutJson"], false, complete);
    fixture.log.probe(client, [], true, complete);
    client.siteUrl = "";
    fixture.log.probe(client, [], false, complete);
    assert.strictEqual(done, 4);
    assert.strictEqual(fixture.log.entries[1].status, "error");
    assert.strictEqual(fixture.log.entries[2].status, "skip");
    assert.strictEqual(fixture.log.entries[3].status, "skip");
});

test("接続確認のタイムアウトで再確認が可能になり、遅延応答で二重完了しない", function () {
    var fixture = makeLogFixture(), callback, completed = 0;
    var client = {siteUrl: "/site", listTitle: "設定", getListPath: function () { return "list"; }, getApiUrl: function (path) { return path; },
        request: function (method, url, headers, body, success) { callback = success; }};
    fixture.log.probe(client, ["ID"], false, function () { completed += 1; });
    fixture.timers.slice().forEach(function (timer) { if (timer) { timer(); } });
    assert.strictEqual(completed, 1);
    assert.strictEqual(fixture.log.entries[0].status, "warning");
    callback({responseText: '{"d":{"results":[]}}'});
    assert.strictEqual(completed, 1);
});

process.stdout.write(passed + " tests passed\n");
