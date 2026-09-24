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
    counter.api.getEntityType = function (success) { success("SP.Data.CounterListItem"); };
    counter.api.getDigest = function (success) { success("digest"); };
    counter.api.request = function (method, url, headers, body, success, failure) {
        if (method === "GET") {
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

test("上部に更新切替と小さな共有カウンターを置き縦余白を半減する", function () {
    var html = fs.readFileSync(path.join(root, "index.html"), "utf8");
    var css = fs.readFileSync(path.join(root, "css/style.css"), "utf8");
    var app = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    assert.ok(html.indexOf('id="edit-mode-status"') >= 0);
    assert.ok(html.indexOf('id="toggle-schedule-edit"') >= 0);
    assert.ok(html.indexOf('id="access-counter"') >= 0);
    assert.ok(html.indexOf('>アクセス 123</span>') >= 0);
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

test("週間のグループ内横罫線と予定内の横線を表示しない", function () {
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
        "weekly-table": {offsetWidth: 1400},
        "month-title": {}, "print-heading": {}, "weekly-head": {}, "weekly-body": {}
    };
    var weekendItems = {};
    var scope = vm.createContext({
        state: {displayDate: new Date(2026, 8, 24)},
        weeklyColumnWeights: [],
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
    assert.ok(/\.weekly-schedule thead th\.saturday\s*\{[^}]*color:\s*#1f5794;/.test(css));
    assert.ok(/\.weekly-schedule thead th\.sunday\s*\{[^}]*color:\s*#b32929;/.test(css));
    assert.ok(/body\.dark-mode \.weekly-schedule thead th\.saturday\s*\{[^}]*color:\s*#8dbdff;/.test(css));
    assert.ok(/body\.dark-mode \.weekly-schedule thead th\.sunday\s*\{[^}]*color:\s*#ff9696;/.test(css));
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
    var scope = vm.createContext({Math: Math});
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
    var scope = vm.createContext({Math: Math});
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
    var scope = vm.createContext({Math: Math});
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
    assert.ok(button.children[3].className.indexOf("period-resize-start") >= 0);
    assert.ok(button.children[4].className.indexOf("period-resize-end") >= 0);
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
    assert.strictEqual(button.children.length, 4);
});

test("日々予定にドラッグ操作とコピー操作のUIがある", function () {
    var appSource = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var css = fs.readFileSync(path.join(root, "css/style.css"), "utf8");
    var html = fs.readFileSync(path.join(root, "index.html"), "utf8");
    assert.ok(appSource.indexOf('beginDailyDrag(event || window.event, item, button, "move")') >= 0);
    assert.ok(appSource.indexOf('beginDailyDrag(event || window.event, item, button, "start")') >= 0);
    assert.ok(appSource.indexOf('beginDailyDrag(event || window.event, item, button, "end")') >= 0);
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

process.stdout.write(passed + " tests passed\n");
