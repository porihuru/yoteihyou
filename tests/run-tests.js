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
    "js/sharepoint-settings-data-source.js"
]);
var util = context.window.YoteihyouUtil;
var CsvDataSource = context.window.YoteihyouCsvDataSource;
var SharePointDataSource = context.window.YoteihyouSharePointDataSource;
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
    var converted = source.toOrganizationConfig([
        {id: 3, groupName: "GP2", teamName: "", monthlyRows: 1, weeklyRows: 2, dailyRows: 3, sortOrder: 20, isActive: true},
        {id: 1, groupName: "GP1科", teamName: "GS班", monthlyRows: 5, weeklyRows: 2, dailyRows: 2, sortOrder: 10, isActive: true},
        {id: 2, groupName: "GP1科", teamName: "JN班", monthlyRows: 4, weeklyRows: 3, dailyRows: 2, sortOrder: 11, isActive: true},
        {id: 4, groupName: "無効", teamName: "", monthlyRows: 1, weeklyRows: 1, dailyRows: 1, sortOrder: 30, isActive: false}
    ], {separator: "／", metadataSeparator: "｜", targetSeparator: "・"});
    assert.strictEqual(converted.groups.length, 2);
    assert.strictEqual(converted.groups[0].name, "GP1科");
    assert.strictEqual(converted.groups[0].teams.length, 2);
    assert.strictEqual(converted.groups[0].teams[1].name, "JN班");
    assert.strictEqual(converted.groups[1].name, "GP2");
    assert.strictEqual(converted.groups[1].dailyRows, 3);
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

test("日々表示が時間枠設定を使用する", function () {
    var appSource = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var html = fs.readFileSync(path.join(root, "index.html"), "utf8");
    assert.ok(appSource.indexOf("dailyViewConfig.slotMinutes") >= 0);
    assert.ok(appSource.indexOf("dailyViewConfig.startHour") >= 0);
    assert.ok(appSource.indexOf("dailyViewConfig.endHour") >= 0);
    assert.ok(html.indexOf("空白の時間帯をクリックして入力") >= 0);
    assert.strictEqual(html.indexOf('id="daily-head"'), -1);
});

process.stdout.write(passed + " tests passed\n");
