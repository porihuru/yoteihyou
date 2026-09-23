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

test("日々表示が組織縦・時刻横の時間枠を使用する", function () {
    var appSource = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var html = fs.readFileSync(path.join(root, "index.html"), "utf8");
    assert.ok(appSource.indexOf("dailyViewConfig.slotMinutes") >= 0);
    assert.ok(appSource.indexOf("dailyViewConfig.startHour") >= 0);
    assert.ok(appSource.indexOf("dailyViewConfig.endHour") >= 0);
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
});

test("日々予定にドラッグ操作とコピー操作のUIがある", function () {
    var appSource = fs.readFileSync(path.join(root, "js/app.js"), "utf8");
    var html = fs.readFileSync(path.join(root, "index.html"), "utf8");
    assert.ok(appSource.indexOf('beginDailyDrag(event || window.event, item, button, "move")') >= 0);
    assert.ok(appSource.indexOf('beginDailyDrag(event || window.event, item, button, "start")') >= 0);
    assert.ok(appSource.indexOf('beginDailyDrag(event || window.event, item, button, "end")') >= 0);
    assert.ok(appSource.indexOf("event.keyCode === 67") >= 0);
    assert.ok(appSource.indexOf("event.keyCode === 88") >= 0);
    assert.ok(appSource.indexOf("event.keyCode === 86") >= 0);
    assert.ok(html.indexOf('id="daily-context-menu"') >= 0);
});

test("試験用CSVは9月と10月の全グループを毎日収録する", function () {
    var generator = require(path.join(root, "scripts/generate-sample-schedule.js"));
    var rows = generator.buildRows();
    var csvText = fs.readFileSync(path.join(root, "data/schedule.csv"), "utf8");
    var dailyCoverage = {};
    var groupCounts = {};
    var current = new Date(Date.UTC(2026, 8, 1));
    var last = new Date(Date.UTC(2026, 9, 31));
    var dateText;
    var group;
    var i;

    rows.forEach(function (row) {
        group = row.Purpose.split("｜")[0];
        groupCounts[group] = (groupCounts[group] || 0) + 1;
        dailyCoverage[row.EventDate.substring(0, 10) + "\u001f" + group] = true;
    });

    while (current.getTime() <= last.getTime()) {
        dateText = current.getUTCFullYear() + "-" +
            (current.getUTCMonth() + 1 < 10 ? "0" : "") + (current.getUTCMonth() + 1) + "-" +
            (current.getUTCDate() < 10 ? "0" : "") + current.getUTCDate();
        for (i = 0; i < generator.groups.length; i += 1) {
            assert.strictEqual(dailyCoverage[dateText + "\u001f" + generator.groups[i]], true,
                dateText + " " + generator.groups[i]);
        }
        current = new Date(current.getTime() + 24 * 60 * 60 * 1000);
    }

    assert.strictEqual(rows.length, 2198);
    assert.strictEqual(csvText, generator.toCsv(rows));
    assert.strictEqual(groupCounts.GP1, 63);
    generator.busyGroups.forEach(function (busyGroup) {
        assert.strictEqual(groupCounts[busyGroup], 244);
    });
    assert.ok(rows.some(function (row) {
        return row.Title === "GP1 月またぎ計画" && row.EventDate < "2026-10-01" && row.EndDate >= "2026-10-01";
    }));
    assert.ok(rows.some(function (row) {
        return row.Title === "GP1 週またぎ対応" && row.EventDate === "2026-10-09 13:00" && row.EndDate === "2026-10-13 12:00";
    }));
});

process.stdout.write(passed + " tests passed\n");
