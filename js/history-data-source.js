(function (window) {
    "use strict";

    var util = window.YoteihyouUtil;
    var MAX_HISTORY = 300;

    function snapshot(item) {
        if (!item) { return null; }
        return {
            id: String(item.id || ""),
            title: item.title || "",
            startDate: item.startDate ? util.toIsoString(item.startDate) : "",
            endDate: item.endDate ? util.toIsoString(item.endDate) : "",
            allDay: !!item.allDay,
            purpose: item.purpose || "",
            category: item.category || "",
            location: item.location || "",
            description: item.description || "",
            lineStyle: util.normalizeLineStyle(item.lineStyle),
            lineColor: util.normalizeLineColor(item.lineColor),
            textColor: util.normalizeTextColor(item.textColor)
        };
    }

    function buildSampleItems() {
        var titles = ["日次確認", "進捗会議", "資料レビュー", "作業調整", "設備点検",
            "引継ぎ", "週次報告", "現地確認", "打合せ", "月次準備"];
        var actors = ["佐藤", "鈴木", "田中", "高橋", "伊藤"];
        var locations = ["オンライン", "第１会議室", "作業室", "現地", "第２会議室"];
        var actions = ["create", "update", "delete"];
        var entries = [];
        var i;
        var group;
        var previousGroup;
        var action;
        var start;
        var before;
        var after;
        for (i = 0; i < 30; i += 1) {
            group = "GP" + String(i % 5 + 1);
            previousGroup = "GP" + String((i + 3) % 5 + 1);
            action = actions[i % actions.length];
            start = new Date(2026, 9, 11 - i, 9 + i % 4, 0);
            before = {
                id: String(101 + i),
                title: group + " " + titles[i % titles.length],
                startDate: start,
                endDate: new Date(start.getTime() + 60 * 60000),
                allDay: false,
                purpose: (action === "update" ? previousGroup : group) + "｜日々・週間・月間",
                location: locations[i % locations.length],
                description: "サンプル予定 " + String(i + 1),
                lineStyle: i % 4 === 0 ? "dotted" : "solid",
                lineColor: ["default", "red", "green", "brown"][i % 4],
                textColor: ["default", "red", "green", "brown"][i % 4]
            };
            after = {
                id: before.id,
                title: before.title,
                startDate: action === "update" ? new Date(start.getTime() + 30 * 60000) : start,
                endDate: new Date(start.getTime() + (action === "update" ? 90 : 60) * 60000),
                allDay: false,
                purpose: group + "｜日々・週間・月間",
                location: action === "update" ? locations[(i + 1) % locations.length] : before.location,
                description: before.description,
                lineStyle: before.lineStyle,
                lineColor: before.lineColor,
                textColor: before.textColor
            };
            entries.push({
                id: "sample-" + String(i + 1),
                action: action,
                actor: "サンプル：" + actors[i % actors.length],
                at: new Date(2026, 9, 10 - i, 16 - i % 5, 10),
                before: action === "create" ? null : before,
                after: action === "delete" ? null : after
            });
        }
        return entries;
    }

    function HistoryDataSource(options) {
        var history = options.history || {};
        this.listTitle = history.listTitle || "予定表操作履歴";
        this.scheduleListTitle = options.listTitle || "予定表";
        this.api = new window.YoteihyouSharePointDataSource({
            siteUrl: options.siteUrl,
            listTitle: this.listTitle,
            pageSize: 500,
            fields: {id: "ID", title: "Title"}
        });
        this.sessionItems = [];
        this.sampleItems = buildSampleItems();
    }

    HistoryDataSource.prototype.getSample = function () {
        return this.sessionItems.concat(this.sampleItems).slice(0, MAX_HISTORY);
    };

    HistoryDataSource.prototype.recordSample = function (action, before, after) {
        this.sessionItems.unshift({
            id: "session-" + String(new Date().getTime()) + "-" + String(this.sessionItems.length),
            action: action,
            actor: "試験用CSV（この画面）",
            at: new Date(),
            before: snapshot(before),
            after: snapshot(after)
        });
        if (this.sessionItems.length > MAX_HISTORY) { this.sessionItems.length = MAX_HISTORY; }
    };

    HistoryDataSource.prototype.toEntry = function (row) {
        var before;
        var after;
        try {
            before = row.BeforeJson ? JSON.parse(row.BeforeJson) : null;
            after = row.AfterJson ? JSON.parse(row.AfterJson) : null;
        } catch (ignore) {
            before = null;
            after = null;
        }
        return {
            id: row.ID,
            action: row.Action || "",
            actor: row.Author && row.Author.Title ? row.Author.Title : "情報なし",
            at: row.Created ? new Date(row.Created) : null,
            before: before,
            after: after,
            title: row.Title || "",
            scheduleList: row.ScheduleList || this.scheduleListTitle,
            itemId: row.ScheduleItemId || ""
        };
    };

    HistoryDataSource.prototype.load = function (success, failure) {
        var self = this;
        var url = this.api.getApiUrl(this.api.getListPath() +
            "/items?$select=" + encodeURIComponent("ID,Title,Action,ScheduleItemId,ScheduleList,BeforeJson,AfterJson,Created,Author/Title") +
            "&$expand=Author&$orderby=" + encodeURIComponent("ID desc") + "&$top=" + MAX_HISTORY);
        this.api.request("GET", url, {}, null, function (xhr) {
            var data;
            var rows;
            var entries = [];
            var i;
            try {
                data = util.getJson(xhr);
                rows = data.d && data.d.results ? data.d.results : [];
                for (i = 0; i < rows.length; i += 1) { entries.push(self.toEntry(rows[i])); }
                success(entries);
            } catch (error) { failure("履歴の読込に失敗しました。" + error.message); }
        }, failure);
    };

    HistoryDataSource.prototype.prune = function (success, failure) {
        var self = this;
        var ids = [];
        var url = this.api.getApiUrl(this.api.getListPath() +
            "/items?$select=ID&$orderby=" + encodeURIComponent("ID desc") + "&$top=500");
        function readPage(pageUrl) {
            self.api.request("GET", pageUrl, {}, null, function (xhr) {
                var data;
                var rows;
                var i;
                try {
                    data = util.getJson(xhr);
                    rows = data.d && data.d.results ? data.d.results : [];
                    for (i = 0; i < rows.length; i += 1) { ids.push(rows[i].ID); }
                    if (data.d && data.d.__next) { readPage(data.d.__next); }
                    else { deleteOld(); }
                } catch (error) { failure(error.message); }
            }, failure);
        }
        function deleteOld() {
            var old = ids.slice(MAX_HISTORY);
            var digest;
            function next() {
                var id;
                var deleteUrl;
                if (!old.length) { success(); return; }
                id = old.shift();
                deleteUrl = self.api.getApiUrl(self.api.getListPath() + "/items(" + encodeURIComponent(id) + ")");
                self.api.request("POST", deleteUrl, {
                    "X-RequestDigest": digest,
                    "IF-MATCH": "*",
                    "X-HTTP-Method": "DELETE"
                }, null, next, failure);
            }
            if (!old.length) { success(); return; }
            self.api.getDigest(function (value) { digest = value; next(); }, failure);
        }
        readPage(url);
    };

    HistoryDataSource.prototype.record = function (action, before, after, success, failure) {
        var self = this;
        var prior = snapshot(before);
        var current = snapshot(after);
        var target = current || prior || {};
        this.api.getEntityType(function (entityType) {
            self.api.getDigest(function (digest) {
                var payload = {
                    __metadata: {type: entityType},
                    Title: target.title || "（件名なし）",
                    Action: action,
                    ScheduleItemId: target.id || "",
                    ScheduleList: self.scheduleListTitle,
                    BeforeJson: prior ? JSON.stringify(prior) : "",
                    AfterJson: current ? JSON.stringify(current) : ""
                };
                var url = self.api.getApiUrl(self.api.getListPath() + "/items");
                self.api.request("POST", url, {
                    "Content-Type": "application/json;odata=verbose",
                    "X-RequestDigest": digest
                }, JSON.stringify(payload), function () {
                    self.prune(success, failure);
                }, failure);
            }, failure);
        }, failure);
    };

    window.YoteihyouHistoryDataSource = HistoryDataSource;
}(window));
