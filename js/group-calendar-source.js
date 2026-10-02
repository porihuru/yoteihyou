(function (window) {
    "use strict";
    var util = window.YoteihyouUtil;
    var own = Object.prototype.hasOwnProperty;

    function copy(value) {
        var result = {}, key;
        for (key in value) { if (own.call(value, key)) { result[key] = value[key]; } }
        return result;
    }
    function groupKey(section, team) { return JSON.stringify([section || "", team || ""]); }
    function standardClient(siteUrl, listTitle) {
        return new window.YoteihyouSharePointDataSource({siteUrl: siteUrl, listTitle: listTitle, timeout: 30000, fields: {
            id: "ID", title: "Title", startDate: "EventDate", endDate: "EndDate",
            allDay: "fAllDayEvent", category: "Category", location: "Location", description: "Description",
            recurrence: "fRecurrence", eventType: "EventType"
        }});
    }
    function resolve(client, success, failure) {
        var url;
        try { url = client.getApiUrl(client.getListPath() + "?$select=Id,BaseTemplate"); }
        catch (error) { failure(error.message); return; }
        client.request("GET", url, {}, null, function (xhr) {
            var data;
            try {
                data = util.getJson(xhr).d;
                if (Number(data.BaseTemplate) !== 106) { throw new Error("標準の予定表リストを指定してください。"); }
                if (!/^[0-9a-f-]{36}$/i.test(data.Id)) { throw new Error("予定表の識別番号を取得できません。"); }
            } catch (error) { failure(error.message); return; }
            success(String(data.Id).toLowerCase());
        }, failure);
    }

    // Auxiliary data uses its own records in the existing layout list, never calendar columns.
    function Metadata(options) {
        this.client = new window.YoteihyouSharePointDataSource({siteUrl: options.siteUrl,
            listTitle: (options.layoutSettings || {}).listTitle || "予定表配置設定", timeout: 30000, fields: {}});
    }
    Metadata.prototype.load = function (guid, success, failure) {
        var client = this.client, key = "events:" + guid, url;
        try {
            url = client.getApiUrl(client.getListPath() + "/items?$select=ID,Title,LayoutJson&$filter=" +
                encodeURIComponent("Title eq '" + key + "'") + "&$top=2");
        } catch (error) { failure(error.message); return; }
        client.request("GET", url, {}, null, function (xhr) {
            var rows, row, data, id, entry;
            try {
                rows = util.getJson(xhr).d.results;
                if (rows.length > 1) { throw new Error("補助設定のTitleが重複しています。"); }
                row = rows[0];
                data = row ? JSON.parse(row.LayoutJson || "{}") : {};
                if (!data || typeof data !== "object" || Array.isArray(data)) { throw new Error("補助設定が不正です。"); }
                for (id in data) {
                    if (!own.call(data, id)) { continue; }
                    entry = data[id];
                    if (!/^\d+$/.test(id) || !entry || typeof entry !== "object" ||
                            typeof entry.targets !== "string" || !/^(日々|週間|月間)(・(日々|週間|月間))*$/.test(entry.targets)) {
                        throw new Error("補助設定の予定情報が不正です。");
                    }
                }
            } catch (error) { failure("予定の補助設定を読み込めません。" + error.message); return; }
            success({key: key, id: row && row.ID, etag: row && row.__metadata && row.__metadata.etag, data: data});
        }, failure);
    };
    Metadata.prototype.save = function (record, success, failure) {
        var client = this.client;
        if (record.id && !record.etag) { failure("補助設定の更新情報がありません。再読込してください。"); return; }
        client.getEntityType(function (type) {
            client.getDigest(function (digest) {
                var headers = {"Content-Type": "application/json;odata=verbose", "X-RequestDigest": digest};
                if (record.id) { headers["IF-MATCH"] = record.etag; headers["X-HTTP-Method"] = "MERGE"; }
                client.request("POST", client.getApiUrl(client.getListPath() + "/items" +
                    (record.id ? "(" + encodeURIComponent(record.id) + ")" : "")), headers,
                    JSON.stringify({__metadata: {type: type}, Title: record.key, LayoutJson: JSON.stringify(record.data)}),
                    success, failure);
            }, failure);
        }, failure);
    };

    function Source(options) {
        this.options = options;
        this.metadata = new Metadata(options);
        this.siteUrl = this.metadata.client.siteUrl;
        this.listTitle = "グループ別予定表";
        this.fields = {};
        this.readOnly = false;
        this.configure(window.YOTEIHYOU_ORGANIZATIONS || {});
    }
    Source.prototype.configure = function (config) {
        var self = this, groups = config.groups || config.sections || [], i, j, teams;
        this.config = config;
        this.generation = (this.generation || 0) + 1;
        this.entries = [];
        this.statuses = [];
        this.items = {};
        this.loading = true;
        function add(group, team, setting) {
            self.entries.push({section: group, team: team, key: groupKey(group, team),
                label: group + (team ? "／" + team : ""), guid: "",
                client: setting.calendarListTitle ? standardClient(setting.calendarSiteUrl || self.options.siteUrl,
                    setting.calendarListTitle) : null});
        }
        for (i = 0; i < groups.length; i += 1) {
            teams = groups[i].teams || [];
            if (!teams.length || groups[i].calendarListTitle) { add(groups[i].name, "", groups[i]); }
            for (j = 0; j < teams.length; j += 1) {
                add(groups[i].name, typeof teams[j] === "string" ? teams[j] : teams[j].name, teams[j]);
            }
        }
    };
    Source.prototype.purpose = function (entry, targets) {
        return entry.section + (entry.team ? (this.config.separator || "／") + entry.team : "") +
            (this.config.metadataSeparator || "｜") + (targets || "日々・週間・月間");
    };
    Source.prototype.decorate = function (item, entry, metadata) {
        var result = copy(item), extra = metadata.data[String(item.id)];
        result.id = entry.guid + ":" + item.id;
        result.calendarItemId = item.id;
        result.calendarGuid = entry.guid;
        result.calendarSiteUrl = entry.client.siteUrl;
        result.calendarListTitle = entry.client.listTitle;
        result.purpose = this.purpose(entry, extra && extra.targets);
        if (extra) {
            result.lineStyle = util.normalizeLineStyle(extra.lineStyle);
            result.lineColor = util.normalizeLineColor(extra.lineColor);
            result.textColor = util.normalizeTextColor(extra.textColor);
        }
        return result;
    };
    Source.prototype.load = function (range, success, failure) {
        var self = this, entries = this.entries.slice(0), results = [], statuses = [], items = {}, used = {}, index = 0;
        var generation = ++this.generation;
        this.loading = true;
        this.items = {};
        if (!entries.length) { statuses.push({key: "", label: "グループ設定", error: "有効なグループがありません。設定を登録してください。"}); }
        // Sequential requests keep load bounded on the intranet. Each request owns its result state.
        function next() {
            var entry;
            if (generation !== self.generation) { return; }
            if (index === entries.length) {
                self.statuses = statuses;
                self.items = items;
                self.loading = false;
                success(results, statuses.filter(function (s) { return s.error || s.warning; }).map(function (s) {
                    return s.label + "：" + (s.error || s.warning);
                }).join("\n"));
                return;
            }
            entry = entries[index++];
            function failed(message) { statuses.push({key: entry.key, label: entry.label, error: message}); next(); }
            if (!entry.client) { failed("接続先未設定"); return; }
            resolve(entry.client, function (guid) {
                entry.guid = guid;
                if (used[guid]) { failed("同じ予定表が複数のグループに設定されています。"); return; }
                used[guid] = true;
                entry.client.load(range, function (loaded) {
                    self.metadata.load(guid, function (metadata) {
                        var i, item, recurringCount = 0;
                        for (i = 0; i < loaded.length; i += 1) {
                            if (loaded[i].recurring) { recurringCount += 1; continue; }
                            item = self.decorate(loaded[i], entry, metadata);
                            results.push(item);
                            items[item.id] = {entry: entry, item: item};
                        }
                        statuses.push({key: entry.key, label: entry.label, error: "", warning: recurringCount ?
                            "繰り返し予定は未対応のため表示対象外です。SharePoint側で確認してください。" : ""});
                        next();
                    }, failed);
                }, failed);
            }, failed);
        }
        next();
    };
    Source.prototype.destination = function (item) {
        var org = String(item.purpose || "").split(this.config.metadataSeparator || "｜")[0];
        var i, entry;
        for (i = 0; i < this.entries.length; i += 1) {
            entry = this.entries[i];
            if (this.purpose(entry).split(this.config.metadataSeparator || "｜")[0] === org) { return entry; }
        }
        return null;
    };
    Source.prototype.change = function (method, item, success, failure) {
        var self = this, original = this.items[String(item.id)], entry, plain, extra, targets;
        if (this.loading) { failure("グループ別予定表の読込が終わってから保存してください。"); return; }
        if (method !== "create" && !original) { failure("元の予定表を確認できません。再読込してください。"); return; }
        entry = method === "remove" ? original.entry : this.destination(item);
        if (!entry || !entry.client || !entry.guid || this.statuses.some(function (status) {
            return status.key === entry.key && !!status.error;
        })) { failure("保存先グループの予定表を設定し、再読込してください。"); return; }
        if (original && method === "update" && original.entry.guid !== entry.guid) {
            failure("別の予定表への移動はできません。移動先へコピーし、確認後に元の予定を削除してください。"); return;
        }
        targets = String(item.purpose || "").split(this.config.metadataSeparator || "｜")[1] || "日々・週間・月間";
        extra = {targets: targets, lineStyle: util.normalizeLineStyle(item.lineStyle),
            lineColor: util.normalizeLineColor(item.lineColor), textColor: util.normalizeTextColor(item.textColor)};
        plain = copy(item);
        plain.id = method === "create" ? "" : original.item.calendarItemId;
        plain.lineStyle = "solid"; plain.lineColor = "default"; plain.textColor = "default";
        self.metadata.load(entry.guid, function (record) {
            entry.client[method](plain, function (saved, savedItems, calendarWarning) {
                var result;
                if (method === "remove") { delete record.data[String(plain.id)]; result = item; }
                else { record.data[String(saved.id)] = extra; result = self.decorate(saved, entry, record); }
                self.metadata.save(record, function () { success(result, null, calendarWarning || ""); }, function (message) {
                    // The calendar write succeeded; never report it as a failed create (duplicate risk).
                    success(result, null, (calendarWarning || "") + " 予定本体の変更は完了しましたが、反映先・色などの補助設定は保存できませんでした。" + message);
                });
            }, failure);
        }, failure);
    };
    Source.prototype.create = function (item, success, failure) { this.change("create", item, success, failure); };
    Source.prototype.update = function (item, success, failure) { this.change("update", item, success, failure); };
    Source.prototype.remove = function (item, success, failure) { this.change("remove", item, success, failure); };
    Source.check = function (siteUrl, listTitle, success, failure) {
        var client = standardClient(siteUrl, listTitle);
        resolve(client, function () {
            client.load({startDate: new Date(), endDate: new Date()}, function () { success(); }, failure);
        }, failure);
    };
    window.YoteihyouGroupCalendarSource = Source;
}(window));
