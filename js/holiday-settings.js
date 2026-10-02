(function (window, document) {
    "use strict";
    var util = window.YoteihyouUtil;
    function validate(rows) {
        var seen = {}, result = [];
        if (!Array.isArray(rows) || rows.length > 1000) { throw new Error("休日は1000件以内で入力してください。"); }
        rows.forEach(function (row) {
            if (!Array.isArray(row) || row.length !== 2 || typeof row[0] !== "string" || typeof row[1] !== "string" ||
                    !/^\d{4}-\d{2}-\d{2}$/.test(row[0]) || !util.parseDate(row[0]) ||
                    !util.trim(row[1]) || row[1].length > 100 || /[\r\n|]/.test(row[1])) {
                throw new Error("休日は「YYYY-MM-DD | 名称」で入力してください。日付は実在する日、名称は100文字以内です。");
            }
            if (seen[row[0]]) { throw new Error(row[0] + " が重複しています。"); }
            seen[row[0]] = true; result.push([row[0], util.trim(row[1])]);
        });
        return result.sort(function (a, b) { return a[0] < b[0] ? -1 : 1; });
    }
    function parse(text) {
        var rows = [];
        String(text).split(/\r?\n/).forEach(function (line) {
            if (!util.trim(line)) { return; }
            rows.push(line.split("|").map(function (part) { return util.trim(part); }));
        });
        return validate(rows);
    }
    function Source(options) {
        this.client = new window.YoteihyouSharePointDataSource({siteUrl: options.siteUrl,
            listTitle: (options.layoutSettings || {}).listTitle || "予定表配置設定", fields: {}, timeout: 30000});
    }
    Source.prototype.load = function (success, failure) {
        var c = this.client;
        if (!c.siteUrl) { failure("SharePoint未接続です。初期休日を表示しています。"); return; }
        c.request("GET", c.getApiUrl(c.getListPath() + "/items?$select=ID,Title,LayoutJson&$filter=" +
            encodeURIComponent("Title eq 'holidays-v1'") + "&$top=2"), {}, null, function (xhr) {
            try {
                var rows = util.getJson(xhr).d.results, row = rows[0];
                if (rows.length > 1) { throw new Error("holidays-v1が重複しています。"); }
                success(row ? {id: row.ID, etag: row.__metadata && row.__metadata.etag,
                    rows: validate(JSON.parse(row.LayoutJson))} : {rows: validate(window.YOTEIHYOU_HOLIDAYS || [])});
            } catch (error) { failure("休日設定を読めません。" + error.message); }
        }, failure);
    };
    Source.prototype.save = function (record, rows, success, failure) {
        var c = this.client, data;
        try { data = validate(rows); } catch (error) { failure(error.message); return; }
        if (!record || (record.id && !record.etag)) { failure("休日設定を再読込してから保存してください。"); return; }
        c.getEntityType(function (type) { c.getDigest(function (digest) {
            var headers = {"Content-Type": "application/json;odata=verbose", "X-RequestDigest": digest};
            if (record.id) { headers["IF-MATCH"] = record.etag; headers["X-HTTP-Method"] = "MERGE"; }
            c.request("POST", c.getApiUrl(c.getListPath() + "/items" + (record.id ? "(" + record.id + ")" : "")), headers,
                JSON.stringify({__metadata: {type: type}, Title: "holidays-v1", LayoutJson: JSON.stringify(data)}), success, failure);
        }, failure); }, failure);
    };
    function Holidays(options, canEdit, onChange) {
        this.source = new Source(options); this.canEdit = canEdit; this.onChange = onChange;
        this.rows = validate(window.YOTEIHYOU_HOLIDAYS || []); this.record = null; this.busy = false;
    }
    Holidays.prototype.name = function (day) {
        var key = util.formatDateKey(day), i;
        for (i = 0; i < this.rows.length; i += 1) { if (this.rows[i][0] === key) { return this.rows[i][1]; } }
        return "";
    };
    Holidays.prototype.dayClass = function (day) {
        return this.name(day) || day.getDay() === 0 ? "sunday" : day.getDay() === 6 ? "saturday" : "";
    };
    Holidays.prototype.initialize = function () {
        var self = this, input = document.getElementById("holiday-input"), status = document.getElementById("holiday-status");
        var save = document.getElementById("holiday-save"), reload = document.getElementById("holiday-reload");
        function show() { input.value = self.rows.map(function (r) { return r[0] + " | " + r[1]; }).join("\n"); }
        function busy(value) { self.busy = value; save.disabled = value; reload.disabled = value; input.disabled = value; }
        function load(message) {
            busy(true); self.record = null;
            self.source.load(function (record) {
                self.record = record; self.rows = record.rows; show(); busy(false); self.onChange();
                status.textContent = message || (record.id ? "共有休日を読込みました。" : "初期休日を表示中。保存すると全利用者で共有します。");
            }, function (error) { busy(false); status.textContent = error + " 表示中の休日は維持します。"; });
        }
        save.onclick = function () {
            if (self.busy) { return; }
            if (!self.canEdit()) { status.textContent = "「予定表更新」を押してから保存してください。"; return; }
            var rows;
            try { rows = parse(input.value); } catch (error) { status.textContent = error.message; return; }
            if (!window.confirm("休日" + rows.length + "件を共有保存します。削除した行は休日から外れます。よろしいですか？")) { return; }
            busy(true);
            self.source.save(self.record, rows, function () {
                self.rows = rows; self.record = null; self.onChange();
                load("休日を共有保存しました。");
            }, function (error) { busy(false); status.textContent = error + " 入力内容は残しています。競合時は再読込してやり直してください。"; });
        };
        reload.onclick = function () {
            if (self.busy) { return; }
            if (window.confirm("未保存の休日編集を破棄して再読込しますか？")) { load(); }
        };
        show(); load();
    };
    Holidays.validate = validate; Holidays.parse = parse; Holidays.Source = Source;
    window.YoteihyouHolidays = Holidays;
}(window, document));
