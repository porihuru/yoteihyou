(function (window, document) {
    "use strict";
    var util = window.YoteihyouUtil;
    function validate(rows) {
        var seen = {};
        if (!Array.isArray(rows) || rows.length > 1000) { throw new Error("グループ設定は1000件以内です。"); }
        return rows.map(function (r) {
            if (!Array.isArray(r) || r.length !== 4 || typeof r[0] !== "string" || typeof r[1] !== "string" ||
                    typeof r[2] !== "boolean" || typeof r[3] !== "boolean") { throw new Error("グループ設定の形式が不正です。"); }
            var key = JSON.stringify(r.slice(0, 2));
            if (seen[key]) { throw new Error("グループ設定が重複しています。"); }
            seen[key] = true; return r.slice(0);
        });
    }
    function Settings(options, defaults, groups, canEdit, changed) {
        this.client = new window.YoteihyouSharePointDataSource({siteUrl: options.siteUrl,
            listTitle: (options.layoutSettings || {}).listTitle || "予定表配置設定", fields: {}, timeout: 20000});
        this.defaults = defaults; this.groups = groups; this.canEdit = canEdit; this.changed = changed;
        this.rows = []; this.record = null; this.busy = false;
    }
    Settings.prototype.get = function (section, team) {
        var i;
        for (i = 0; i < this.rows.length; i += 1) {
            if (this.rows[i][0] === section && this.rows[i][1] === (team || "")) {
                return {showTime: this.rows[i][2], showMultiDayLine: this.rows[i][3]};
            }
        }
        return this.defaults();
    };
    Settings.prototype.refresh = function () {
        var select = document.getElementById("weekly-group-target"), self = this, old = select.value;
        if (this.busy) { return; }
        select.innerHTML = "";
        this.groups().forEach(function (g) {
            var option = document.createElement("option");
            option.value = JSON.stringify([g.section, g.team || ""]);
            option.appendChild(document.createTextNode(g.label)); select.appendChild(option);
        });
        select.value = old;
        if (select.selectedIndex < 0 && select.options.length) { select.selectedIndex = 0; }
        if (self.show) { self.show(); }
    };
    Settings.prototype.initialize = function () {
        var self = this, c = this.client;
        function el(id) { return document.getElementById("weekly-group-" + id); }
        function status(text) { el("status").textContent = text; }
        function controls() {
            el("target").disabled = self.busy;
            el("reload").disabled = self.busy;
            el("save").disabled = self.busy || !self.record || !self.canEdit() || !el("target").value;
            el("inherit").disabled = self.busy || !self.canEdit();
            el("time").disabled = el("line").disabled = self.busy || !self.canEdit() || el("inherit").checked;
        }
        this.show = function () {
            if (!el("target").value) { controls(); return; }
            var key = JSON.parse(el("target").value), value = self.get(key[0], key[1]);
            el("time").checked = value.showTime; el("line").checked = value.showMultiDayLine;
            el("inherit").checked = !self.rows.some(function (r) { return r[0] === key[0] && r[1] === key[1]; });
            controls();
        };
        function load(message) {
            if (self.busy) { return; }
            self.record = null;
            if (!c.siteUrl) { status("SharePoint未接続です。グループ別設定は保存できません。"); controls(); return; }
            self.busy = true; controls();
            c.request("GET", c.getApiUrl(c.getListPath() + "/items?$select=ID,Title,LayoutJson&$filter=" +
                encodeURIComponent("Title eq 'weekly-groups-v1'") + "&$top=2"), {}, null, function (xhr) {
                try {
                    var rows = util.getJson(xhr).d.results, r = rows[0];
                    if (rows.length > 1) { throw new Error("設定が重複しています。"); }
                    var data = r ? validate(JSON.parse(r.LayoutJson)) : [];
                    self.rows = data; self.record = r ? {id: r.ID, etag: r.__metadata && r.__metadata.etag} : {};
                } catch (e) { failed(e.message); return; }
                self.busy = false; self.refresh(); self.changed(); status(message || "グループ別設定を読み込みました。");
            }, failed);
        }
        function failed(error) { self.busy = false; self.record = null; controls(); status(error + " 再読込してください。表示は維持します。"); }
        el("target").onchange = this.show;
        el("inherit").onchange = controls;
        el("reload").onclick = function () { load(); };
        el("save").onclick = function () {
            if (self.busy || !self.record || !self.canEdit() || !el("target").value) { return; }
            var key = JSON.parse(el("target").value), record = self.record;
            var rows = self.rows.filter(function (r) { return r[0] !== key[0] || r[1] !== key[1]; });
            if (!el("inherit").checked) { rows.push([key[0], key[1], el("time").checked, el("line").checked]); }
            try { rows = validate(rows); } catch (error) { status(error.message); return; }
            if (record.id && !record.etag) { failed("更新情報がありません。"); return; }
            self.busy = true; controls();
            c.getEntityType(function (type) { c.getDigest(function (digest) {
                var headers = {"Content-Type": "application/json;odata=verbose", "X-RequestDigest": digest};
                if (record.id) { headers["IF-MATCH"] = record.etag; headers["X-HTTP-Method"] = "MERGE"; }
                c.request("POST", c.getApiUrl(c.getListPath() + "/items" + (record.id ? "(" + record.id + ")" : "")), headers,
                    JSON.stringify({__metadata: {type: type}, Title: "weekly-groups-v1", LayoutJson: JSON.stringify(validate(rows))}), function () {
                        self.rows = rows; self.changed(); self.busy = false; load("グループ別設定を保存しました。");
                    }, failed);
            }, failed); }, failed);
        };
        this.refresh(); load();
    };
    Settings.validate = validate;
    window.YoteihyouWeeklyGroups = Settings;
}(window, document));
