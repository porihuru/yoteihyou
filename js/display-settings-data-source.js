(function (window) {
    "use strict";

    function DisplaySettingsDataSource(options) {
        var settings = options.displaySettings || {};
        this.client = new window.YoteihyouSharePointDataSource({
            siteUrl: options.siteUrl,
            listTitle: settings.listTitle || "予定表表示設定",
            timeout: 20000,
            fields: {}
        });
    }

    DisplaySettingsDataSource.prototype.load = function (success, failure) {
        var client = this.client;
        if (!client.siteUrl) {
            failure("文字設定を読み込むにはSharePointへ接続してください。");
            return;
        }
        client.request("GET", client.getApiUrl(client.getListPath() +
            "/items?$select=ID,FontFamily,FontSize&$filter=" + encodeURIComponent("Title eq 'default'") + "&$top=2"),
        {}, null, function (xhr) {
            var rows;
            try {
                rows = window.YoteihyouUtil.getJson(xhr).d.results;
                if (rows.length !== 1) {
                    failure("予定表表示設定リストにTitleがdefaultの項目を1件用意してください。");
                    return;
                }
                success({id: rows[0].ID, etag: rows[0].__metadata && rows[0].__metadata.etag,
                    family: rows[0].FontFamily, size: rows[0].FontSize});
            } catch (error) {
                failure("文字設定の応答を解析できませんでした。");
            }
        }, failure);
    };

    DisplaySettingsDataSource.prototype.save = function (item, family, size, success, failure) {
        var client = this.client;
        if (!item || !item.id || !item.etag) {
            failure("文字設定をリストから再読込してから保存してください。");
            return;
        }
        if (["default", "meiryo", "gothic", "mincho"].indexOf(family) < 0 ||
                [12, 14, 16, 18, 20].indexOf(Number(size)) < 0) {
            failure("フォントと文字サイズを選択してください。");
            return;
        }
        client.getEntityType(function (entityType) {
            client.getDigest(function (digest) {
                client.request("POST", client.getApiUrl(client.getListPath() + "/items(" + encodeURIComponent(item.id) + ")"), {
                    "Content-Type": "application/json;odata=verbose",
                    "X-RequestDigest": digest,
                    "IF-MATCH": item.etag,
                    "X-HTTP-Method": "MERGE"
                }, JSON.stringify({__metadata: {type: entityType}, FontFamily: family, FontSize: Number(size)}),
                success, failure);
            }, failure);
        }, failure);
    };

    DisplaySettingsDataSource.prototype.loadLogging = function (success, failure) {
        var client = this.client;
        if (!client.siteUrl) { failure("ログ記録設定を読み込むにはSharePointへ接続してください。"); return; }
        client.request("GET", client.getApiUrl(client.getListPath() +
            "/items?$select=ID,SystemLogEnabled&$filter=" + encodeURIComponent("Title eq 'default'") + "&$top=2"),
        {}, null, function (xhr) {
            var rows;
            try { rows = window.YoteihyouUtil.getJson(xhr).d.results; }
            catch (error) { failure("ログ記録設定の応答を解析できませんでした。"); return; }
            if (!Array.isArray(rows) || rows.length !== 1 || typeof rows[0].SystemLogEnabled !== "boolean" ||
                    !rows[0].ID || !rows[0].__metadata || !rows[0].__metadata.etag) {
                failure("予定表表示設定のTitle=defaultの項目を1件用意し、SystemLogEnabled列（はい/いいえ）を確認してください。");
                return;
            }
            success({id: rows[0].ID, etag: rows[0].__metadata.etag, enabled: rows[0].SystemLogEnabled});
        }, failure);
    };

    DisplaySettingsDataSource.prototype.saveLogging = function (item, enabled, success, failure) {
        var client = this.client;
        if (!client.siteUrl || !item || !item.id || !item.etag || typeof enabled !== "boolean") {
            failure("ログ記録設定をSharePointから再読込してから保存してください。"); return;
        }
        client.getEntityType(function (entityType) {
            client.getDigest(function (digest) {
                client.request("POST", client.getApiUrl(client.getListPath() + "/items(" + encodeURIComponent(item.id) + ")"), {
                    "Content-Type": "application/json;odata=verbose", "X-RequestDigest": digest,
                    "IF-MATCH": item.etag, "X-HTTP-Method": "MERGE"
                }, JSON.stringify({__metadata: {type: entityType}, SystemLogEnabled: enabled}), success, failure);
            }, failure);
        }, failure);
    };

    window.YoteihyouDisplaySettingsDataSource = DisplaySettingsDataSource;
}(window));
