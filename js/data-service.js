(function (window) {
    "use strict";

    function DataService(config) {
        this.sources = {
            csv: new window.YoteihyouCsvDataSource(config.csv),
            sharepoint: new window.YoteihyouSharePointDataSource(config.sharePoint)
        };
        this.history = new window.YoteihyouHistoryDataSource(config.sharePoint);
        this.mode = config.defaultDataSource === "sharepoint" ? "sharepoint" : "csv";
        this.editingEnabled = false;
    }

    DataService.prototype.setMode = function (mode) {
        if (!this.sources[mode]) {
            throw new Error("不明なデータ接続先です。");
        }
        this.mode = mode;
    };

    DataService.prototype.getMode = function () {
        return this.mode;
    };

    DataService.prototype.isReadOnly = function () {
        return !this.editingEnabled || this.sources[this.mode].readOnly === true;
    };

    DataService.prototype.setEditingEnabled = function (enabled) {
        this.editingEnabled = enabled === true;
    };

    DataService.prototype.isEditingEnabled = function () {
        return this.editingEnabled;
    };

    DataService.prototype.load = function (range, success, failure) {
        this.sources[this.mode].load(range, success, failure);
    };

    DataService.prototype.create = function (item, success, failure) {
        this.change("create", item, null, success, failure);
    };

    DataService.prototype.update = function (item, success, failure, before) {
        this.change("update", item, before, success, failure);
    };

    DataService.prototype.remove = function (item, success, failure) {
        this.change("delete", item, item, success, failure);
    };

    DataService.prototype.change = function (action, item, before, success, failure) {
        var self = this;
        var mode = this.mode;
        var method = action === "delete" ? "remove" : action;
        if (this.isReadOnly()) {
            failure("予定表は読取専用です。作業するには「予定表更新」を押してください。");
            return;
        }
        this.sources[mode][method](item, function (saved, items) {
            var after = action === "delete" ? null : saved;
            if (mode === "csv") {
                self.history.recordSample(action, before, after);
                success(saved, items, "");
            } else {
                self.history.record(action, before, after, function () {
                    success(saved, items, "");
                }, function (message) {
                    success(saved, items, "予定の変更は完了しましたが、履歴の記録または整理に失敗しました。" + message);
                });
            }
        }, failure);
    };

    DataService.prototype.loadHistory = function (success, failure) {
        if (this.mode === "csv") {
            var self = this;
            var local = this.history.getSample();
            if (!this.history.api.siteUrl) {
                success(local, "SharePoint未接続のため、組織設定の共有履歴は取得できません。");
                return;
            }
            this.history.load(function (entries) {
                var settings = [];
                var i;
                for (i = 0; i < entries.length; i += 1) {
                    if (((entries[i].after || entries[i].before) || {}).kind === "organizationSettings") {
                        settings.push(entries[i]);
                    }
                }
                local = self.history.getSample().filter(function (entry) {
                    return ((entry.after || entry.before) || {}).kind !== "organizationSettings";
                });
                success(settings.concat(local).slice(0, 300), "");
            }, function (message) {
                success(local, "組織設定の共有履歴を取得できませんでした。" + message);
            });
        } else {
            this.history.load(success, failure);
        }
    };

    DataService.prototype.recordSettingsChange = function (action, before, after, done) {
        var self = this;
        this.history.recordSettings(action, before, after, function () {
            if (self.mode === "csv") { self.history.recordSampleSettings(action, before, after); }
            done("");
        }, function (message) {
            done(message);
        });
    };

    window.YoteihyouDataService = DataService;
}(window));
