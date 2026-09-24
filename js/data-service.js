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
            success(this.history.getSample());
        } else {
            this.history.load(success, failure);
        }
    };

    window.YoteihyouDataService = DataService;
}(window));
