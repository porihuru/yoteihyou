(function (window) {
    "use strict";

    function DataService(config) {
        this.sources = {
            csv: new window.YoteihyouCsvDataSource(config.csv),
            sharepoint: new window.YoteihyouSharePointDataSource(config.sharePoint)
        };
        this.mode = config.defaultDataSource === "sharepoint" ? "sharepoint" : "csv";
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
        return this.sources[this.mode].readOnly === true;
    };

    DataService.prototype.load = function (success, failure) {
        this.sources[this.mode].load(success, failure);
    };

    DataService.prototype.create = function (item, success, failure) {
        this.sources[this.mode].create(item, success, failure);
    };

    DataService.prototype.update = function (item, success, failure) {
        this.sources[this.mode].update(item, success, failure);
    };

    DataService.prototype.remove = function (item, success, failure) {
        this.sources[this.mode].remove(item, success, failure);
    };

    window.YoteihyouDataService = DataService;
}(window));
