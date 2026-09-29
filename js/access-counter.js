(function (window) {
    "use strict";

    var util = window.YoteihyouUtil;

    function AccessCounter(options) {
        var settings = options.accessCounter || {};
        this.itemTitle = settings.itemTitle || "予定表";
        this.readSequence = 0;
        this.readToken = new Date().getTime() + "-" + Math.random().toString(36).slice(2);
        this.api = new window.YoteihyouSharePointDataSource({
            siteUrl: options.siteUrl,
            listTitle: settings.listTitle || "予定表アクセスカウンター",
            pageSize: 1,
            fields: {id: "ID", title: "Title"}
        });
    }

    AccessCounter.prototype.increment = function (success, failure) {
        var self = this;
        var digest = "";
        var entityType = "";
        var attempts = 0;

        function failedAt(stage) {
            return function (message, status) {
                failure(stage + ": " + message, status);
            };
        }

        function readAndUpdate() {
            var filter = "Title eq '" + self.itemTitle.replace(/'/g, "''") + "'";
            var url = self.api.getApiUrl(self.api.getListPath() +
                "/items?$select=ID,Title,VisitCount&$filter=" + encodeURIComponent(filter) + "&$top=1&_=" +
                encodeURIComponent(self.readToken + "-" + (++self.readSequence)));
            self.api.request("GET", url, {"Cache-Control": "no-cache", "Pragma": "no-cache"}, null, function (xhr) {
                var data;
                var rows;
                var row;
                var count;
                try {
                    data = util.getJson(xhr);
                    rows = data.d && data.d.results ? data.d.results : [];
                    row = rows[0];
                    if (!row) {
                        failure("アクセスカウンターの初期項目がありません。リスト「" + self.api.listTitle +
                            "」に Title=「" + self.itemTitle + "」、VisitCount=0 の項目を作成してください。");
                        return;
                    }
                    if (!row.__metadata || !row.__metadata.etag) {
                        failure("アクセスカウンターの更新情報（ETag）を取得できません。");
                        return;
                    }
                    count = Number(row.VisitCount);
                    if (row.VisitCount === null || row.VisitCount === undefined || row.VisitCount === "" ||
                            !isFinite(count) || count < 0 || Math.floor(count) !== count) {
                        failure("アクセスカウンターのVisitCountが不正です。0以上の整数を設定してください。");
                        return;
                    }
                    if (digest) { update(row, count); }
                    else {
                        self.api.getEntityType(function (type) {
                            entityType = type;
                            self.api.getDigest(function (value) {
                                digest = value;
                                update(row, count);
                            }, failedAt("更新用トークンの取得"));
                        }, failedAt("リスト情報の取得"));
                    }
                } catch (error) { failure(error.message); }
            }, failedAt("カウンター項目の読込"));
        }

        function update(row, count) {
            var url = self.api.getApiUrl(self.api.getListPath() + "/items(" + encodeURIComponent(row.ID) + ")");
            self.api.request("POST", url, {
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": digest,
                "IF-MATCH": row.__metadata.etag,
                "X-HTTP-Method": "MERGE"
            }, JSON.stringify({__metadata: {type: entityType}, VisitCount: count + 1}), function () {
                success(count + 1);
            }, function (message, status) {
                attempts += 1;
                if (status === 412 && attempts < 5) { readAndUpdate(); }
                else if (status === 412) { failure("アクセス数の同時更新が5回続いたため集計できませんでした。", status); }
                else { failedAt("アクセス数の更新")(message, status); }
            });
        }

        try { readAndUpdate(); }
        catch (error) { failure(error.message); }
    };

    window.YoteihyouAccessCounter = AccessCounter;
}(window));
