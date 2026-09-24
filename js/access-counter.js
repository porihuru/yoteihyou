(function (window) {
    "use strict";

    var util = window.YoteihyouUtil;

    function AccessCounter(options) {
        var settings = options.accessCounter || {};
        this.itemTitle = settings.itemTitle || "予定表";
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

        function readAndUpdate() {
            var filter = "Title eq '" + self.itemTitle.replace(/'/g, "''") + "'";
            var url = self.api.getApiUrl(self.api.getListPath() +
                "/items?$select=ID,Title,VisitCount&$filter=" + encodeURIComponent(filter) + "&$top=1");
            self.api.request("GET", url, {}, null, function (xhr) {
                var data;
                var rows;
                var row;
                var count;
                try {
                    data = util.getJson(xhr);
                    rows = data.d && data.d.results ? data.d.results : [];
                    row = rows[0];
                    if (!row || !row.__metadata || !row.__metadata.etag) {
                        failure("アクセスカウンターの項目が見つかりません。");
                        return;
                    }
                    count = parseInt(row.VisitCount, 10);
                    if (isNaN(count) || count < 0) { count = 0; }
                    if (digest) { update(row, count); }
                    else {
                        self.api.getEntityType(function (type) {
                            entityType = type;
                            self.api.getDigest(function (value) {
                                digest = value;
                                update(row, count);
                            }, failure);
                        }, failure);
                    }
                } catch (error) { failure(error.message); }
            }, failure);
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
                else { failure(message); }
            });
        }

        try { readAndUpdate(); }
        catch (error) { failure(error.message); }
    };

    window.YoteihyouAccessCounter = AccessCounter;
}(window));
