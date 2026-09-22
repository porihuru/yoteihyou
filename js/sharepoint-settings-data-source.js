(function (window) {
    "use strict";

    function numberOrDefault(value, defaultValue) {
        var parsed = parseInt(value, 10);
        return isNaN(parsed) ? defaultValue : parsed;
    }

    function OrganizationSettingsDataSource(sharePointOptions) {
        var options = sharePointOptions.organizationSettings || {};
        var clientOptions = {
            siteUrl: sharePointOptions.siteUrl,
            listTitle: options.listTitle || "予定表組織設定",
            pageSize: options.pageSize || 500,
            fields: options.fields || {}
        };
        this.enabled = options.enabled !== false;
        this.fields = clientOptions.fields;
        this.client = new window.YoteihyouSharePointDataSource(clientOptions);
    }

    OrganizationSettingsDataSource.prototype.canConnect = function () {
        return this.enabled && !!this.client.siteUrl;
    };

    OrganizationSettingsDataSource.prototype.toItem = function (row) {
        var f = this.fields;
        return {
            id: row[f.id],
            etag: row.__metadata && row.__metadata.etag ? row.__metadata.etag : "",
            groupName: row[f.groupName] || "",
            teamName: row[f.teamName] || "",
            monthlyRows: numberOrDefault(row[f.monthlyRows], 1),
            weeklyRows: numberOrDefault(row[f.weeklyRows], 1),
            dailyRows: numberOrDefault(row[f.dailyRows], 1),
            sortOrder: numberOrDefault(row[f.sortOrder], 0),
            isActive: row[f.isActive] !== false
        };
    };

    OrganizationSettingsDataSource.prototype.load = function (success, failure) {
        var self = this;
        var f = this.fields;
        var select = [
            f.id,
            f.groupName,
            f.teamName,
            f.monthlyRows,
            f.weeklyRows,
            f.dailyRows,
            f.sortOrder,
            f.isActive
        ].join(",");
        var url;
        var items = [];

        if (!this.canConnect()) {
            failure("SharePointサイト上で開くか、config.jsのsiteUrlを設定してください。");
            return;
        }
        try {
            url = this.client.getApiUrl(this.client.getListPath() + "/items?$select=" + encodeURIComponent(select) +
                "&$orderby=" + encodeURIComponent(f.sortOrder + " asc," + f.id + " asc") +
                "&$top=" + encodeURIComponent(this.client.pageSize));
        } catch (error) {
            failure(error.message);
            return;
        }

        function loadPage(pageUrl) {
            self.client.request("GET", pageUrl, {}, null, function (xhr) {
                var data;
                var results;
                var i;
                try {
                    data = window.YoteihyouUtil.getJson(xhr);
                    results = data.d && data.d.results ? data.d.results : [];
                    for (i = 0; i < results.length; i += 1) {
                        items.push(self.toItem(results[i]));
                    }
                    if (data.d && data.d.__next) {
                        loadPage(data.d.__next);
                    } else {
                        success(items);
                    }
                } catch (error) {
                    failure("組織設定の応答を解析できませんでした。" + (error.message ? " " + error.message : ""));
                }
            }, failure);
        }

        loadPage(url);
    };

    OrganizationSettingsDataSource.prototype.toPayload = function (item, entityType) {
        var f = this.fields;
        var payload = {__metadata: {type: entityType}};
        payload[f.groupName] = item.groupName;
        payload[f.teamName] = item.teamName || "";
        payload[f.monthlyRows] = item.monthlyRows;
        payload[f.weeklyRows] = item.weeklyRows;
        payload[f.dailyRows] = item.dailyRows;
        payload[f.sortOrder] = item.sortOrder;
        payload[f.isActive] = item.isActive !== false;
        return payload;
    };

    OrganizationSettingsDataSource.prototype.write = function (item, isUpdate, success, failure) {
        var self = this;
        if (isUpdate && !item.etag) {
            failure("設定の更新情報を確認できません。再読込してから、もう一度操作してください。");
            return;
        }
        this.client.getEntityType(function (entityType) {
            self.client.getDigest(function (digest) {
                var path = self.client.getListPath() + "/items";
                var headers = {
                    "Content-Type": "application/json;odata=verbose",
                    "X-RequestDigest": digest
                };
                if (isUpdate) {
                    path += "(" + encodeURIComponent(item.id) + ")";
                    headers["IF-MATCH"] = item.etag;
                    headers["X-HTTP-Method"] = "MERGE";
                }
                self.client.request(
                    "POST",
                    self.client.getApiUrl(path),
                    headers,
                    JSON.stringify(self.toPayload(item, entityType)),
                    function (xhr) {
                        var data;
                        if (!isUpdate && xhr.responseText) {
                            data = window.YoteihyouUtil.getJson(xhr);
                            item.id = data.d[self.fields.id];
                            item.etag = data.d.__metadata && data.d.__metadata.etag ? data.d.__metadata.etag : "";
                        }
                        success(item);
                    },
                    failure
                );
            }, failure);
        }, failure);
    };

    OrganizationSettingsDataSource.prototype.create = function (item, success, failure) {
        this.write(item, false, success, failure);
    };

    OrganizationSettingsDataSource.prototype.update = function (item, success, failure) {
        this.write(item, true, success, failure);
    };

    OrganizationSettingsDataSource.prototype.remove = function (item, success, failure) {
        var self = this;
        if (!item.etag) {
            failure("設定の更新情報を確認できません。再読込してから、もう一度操作してください。");
            return;
        }
        this.client.getDigest(function (digest) {
            self.client.request(
                "POST",
                self.client.getApiUrl(self.client.getListPath() + "/items(" + encodeURIComponent(item.id) + ")"),
                {
                    "X-RequestDigest": digest,
                    "IF-MATCH": item.etag,
                    "X-HTTP-Method": "DELETE"
                },
                null,
                function () {
                    success(item);
                },
                failure
            );
        }, failure);
    };

    OrganizationSettingsDataSource.prototype.toOrganizationConfig = function (items, baseConfig) {
        var groups = [];
        var groupByName = {};
        var sorted = items.slice(0);
        var item;
        var group;
        var i;
        sorted.sort(function (a, b) {
            if (a.sortOrder !== b.sortOrder) {
                return a.sortOrder - b.sortOrder;
            }
            return Number(a.id || 0) - Number(b.id || 0);
        });
        for (i = 0; i < sorted.length; i += 1) {
            item = sorted[i];
            if (item.isActive === false || !item.groupName) {
                continue;
            }
            group = groupByName[item.groupName];
            if (!group) {
                group = {name: item.groupName};
                groupByName[item.groupName] = group;
                groups.push(group);
            }
            if (item.teamName) {
                if (!group.teams) {
                    group.teams = [];
                    delete group.monthlyRows;
                    delete group.weeklyRows;
                    delete group.dailyRows;
                }
                group.teams.push({
                    name: item.teamName,
                    monthlyRows: item.monthlyRows,
                    weeklyRows: item.weeklyRows,
                    dailyRows: item.dailyRows
                });
            } else if (!group.teams) {
                group.monthlyRows = item.monthlyRows;
                group.weeklyRows = item.weeklyRows;
                group.dailyRows = item.dailyRows;
            }
        }
        return {
            separator: baseConfig.separator || "／",
            metadataSeparator: baseConfig.metadataSeparator || "｜",
            targetSeparator: baseConfig.targetSeparator || "・",
            groups: groups
        };
    };

    window.YoteihyouOrganizationSettingsDataSource = OrganizationSettingsDataSource;
}(window));
