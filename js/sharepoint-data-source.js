(function (window) {
    "use strict";

    var util = window.YoteihyouUtil;

    function trimTrailingSlash(value) {
        return String(value || "").replace(/\/$/, "");
    }

    function inferSiteUrl() {
        var path = window.location.pathname || "";
        var lower = path.toLowerCase();
        var markers = ["/siteassets/", "/style library/", "/shared documents/", "/documents/", "/doclib/", "/_layouts/"];
        var found = -1;
        var i;
        var position;
        for (i = 0; i < markers.length; i += 1) {
            position = lower.indexOf(markers[i]);
            if (position >= 0 && (found < 0 || position < found)) {
                found = position;
            }
        }
        if (found >= 0) {
            return window.location.protocol + "//" + window.location.host + path.substring(0, found);
        }
        return "";
    }

    function SharePointDataSource(options) {
        this.options = options;
        this.fields = options.fields;
        this.siteUrl = trimTrailingSlash(options.siteUrl ||
            (window._spPageContextInfo && window._spPageContextInfo.webAbsoluteUrl) ||
            inferSiteUrl());
        this.listTitle = options.listTitle;
        this.pageSize = options.pageSize || 500;
        this.readOnly = false;
        this.entityType = "";
    }

    SharePointDataSource.prototype.getApiUrl = function (path) {
        if (!this.siteUrl) {
            throw new Error("config.js の SharePoint siteUrl を設定してください。");
        }
        return this.siteUrl + "/_api/" + path;
    };

    SharePointDataSource.prototype.getListPath = function () {
        return "web/lists/getbytitle('" + String(this.listTitle).replace(/'/g, "''") + "')";
    };

    SharePointDataSource.prototype.request = function (method, url, headers, body, success, failure) {
        var xhr = new XMLHttpRequest();
        var key;
        xhr.open(method, url, true);
        xhr.setRequestHeader("Accept", "application/json;odata=verbose");
        for (key in headers) {
            if (Object.prototype.hasOwnProperty.call(headers, key)) {
                xhr.setRequestHeader(key, headers[key]);
            }
        }
        xhr.onreadystatechange = function () {
            if (xhr.readyState !== 4) {
                return;
            }
            if ((xhr.status >= 200 && xhr.status < 300) || xhr.status === 304) {
                success(xhr);
            } else {
                failure(util.getErrorMessage(xhr, "SharePointとの通信に失敗しました"));
            }
        };
        xhr.send(body || null);
    };

    SharePointDataSource.prototype.toItem = function (row) {
        var f = this.fields;
        var startDate = row[f.startDate] ? new Date(row[f.startDate]) : null;
        var endDate = row[f.endDate] ? new Date(row[f.endDate]) : null;
        var active = row[f.isActive];
        return {
            id: row[f.id],
            title: row[f.title] || "",
            startDate: startDate,
            endDate: endDate || startDate,
            allDay: util.isTrue(row[f.allDay]),
            category: row[f.category] || "",
            location: row[f.location] || "",
            description: row[f.description] || "",
            sortOrder: parseInt(row[f.sortOrder], 10) || 0,
            isActive: active === null || typeof active === "undefined" ? true : util.isTrue(active),
            source: "sharepoint"
        };
    };

    SharePointDataSource.prototype.load = function (success, failure) {
        var self = this;
        var f = this.fields;
        var select = [f.id, f.title, f.startDate, f.endDate, f.allDay, f.category, f.location, f.description, f.sortOrder, f.isActive].join(",");
        var url;
        var items = [];

        try {
            url = this.getApiUrl(this.getListPath() + "/items?$select=" + encodeURIComponent(select) +
                "&$orderby=" + encodeURIComponent(f.startDate + " asc," + f.sortOrder + " asc") +
                "&$top=" + encodeURIComponent(this.pageSize));
        } catch (error) {
            failure(error.message);
            return;
        }

        function loadPage(pageUrl) {
            self.request("GET", pageUrl, {}, null, function (xhr) {
                var data;
                var results;
                var i;
                try {
                    data = util.getJson(xhr);
                    results = data.d && data.d.results ? data.d.results : [];
                    for (i = 0; i < results.length; i += 1) {
                        if (self.toItem(results[i]).isActive) {
                            items.push(self.toItem(results[i]));
                        }
                    }
                    if (data.d && data.d.__next) {
                        loadPage(data.d.__next);
                    } else {
                        success(items);
                    }
                } catch (error) {
                    failure("SharePointの応答を解析できませんでした。" + (error.message ? " " + error.message : ""));
                }
            }, failure);
        }

        loadPage(url);
    };

    SharePointDataSource.prototype.getDigest = function (success, failure) {
        var url;
        try {
            url = this.getApiUrl("contextinfo");
        } catch (error) {
            failure(error.message);
            return;
        }
        this.request("POST", url, {"Content-Type": "application/json;odata=verbose"}, null, function (xhr) {
            var data = util.getJson(xhr);
            success(data.d.GetContextWebInformation.FormDigestValue);
        }, failure);
    };

    SharePointDataSource.prototype.getEntityType = function (success, failure) {
        var self = this;
        var url;
        if (this.entityType) {
            success(this.entityType);
            return;
        }
        try {
            url = this.getApiUrl(this.getListPath() + "?$select=ListItemEntityTypeFullName");
        } catch (error) {
            failure(error.message);
            return;
        }
        this.request("GET", url, {}, null, function (xhr) {
            var data = util.getJson(xhr);
            self.entityType = data.d.ListItemEntityTypeFullName;
            success(self.entityType);
        }, failure);
    };

    SharePointDataSource.prototype.toPayload = function (item, entityType) {
        var f = this.fields;
        var payload = {__metadata: {type: entityType}};
        payload[f.title] = item.title;
        payload[f.startDate] = util.toIsoString(item.startDate);
        payload[f.endDate] = util.toIsoString(item.endDate || item.startDate);
        payload[f.allDay] = item.allDay === true;
        payload[f.category] = item.category || "";
        payload[f.location] = item.location || "";
        payload[f.description] = item.description || "";
        payload[f.sortOrder] = parseInt(item.sortOrder, 10) || 0;
        payload[f.isActive] = item.isActive !== false;
        return payload;
    };

    SharePointDataSource.prototype.write = function (item, isUpdate, success, failure) {
        var self = this;
        this.getEntityType(function (entityType) {
            self.getDigest(function (digest) {
                var path = self.getListPath() + "/items";
                var headers = {
                    "Content-Type": "application/json;odata=verbose",
                    "X-RequestDigest": digest
                };
                var url;
                if (isUpdate) {
                    path += "(" + encodeURIComponent(item.id) + ")";
                    headers["IF-MATCH"] = "*";
                    headers["X-HTTP-Method"] = "MERGE";
                }
                url = self.getApiUrl(path);
                self.request("POST", url, headers, JSON.stringify(self.toPayload(item, entityType)), function (xhr) {
                    var data;
                    if (isUpdate || !xhr.responseText) {
                        success(item);
                        return;
                    }
                    data = util.getJson(xhr);
                    item.id = data.d[self.fields.id];
                    success(item);
                }, failure);
            }, failure);
        }, failure);
    };

    SharePointDataSource.prototype.create = function (item, success, failure) {
        this.write(item, false, success, failure);
    };

    SharePointDataSource.prototype.update = function (item, success, failure) {
        this.write(item, true, success, failure);
    };

    SharePointDataSource.prototype.remove = function (item, success, failure) {
        var self = this;
        this.getDigest(function (digest) {
            var url = self.getApiUrl(self.getListPath() + "/items(" + encodeURIComponent(item.id) + ")");
            self.request("POST", url, {
                "X-RequestDigest": digest,
                "IF-MATCH": "*",
                "X-HTTP-Method": "DELETE"
            }, null, function () {
                success(item);
            }, failure);
        }, failure);
    };

    window.YoteihyouSharePointDataSource = SharePointDataSource;
}(window));
