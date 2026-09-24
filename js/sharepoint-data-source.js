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
        this.styleFieldsAvailable = null;
        this.textColorFieldAvailable = null;
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
            } else if (xhr.status === 412) {
                failure("ほかのユーザーがこの予定を更新しました。再読込してから、もう一度操作してください。", 412);
            } else {
                failure(util.getErrorMessage(xhr, "SharePointとの通信に失敗しました"), xhr.status);
            }
        };
        xhr.send(body || null);
    };

    SharePointDataSource.prototype.toItem = function (row) {
        var f = this.fields;
        var startDate = row[f.startDate] ? new Date(row[f.startDate]) : null;
        var endDate = row[f.endDate] ? new Date(row[f.endDate]) : null;
        var created = row[f.created || "Created"] ? new Date(row[f.created || "Created"]) : null;
        var modified = row[f.modified || "Modified"] ? new Date(row[f.modified || "Modified"]) : null;
        var author = row[f.author || "Author"];
        var editor = row[f.editor || "Editor"];
        return {
            id: row[f.id],
            etag: row.__metadata && row.__metadata.etag ? row.__metadata.etag : "",
            title: row[f.title] || "",
            startDate: startDate,
            endDate: endDate || startDate,
            allDay: f.allDay ? row[f.allDay] === true || row[f.allDay] === 1 : false,
            category: row[f.category] || "",
            location: row[f.location] || "",
            description: row[f.description] || "",
            purpose: row[f.purpose] || "",
            lineStyle: util.normalizeLineStyle(f.lineStyle ? row[f.lineStyle] : ""),
            lineColor: util.normalizeLineColor(f.lineColor ? row[f.lineColor] : ""),
            textColor: util.normalizeTextColor(f.textColor ? row[f.textColor] : ""),
            createdBy: author && author.Title ? author.Title : "",
            createdAt: created && !isNaN(created.getTime()) ? created : null,
            modifiedBy: editor && editor.Title ? editor.Title : "",
            modifiedAt: modified && !isNaN(modified.getTime()) ? modified : null,
            sortOrder: 0,
            isActive: true,
            source: "sharepoint"
        };
    };

    SharePointDataSource.prototype.load = function (range, success, failure) {
        var self = this;
        var f = this.fields;
        var authorField = f.author || "Author";
        var editorField = f.editor || "Editor";
        var baseSelect = [f.id, f.title, f.startDate, f.endDate, f.category, f.location,
            f.description, f.purpose, f.created || "Created", f.modified || "Modified",
            authorField + "/Title", editorField + "/Title"];
        var includeStyle = !!(f.lineStyle && f.lineColor && this.styleFieldsAvailable !== false);
        var includeTextColor = !!(f.textColor && this.textColorFieldAvailable !== false);
        var attempts = [{style: includeStyle, text: includeTextColor}];
        var attemptIndex = 0;
        var filter = "";
        var url;
        var items = [];
        if (f.allDay) { baseSelect.push(f.allDay); }
        if (includeTextColor) { attempts.push({style: includeStyle, text: false}); }
        if (includeStyle && includeTextColor) { attempts.push({style: false, text: true}); }
        if (includeStyle || includeTextColor) { attempts.push({style: false, text: false}); }

        function buildUrl() {
            var select = baseSelect.slice(0);
            if (attempts[attemptIndex].style) {
                select.push(f.lineStyle, f.lineColor);
            }
            if (attempts[attemptIndex].text) {
                select.push(f.textColor);
            }
            return self.getApiUrl(self.getListPath() + "/items?$select=" + encodeURIComponent(select.join(",")) +
                "&$expand=" + encodeURIComponent(authorField + "," + editorField) +
                (filter ? "&$filter=" + encodeURIComponent(filter) : "") +
                "&$orderby=" + encodeURIComponent(f.startDate + " asc") +
                "&$top=" + encodeURIComponent(self.pageSize));
        }

        try {
            if (range && range.startDate && range.endDate) {
                filter = f.startDate + " lt datetime'" + util.toIsoString(range.endDate) + "' and " +
                    f.endDate + " ge datetime'" + util.toIsoString(range.startDate) + "'";
            }
            url = buildUrl();
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
                    if (f.lineStyle && f.lineColor) {
                        self.styleFieldsAvailable = attempts[attemptIndex].style;
                    }
                    if (f.textColor) {
                        self.textColorFieldAvailable = attempts[attemptIndex].text;
                    }
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
                    failure("SharePointの応答を解析できませんでした。" + (error.message ? " " + error.message : ""));
                }
            }, function (message, status) {
                if (status === 400 && items.length === 0 && attemptIndex + 1 < attempts.length) {
                    attemptIndex += 1;
                    try {
                        loadPage(buildUrl());
                    } catch (error) {
                        failure(error.message);
                    }
                } else {
                    failure(message);
                }
            });
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
        payload[f.category] = item.category || "";
        payload[f.location] = item.location || "";
        payload[f.description] = item.description || "";
        payload[f.purpose] = item.purpose || "";
        if (f.allDay) { payload[f.allDay] = !!item.allDay; }
        if (f.lineStyle && f.lineColor && this.styleFieldsAvailable !== false &&
                (this.styleFieldsAvailable === true ||
                    util.normalizeLineStyle(item.lineStyle) !== "solid" ||
                    util.normalizeLineColor(item.lineColor) !== "default")) {
            payload[f.lineStyle] = util.normalizeLineStyle(item.lineStyle);
            payload[f.lineColor] = util.normalizeLineColor(item.lineColor);
        }
        if (f.textColor && this.textColorFieldAvailable !== false &&
                (this.textColorFieldAvailable === true ||
                    util.normalizeTextColor(item.textColor) !== "default")) {
            payload[f.textColor] = util.normalizeTextColor(item.textColor);
        }
        return payload;
    };

    SharePointDataSource.prototype.write = function (item, isUpdate, success, failure) {
        var self = this;
        if (isUpdate && !item.etag) {
            failure("予定の更新情報を確認できません。再読込してから、もう一度操作してください。");
            return;
        }
        if ((util.normalizeLineStyle(item.lineStyle) !== "solid" ||
                util.normalizeLineColor(item.lineColor) !== "default") &&
                (!this.fields.lineStyle || !this.fields.lineColor || this.styleFieldsAvailable === false)) {
            failure("線の設定を保存するには、SharePoint予定リストにLineStyle列とLineColor列を追加してください。");
            return;
        }
        if (util.normalizeTextColor(item.textColor) !== "default" &&
                (!this.fields.textColor || this.textColorFieldAvailable === false)) {
            failure("文字色を保存するには、SharePoint予定リストにTextColor列を追加してください。");
            return;
        }
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
                    headers["IF-MATCH"] = item.etag;
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
                    item.etag = data.d.__metadata && data.d.__metadata.etag ? data.d.__metadata.etag : "";
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
        if (!item.etag) {
            failure("予定の更新情報を確認できません。再読込してから、もう一度操作してください。");
            return;
        }
        this.getDigest(function (digest) {
            var url = self.getApiUrl(self.getListPath() + "/items(" + encodeURIComponent(item.id) + ")");
            self.request("POST", url, {
                "X-RequestDigest": digest,
                "IF-MATCH": item.etag,
                "X-HTTP-Method": "DELETE"
            }, null, function () {
                success(item);
            }, failure);
        }, failure);
    };

    window.YoteihyouSharePointDataSource = SharePointDataSource;
}(window));
