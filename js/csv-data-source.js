(function (window) {
    "use strict";

    var util = window.YoteihyouUtil;

    function parseCsv(text) {
        var rows = [];
        var row = [];
        var field = "";
        var quoted = false;
        var i;
        var ch;
        var next;

        text = String(text || "").replace(/^\uFEFF/, "");
        for (i = 0; i < text.length; i += 1) {
            ch = text.charAt(i);
            next = text.charAt(i + 1);

            if (quoted) {
                if (ch === '"' && next === '"') {
                    field += '"';
                    i += 1;
                } else if (ch === '"') {
                    quoted = false;
                } else {
                    field += ch;
                }
            } else if (ch === '"') {
                quoted = true;
            } else if (ch === ",") {
                row.push(field);
                field = "";
            } else if (ch === "\r" && next === "\n") {
                row.push(field);
                rows.push(row);
                row = [];
                field = "";
                i += 1;
            } else if (ch === "\n" || ch === "\r") {
                row.push(field);
                rows.push(row);
                row = [];
                field = "";
            } else {
                field += ch;
            }
        }

        if (field !== "" || row.length > 0) {
            row.push(field);
            rows.push(row);
        }
        return rows;
    }

    function rowToObject(headers, row) {
        var result = {};
        var i;
        for (i = 0; i < headers.length; i += 1) {
            result[util.trim(headers[i])] = typeof row[i] === "undefined" ? "" : row[i];
        }
        return result;
    }

    function normalizeItem(source, rowNumber) {
        var startDate = util.parseDate(source.EventDate);
        var endDate = util.parseDate(source.EndDate);
        return {
            id: util.trim(source.ID) || "csv-" + rowNumber,
            title: util.trim(source.Title),
            startDate: startDate,
            endDate: endDate || startDate,
            allDay: false,
            category: util.trim(source.Category),
            location: util.trim(source.Location),
            description: util.trim(source.Description),
            purpose: util.trim(source.Purpose),
            sortOrder: 0,
            isActive: true,
            source: "csv"
        };
    }

    function CsvDataSource(options) {
        this.url = options.url;
        this.readOnly = true;
    }

    CsvDataSource.prototype.load = function (success, failure) {
        var xhr = new XMLHttpRequest();
        var url = this.url + (this.url.indexOf("?") >= 0 ? "&" : "?") + "_=" + new Date().getTime();
        xhr.open("GET", url, true);
        xhr.onreadystatechange = function () {
            var rows;
            var headers;
            var items = [];
            var i;
            var item;

            if (xhr.readyState !== 4) {
                return;
            }
            if (xhr.status !== 200 && xhr.status !== 0) {
                failure("試験用CSVを読み込めませんでした（HTTP " + xhr.status + "）。");
                return;
            }

            try {
                rows = parseCsv(xhr.responseText);
                if (rows.length < 1) {
                    success([]);
                    return;
                }
                headers = rows[0];
                for (i = 1; i < rows.length; i += 1) {
                    if (rows[i].length === 1 && util.trim(rows[i][0]) === "") {
                        continue;
                    }
                    item = normalizeItem(rowToObject(headers, rows[i]), i + 1);
                    if (item.title && item.startDate && item.isActive) {
                        items.push(item);
                    }
                }
                success(items);
            } catch (error) {
                failure("試験用CSVの形式を確認してください。" + (error.message ? " " + error.message : ""));
            }
        };
        xhr.send(null);
    };

    CsvDataSource.prototype.create = function (item, success, failure) {
        failure("試験用CSVは閲覧専用です。SharePointへ切り替えてください。");
    };

    CsvDataSource.prototype.update = CsvDataSource.prototype.create;
    CsvDataSource.prototype.remove = CsvDataSource.prototype.create;

    window.YoteihyouCsvDataSource = CsvDataSource;
}(window));
