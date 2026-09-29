(function (window) {
    "use strict";

    // Saved rows hold IDs, not dates or copies of schedule items.
    function packLegacy(entries, savedRows) {
        var rows = [], placed = {}, byId = {}, result = {}, i, j, row, entry, base;
        function fits(index, value) {
            var occupants = rows[index] || [], k;
            for (k = 0; k < occupants.length; k += 1) {
                if (value.start < occupants[k].end && occupants[k].start < value.end) { return false; }
            }
            return true;
        }
        function place(value, index) {
            if (!rows[index]) { rows[index] = []; }
            rows[index].push(value);
            placed[String(value.id)] = true;
            result[String(value.id)] = index;
        }
        for (i = 0; i < entries.length; i += 1) { byId[String(entries[i].id)] = entries[i]; }
        for (i = 0; i < (savedRows || []).length; i += 1) {
            base = rows.length;
            for (j = 0; j < savedRows[i].length; j += 1) {
                entry = byId[String(savedRows[i][j])];
                if (!entry || placed[String(entry.id)]) { continue; }
                row = base;
                while (!fits(row, entry)) { row += 1; }
                place(entry, row);
            }
        }
        for (i = 0; i < entries.length; i += 1) {
            entry = entries[i];
            if (placed[String(entry.id)]) { continue; }
            row = 0;
            while (!fits(row, entry)) { row += 1; }
            place(entry, row);
        }
        return {lanes: result, count: rows.length};
    }

    function overlaps(a, b) {
        return a.start < b.end && b.start < a.end;
    }

    function fits(entries, positions, entry, row) {
        var i, other;
        for (i = 0; i < entries.length; i += 1) {
            other = entries[i];
            if (String(other.id) !== String(entry.id) && positions[String(other.id)] === row && overlaps(entry, other)) { return false; }
        }
        return true;
    }

    function pack(entries, saved) {
        var positions = {}, pending = [], desired = saved && saved.positions || {}, count = 0, i, entry, row;
        // Old records retain their rendered positions until the next drag saves v2.
        if (Array.isArray(saved)) { return packLegacy(entries, saved); }
        for (i = 0; i < entries.length; i += 1) {
            entry = entries[i]; row = desired[String(entry.id)];
            if (typeof row === "number" && fits(entries, positions, entry, row)) {
                positions[String(entry.id)] = row;
                count = Math.max(count, row + 1);
            } else { pending.push(entry); }
        }
        // Reserve unaffected positions first; edited/new items use remaining space.
        for (i = 0; i < pending.length; i += 1) {
            entry = pending[i]; row = 0;
            while (!fits(entries, positions, entry, row)) { row += 1; }
            positions[String(entry.id)] = row;
            count = Math.max(count, row + 1);
        }
        return {lanes: positions, count: count};
    }

    function move(entries, current, id, destination) {
        var positions = {}, conflicts = [], moved, origin, i, entry, row;
        id = String(id);
        for (i = 0; i < entries.length; i += 1) {
            entry = entries[i];
            positions[String(entry.id)] = current[String(entry.id)];
            if (String(entry.id) === id) { moved = entry; }
        }
        if (!moved) { return {version: 2, positions: positions}; }
        origin = positions[id];
        positions[id] = destination;
        for (i = 0; i < entries.length; i += 1) {
            entry = entries[i];
            if (String(entry.id) !== id && positions[String(entry.id)] === destination && overlaps(moved, entry)) {
                conflicts.push(entry);
                delete positions[String(entry.id)];
            }
        }
        // Only direct collisions move. Try the vacated row, then another free row.
        // Never push unrelated occupants to make space.
        for (i = 0; i < conflicts.length; i += 1) {
            entry = conflicts[i]; row = origin;
            if (!fits(entries, positions, entry, row)) {
                row = 0;
                while (!fits(entries, positions, entry, row)) { row += 1; }
            }
            positions[String(entry.id)] = row;
        }
        return {version: 2, positions: positions};
    }

    function validate(orders) {
        var key, i, j, value, id, row;
        if (!orders || typeof orders !== "object" || Array.isArray(orders)) { throw new Error("配置データが不正です。"); }
        for (key in orders) {
            if (!Object.prototype.hasOwnProperty.call(orders, key)) { continue; }
            value = orders[key];
            if (!Array.isArray(value)) {
                if (!value || value.version !== 2 || !value.positions || typeof value.positions !== "object" || Array.isArray(value.positions)) {
                    throw new Error("配置データが不正です。");
                }
                for (id in value.positions) {
                    if (!Object.prototype.hasOwnProperty.call(value.positions, id)) { continue; }
                    row = value.positions[id];
                    if (typeof row !== "number" || !isFinite(row) || row < 0 || row !== Math.floor(row) || row > 10000) {
                        throw new Error("配置データが不正です。");
                    }
                }
                continue;
            }
            for (i = 0; i < orders[key].length; i += 1) {
                if (!Array.isArray(orders[key][i])) { throw new Error("配置データが不正です。"); }
                for (j = 0; j < orders[key][i].length; j += 1) {
                    if (typeof orders[key][i][j] !== "string") { throw new Error("配置データが不正です。"); }
                }
            }
        }
        return orders;
    }

    function Source(options) {
        this.client = new window.YoteihyouSharePointDataSource({siteUrl: options.siteUrl,
            listTitle: (options.layoutSettings || {}).listTitle || "予定表配置設定", fields: {}});
    }
    Source.prototype.load = function (key, success, failure) {
        var client = this.client;
        client.request("GET", client.getApiUrl(client.getListPath() +
            "/items?$select=ID,Title,LayoutJson&$filter=" + encodeURIComponent("Title eq '" + key.replace(/'/g, "''") + "'") + "&$top=2"),
        {}, null, function (xhr) {
            var rows, record;
            try {
                rows = window.YoteihyouUtil.getJson(xhr).d.results;
                if (rows.length > 1) { throw new Error("配置設定のTitleが重複しています。"); }
                record = rows[0];
                success(record ? {key: key, id: record.ID, etag: record.__metadata.etag,
                    orders: validate(JSON.parse(record.LayoutJson || "{}"))} : {key: key, orders: {}});
            } catch (error) { failure("配置設定を読み込めませんでした。" + error.message); }
        }, failure);
    };
    Source.prototype.save = function (record, orders, success, failure) {
        var client = this.client;
        if (record.id && !record.etag) { failure("配置設定を再読込してください。"); return; }
        try { validate(orders); } catch (error) { failure(error.message); return; }
        client.getEntityType(function (entityType) {
            client.getDigest(function (digest) {
                var headers = {"Content-Type": "application/json;odata=verbose", "X-RequestDigest": digest};
                if (record.id) { headers["IF-MATCH"] = record.etag; headers["X-HTTP-Method"] = "MERGE"; }
                client.request("POST", client.getApiUrl(client.getListPath() +
                    (record.id ? "/items(" + encodeURIComponent(record.id) + ")" : "/items")), headers,
                JSON.stringify({__metadata: {type: entityType}, Title: record.key, LayoutJson: JSON.stringify(orders)}),
                success, failure);
            }, failure);
        }, failure);
    };
    window.YoteihyouScheduleLayout = {pack: pack, move: move, Source: Source};
}(window));
