(function (window, document) {
    "use strict";
    var LIMIT = 300;
    function SystemLog(options) {
        options = options || {};
        this.storage = options.storage || null;
        this.key = options.key || "yoteihyou.systemLog";
        this.entries = [];
        this.persistent = !!this.storage;
        this.onChange = function () {};
        this.sequence = 0;
        this.session = new Date().getTime() + "-" + Math.random().toString(36).slice(2);
        try {
            var saved = this.storage && JSON.parse(this.storage.getItem(this.key) || "[]"), i, entry;
            if (Array.isArray(saved)) {
                for (i = 0; i < saved.length; i += 1) {
                    entry = saved[i];
                    if (!entry || typeof entry.time !== "string" || typeof entry.target !== "string" || typeof entry.detail !== "string") { continue; }
                    if (entry.status === "pending") { entry.status = "warning"; entry.detail = "前回終了時点で処理の完了を確認できませんでした。"; }
                    this.entries.push(entry);
                }
                this.entries = this.entries.slice(-LIMIT);
            }
        } catch (ignore) { this.persistent = false; }
    }
    SystemLog.prototype.changed = function () {
        this.entries = this.entries.slice(-LIMIT);
        if (this.storage) {
            try { this.storage.setItem(this.key, JSON.stringify(this.entries)); this.persistent = true; }
            catch (ignore) { this.persistent = false; }
        }
        this.onChange();
    };
    SystemLog.prototype.add = function (target, operation, status, detail) {
        var entry = {id: this.session + ":" + (++this.sequence), time: new Date().toISOString(),
            target: String(target).slice(0, 300), operation: String(operation).slice(0, 100),
            status: status, detail: String(detail || "").slice(0, 1500)};
        this.entries.push(entry); this.changed(); return entry;
    };
    SystemLog.prototype.begin = function (target, operation) {
        var self = this, start = new Date().getTime();
        var entry = this.add(target, operation, "pending", "処理中");
        var finished = false;
        var timer = window.setTimeout(function () {
            if (finished) { return; }
            entry.status = "warning"; entry.detail = "20秒経過しても完了を確認できません。通信が継続している可能性があります。";
            self.changed();
        }, 20000);
        return function (status, detail) {
            if (finished) { return; }
            finished = true; window.clearTimeout(timer);
            entry.status = status;
            entry.detail = String(detail || "").slice(0, 1500);
            entry.duration = new Date().getTime() - start;
            self.changed();
        };
    };
    SystemLog.prototype.observe = function (object, method, target, successIndex, failureIndex, skip) {
        var self = this, original = object[method];
        object[method] = function () {
            var args = Array.prototype.slice.call(arguments), receiver = this;
            if (skip && skip()) { return original.apply(receiver, args); }
            var finish = self.begin(target, method === "load" ? "読込" : "保存・更新");
            var success = args[successIndex], failure = args[failureIndex];
            args[successIndex] = function () {
                finish("success", "完了しました。");
                if (success) { return success.apply(this, arguments); }
            };
            args[failureIndex] = function (message, status) {
                finish("error", (status ? "HTTP " + status + " / " : "") + String(message || "処理に失敗しました。"));
                if (failure) { return failure.apply(this, arguments); }
            };
            try { return original.apply(receiver, args); }
            catch (error) { finish("error", error.message); throw error; }
        };
    };
    SystemLog.prototype.probe = function (client, fields, disabled, complete) {
        var self = this, target = client.listTitle + (client.siteUrl ? " / " + client.siteUrl : "");
        if (disabled || !client.siteUrl) {
            this.add(target, "起動・接続確認", "skip", disabled ? "設定で無効のため確認対象外です。" : "SharePoint接続先が未設定のため確認対象外です。");
            complete(); return;
        }
        var finish = this.begin(target, "起動・接続確認"), completed = false;
        var timer = window.setTimeout(function () { done("warning", "接続確認が20秒以内に完了しませんでした。再確認してください。"); }, 20000);
        function done(status, detail) {
            if (completed) { return; } completed = true;
            window.clearTimeout(timer); finish(status, detail); complete();
        }
        try {
            client.request("GET", client.getApiUrl(client.getListPath() + "/items?$select=" + encodeURIComponent(fields.join(",")) + "&$top=1"), {}, null,
                function (xhr) {
                    try {
                        var data = JSON.parse(xhr.responseText);
                        if (!data.d || !Array.isArray(data.d.results)) { throw new Error("リスト応答の形式が不正です。"); }
                        done("success", "リスト・指定列の読取を確認しました。書込権限はこの確認では検証していません。");
                    } catch (error) { done("error", error.message); }
                }, function (message, status) { done("error", (status ? "HTTP " + status + " / " : "") + message); });
        } catch (error) { done("error", error.message); }
    };
    SystemLog.prototype.probeCsv = function (url, complete) {
        var finish = this.begin("試験用CSV / " + url, "起動・接続確認"), xhr = new window.XMLHttpRequest(), completed = false;
        function done(status, message) {
            if (completed) { return; } completed = true;
            window.clearTimeout(timer); finish(status, message); complete();
        }
        var timer = window.setTimeout(function () { done("error", "CSVの接続確認がタイムアウトしました。"); xhr.abort(); }, 20000);
        try {
            xhr.open("GET", url, true);
            xhr.onreadystatechange = function () {
                if (xhr.readyState !== 4) { return; }
                if (xhr.status === 200 || (xhr.status === 0 && window.location.protocol === "file:" && xhr.responseText)) {
                    done("success", "CSVを取得できました。予定の形式は通常の読込処理で検証します。");
                } else { done("error", "CSVを取得できません（HTTP " + xhr.status + "）。"); }
            };
            xhr.onerror = function () { done("error", "CSVとの通信に失敗しました。"); };
            xhr.send(null);
        } catch (error) { done("error", error.message); }
    };
    window.YoteihyouSystemLog = SystemLog;

    // Initialize before other application scripts so loading/runtime errors are visible.
    if (!document || !document.getElementById("open-system-log")) { return; }
    var storage = null;
    try { storage = window.localStorage; } catch (ignore) {}
    var log = new SystemLog({storage: storage, key: "yoteihyou.systemLog:" + window.location.pathname});
    window.yoteihyouSystemLog = log;
    function byId(id) { return document.getElementById(id); }
    function textCell(row, value) { var cell = document.createElement("td"); cell.appendChild(document.createTextNode(value)); row.appendChild(cell); }
    function render() {
        var entries = log.entries, i, entry, row, errors = 0, pending = 0;
        var labels = {pending: "処理中", success: "成功", error: "エラー", warning: "注意", skip: "対象外", info: "情報"};
        for (i = 0; i < entries.length; i += 1) { if (entries[i].status === "error") { errors += 1; } if (entries[i].status === "pending") { pending += 1; } }
        byId("open-system-log").textContent = "システムログ" + (errors ? "（エラー " + errors + "）" : pending ? "（確認中）" : "");
        byId("open-system-log").className = "button button-small system-log-button screen-only" + (errors ? " system-log-errors" : "");
        if (byId("system-log-panel").style.display === "none") { return; }
        byId("system-log-status").textContent = "最新 " + entries.length + " / " + LIMIT + "件・処理中 " + pending + "件・エラー記録 " + errors + "件。" +
            (log.persistent ? "このブラウザーに保持しています。" : "ブラウザー保存が利用できないため、現在の画面内だけに保持しています。");
        var body = byId("system-log-rows");
        while (body.firstChild) { body.removeChild(body.firstChild); }
        for (i = entries.length - 1; i >= 0; i -= 1) {
            entry = entries[i]; row = document.createElement("tr"); row.className = "system-log-" + (labels[entry.status] ? entry.status : "info");
            textCell(row, new Date(entry.time).toLocaleString()); textCell(row, labels[entry.status] || "情報");
            textCell(row, entry.target); textCell(row, entry.operation || ""); textCell(row, entry.detail);
            body.appendChild(row);
        }
    }
    log.onChange = render;
    var returnFocus;
    function close() { byId("system-log-panel").style.display = "none"; if (returnFocus) { returnFocus.focus(); } }
    byId("open-system-log").onclick = function () {
        returnFocus = document.activeElement; byId("system-log-panel").style.display = "block"; render(); byId("close-system-log").focus();
    };
    byId("close-system-log").onclick = close;
    byId("clear-system-log").onclick = function () { log.entries = []; log.changed(); };
    document.addEventListener("keydown", function (event) {
        if (byId("system-log-panel").style.display === "none") { return; }
        if (event.keyCode === 27) { close(); event.preventDefault(); event.stopImmediatePropagation(); }
        if (event.keyCode === 9) {
            var controls = byId("system-log-panel").querySelectorAll("button:not([disabled])");
            if (event.shiftKey && document.activeElement === controls[0]) { controls[controls.length - 1].focus(); event.preventDefault(); }
            else if (!event.shiftKey && document.activeElement === controls[controls.length - 1]) { controls[0].focus(); event.preventDefault(); }
        }
    }, true);
    window.addEventListener("error", function (event) {
        var element = event.target;
        if (element && element !== window && (element.tagName === "SCRIPT" || element.tagName === "LINK")) {
            log.add("アプリファイル", "読込", "error", String(element.src || element.href || "").split("?")[0]);
        } else if (event.message) { log.add("JavaScript", "実行", "error", event.message + "（行 " + (event.lineno || "不明") + "）"); }
    }, true);
    log.add("アプリ", "起動", "info", "起動しました。接続確認を開始します。");
}(window, typeof document === "undefined" ? null : document));
