(function (window, document) {
    "use strict";

    var util = window.YoteihyouUtil;

    function byId(id) {
        return document.getElementById(id);
    }

    function addEvent(element, eventName, handler) {
        util.addEvent(element, eventName, handler);
    }

    function createCell(text) {
        var cell = document.createElement("td");
        cell.appendChild(document.createTextNode(text === null || typeof text === "undefined" ? "" : String(text)));
        return cell;
    }

    function SettingsController(options) {
        this.source = options.source;
        this.baseConfig = options.baseConfig;
        this.onApply = options.onApply;
        this.onOpen = options.onOpen;
        this.items = [];
        this.editingItem = null;
        this.loaded = false;
        this.loadRequestId = 0;
    }

    SettingsController.prototype.setStatus = function (text, isError) {
        var element = byId("settings-status");
        element.className = isError ? "settings-status error" : "settings-status";
        element.innerHTML = util.escapeHtml(text || "");
    };

    SettingsController.prototype.setBusy = function (busy) {
        byId("settings-save").disabled = busy;
        byId("settings-delete").disabled = busy;
        byId("settings-reload").disabled = busy;
    };

    SettingsController.prototype.renderItems = function () {
        var self = this;
        var body = byId("settings-body");
        var row;
        var actionCell;
        var button;
        var item;
        var i;
        while (body.firstChild) {
            body.removeChild(body.firstChild);
        }
        if (this.items.length === 0) {
            row = document.createElement("tr");
            actionCell = createCell("設定がありません。下のフォームから追加してください。");
            actionCell.colSpan = 9;
            actionCell.className = "empty-schedule";
            row.appendChild(actionCell);
            body.appendChild(row);
            return;
        }
        for (i = 0; i < this.items.length; i += 1) {
            item = this.items[i];
            row = document.createElement("tr");
            if (item.isActive === false) {
                row.className = "settings-inactive";
            }
            row.appendChild(createCell(item.groupName));
            row.appendChild(createCell(item.teamName || "－"));
            row.appendChild(createCell(item.monthlyRows));
            row.appendChild(createCell(item.weeklyRows));
            row.appendChild(createCell(item.dailyRows));
            row.appendChild(createCell(item.autoRows === false ? "固定" : "自動"));
            row.appendChild(createCell(item.sortOrder));
            row.appendChild(createCell(item.isActive === false ? "無効" : "有効"));
            actionCell = document.createElement("td");
            button = document.createElement("button");
            button.type = "button";
            button.className = "button button-small";
            button.appendChild(document.createTextNode("編集"));
            (function (targetItem) {
                button.onclick = function () {
                    self.edit(targetItem);
                };
            }(item));
            actionCell.appendChild(button);
            row.appendChild(actionCell);
            body.appendChild(row);
        }
    };

    SettingsController.prototype.clearForm = function () {
        this.editingItem = null;
        byId("setting-id").value = "";
        byId("setting-group-name").value = "";
        byId("setting-team-name").value = "";
        byId("setting-monthly-rows").value = "5";
        byId("setting-weekly-rows").value = "5";
        byId("setting-daily-rows").value = "5";
        byId("setting-auto-rows").checked = true;
        byId("setting-sort-order").value = "0";
        byId("setting-active").checked = true;
        byId("settings-form-title").innerHTML = "設定を追加";
        byId("settings-delete").style.display = "none";
    };

    SettingsController.prototype.edit = function (item) {
        this.editingItem = item;
        byId("setting-id").value = item.id;
        byId("setting-group-name").value = item.groupName;
        byId("setting-team-name").value = item.teamName;
        byId("setting-monthly-rows").value = item.monthlyRows;
        byId("setting-weekly-rows").value = item.weeklyRows;
        byId("setting-daily-rows").value = item.dailyRows;
        byId("setting-auto-rows").checked = item.autoRows !== false;
        byId("setting-sort-order").value = item.sortOrder;
        byId("setting-active").checked = item.isActive !== false;
        byId("settings-form-title").innerHTML = "設定を編集";
        byId("settings-delete").style.display = "inline-block";
        byId("setting-group-name").focus();
    };

    SettingsController.prototype.readPositiveNumber = function (id, label) {
        var value = parseInt(byId(id).value, 10);
        if (isNaN(value) || value < 1 || value > 50) {
            throw new Error(label + "は1～50で入力してください。");
        }
        return value;
    };

    SettingsController.prototype.readForm = function () {
        var item = {
            id: byId("setting-id").value,
            etag: this.editingItem ? this.editingItem.etag : "",
            groupName: util.trim(byId("setting-group-name").value),
            teamName: util.trim(byId("setting-team-name").value),
            monthlyRows: this.readPositiveNumber("setting-monthly-rows", "月間行数"),
            weeklyRows: this.readPositiveNumber("setting-weekly-rows", "週間行数"),
            dailyRows: this.readPositiveNumber("setting-daily-rows", "日々行数"),
            autoRows: byId("setting-auto-rows").checked,
            sortOrder: parseInt(byId("setting-sort-order").value, 10),
            isActive: byId("setting-active").checked
        };
        var existing;
        var i;
        if (!item.groupName) {
            throw new Error("大グループ名を入力してください。");
        }
        if (isNaN(item.sortOrder)) {
            item.sortOrder = 0;
        }
        for (i = 0; i < this.items.length; i += 1) {
            existing = this.items[i];
            if (String(existing.id) === String(item.id) || existing.groupName !== item.groupName) {
                continue;
            }
            if (existing.teamName === item.teamName) {
                throw new Error("同じ大グループと小グループの設定が既にあります。");
            }
            if (!existing.teamName || !item.teamName) {
                throw new Error("同じ大グループで「小グループなし」と小グループ設定は併用できません。");
            }
        }
        return item;
    };

    SettingsController.prototype.load = function (showMessage) {
        var self = this;
        var requestId = this.loadRequestId + 1;
        this.loadRequestId = requestId;
        if (!this.source.canConnect()) {
            this.setStatus("SharePointへ接続できないため、organizations.jsの設定を使用しています。", true);
            return;
        }
        this.setBusy(true);
        this.setStatus("SharePointから組織設定を読み込んでいます。", false);
        this.source.load(function (items) {
            if (requestId !== self.loadRequestId) {
                return;
            }
            self.items = items;
            self.loaded = true;
            self.renderItems();
            self.onApply(self.source.toOrganizationConfig(items, self.baseConfig));
            self.setStatus(showMessage ? "組織設定を再読込しました。" : "SharePointの組織設定を使用しています。", false);
            self.setBusy(false);
        }, function (message) {
            if (requestId !== self.loadRequestId) {
                return;
            }
            self.setStatus(
                message + (self.loaded ? " 現在表示中の設定を維持します。" : " organizations.jsの設定を使用します。"),
                true
            );
            self.setBusy(false);
        });
    };

    SettingsController.prototype.open = function () {
        this.onOpen();
        byId("settings-panel").style.display = "block";
        this.clearForm();
        this.renderItems();
        this.load(false);
    };

    SettingsController.prototype.close = function () {
        byId("settings-panel").style.display = "none";
        this.clearForm();
    };

    SettingsController.prototype.save = function (event) {
        var self = this;
        var item;
        var isUpdate;
        if (event && event.preventDefault) {
            event.preventDefault();
        } else if (window.event) {
            window.event.returnValue = false;
        }
        if (!this.source.canConnect()) {
            this.setStatus("SharePointへ接続できないため保存できません。", true);
            return false;
        }
        try {
            item = this.readForm();
        } catch (error) {
            this.setStatus(error.message, true);
            return false;
        }
        isUpdate = !!item.id;
        this.setBusy(true);
        this.setStatus("設定を保存しています。", false);
        (isUpdate ? this.source.update : this.source.create).call(this.source, item, function () {
            self.clearForm();
            self.load(true);
        }, function (message) {
            self.setStatus(message, true);
            self.setBusy(false);
        });
        return false;
    };

    SettingsController.prototype.remove = function () {
        var self = this;
        var item = this.editingItem;
        if (!item || !window.confirm("この組織設定を削除しますか？")) {
            return;
        }
        this.setBusy(true);
        this.setStatus("設定を削除しています。", false);
        this.source.remove(item, function () {
            self.clearForm();
            self.load(true);
        }, function (message) {
            self.setStatus(message, true);
            self.setBusy(false);
        });
    };

    SettingsController.prototype.initialize = function () {
        var self = this;
        addEvent(byId("open-settings"), "click", function () {
            self.open();
        });
        addEvent(byId("close-settings"), "click", function () {
            self.close();
        });
        addEvent(byId("settings-new"), "click", function () {
            self.clearForm();
            byId("setting-group-name").focus();
        });
        addEvent(byId("settings-cancel-edit"), "click", function () {
            self.clearForm();
        });
        addEvent(byId("settings-reload"), "click", function () {
            self.load(true);
        });
        addEvent(byId("settings-delete"), "click", function () {
            self.remove();
        });
        addEvent(byId("settings-form"), "submit", function (event) {
            return self.save(event);
        });
        addEvent(document, "keydown", function (event) {
            event = event || window.event;
            if (event.keyCode === 27 && byId("settings-panel").style.display !== "none") {
                self.close();
            }
        });
        this.clearForm();
        this.renderItems();
        if (this.source.canConnect()) {
            this.load(false);
        } else {
            this.setStatus("SharePointへ接続できないため、organizations.jsの設定を使用しています。", true);
        }
    };

    window.YoteihyouSettingsController = SettingsController;
}(window, document));
