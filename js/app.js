(function (window, document) {
    "use strict";

    var config = window.YOTEIHYOU_CONFIG;
    var organizationConfig = window.YOTEIHYOU_ORGANIZATIONS || {separator: "／", sections: []};
    var util = window.YoteihyouUtil;
    var service = new window.YoteihyouDataService(config);
    var state = {
        displayMonth: new Date(new Date().getFullYear(), new Date().getMonth(), 1),
        items: [],
        editingItem: null
    };

    function byId(id) {
        return document.getElementById(id);
    }

    function setMessage(text, isError) {
        var element = byId("message");
        element.className = isError ? "message error" : "message";
        element.innerHTML = util.escapeHtml(text || "");
        element.style.display = text ? "block" : "none";
    }

    function setConnectionStatus(text, statusClass) {
        var element = byId("connection-status");
        element.className = "connection-status" + (statusClass ? " " + statusClass : "");
        element.innerHTML = util.escapeHtml(text);
    }

    function getStoredMode() {
        var value = "";
        if (!config.rememberDataSource) {
            return "";
        }
        try {
            value = window.localStorage.getItem(config.storageKey) || "";
        } catch (ignore) {
            value = "";
        }
        return value === "csv" || value === "sharepoint" ? value : "";
    }

    function storeMode(mode) {
        if (!config.rememberDataSource) {
            return;
        }
        try {
            window.localStorage.setItem(config.storageKey, mode);
        } catch (ignore) {
            /* IE11のセキュリティ設定で保存できない場合も動作を継続します。 */
        }
    }

    function setEditorReadOnly(readOnly) {
        var formElements = byId("event-form").elements;
        var i;
        for (i = 0; i < formElements.length; i += 1) {
            if (formElements[i].id !== "cancel-edit") {
                formElements[i].disabled = readOnly;
            }
        }
        byId("csv-readonly-note").style.display = readOnly ? "block" : "none";
    }

    function updateSourceControls() {
        var mode = service.getMode();
        byId("source-csv").checked = mode === "csv";
        byId("source-sharepoint").checked = mode === "sharepoint";
        byId("new-event").disabled = service.isReadOnly();
    }

    function sameDate(a, b) {
        return a && b && a.getFullYear() === b.getFullYear() &&
            a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
    }

    function startOfDay(date) {
        return new Date(date.getFullYear(), date.getMonth(), date.getDate());
    }

    function itemOccursOn(item, day) {
        var dayTime = startOfDay(day).getTime();
        var startTime;
        var endTime;
        if (!item.startDate || item.isActive === false) {
            return false;
        }
        startTime = startOfDay(item.startDate).getTime();
        endTime = startOfDay(item.endDate || item.startDate).getTime();
        return dayTime >= startTime && dayTime <= endTime;
    }

    function compareItems(a, b) {
        var dateDifference = a.startDate.getTime() - b.startDate.getTime();
        if (dateDifference !== 0) {
            return dateDifference;
        }
        if (a.sortOrder !== b.sortOrder) {
            return a.sortOrder - b.sortOrder;
        }
        return String(a.title).localeCompare(String(b.title));
    }

    function getItemsForDay(day) {
        var result = [];
        var i;
        for (i = 0; i < state.items.length; i += 1) {
            if (itemOccursOn(state.items[i], day)) {
                result.push(state.items[i]);
            }
        }
        result.sort(compareItems);
        return result;
    }

    function getTimeLabel(item, day) {
        if (item.allDay || !sameDate(item.startDate, day)) {
            return "";
        }
        return util.pad2(item.startDate.getHours()) + ":" + util.pad2(item.startDate.getMinutes());
    }

    function createEventButton(item, day) {
        var button = document.createElement("button");
        var time = getTimeLabel(item, day);
        var timeSpan;
        var titleSpan;
        button.type = "button";
        button.className = "event-item";
        button.title = item.title + (item.purpose ? " / " + item.purpose : "") +
            (item.location ? " / " + item.location : "");
        if (time) {
            timeSpan = document.createElement("span");
            timeSpan.className = "event-time";
            timeSpan.appendChild(document.createTextNode(time));
            button.appendChild(timeSpan);
        }
        titleSpan = document.createElement("span");
        titleSpan.appendChild(document.createTextNode(item.title));
        button.appendChild(titleSpan);
        button.onclick = function () {
            openEditor(item);
        };
        return button;
    }

    function renderCalendar() {
        var body = byId("calendar-body");
        var year = state.displayMonth.getFullYear();
        var month = state.displayMonth.getMonth();
        var firstCell = new Date(year, month, 1 - new Date(year, month, 1).getDay());
        var today = new Date();
        var row;
        var cell;
        var date;
        var dayNumber;
        var dayItems;
        var rowIndex;
        var columnIndex;
        var itemIndex;

        byId("month-title").innerHTML = year + "年" + (month + 1) + "月";
        while (body.firstChild) {
            body.removeChild(body.firstChild);
        }

        for (rowIndex = 0; rowIndex < 6; rowIndex += 1) {
            row = document.createElement("tr");
            for (columnIndex = 0; columnIndex < 7; columnIndex += 1) {
                date = new Date(firstCell.getFullYear(), firstCell.getMonth(), firstCell.getDate() + (rowIndex * 7) + columnIndex);
                cell = document.createElement("td");
                cell.className = columnIndex === 0 ? "sunday" : (columnIndex === 6 ? "saturday" : "");
                if (date.getMonth() !== month) {
                    cell.className += (cell.className ? " " : "") + "other-month";
                }
                if (sameDate(date, today)) {
                    cell.className += (cell.className ? " " : "") + "today";
                }
                dayNumber = document.createElement("div");
                dayNumber.className = "day-number";
                dayNumber.appendChild(document.createTextNode(date.getDate()));
                cell.appendChild(dayNumber);
                dayItems = getItemsForDay(date);
                for (itemIndex = 0; itemIndex < dayItems.length; itemIndex += 1) {
                    cell.appendChild(createEventButton(dayItems[itemIndex], date));
                }
                row.appendChild(cell);
            }
            body.appendChild(row);
        }
    }

    function findItemById(id) {
        var i;
        for (i = 0; i < state.items.length; i += 1) {
            if (String(state.items[i].id) === String(id)) {
                return state.items[i];
            }
        }
        return null;
    }

    function clearOptions(selectElement) {
        while (selectElement.options.length > 0) {
            selectElement.remove(0);
        }
    }

    function addOption(selectElement, text, value) {
        var option = document.createElement("option");
        option.text = text;
        option.value = value;
        selectElement.add(option);
    }

    function findOrganizationSection(sectionName) {
        var sections = organizationConfig.sections || [];
        var i;
        for (i = 0; i < sections.length; i += 1) {
            if (sections[i].name === sectionName) {
                return sections[i];
            }
        }
        return null;
    }

    function loadOrganizationTeams(sectionName, selectedTeam) {
        var teamSelect = byId("organization-team");
        var section = findOrganizationSection(sectionName);
        var teams = section && section.teams ? section.teams : [];
        var i;
        clearOptions(teamSelect);
        addOption(teamSelect, "班を選択", "");
        for (i = 0; i < teams.length; i += 1) {
            addOption(teamSelect, teams[i], teams[i]);
        }
        if (selectedTeam && teams.indexOf(selectedTeam) < 0) {
            addOption(teamSelect, "設定外：" + selectedTeam, selectedTeam);
        }
        teamSelect.value = selectedTeam || "";
    }

    function loadOrganizationSections(selectedSection, selectedTeam) {
        var sectionSelect = byId("organization-section");
        var sections = organizationConfig.sections || [];
        var i;
        clearOptions(sectionSelect);
        addOption(sectionSelect, "科を選択", "");
        for (i = 0; i < sections.length; i += 1) {
            addOption(sectionSelect, sections[i].name, sections[i].name);
        }
        if (selectedSection && !findOrganizationSection(selectedSection)) {
            addOption(sectionSelect, "設定外：" + selectedSection, selectedSection);
        }
        sectionSelect.value = selectedSection || "";
        loadOrganizationTeams(sectionSelect.value, selectedTeam || "");
    }

    function splitOrganization(value) {
        var separator = organizationConfig.separator || "／";
        var text = util.trim(value);
        var position = text.indexOf(separator);
        if (position < 0) {
            return {section: text, team: ""};
        }
        return {
            section: util.trim(text.substring(0, position)),
            team: util.trim(text.substring(position + separator.length))
        };
    }

    function joinOrganization(section, team) {
        var separator = organizationConfig.separator || "／";
        if (!section) {
            return "";
        }
        return team ? section + separator + team : section;
    }

    function clearEditor() {
        byId("event-id").value = "";
        byId("event-name").value = "";
        byId("start-date").value = util.formatDateKey(new Date()) + " 09:00";
        byId("end-date").value = util.formatDateKey(new Date()) + " 10:00";
        loadOrganizationSections("", "");
        byId("category").value = "";
        byId("location").value = "";
        byId("description").value = "";
    }

    function openEditor(item) {
        var readOnly = service.isReadOnly();
        var organization;
        state.editingItem = item || null;
        clearEditor();
        if (item) {
            byId("editor-title").innerHTML = readOnly ? "予定の詳細" : "予定を編集";
            byId("event-id").value = item.id;
            byId("event-name").value = item.title;
            byId("start-date").value = util.formatDateTime(item.startDate);
            byId("end-date").value = util.formatDateTime(item.endDate);
            organization = splitOrganization(item.purpose || "");
            loadOrganizationSections(organization.section, organization.team);
            byId("category").value = item.category;
            byId("location").value = item.location;
            byId("description").value = item.description;
        } else {
            byId("editor-title").innerHTML = "予定を追加";
        }
        setEditorReadOnly(readOnly);
        byId("delete-event").style.display = item && !readOnly ? "inline-block" : "none";
        byId("event-editor").style.display = "block";
        if (!readOnly) {
            byId("event-name").focus();
        }
    }

    function closeEditor() {
        state.editingItem = null;
        byId("event-editor").style.display = "none";
        setMessage("", false);
    }

    function readForm() {
        var startDate = util.parseDate(byId("start-date").value);
        var endDate = util.parseDate(byId("end-date").value);
        var title = util.trim(byId("event-name").value);
        if (!title) {
            throw new Error("件名を入力してください。");
        }
        if (!startDate) {
            throw new Error("開始日時を「2026-09-16 09:00」の形式で入力してください。");
        }
        if (!endDate) {
            endDate = new Date(startDate.getTime());
        }
        if (endDate.getTime() < startDate.getTime()) {
            throw new Error("終了日時は開始日時以降にしてください。");
        }
        return {
            id: byId("event-id").value,
            title: title,
            startDate: startDate,
            endDate: endDate,
            allDay: false,
            category: util.trim(byId("category").value),
            location: util.trim(byId("location").value),
            description: util.trim(byId("description").value),
            purpose: joinOrganization(
                util.trim(byId("organization-section").value),
                util.trim(byId("organization-team").value)
            ),
            sortOrder: 0,
            isActive: true,
            source: "sharepoint"
        };
    }

    function reloadData(successMessage) {
        var mode = service.getMode();
        setMessage("", false);
        setConnectionStatus((mode === "csv" ? "CSV" : "SharePoint") + " 読込中", "");
        service.load(function (items) {
            state.items = items;
            renderCalendar();
            setConnectionStatus((mode === "csv" ? "試験用CSV" : "SharePoint") + " 接続済（" + items.length + "件）", "connected");
            if (successMessage) {
                setMessage(successMessage, false);
            }
        }, function (message) {
            state.items = [];
            renderCalendar();
            setConnectionStatus((mode === "csv" ? "CSV" : "SharePoint") + " 接続失敗", "error");
            setMessage(message, true);
        });
    }

    function switchMode(mode) {
        closeEditor();
        service.setMode(mode);
        storeMode(mode);
        updateSourceControls();
        reloadData("");
    }

    function saveEvent(event) {
        var item;
        var current;
        if (event && event.preventDefault) {
            event.preventDefault();
        } else if (window.event) {
            window.event.returnValue = false;
        }
        if (service.isReadOnly()) {
            setMessage("試験用CSVは閲覧専用です。", true);
            return false;
        }
        try {
            item = readForm();
        } catch (error) {
            setMessage(error.message, true);
            return false;
        }
        current = item.id ? findItemById(item.id) : null;
        setConnectionStatus("SharePoint 保存中", "");
        (current ? service.update : service.create).call(service, item, function () {
            closeEditor();
            reloadData(current ? "予定を更新しました。" : "予定を登録しました。");
        }, function (message) {
            setConnectionStatus("SharePoint 保存失敗", "error");
            setMessage(message, true);
        });
        return false;
    }

    function deleteEvent() {
        var item = state.editingItem;
        if (!item || service.isReadOnly()) {
            return;
        }
        if (!window.confirm("「" + item.title + "」を削除しますか？")) {
            return;
        }
        setConnectionStatus("SharePoint 削除中", "");
        service.remove(item, function () {
            closeEditor();
            reloadData("予定を削除しました。");
        }, function (message) {
            setConnectionStatus("SharePoint 削除失敗", "error");
            setMessage(message, true);
        });
    }

    function moveMonth(amount) {
        state.displayMonth = new Date(state.displayMonth.getFullYear(), state.displayMonth.getMonth() + amount, 1);
        renderCalendar();
    }

    function bindEvents() {
        util.addEvent(byId("source-csv"), "change", function () {
            if (this.checked) {
                switchMode("csv");
            }
        });
        util.addEvent(byId("source-sharepoint"), "change", function () {
            if (this.checked) {
                switchMode("sharepoint");
            }
        });
        util.addEvent(byId("organization-section"), "change", function () {
            loadOrganizationTeams(this.value, "");
        });
        util.addEvent(byId("reload-button"), "click", function () {
            reloadData("");
        });
        util.addEvent(byId("previous-month"), "click", function () {
            moveMonth(-1);
        });
        util.addEvent(byId("current-month"), "click", function () {
            state.displayMonth = new Date(new Date().getFullYear(), new Date().getMonth(), 1);
            renderCalendar();
        });
        util.addEvent(byId("next-month"), "click", function () {
            moveMonth(1);
        });
        util.addEvent(byId("new-event"), "click", function () {
            openEditor(null);
        });
        util.addEvent(byId("cancel-edit"), "click", closeEditor);
        util.addEvent(byId("delete-event"), "click", deleteEvent);
        util.addEvent(byId("event-form"), "submit", saveEvent);
    }

    function initialize() {
        var mode = getStoredMode() || config.defaultDataSource;
        byId("app-title").innerHTML = util.escapeHtml(config.appTitle);
        document.title = config.appTitle;
        loadOrganizationSections("", "");
        service.setMode(mode);
        updateSourceControls();
        bindEvents();
        renderCalendar();
        reloadData("");
    }

    util.addEvent(window, "load", initialize);
}(window, document));
