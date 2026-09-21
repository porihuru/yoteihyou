(function (window, document) {
    "use strict";

    var config = window.YOTEIHYOU_CONFIG;
    var dailyViewConfig = config.dailyView || {
        startHour: 6,
        endHour: 22,
        slotMinutes: 60,
        defaultStartHour: 9,
        defaultDurationMinutes: 60
    };
    var printConfig = config.print || {marginMm: 10, minimumScale: 0.65};
    var organizationConfig = window.YOTEIHYOU_ORGANIZATIONS || {separator: "／", sections: []};
    var util = window.YoteihyouUtil;
    var service = new window.YoteihyouDataService(config);
    var state = {
        viewMode: "monthly",
        displayDate: startOfDay(new Date()),
        displayMonth: new Date(new Date().getFullYear(), new Date().getMonth(), 1),
        items: [],
        editingItem: null
    };
    var printState = null;

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

    function setEditorLayoutOpen(isOpen) {
        var className = document.body.className.replace(/(^|\s)editor-open(?=\s|$)/g, " ");
        className = className.replace(/^\s+|\s+$/g, "").replace(/\s+/g, " ");
        document.body.className = isOpen ? (className ? className + " editor-open" : "editor-open") : className;
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

    function startOfWeek(date) {
        var day = startOfDay(date);
        return new Date(day.getFullYear(), day.getMonth(), day.getDate() - day.getDay());
    }

    function formatJapaneseDate(date, includeYear) {
        var weekdays = ["日", "月", "火", "水", "木", "金", "土"];
        var text = includeYear ? date.getFullYear() + "年" : "";
        return text + (date.getMonth() + 1) + "月" + date.getDate() + "日（" + weekdays[date.getDay()] + "）";
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

    function getItemsForDay(day, viewMode) {
        var result = [];
        var i;
        for (i = 0; i < state.items.length; i += 1) {
            if (itemOccursOn(state.items[i], day) && itemReflectsInView(state.items[i], viewMode)) {
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

    function stopEvent(event) {
        event = event || window.event;
        if (event.stopPropagation) {
            event.stopPropagation();
        }
        event.cancelBubble = true;
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
        button.onclick = function (event) {
            stopEvent(event);
            openEditor(item);
            return false;
        };
        return button;
    }

    function getOrganizationGroups() {
        return organizationConfig.groups || organizationConfig.sections || [];
    }

    function getOrganizationTeams(group) {
        var source = group && group.teams ? group.teams : [];
        var result = [];
        var i;
        for (i = 0; i < source.length; i += 1) {
            result.push(typeof source[i] === "string" ? {name: source[i]} : source[i]);
        }
        return result;
    }

    function getConfiguredRowCount(entry, viewMode) {
        var value = parseInt(entry ? entry[viewMode + "Rows"] : 0, 10);
        return isNaN(value) || value < 1 ? 1 : value;
    }

    function getOrganizationKey(section, team) {
        return String(section || "") + "\u001f" + String(team || "");
    }

    function getConfiguredOrganizationBlocks(viewMode) {
        var groups = getOrganizationGroups();
        var blocks = [];
        var teams;
        var i;
        var j;
        for (i = 0; i < groups.length; i += 1) {
            teams = getOrganizationTeams(groups[i]);
            if (teams.length === 0) {
                blocks.push({
                    section: groups[i].name,
                    team: "",
                    label: groups[i].name,
                    rowCount: getConfiguredRowCount(groups[i], viewMode)
                });
            } else {
                for (j = 0; j < teams.length; j += 1) {
                    blocks.push({
                        section: groups[i].name,
                        team: teams[j].name,
                        label: joinOrganization(groups[i].name, teams[j].name),
                        rowCount: getConfiguredRowCount(teams[j], viewMode)
                    });
                }
            }
        }
        return blocks;
    }

    function itemOccursInDates(item, dates) {
        var i;
        for (i = 0; i < dates.length; i += 1) {
            if (itemOccursOn(item, dates[i])) {
                return true;
            }
        }
        return false;
    }

    function getOrganizationBlocks(viewMode, dates) {
        var blocks = getConfiguredOrganizationBlocks(viewMode);
        var configured = {};
        var purpose;
        var key;
        var i;
        for (i = 0; i < blocks.length; i += 1) {
            configured[getOrganizationKey(blocks[i].section, blocks[i].team)] = true;
        }
        for (i = 0; i < state.items.length; i += 1) {
            if (!itemReflectsInView(state.items[i], viewMode) || !itemOccursInDates(state.items[i], dates)) {
                continue;
            }
            purpose = splitPurpose(state.items[i].purpose || "");
            key = getOrganizationKey(purpose.section, purpose.team);
            if (!configured[key]) {
                blocks.push({
                    section: purpose.section,
                    team: purpose.team,
                    label: joinOrganization(purpose.section, purpose.team) || "グループ未設定",
                    rowCount: 1
                });
                configured[key] = true;
            }
        }
        return blocks;
    }

    function getItemsForOrganizationDay(section, team, day, viewMode) {
        var dayItems = getItemsForDay(day, viewMode);
        var result = [];
        var purpose;
        var i;
        for (i = 0; i < dayItems.length; i += 1) {
            purpose = splitPurpose(dayItems[i].purpose || "");
            if (purpose.section === section && purpose.team === team) {
                result.push(dayItems[i]);
            }
        }
        return result;
    }

    function createHeaderCell(text, className) {
        var cell = document.createElement("th");
        cell.className = className || "";
        cell.appendChild(document.createTextNode(text));
        return cell;
    }

    function renderOrganizationHeader(head, dates, viewMode) {
        var weekdays = ["日", "月", "火", "水", "木", "金", "土"];
        var row = document.createElement("tr");
        var label;
        var className;
        var i;
        while (head.firstChild) {
            head.removeChild(head.firstChild);
        }
        row.appendChild(createHeaderCell("グループ", "organization-column"));
        for (i = 0; i < dates.length; i += 1) {
            if (viewMode === "daily") {
                label = "予定（空白をクリックして入力）";
            } else if (viewMode === "weekly") {
                label = (dates[i].getMonth() + 1) + "/" + dates[i].getDate() + "（" + weekdays[dates[i].getDay()] + "）";
            } else {
                label = dates[i].getDate() + "（" + weekdays[dates[i].getDay()] + "）";
            }
            className = dates[i].getDay() === 0 ? "sunday" : (dates[i].getDay() === 6 ? "saturday" : "");
            row.appendChild(createHeaderCell(label, className));
        }
        head.appendChild(row);
    }

    function renderOrganizationSchedule(viewMode, dates, head, body) {
        var blocks = getOrganizationBlocks(viewMode, dates);
        var itemsByDate;
        var rowCount;
        var row;
        var groupCell;
        var scheduleCell;
        var date;
        var item;
        var i;
        var j;
        var rowIndex;
        renderOrganizationHeader(head, dates, viewMode);
        while (body.firstChild) {
            body.removeChild(body.firstChild);
        }
        for (i = 0; i < blocks.length; i += 1) {
            itemsByDate = [];
            rowCount = blocks[i].rowCount;
            for (j = 0; j < dates.length; j += 1) {
                itemsByDate[j] = getItemsForOrganizationDay(blocks[i].section, blocks[i].team, dates[j], viewMode);
                rowCount = Math.max(rowCount, itemsByDate[j].length);
            }
            for (rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
                row = document.createElement("tr");
                if (rowIndex === 0) {
                    groupCell = document.createElement("th");
                    groupCell.className = "organization-name";
                    groupCell.rowSpan = rowCount;
                    groupCell.appendChild(document.createTextNode(blocks[i].label));
                    row.appendChild(groupCell);
                }
                for (j = 0; j < dates.length; j += 1) {
                    date = new Date(dates[j].getTime());
                    item = itemsByDate[j][rowIndex] || null;
                    scheduleCell = document.createElement("td");
                    scheduleCell.className = "organization-schedule-cell clickable-date";
                    scheduleCell.title = formatJapaneseDate(date, true) + "の" + blocks[i].label + "に予定を追加";
                    if (item) {
                        scheduleCell.appendChild(createEventButton(item, date));
                    } else {
                        scheduleCell.appendChild(document.createTextNode("\u00a0"));
                    }
                    (function (cell, targetDate, section, team) {
                        cell.onclick = function () {
                            openEditor(null, targetDate, false, {section: section, team: team});
                        };
                    }(scheduleCell, date, blocks[i].section, blocks[i].team));
                    row.appendChild(scheduleCell);
                }
                body.appendChild(row);
            }
        }
    }

    function renderCalendar() {
        var year = state.displayMonth.getFullYear();
        var month = state.displayMonth.getMonth();
        var lastDate = new Date(year, month + 1, 0).getDate();
        var dates = [];
        var i;
        byId("month-title").innerHTML = year + "年" + (month + 1) + "月";
        byId("print-heading").innerHTML = year + "年" + (month + 1) + "月　月間予定表";
        for (i = 1; i <= lastDate; i += 1) {
            dates.push(new Date(year, month, i));
        }
        renderOrganizationSchedule("monthly", dates, byId("monthly-head"), byId("calendar-body"));
    }

    function createTextCell(text, className) {
        var cell = document.createElement("td");
        if (className) {
            cell.className = className;
        }
        cell.appendChild(document.createTextNode(text || ""));
        return cell;
    }

    function createEventTitleButton(item) {
        var button = document.createElement("button");
        button.type = "button";
        button.className = "event-detail-button";
        button.appendChild(document.createTextNode(item.title));
        button.onclick = function (event) {
            stopEvent(event);
            openEditor(item);
            return false;
        };
        return button;
    }

    function getDailyTimeRange(item, day) {
        var startText = sameDate(item.startDate, day) ?
            util.pad2(item.startDate.getHours()) + ":" + util.pad2(item.startDate.getMinutes()) : "継続";
        var endDate = item.endDate || item.startDate;
        var endText = sameDate(endDate, day) ?
            util.pad2(endDate.getHours()) + ":" + util.pad2(endDate.getMinutes()) : "継続";
        return startText + "～" + endText;
    }

    function getOrganizationFromItem(item) {
        var purpose = splitPurpose(item.purpose || "");
        return joinOrganization(purpose.section, purpose.team);
    }

    function renderDaily() {
        var day = state.displayDate;
        byId("month-title").innerHTML = formatJapaneseDate(day, true);
        byId("print-heading").innerHTML = formatJapaneseDate(day, true) + "　日々予定表";
        renderOrganizationSchedule("daily", [new Date(day.getTime())], byId("daily-head"), byId("daily-body"));
    }

    function renderWeekly() {
        var firstDay = startOfWeek(state.displayDate);
        var lastDay = new Date(firstDay.getFullYear(), firstDay.getMonth(), firstDay.getDate() + 6);
        var dates = [];
        var i;
        byId("month-title").innerHTML = (firstDay.getMonth() + 1) + "月" + firstDay.getDate() + "日～" +
            (lastDay.getMonth() + 1) + "月" + lastDay.getDate() + "日";
        byId("print-heading").innerHTML = firstDay.getFullYear() + "年" +
            (firstDay.getMonth() + 1) + "月" + firstDay.getDate() + "日～" +
            (lastDay.getMonth() + 1) + "月" + lastDay.getDate() + "日　週間予定表";
        for (i = 0; i < 7; i += 1) {
            dates.push(new Date(firstDay.getFullYear(), firstDay.getMonth(), firstDay.getDate() + i));
        }
        renderOrganizationSchedule("weekly", dates, byId("weekly-head"), byId("weekly-body"));
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
        var sections = getOrganizationGroups();
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
        var teams = getOrganizationTeams(section);
        var i;
        clearOptions(teamSelect);
        addOption(teamSelect, teams.length > 0 ? "小グループを選択" : "小グループなし", "");
        for (i = 0; i < teams.length; i += 1) {
            addOption(teamSelect, teams[i].name, teams[i].name);
        }
        if (selectedTeam && !findOrganizationTeam(section, selectedTeam)) {
            addOption(teamSelect, "設定外：" + selectedTeam, selectedTeam);
        }
        teamSelect.value = selectedTeam || "";
    }

    function findOrganizationTeam(section, teamName) {
        var teams = getOrganizationTeams(section);
        var i;
        for (i = 0; i < teams.length; i += 1) {
            if (teams[i].name === teamName) {
                return teams[i];
            }
        }
        return null;
    }

    function loadOrganizationSections(selectedSection, selectedTeam) {
        var sectionSelect = byId("organization-section");
        var sections = getOrganizationGroups();
        var i;
        clearOptions(sectionSelect);
        addOption(sectionSelect, "大グループを選択", "");
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

    function getAllTargetKeys() {
        return ["daily", "weekly", "monthly"];
    }

    function getTargetLabel(targetKey) {
        if (targetKey === "daily") {
            return "日々";
        }
        if (targetKey === "weekly") {
            return "週間";
        }
        return "月間";
    }

    function getTargetKey(label) {
        if (label === "日々" || label === "daily") {
            return "daily";
        }
        if (label === "週間" || label === "weekly") {
            return "weekly";
        }
        if (label === "月間" || label === "monthly") {
            return "monthly";
        }
        return "";
    }

    function splitPurpose(value) {
        var metadataSeparator = organizationConfig.metadataSeparator || "｜";
        var targetSeparator = organizationConfig.targetSeparator || "・";
        var text = util.trim(value);
        var position = text.indexOf(metadataSeparator);
        var organizationText = position >= 0 ? text.substring(0, position) : text;
        var targetText = position >= 0 ? text.substring(position + metadataSeparator.length) : "";
        var organization = splitOrganization(organizationText);
        var labels;
        var targets = [];
        var key;
        var i;

        if (targetText) {
            labels = targetText.split(targetSeparator);
            for (i = 0; i < labels.length; i += 1) {
                key = getTargetKey(util.trim(labels[i]));
                if (key && targets.indexOf(key) < 0) {
                    targets.push(key);
                }
            }
        }
        if (targets.length === 0) {
            targets = getAllTargetKeys();
        }
        return {
            section: organization.section,
            team: organization.team,
            targets: targets
        };
    }

    function joinPurpose(section, team, targets) {
        var organization = joinOrganization(section, team);
        var metadataSeparator = organizationConfig.metadataSeparator || "｜";
        var targetSeparator = organizationConfig.targetSeparator || "・";
        var labels = [];
        var i;
        for (i = 0; i < targets.length; i += 1) {
            labels.push(getTargetLabel(targets[i]));
        }
        return organization + metadataSeparator + labels.join(targetSeparator);
    }

    function itemReflectsInView(item, viewMode) {
        var purpose = splitPurpose(item.purpose || "");
        return purpose.targets.indexOf(viewMode) >= 0;
    }

    function getSelectedTargets() {
        var targets = [];
        if (byId("target-daily").checked) {
            targets.push("daily");
        }
        if (byId("target-weekly").checked) {
            targets.push("weekly");
        }
        if (byId("target-monthly").checked) {
            targets.push("monthly");
        }
        return targets;
    }

    function setSelectedTargets(targets) {
        byId("target-daily").checked = targets.indexOf("daily") >= 0;
        byId("target-weekly").checked = targets.indexOf("weekly") >= 0;
        byId("target-monthly").checked = targets.indexOf("monthly") >= 0;
    }

    function clearEditor(selectedDate, useSelectedTime, selectedOrganization) {
        var date = selectedDate ? new Date(selectedDate.getTime()) : startOfDay(state.displayDate || new Date());
        var durationMinutes = parseInt(dailyViewConfig.defaultDurationMinutes, 10) || 60;
        var endDate;
        if (!useSelectedTime) {
            date.setHours(parseInt(dailyViewConfig.defaultStartHour, 10) || 9, 0, 0, 0);
        }
        endDate = new Date(date.getTime() + durationMinutes * 60000);
        byId("event-id").value = "";
        byId("event-name").value = "";
        byId("start-date").value = util.formatDateTime(date);
        byId("end-date").value = util.formatDateTime(endDate);
        setSelectedTargets(getAllTargetKeys());
        loadOrganizationSections(
            selectedOrganization ? selectedOrganization.section : "",
            selectedOrganization ? selectedOrganization.team : ""
        );
        byId("category").value = "";
        byId("location").value = "";
        byId("description").value = "";
    }

    function openEditor(item, selectedDate, useSelectedTime, selectedOrganization) {
        var readOnly = service.isReadOnly();
        var purpose;
        if (!item && readOnly) {
            setMessage("試験用CSVは閲覧専用です。新規入力はSharePointへ切り替えてください。", true);
            return;
        }
        state.editingItem = item || null;
        clearEditor(selectedDate, useSelectedTime === true, selectedOrganization || null);
        if (item) {
            byId("editor-title").innerHTML = readOnly ? "予定の詳細" : "予定を編集";
            byId("event-id").value = item.id;
            byId("event-name").value = item.title;
            byId("start-date").value = util.formatDateTime(item.startDate);
            byId("end-date").value = util.formatDateTime(item.endDate);
            purpose = splitPurpose(item.purpose || "");
            setSelectedTargets(purpose.targets);
            loadOrganizationSections(purpose.section, purpose.team);
            byId("category").value = item.category;
            byId("location").value = item.location;
            byId("description").value = item.description;
        } else {
            byId("editor-title").innerHTML = "予定を追加";
        }
        setEditorReadOnly(readOnly);
        byId("delete-event").style.display = item && !readOnly ? "inline-block" : "none";
        byId("event-editor").style.display = "block";
        setEditorLayoutOpen(true);
        if (!readOnly) {
            byId("event-name").focus();
        }
    }

    function closeEditor() {
        state.editingItem = null;
        byId("event-editor").style.display = "none";
        setEditorLayoutOpen(false);
        setMessage("", false);
    }

    function readForm() {
        var startDate = util.parseDate(byId("start-date").value);
        var endDate = util.parseDate(byId("end-date").value);
        var title = util.trim(byId("event-name").value);
        var targets = getSelectedTargets();
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
        if (targets.length === 0) {
            throw new Error("反映先を一つ以上選択してください。");
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
            purpose: joinPurpose(
                util.trim(byId("organization-section").value),
                util.trim(byId("organization-team").value),
                targets
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
            renderCurrentView();
            setConnectionStatus((mode === "csv" ? "試験用CSV" : "SharePoint") + " 接続済（" + items.length + "件）", "connected");
            if (successMessage) {
                setMessage(successMessage, false);
            }
        }, function (message) {
            state.items = [];
            renderCurrentView();
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

    function updatePrintPageStyle() {
        var style = byId("print-page-style");
        var size = state.viewMode === "monthly" ? "A3 landscape" : "A4 portrait";
        var margin = parseFloat(printConfig.marginMm) || 10;
        var cssText = "@media print { @page { size: " + size + "; margin: " + margin + "mm; } }";
        if (style.styleSheet) {
            style.styleSheet.cssText = cssText;
        } else {
            style.innerHTML = cssText;
        }
    }

    function getActiveViewElement() {
        if (state.viewMode === "daily") {
            return byId("daily-view");
        }
        if (state.viewMode === "weekly") {
            return byId("weekly-view");
        }
        return byId("monthly-view");
    }

    function resetPrintLayout() {
        var i;
        if (!printState) {
            return;
        }
        printState.area.style.width = printState.areaWidth;
        printState.area.style.zoom = printState.areaZoom;
        if (printState.calendar) {
            printState.calendar.style.minWidth = printState.calendarMinWidth;
        }
        for (i = 0; i < printState.screenOnly.length; i += 1) {
            printState.screenOnly[i].element.style.display = printState.screenOnly[i].display;
        }
        printState = null;
    }

    function preparePrintLayout() {
        var area = byId("print-area");
        var activeView = getActiveViewElement();
        var calendar = activeView.getElementsByTagName("table")[0];
        var screenOnlyElements = activeView.getElementsByClassName("screen-only");
        var screenOnly = [];
        var pageWidthMm = state.viewMode === "monthly" ? 420 : 210;
        var pageHeightMm = 297;
        var marginMm = parseFloat(printConfig.marginMm) || 10;
        var pixelsPerMm = 96 / 25.4;
        var availableWidth = (pageWidthMm - marginMm * 2) * pixelsPerMm * 0.98;
        var availableHeight = (pageHeightMm - marginMm * 2) * pixelsPerMm * 0.98;
        var contentWidth;
        var contentHeight;
        var requiredScale;
        var minimumScale = parseFloat(printConfig.minimumScale) || 0.65;
        var i;

        resetPrintLayout();
        printState = {
            area: area,
            areaWidth: area.style.width,
            areaZoom: area.style.zoom,
            calendar: calendar,
            calendarMinWidth: calendar ? calendar.style.minWidth : "",
            screenOnly: screenOnly
        };

        for (i = 0; i < screenOnlyElements.length; i += 1) {
            screenOnly.push({element: screenOnlyElements[i], display: screenOnlyElements[i].style.display});
            screenOnlyElements[i].style.display = "none";
        }
        area.style.zoom = "1";
        area.style.width = Math.floor(availableWidth) + "px";
        if (calendar) {
            calendar.style.minWidth = "0";
        }

        contentWidth = Math.max(activeView.scrollWidth, area.scrollWidth);
        contentHeight = activeView.scrollHeight + 60;
        requiredScale = Math.min(1, availableWidth / contentWidth, availableHeight / contentHeight) * 0.98;

        if (requiredScale < minimumScale) {
            resetPrintLayout();
            window.alert(
                "予定が多いため、1枚に収まりません。\n" +
                "1枚に収めるには約" + Math.floor(requiredScale * 100) + "%まで縮小する必要があります。\n" +
                "表示する予定を減らすか、内容を短くしてから印刷してください。"
            );
            return false;
        }

        area.style.zoom = String(Math.floor(requiredScale * 100) / 100);
        return true;
    }

    function printCurrentView() {
        closeEditor();
        updatePrintPageStyle();
        if (!preparePrintLayout()) {
            return;
        }
        window.print();
        window.setTimeout(resetPrintLayout, 0);
    }

    function updateViewControls() {
        var mode = state.viewMode;
        byId("daily-view").style.display = mode === "daily" ? "block" : "none";
        byId("weekly-view").style.display = mode === "weekly" ? "block" : "none";
        byId("monthly-view").style.display = mode === "monthly" ? "block" : "none";
        byId("view-daily").className = "button" + (mode === "daily" ? " active" : "");
        byId("view-weekly").className = "button" + (mode === "weekly" ? " active" : "");
        byId("view-monthly").className = "button" + (mode === "monthly" ? " active" : "");
        byId("previous-month").innerHTML = mode === "daily" ? "前日" : (mode === "weekly" ? "前週" : "前月");
        byId("current-month").innerHTML = mode === "daily" ? "今日" : (mode === "weekly" ? "今週" : "今月");
        byId("next-month").innerHTML = mode === "daily" ? "翌日" : (mode === "weekly" ? "翌週" : "翌月");
        updatePrintPageStyle();
    }

    function renderCurrentView() {
        updateViewControls();
        if (state.viewMode === "daily") {
            renderDaily();
        } else if (state.viewMode === "weekly") {
            renderWeekly();
        } else {
            renderCalendar();
        }
    }

    function setViewMode(mode) {
        if (mode !== "daily" && mode !== "weekly" && mode !== "monthly") {
            return;
        }
        closeEditor();
        state.viewMode = mode;
        if (mode === "monthly") {
            state.displayMonth = new Date(state.displayDate.getFullYear(), state.displayDate.getMonth(), 1);
        }
        renderCurrentView();
    }

    function moveView(amount) {
        if (state.viewMode === "daily") {
            state.displayDate = new Date(state.displayDate.getFullYear(), state.displayDate.getMonth(), state.displayDate.getDate() + amount);
        } else if (state.viewMode === "weekly") {
            state.displayDate = new Date(state.displayDate.getFullYear(), state.displayDate.getMonth(), state.displayDate.getDate() + (amount * 7));
        } else {
            state.displayMonth = new Date(state.displayMonth.getFullYear(), state.displayMonth.getMonth() + amount, 1);
            state.displayDate = new Date(state.displayMonth.getTime());
        }
        renderCurrentView();
    }

    function goCurrentPeriod() {
        var today = startOfDay(new Date());
        state.displayDate = today;
        state.displayMonth = new Date(today.getFullYear(), today.getMonth(), 1);
        renderCurrentView();
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
        util.addEvent(byId("view-daily"), "click", function () {
            setViewMode("daily");
        });
        util.addEvent(byId("view-weekly"), "click", function () {
            setViewMode("weekly");
        });
        util.addEvent(byId("view-monthly"), "click", function () {
            setViewMode("monthly");
        });
        util.addEvent(byId("previous-month"), "click", function () {
            moveView(-1);
        });
        util.addEvent(byId("current-month"), "click", function () {
            goCurrentPeriod();
        });
        util.addEvent(byId("next-month"), "click", function () {
            moveView(1);
        });
        util.addEvent(byId("print-view"), "click", function () {
            printCurrentView();
        });
        util.addEvent(byId("new-event"), "click", function () {
            openEditor(null, state.displayDate);
        });
        util.addEvent(document, "keydown", function (event) {
            event = event || window.event;
            if (event.ctrlKey && event.keyCode === 80) {
                if (event.preventDefault) {
                    event.preventDefault();
                }
                event.returnValue = false;
                printCurrentView();
                return false;
            }
            if (event.keyCode === 27 && byId("event-editor").style.display !== "none") {
                closeEditor();
            }
        });
        util.addEvent(window, "afterprint", resetPrintLayout);
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
        renderCurrentView();
        reloadData("");
    }

    util.addEvent(window, "load", initialize);
}(window, document));
