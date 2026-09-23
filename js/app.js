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
    var organizationSettingsSource = new window.YoteihyouOrganizationSettingsDataSource(config.sharePoint);
    var settingsController = null;
    var state = {
        viewMode: "monthly",
        displayDate: startOfDay(new Date()),
        displayMonth: new Date(new Date().getFullYear(), new Date().getMonth(), 1),
        items: [],
        editingItem: null
    };
    var printState = null;
    var loadRequestId = 0;
    var DAILY_SNAP_MINUTES = 15;
    var dailyInteraction = {
        selectedItemId: "",
        clipboard: null,
        target: null,
        drag: null,
        indicator: null,
        suppressItemId: ""
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
        byId("csv-unsaved-note").style.display = service.getMode() === "csv" ? "block" : "none";
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

    function preventEvent(event) {
        event = event || window.event;
        stopEvent(event);
        if (event.preventDefault) {
            event.preventDefault();
        }
        event.returnValue = false;
        return false;
    }

    function hasClass(element, className) {
        return element && (" " + element.className + " ").indexOf(" " + className + " ") >= 0;
    }

    function addClass(element, className) {
        if (element && !hasClass(element, className)) {
            element.className += (element.className ? " " : "") + className;
        }
    }

    function removeClass(element, className) {
        if (element) {
            element.className = (" " + element.className + " ")
                .replace(new RegExp("\\s" + className + "(?=\\s)", "g"), " ")
                .replace(/^\s+|\s+$/g, "").replace(/\s+/g, " ");
        }
    }

    function findParentByClass(element, className) {
        while (element && element !== document.body) {
            if (hasClass(element, className)) {
                return element;
            }
            element = element.parentNode;
        }
        return null;
    }

    function cloneScheduleItem(item) {
        return {
            id: item.id,
            etag: item.etag || "",
            title: item.title,
            startDate: new Date(item.startDate.getTime()),
            endDate: new Date((item.endDate || item.startDate).getTime()),
            allDay: !!item.allDay,
            category: item.category || "",
            location: item.location || "",
            description: item.description || "",
            purpose: item.purpose || "",
            sortOrder: item.sortOrder || 0,
            isActive: item.isActive !== false,
            source: service.getMode()
        };
    }

    function getSelectedDailyItem() {
        return dailyInteraction.selectedItemId ? findItemById(dailyInteraction.selectedItemId) : null;
    }

    function selectDailyItem(item, button) {
        var selected = document.getElementsByClassName("daily-event-selected");
        while (selected.length > 0) {
            removeClass(selected[0], "daily-event-selected");
        }
        dailyInteraction.selectedItemId = item ? String(item.id) : "";
        if (button) {
            addClass(button, "daily-event-selected");
            dailyInteraction.target = getItemTarget(item, findParentByClass(button, "daily-timeline-cell"));
        }
    }

    function hideDailyContextMenu() {
        byId("daily-context-menu").style.display = "none";
    }

    function getDailyTarget(cell, clientX, allowRangeEnd) {
        var meta = cell && cell._dailyMeta;
        var rect;
        var ratio;
        var rawMinute;
        var minute;
        var maximum;
        if (!meta) {
            return null;
        }
        rect = meta.timeline.getBoundingClientRect();
        ratio = rect.width > 0 ? (clientX - rect.left) / rect.width : 0;
        ratio = Math.max(0, Math.min(1, ratio));
        rawMinute = meta.rangeStart + ratio * (meta.rangeEnd - meta.rangeStart);
        minute = Math.round(rawMinute / DAILY_SNAP_MINUTES) * DAILY_SNAP_MINUTES;
        maximum = allowRangeEnd ? meta.rangeEnd : meta.rangeEnd - DAILY_SNAP_MINUTES;
        minute = Math.max(meta.rangeStart, Math.min(maximum, minute));
        return {
            cell: cell,
            timeline: meta.timeline,
            section: meta.section,
            team: meta.team,
            day: new Date(meta.day.getTime()),
            rangeStart: meta.rangeStart,
            rangeEnd: meta.rangeEnd,
            minute: minute
        };
    }

    function getItemTarget(item, cell) {
        var purpose = splitPurpose(item.purpose || "");
        var meta = cell && cell._dailyMeta;
        var minute = item.startDate.getHours() * 60 + item.startDate.getMinutes();
        return {
            cell: cell || null,
            timeline: meta ? meta.timeline : null,
            section: purpose.section,
            team: purpose.team,
            day: startOfDay(state.displayDate),
            rangeStart: meta ? meta.rangeStart : 0,
            rangeEnd: meta ? meta.rangeEnd : 24 * 60,
            minute: minute
        };
    }

    function showDailyContextMenu(event, item, cell) {
        var menu = byId("daily-context-menu");
        var documentWidth = document.documentElement.clientWidth;
        var documentHeight = document.documentElement.clientHeight;
        var left;
        var top;
        if (item) {
            selectDailyItem(item, findParentByClass(event.srcElement || event.target, "daily-event-bar"));
        }
        if (cell) {
            dailyInteraction.target = getDailyTarget(cell, event.clientX, false);
        } else if (item) {
            dailyInteraction.target = getItemTarget(item, findParentByClass(event.srcElement || event.target, "daily-timeline-cell"));
        }
        byId("daily-menu-copy").disabled = !getSelectedDailyItem();
        byId("daily-menu-cut").disabled = !getSelectedDailyItem() || service.isReadOnly();
        byId("daily-menu-paste").disabled = !dailyInteraction.clipboard || !dailyInteraction.target || service.isReadOnly();
        menu.style.display = "block";
        left = Math.max(0, Math.min(event.clientX, documentWidth - menu.offsetWidth - 4));
        top = Math.max(0, Math.min(event.clientY, documentHeight - menu.offsetHeight - 4));
        menu.style.left = left + "px";
        menu.style.top = top + "px";
        return preventEvent(event);
    }

    function copySelectedDailyItem(cut) {
        var item = getSelectedDailyItem();
        if (!item) {
            setMessage("コピーまたは切り取りする予定を選択してください。", true);
            return false;
        }
        if (cut && service.isReadOnly()) {
            setMessage("現在の接続先では予定を移動できません。", true);
            return false;
        }
        dailyInteraction.clipboard = {
            item: cloneScheduleItem(item),
            cut: cut === true,
            sourceId: String(item.id)
        };
        renderCurrentView();
        setMessage("「" + item.title + "」を" + (cut ? "切り取りました。貼り付け先を右クリックしてください。" : "コピーしました。貼り付け先を右クリックしてください。"), false);
        return true;
    }

    function persistDailyItem(item, createNew, successMessage, completed) {
        var mode = service.getMode();
        var sourceLabel = mode === "csv" ? "CSV" : "SharePoint";
        if (service.isReadOnly()) {
            setMessage("現在の接続先では予定を変更できません。", true);
            return;
        }
        setConnectionStatus(sourceLabel + " 保存中", "");
        (createNew ? service.create : service.update).call(service, item, function (savedItem, items) {
            if (completed) {
                completed(savedItem);
            }
            if (mode === "csv") {
                state.items = items;
                renderCurrentView();
                setConnectionStatus("試験用CSV 編集中（未保存）", "connected");
                setMessage(successMessage + " 再読込すると変更は消えます。", false);
            } else {
                reloadData(successMessage);
            }
        }, function (message) {
            setConnectionStatus(sourceLabel + " 保存失敗", "error");
            setMessage(message, true);
        });
    }

    function moveItemToTarget(item, target) {
        var changed = cloneScheduleItem(item);
        var purpose = splitPurpose(changed.purpose || "");
        var duration = changed.endDate.getTime() - changed.startDate.getTime();
        var daySpan;
        changed.purpose = joinPurpose(target.section, target.team, purpose.targets);
        if (changed.allDay) {
            daySpan = Math.max(0, Math.round((startOfDay(changed.endDate).getTime() -
                startOfDay(changed.startDate).getTime()) / 86400000));
            changed.startDate = new Date(target.day.getFullYear(), target.day.getMonth(), target.day.getDate());
            changed.endDate = new Date(target.day.getFullYear(), target.day.getMonth(), target.day.getDate() + daySpan, 23, 59);
        } else {
            changed.startDate = new Date(
                target.day.getFullYear(), target.day.getMonth(), target.day.getDate(),
                Math.floor(target.minute / 60), target.minute % 60, 0, 0
            );
            changed.endDate = new Date(changed.startDate.getTime() + Math.max(duration, DAILY_SNAP_MINUTES * 60000));
        }
        return changed;
    }

    function pasteDailyItem() {
        var clipboard = dailyInteraction.clipboard;
        var source;
        var changed;
        if (!clipboard || !dailyInteraction.target) {
            setMessage("コピー元と貼り付け先を指定してください。", true);
            return false;
        }
        source = clipboard.cut ? (findItemById(clipboard.sourceId) || clipboard.item) : clipboard.item;
        if (!source) {
            setMessage("切り取った予定が見つかりません。もう一度操作してください。", true);
            return false;
        }
        changed = moveItemToTarget(source, dailyInteraction.target);
        if (clipboard.cut) {
            persistDailyItem(changed, false, "予定を移動しました。", function () {
                dailyInteraction.clipboard = null;
                dailyInteraction.selectedItemId = String(changed.id);
            });
        } else {
            changed.id = "";
            changed.etag = "";
            persistDailyItem(changed, true, "予定を貼り付けました。", null);
        }
        hideDailyContextMenu();
        return true;
    }

    function getResizedDailyItem(item, edge, target) {
        var changed = cloneScheduleItem(item);
        var candidate = new Date(
            target.day.getFullYear(), target.day.getMonth(), target.day.getDate(),
            Math.floor(target.minute / 60), target.minute % 60, 0, 0
        );
        if (edge === "start") {
            if (candidate.getTime() > changed.endDate.getTime() - DAILY_SNAP_MINUTES * 60000) {
                throw new Error("開始時刻は終了時刻の15分以上前にしてください。");
            }
            changed.startDate = candidate;
        } else {
            if (candidate.getTime() < changed.startDate.getTime() + DAILY_SNAP_MINUTES * 60000) {
                throw new Error("終了時刻は開始時刻の15分以上後にしてください。");
            }
            changed.endDate = candidate;
        }
        return changed;
    }

    function resizeDailyItem(item, edge, target) {
        var changed;
        try {
            changed = getResizedDailyItem(item, edge, target);
        } catch (error) {
            setMessage(error.message, true);
            return;
        }
        persistDailyItem(changed, false, edge === "start" ? "開始時刻を変更しました。" : "終了時刻を変更しました。", null);
    }

    function showDailyDropIndicator(target) {
        var position;
        if (dailyInteraction.indicator && dailyInteraction.indicator.parentNode) {
            dailyInteraction.indicator.parentNode.removeChild(dailyInteraction.indicator);
        }
        dailyInteraction.indicator = document.createElement("span");
        dailyInteraction.indicator.className = "daily-drop-indicator";
        position = (target.minute - target.rangeStart) / (target.rangeEnd - target.rangeStart) * 100;
        dailyInteraction.indicator.style.left = position + "%";
        target.timeline.appendChild(dailyInteraction.indicator);
    }

    function beginDailyDrag(event, item, button, mode) {
        if (service.isReadOnly() || (item.allDay && mode !== "move")) {
            return;
        }
        selectDailyItem(item, button);
        hideDailyContextMenu();
        dailyInteraction.drag = {
            item: item,
            button: button,
            cell: findParentByClass(button, "daily-timeline-cell"),
            mode: mode,
            startX: event.clientX,
            startY: event.clientY,
            moved: false,
            target: null
        };
        return preventEvent(event);
    }

    function handleDailyDragMove(event) {
        var drag = dailyInteraction.drag;
        var cell;
        var target;
        if (!drag) {
            return;
        }
        event = event || window.event;
        if (!drag.moved && Math.abs(event.clientX - drag.startX) + Math.abs(event.clientY - drag.startY) < 4) {
            return;
        }
        drag.moved = true;
        addClass(drag.button, "daily-event-dragging");
        cell = drag.mode === "move" ? findParentByClass(document.elementFromPoint(event.clientX, event.clientY), "daily-timeline-cell") : drag.cell;
        target = getDailyTarget(cell, event.clientX, drag.mode !== "move");
        if (target) {
            drag.target = target;
            showDailyDropIndicator(target);
        } else {
            drag.target = null;
            if (dailyInteraction.indicator && dailyInteraction.indicator.parentNode) {
                dailyInteraction.indicator.parentNode.removeChild(dailyInteraction.indicator);
            }
            dailyInteraction.indicator = null;
        }
        return preventEvent(event);
    }

    function handleDailyDragEnd(event) {
        var drag = dailyInteraction.drag;
        var changed;
        if (!drag) {
            return;
        }
        if (dailyInteraction.indicator && dailyInteraction.indicator.parentNode) {
            dailyInteraction.indicator.parentNode.removeChild(dailyInteraction.indicator);
        }
        dailyInteraction.indicator = null;
        removeClass(drag.button, "daily-event-dragging");
        dailyInteraction.drag = null;
        if (!drag.moved || !drag.target) {
            return;
        }
        dailyInteraction.suppressItemId = String(drag.item.id);
        if (drag.mode === "move") {
            changed = moveItemToTarget(drag.item, drag.target);
            persistDailyItem(changed, false, "予定を移動しました。", null);
        } else {
            resizeDailyItem(drag.item, drag.mode, drag.target);
        }
        return preventEvent(event);
    }

    function isFormInput(element) {
        var tagName = element && element.tagName ? element.tagName.toLowerCase() : "";
        return tagName === "input" || tagName === "textarea" || tagName === "select" ||
            (element && element.isContentEditable);
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

    function formatDailyTime(minutes) {
        return util.pad2(Math.floor(minutes / 60)) + util.pad2(minutes % 60);
    }

    function getDailyCaptionText(item) {
        return item.title + (item.location ? "（" + item.location + "）" : "");
    }

    function estimateDailyCaptionPixels(item) {
        var text = getDailyCaptionText(item);
        var width = 16;
        var code;
        var i;
        for (i = 0; i < text.length; i += 1) {
            code = text.charCodeAt(i);
            width += code <= 255 ? (text.charAt(i) === " " ? 4 : 7) : 13;
        }
        return Math.max(120, width);
    }

    function getDailyItemRange(item, day, rangeStart, rangeEnd) {
        if (item.allDay) { return {start: rangeStart, end: rangeEnd}; }
        var endDate = item.endDate || item.startDate;
        var startMinutes = sameDate(item.startDate, day) ?
            item.startDate.getHours() * 60 + item.startDate.getMinutes() : rangeStart;
        var endMinutes = sameDate(endDate, day) ?
            endDate.getHours() * 60 + endDate.getMinutes() : rangeEnd;
        startMinutes = Math.max(rangeStart, Math.min(rangeEnd, startMinutes));
        endMinutes = Math.max(startMinutes, Math.min(rangeEnd, endMinutes));
        return {start: startMinutes, end: endMinutes};
    }

    function appendDailyGridLines(container, rangeStart, rangeEnd, slotMinutes) {
        var totalMinutes = rangeEnd - rangeStart;
        var minute;
        var line;
        for (minute = rangeStart; minute <= rangeEnd; minute += slotMinutes) {
            line = document.createElement("span");
            line.className = "daily-grid-line";
            line.style.left = ((minute - rangeStart) / totalMinutes * 100) + "%";
            container.appendChild(line);
        }
    }

    function createDailyEventBar(item, itemRange, displayRange, rangeStart, rangeEnd, laneIndex) {
        var totalMinutes = rangeEnd - rangeStart;
        var left = (displayRange.start - rangeStart) / totalMinutes * 100;
        var width = (displayRange.end - displayRange.start) / totalMinutes * 100;
        var displayDuration = displayRange.end - displayRange.start;
        var lineLeft = (itemRange.start - displayRange.start) / displayDuration * 100;
        var lineWidth = (itemRange.end - itemRange.start) / displayDuration * 100;
        var lineEnd = lineLeft + lineWidth;
        var button = document.createElement("button");
        var times = document.createElement("span");
        var start = document.createElement("span");
        var end = document.createElement("span");
        var line = document.createElement("span");
        var caption = document.createElement("span");
        var startHandle;
        var endHandle;

        button.type = "button";
        button.className = "daily-event-bar";
        if (displayRange.start !== itemRange.start || displayRange.end !== itemRange.end) {
            button.className += " daily-event-short";
        }
        if (dailyInteraction.selectedItemId === String(item.id)) {
            button.className += " daily-event-selected";
        }
        if (dailyInteraction.clipboard && dailyInteraction.clipboard.cut &&
                dailyInteraction.clipboard.sourceId === String(item.id)) {
            button.className += " daily-event-cut";
        }
        button.style.left = left + "%";
        button.style.width = width + "%";
        button.style.top = (laneIndex * 52 + 5) + "px";
        button.title = item.title + " / " + util.formatDateTime(item.startDate) + "～" +
            util.formatDateTime(item.endDate || item.startDate) + (item.location ? " / " + item.location : "");

        times.className = "daily-event-times";
        start.className = "daily-event-start";
        start.style.left = lineLeft + "%";
        start.appendChild(document.createTextNode(formatDailyTime(itemRange.start)));
        end.className = "daily-event-end";
        end.style.right = (100 - lineEnd) + "%";
        end.appendChild(document.createTextNode(formatDailyTime(itemRange.end)));
        if (lineWidth < 45) {
            if (itemRange.start > rangeStart) {
                start.className += " daily-event-time-before";
            }
            if (itemRange.end < rangeEnd) {
                end.className += " daily-event-time-after";
                end.style.right = "auto";
                end.style.left = lineEnd + "%";
            }
        }
        times.appendChild(start);
        times.appendChild(end);

        line.className = "daily-event-line";
        // Text may need extra room, but the line must follow the actual time axis.
        line.style.left = lineLeft + "%";
        line.style.width = lineWidth + "%";
        caption.className = "daily-event-caption";
        caption.appendChild(document.createTextNode(getDailyCaptionText(item)));

        button.appendChild(times);
        button.appendChild(line);
        button.appendChild(caption);
        if (item.allDay) {
            button.removeChild(times);
            button.title = item.title + " / 終日" + (item.location ? " / " + item.location : "");
        } else {
            startHandle = document.createElement("span");
            startHandle.className = "daily-resize-handle daily-resize-start screen-only";
            startHandle.style.left = lineLeft + "%";
            startHandle.title = "ドラッグして開始時刻を変更";
            endHandle = document.createElement("span");
            endHandle.className = "daily-resize-handle daily-resize-end screen-only";
            endHandle.style.left = lineEnd + "%";
            endHandle.title = "ドラッグして終了時刻を変更";
            button.appendChild(startHandle);
            button.appendChild(endHandle);
        }
        button.onclick = function (event) {
            var source = (event || window.event).srcElement || (event || window.event).target;
            stopEvent(event);
            if (hasClass(source, "daily-resize-handle")) {
                return false;
            }
            if (dailyInteraction.suppressItemId === String(item.id)) {
                dailyInteraction.suppressItemId = "";
                return false;
            }
            selectDailyItem(item, button);
            openEditor(item);
            return false;
        };
        button.onfocus = function () {
            selectDailyItem(item, button);
        };
        button.onmousedown = function (event) {
            var source = (event || window.event).srcElement || (event || window.event).target;
            if (hasClass(source, "daily-resize-start")) {
                return beginDailyDrag(event || window.event, item, button, "start");
            }
            if (hasClass(source, "daily-resize-end")) {
                return beginDailyDrag(event || window.event, item, button, "end");
            }
            if (hasClass(source, "daily-event-caption") || source === button) {
                return beginDailyDrag(event || window.event, item, button, "move");
            }
        };
        button.oncontextmenu = function (event) {
            return showDailyContextMenu(event || window.event, item, findParentByClass(button, "daily-timeline-cell"));
        };
        return button;
    }

    function renderDaily() {
        var head = byId("daily-head");
        var body = byId("daily-body");
        var table = byId("daily-view").getElementsByTagName("table")[0];
        var day = state.displayDate;
        var items = getItemsForDay(day, "daily");
        var slotMinutes = parseInt(dailyViewConfig.slotMinutes, 10);
        var startHour = parseInt(dailyViewConfig.startHour, 10);
        var endHour = parseInt(dailyViewConfig.endHour, 10);
        var rangeStart;
        var rangeEnd;
        var totalMinutes;
        var blocks;
        var blockItems;
        var placedItems;
        var laneEnds;
        var rowCount;
        var itemRange;
        var visualDuration;
        var minimumVisualDuration;
        var minimumVisualPixels;
        var maximumCaptionPixels = 120;
        var visualStart;
        var visualEnd;
        var itemMinutes;
        var endDate;
        var laneIndex;
        var headerRow;
        var headerCell;
        var axis;
        var label;
        var row;
        var groupCell;
        var timelineCell;
        var timeline;
        var minute;
        var i;
        var j;

        if (isNaN(slotMinutes) || slotMinutes < 1 || slotMinutes > 24 * 60) {
            slotMinutes = 60;
        }
        if (isNaN(startHour) || startHour < 0 || startHour > 23) {
            startHour = 6;
        }
        if (isNaN(endHour) || endHour < 1 || endHour > 24) {
            endHour = 22;
        }
        if (endHour <= startHour) {
            startHour = 6;
            endHour = 22;
        }
        rangeStart = startHour * 60;
        rangeEnd = endHour * 60;

        while (body.firstChild) {
            body.removeChild(body.firstChild);
        }
        byId("month-title").innerHTML = formatJapaneseDate(day, true);
        byId("print-heading").innerHTML = formatJapaneseDate(day, true) + "　日々予定表";

        for (i = 0; i < items.length; i += 1) {
            maximumCaptionPixels = Math.max(maximumCaptionPixels, estimateDailyCaptionPixels(items[i]));
            if (items[i].allDay) { continue; }
            if (sameDate(items[i].startDate, day)) {
                itemMinutes = items[i].startDate.getHours() * 60 + items[i].startDate.getMinutes();
                rangeStart = Math.min(rangeStart, Math.floor(itemMinutes / slotMinutes) * slotMinutes);
            }
            endDate = items[i].endDate || items[i].startDate;
            if (sameDate(endDate, day)) {
                itemMinutes = endDate.getHours() * 60 + endDate.getMinutes();
                rangeEnd = Math.max(rangeEnd, Math.ceil(itemMinutes / slotMinutes) * slotMinutes);
            }
        }

        rangeStart = Math.max(0, rangeStart);
        rangeEnd = Math.min(24 * 60, Math.max(rangeStart + slotMinutes, rangeEnd));
        totalMinutes = rangeEnd - rangeStart;

        while (head.firstChild) {
            head.removeChild(head.firstChild);
        }
        headerRow = document.createElement("tr");
        headerRow.appendChild(createHeaderCell("グループ", "organization-column"));
        headerCell = document.createElement("th");
        headerCell.className = "daily-timeline-header";
        axis = document.createElement("div");
        axis.className = "daily-timeline-axis";
        for (minute = rangeStart; minute <= rangeEnd; minute += slotMinutes) {
            label = document.createElement("span");
            label.className = "daily-axis-label" +
                (minute === rangeStart ? " first" : (minute === rangeEnd ? " last" : ""));
            label.style.left = ((minute - rangeStart) / totalMinutes * 100) + "%";
            label.appendChild(document.createTextNode(formatDailyTime(minute)));
            axis.appendChild(label);
        }
        headerCell.appendChild(axis);
        headerRow.appendChild(headerCell);
        head.appendChild(headerRow);
        table.style.minWidth = (155 + Math.max(
            Math.ceil(totalMinutes / slotMinutes) * 88,
            maximumCaptionPixels
        )) + "px";

        blocks = getOrganizationBlocks("daily", [day]);
        for (i = 0; i < blocks.length; i += 1) {
            blockItems = getItemsForOrganizationDay(blocks[i].section, blocks[i].team, day, "daily");
            placedItems = [];
            laneEnds = [];
            for (j = 0; j < blockItems.length; j += 1) {
                itemRange = getDailyItemRange(blockItems[j], day, rangeStart, rangeEnd);
                minimumVisualPixels = estimateDailyCaptionPixels(blockItems[j]);
                minimumVisualDuration = Math.min(
                    totalMinutes,
                    Math.ceil(slotMinutes * minimumVisualPixels / 88)
                );
                visualDuration = Math.max(itemRange.end - itemRange.start, minimumVisualDuration);
                visualStart = itemRange.start - (visualDuration - (itemRange.end - itemRange.start)) / 2;
                visualEnd = visualStart + visualDuration;
                if (visualStart < rangeStart) {
                    visualEnd += rangeStart - visualStart;
                    visualStart = rangeStart;
                }
                if (visualEnd > rangeEnd) {
                    visualStart = Math.max(rangeStart, visualStart - (visualEnd - rangeEnd));
                    visualEnd = rangeEnd;
                }
                laneIndex = 0;
                while (laneIndex < laneEnds.length && visualStart < laneEnds[laneIndex]) {
                    laneIndex += 1;
                }
                laneEnds[laneIndex] = visualEnd;
                placedItems.push({
                    item: blockItems[j],
                    range: itemRange,
                    displayRange: {start: visualStart, end: visualEnd},
                    lane: laneIndex
                });
            }
            rowCount = Math.max(blocks[i].rowCount, laneEnds.length, 1);

            row = document.createElement("tr");
            groupCell = document.createElement("th");
            groupCell.className = "organization-name daily-organization-name";
            groupCell.appendChild(document.createTextNode(blocks[i].label));
            row.appendChild(groupCell);

            timelineCell = document.createElement("td");
            timelineCell.className = "daily-timeline-cell clickable-slot";
            timelineCell.title = blocks[i].label + "の時間帯をクリックして予定を追加。右クリックで貼り付け";
            timeline = document.createElement("div");
            timeline.className = "daily-timeline";
            timeline.style.height = (rowCount * 52 + 8) + "px";
            timelineCell._dailyMeta = {
                timeline: timeline,
                section: blocks[i].section,
                team: blocks[i].team,
                day: new Date(day.getTime()),
                rangeStart: rangeStart,
                rangeEnd: rangeEnd
            };
            appendDailyGridLines(timeline, rangeStart, rangeEnd, slotMinutes);
            for (j = 0; j < placedItems.length; j += 1) {
                timeline.appendChild(createDailyEventBar(
                    placedItems[j].item,
                    placedItems[j].range,
                    placedItems[j].displayRange,
                    rangeStart,
                    rangeEnd,
                    placedItems[j].lane
                ));
            }
            timelineCell.appendChild(timeline);
            (function (cell, section, team, targetDay) {
                cell.onclick = function (event) {
                    var sourceEvent = event || window.event;
                    var target = getDailyTarget(cell, sourceEvent.clientX, false);
                    var targetDate;
                    hideDailyContextMenu();
                    dailyInteraction.target = target;
                    targetDate = new Date(
                        targetDay.getFullYear(),
                        targetDay.getMonth(),
                        targetDay.getDate(),
                        Math.floor(target.minute / 60),
                        target.minute % 60,
                        0,
                        0
                    );
                    openEditor(null, targetDate, true, {section: section, team: team});
                };
                cell.oncontextmenu = function (event) {
                    selectDailyItem(null, null);
                    return showDailyContextMenu(event || window.event, null, cell);
                };
            }(timelineCell, blocks[i].section, blocks[i].team, day));
            row.appendChild(timelineCell);
            body.appendChild(row);
        }
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
        window.YoteihyouDateTimeEditor.set("start", date);
        window.YoteihyouDateTimeEditor.set("end", endDate);
        byId("all-day").checked = false;
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
            setMessage("現在の接続先は閲覧専用です。", true);
            return;
        }
        state.editingItem = item || null;
        clearEditor(selectedDate, useSelectedTime === true, selectedOrganization || null);
        if (item) {
            byId("editor-title").innerHTML = readOnly ? "予定の詳細" : "予定を編集";
            byId("event-id").value = item.id;
            byId("event-name").value = item.title;
            window.YoteihyouDateTimeEditor.set("start", item.startDate);
            window.YoteihyouDateTimeEditor.set("end", item.endDate || item.startDate);
            byId("all-day").checked = !!item.allDay;
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
        window.YoteihyouDateTimeEditor.sync();
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
        var allDay = byId("all-day").checked;
        var startDate = window.YoteihyouDateTimeEditor.read("start", allDay);
        var endDate = window.YoteihyouDateTimeEditor.read("end", allDay);
        var title = util.trim(byId("event-name").value);
        var targets = getSelectedTargets();
        if (!title) {
            throw new Error("件名を入力してください。");
        }
        if (!startDate) {
            throw new Error("開始日と時刻を正しく入力してください（日付: 2026-09-16、時刻: 0800）。");
        }
        if (!endDate) {
            throw new Error("終了日と時刻を正しく入力してください（日付: 2026-09-16、時刻: 1700）。");
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
            allDay: allDay,
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
            source: service.getMode()
        };
    }

    function getCurrentLoadRange() {
        var startDate;
        var endDate;
        if (state.viewMode === "daily") {
            startDate = startOfDay(state.displayDate);
            endDate = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate() + 1);
        } else if (state.viewMode === "weekly") {
            startDate = startOfWeek(state.displayDate);
            endDate = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate() + 7);
        } else {
            startDate = new Date(state.displayMonth.getFullYear(), state.displayMonth.getMonth(), 1);
            endDate = new Date(startDate.getFullYear(), startDate.getMonth() + 1, 1);
        }
        return {startDate: startDate, endDate: endDate};
    }

    function reloadData(successMessage) {
        var mode = service.getMode();
        var requestId = loadRequestId + 1;
        var range = getCurrentLoadRange();
        loadRequestId = requestId;
        setMessage("", false);
        setConnectionStatus((mode === "csv" ? "CSV" : "SharePoint") + " 読込中", "");
        service.load(range, function (items) {
            if (requestId !== loadRequestId) {
                return;
            }
            state.items = items;
            renderCurrentView();
            setConnectionStatus((mode === "csv" ? "試験用CSV" : "SharePoint") + " 接続済（" + items.length + "件）", "connected");
            if (successMessage) {
                setMessage(successMessage, false);
            }
        }, function (message) {
            if (requestId !== loadRequestId) {
                return;
            }
            state.items = [];
            renderCurrentView();
            setConnectionStatus((mode === "csv" ? "CSV" : "SharePoint") + " 接続失敗", "error");
            setMessage(message, true);
        });
    }

    function switchMode(mode) {
        closeEditor();
        dailyInteraction.selectedItemId = "";
        dailyInteraction.clipboard = null;
        dailyInteraction.target = null;
        hideDailyContextMenu();
        service.setMode(mode);
        storeMode(mode);
        updateSourceControls();
        reloadData("");
    }

    function saveEvent(event) {
        var item;
        var current;
        var mode = service.getMode();
        var sourceLabel = mode === "csv" ? "CSV" : "SharePoint";
        if (event && event.preventDefault) {
            event.preventDefault();
        } else if (window.event) {
            window.event.returnValue = false;
        }
        if (service.isReadOnly()) {
            setMessage("現在の接続先は閲覧専用です。", true);
            return false;
        }
        try {
            item = readForm();
        } catch (error) {
            setMessage(error.message, true);
            return false;
        }
        current = item.id ? findItemById(item.id) : null;
        if (current) {
            item.etag = current.etag || "";
        }
        setConnectionStatus(sourceLabel + " 保存中", "");
        (current ? service.update : service.create).call(service, item, function (savedItem, items) {
            closeEditor();
            if (mode === "csv") {
                state.items = items;
                renderCurrentView();
                setConnectionStatus("試験用CSV 編集中（未保存）", "connected");
                setMessage(
                    current ? "予定を画面上で更新しました。再読込すると変更は消えます。" :
                        "予定を画面上で追加しました。再読込すると変更は消えます。",
                    false
                );
            } else {
                reloadData(current ? "予定を更新しました。" : "予定を登録しました。");
            }
        }, function (message) {
            setConnectionStatus(sourceLabel + " 保存失敗", "error");
            setMessage(message, true);
        });
        return false;
    }

    function deleteEvent() {
        var item = state.editingItem;
        var mode = service.getMode();
        var sourceLabel = mode === "csv" ? "CSV" : "SharePoint";
        if (!item || service.isReadOnly()) {
            return;
        }
        if (!window.confirm("「" + item.title + "」を削除しますか？")) {
            return;
        }
        setConnectionStatus(sourceLabel + " 削除中", "");
        service.remove(item, function (removedItem, items) {
            closeEditor();
            if (mode === "csv") {
                state.items = items;
                renderCurrentView();
                setConnectionStatus("試験用CSV 編集中（未保存）", "connected");
                setMessage("予定を画面上で削除しました。再読込すると元に戻ります。", false);
            } else {
                reloadData("予定を削除しました。");
            }
        }, function (message) {
            setConnectionStatus(sourceLabel + " 削除失敗", "error");
            setMessage(message, true);
        });
    }

    function updatePrintPageStyle() {
        var style = byId("print-page-style");
        var size = state.viewMode === "weekly" ? "A4 portrait" : "A3 landscape";
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
        var pageWidthMm = state.viewMode === "weekly" ? 210 : 420;
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
        hideDailyContextMenu();
        dailyInteraction.target = null;
        if (mode !== "daily") {
            dailyInteraction.selectedItemId = "";
        }
        state.viewMode = mode;
        if (mode === "monthly") {
            state.displayMonth = new Date(state.displayDate.getFullYear(), state.displayDate.getMonth(), 1);
        }
        renderCurrentView();
        if (service.getMode() === "sharepoint") {
            reloadData("");
        }
    }

    function moveView(amount) {
        hideDailyContextMenu();
        dailyInteraction.target = null;
        dailyInteraction.selectedItemId = "";
        if (state.viewMode === "daily") {
            state.displayDate = new Date(state.displayDate.getFullYear(), state.displayDate.getMonth(), state.displayDate.getDate() + amount);
        } else if (state.viewMode === "weekly") {
            state.displayDate = new Date(state.displayDate.getFullYear(), state.displayDate.getMonth(), state.displayDate.getDate() + (amount * 7));
        } else {
            state.displayMonth = new Date(state.displayMonth.getFullYear(), state.displayMonth.getMonth() + amount, 1);
            state.displayDate = new Date(state.displayMonth.getTime());
        }
        renderCurrentView();
        if (service.getMode() === "sharepoint") {
            reloadData("");
        }
    }

    function goCurrentPeriod() {
        var today = startOfDay(new Date());
        hideDailyContextMenu();
        dailyInteraction.target = null;
        dailyInteraction.selectedItemId = "";
        state.displayDate = today;
        state.displayMonth = new Date(today.getFullYear(), today.getMonth(), 1);
        renderCurrentView();
        if (service.getMode() === "sharepoint") {
            reloadData("");
        }
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
        util.addEvent(byId("daily-menu-copy"), "click", function () {
            copySelectedDailyItem(false);
            hideDailyContextMenu();
        });
        util.addEvent(byId("daily-menu-cut"), "click", function () {
            copySelectedDailyItem(true);
            hideDailyContextMenu();
        });
        util.addEvent(byId("daily-menu-paste"), "click", function () {
            pasteDailyItem();
        });
        util.addEvent(document, "mousemove", handleDailyDragMove);
        util.addEvent(document, "mouseup", handleDailyDragEnd);
        util.addEvent(document, "click", function (event) {
            var target = (event || window.event).srcElement || (event || window.event).target;
            if (!findParentByClass(target, "daily-context-menu")) {
                hideDailyContextMenu();
            }
        });
        util.addEvent(document, "keydown", function (event) {
            var handled = false;
            event = event || window.event;
            if (event.ctrlKey && !event.altKey && !isFormInput(event.srcElement || event.target) && state.viewMode === "daily") {
                if (event.keyCode === 67) {
                    handled = copySelectedDailyItem(false);
                } else if (event.keyCode === 88) {
                    handled = copySelectedDailyItem(true);
                } else if (event.keyCode === 86) {
                    handled = pasteDailyItem();
                }
                if (handled) {
                    return preventEvent(event);
                }
            }
            if (event.ctrlKey && event.keyCode === 80) {
                if (event.preventDefault) {
                    event.preventDefault();
                }
                event.returnValue = false;
                printCurrentView();
                return false;
            }
            if (event.keyCode === 27) {
                hideDailyContextMenu();
                if (dailyInteraction.drag) {
                    removeClass(dailyInteraction.drag.button, "daily-event-dragging");
                    dailyInteraction.drag = null;
                    if (dailyInteraction.indicator && dailyInteraction.indicator.parentNode) {
                        dailyInteraction.indicator.parentNode.removeChild(dailyInteraction.indicator);
                    }
                    dailyInteraction.indicator = null;
                }
                if (byId("event-editor").style.display !== "none") {
                    closeEditor();
                }
            }
        });
        util.addEvent(window, "scroll", hideDailyContextMenu);
        util.addEvent(window, "afterprint", resetPrintLayout);
        util.addEvent(byId("cancel-edit"), "click", closeEditor);
        util.addEvent(byId("delete-event"), "click", deleteEvent);
        util.addEvent(byId("event-form"), "submit", saveEvent);
    }

    function initialize() {
        window.YoteihyouDateTimeEditor.initialize();
        var mode = getStoredMode() || config.defaultDataSource;
        byId("app-title").innerHTML = util.escapeHtml(config.appTitle);
        document.title = config.appTitle;
        loadOrganizationSections("", "");
        service.setMode(mode);
        updateSourceControls();
        bindEvents();
        settingsController = new window.YoteihyouSettingsController({
            source: organizationSettingsSource,
            baseConfig: organizationConfig,
            onOpen: function () {
                closeEditor();
            },
            onApply: function (newOrganizationConfig) {
                organizationConfig = newOrganizationConfig;
                loadOrganizationSections("", "");
                renderCurrentView();
            }
        });
        settingsController.initialize();
        renderCurrentView();
        reloadData("");
    }

    util.addEvent(window, "load", initialize);
}(window, document));
