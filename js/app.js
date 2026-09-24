(function (window, document) {
    "use strict";

    var config = window.YOTEIHYOU_CONFIG;
    var dailyViewConfig = config.dailyView || {
        startHour: 6,
        endHour: 22,
        standardStartHour: 7,
        standardEndHour: 18,
        slotMinutes: 60,
        defaultStartHour: 9,
        defaultDurationMinutes: 60
    };
    var printConfig = config.print || {marginMm: 10, minimumScale: 0.65};
    var organizationConfig = window.YOTEIHYOU_ORGANIZATIONS || {separator: "／", sections: []};
    var util = window.YoteihyouUtil;
    var service = new window.YoteihyouDataService(config);
    var accessCounter = new window.YoteihyouAccessCounter(config.sharePoint);
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
    var printPreferences = null;
    var isCapturingPrint = false;
    var loadRequestId = 0;
    var historyRequestId = 0;
    var DAILY_SNAP_MINUTES = 15;
    var PREFERENCE_COOKIE_PREFIX = "yoteihyou_";
    var DISPLAY_PREFERENCE_KEY = "yoteihyou.display";
    var MIN_DISPLAY_ZOOM = 70;
    var MAX_DISPLAY_ZOOM = 150;
    var DISPLAY_ZOOM_STEP = 10;
    var currentDisplayZoom = 100;
    var weeklyColumnWeights = [];
    var isDarkModeEnabled = false;
    var dailyInteraction = {
        selectedItemId: "",
        clipboard: null,
        target: null,
        drag: null,
        indicator: null,
        targetCell: null,
        suppressItemId: ""
    };
    var monthlyPan = {
        active: false,
        moved: false,
        suppressClick: false,
        startX: 0,
        startY: 0,
        startScrollLeft: 0,
        startScrollTop: 0
    };
    var organizationCollapseState = {
        sections: {},
        blocks: {}
    };
    var fixedHeader = {
        active: false,
        headerPlaceholder: null,
        toolbarPlaceholder: null,
        axisOverlay: null,
        toolbarGap: 0
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
        byId("save-event-top").disabled = readOnly;
        byId("delete-event-top").disabled = readOnly;
        byId("csv-unsaved-note").style.display = service.getMode() === "csv" ? "block" : "none";
    }

    function setEditorLayoutOpen(isOpen) {
        var className = document.body.className.replace(/(^|\s)editor-open(?=\s|$)/g, " ");
        className = className.replace(/^\s+|\s+$/g, "").replace(/\s+/g, " ");
        document.body.className = isOpen ? (className ? className + " editor-open" : "editor-open") : className;
        updateWeeklyColumnWidths();
        refreshFixedTimeAxis();
        updateFixedHeader();
    }

    function initializeFixedHeader() {
        var header = document.getElementsByClassName("app-header")[0];
        var toolbar = document.getElementsByClassName("toolbar")[0];
        var margin = window.getComputedStyle ? window.getComputedStyle(toolbar).marginTop : toolbar.currentStyle.marginTop;
        fixedHeader.toolbarGap = parseFloat(margin);
        if (isNaN(fixedHeader.toolbarGap)) { fixedHeader.toolbarGap = 0; }
        fixedHeader.headerPlaceholder = document.createElement("div");
        fixedHeader.toolbarPlaceholder = document.createElement("div");
        fixedHeader.axisOverlay = document.createElement("div");
        fixedHeader.headerPlaceholder.className = "fixed-header-placeholder screen-only";
        fixedHeader.toolbarPlaceholder.className = "fixed-header-placeholder screen-only";
        fixedHeader.axisOverlay.className = "fixed-daily-axis screen-only";
        header.parentNode.insertBefore(fixedHeader.headerPlaceholder, header);
        toolbar.parentNode.insertBefore(fixedHeader.toolbarPlaceholder, toolbar);
        document.body.appendChild(fixedHeader.axisOverlay);
    }

    function releaseFixedHeader() {
        var header;
        var toolbar;
        if (!fixedHeader.active) { return; }
        header = document.getElementsByClassName("app-header")[0];
        toolbar = document.getElementsByClassName("toolbar")[0];
        header.style.position = "";
        header.style.top = "";
        header.style.left = "";
        header.style.width = "";
        header.style.zIndex = "";
        toolbar.style.position = "";
        toolbar.style.top = "";
        toolbar.style.left = "";
        toolbar.style.width = "";
        toolbar.style.marginTop = "";
        toolbar.style.zIndex = "";
        fixedHeader.headerPlaceholder.style.height = "0";
        fixedHeader.toolbarPlaceholder.style.height = "0";
        fixedHeader.toolbarPlaceholder.style.marginTop = "0";
        fixedHeader.axisOverlay.style.display = "none";
        fixedHeader.active = false;
    }

    function refreshFixedTimeAxis() {
        var overlay = fixedHeader.axisOverlay;
        var viewId = state.viewMode === "daily" ? "daily-view" :
            state.viewMode === "weekly" ? "weekly-view" : "monthly-view";
        var headId = state.viewMode === "daily" ? "daily-head" :
            state.viewMode === "weekly" ? "weekly-head" : "monthly-head";
        var source = byId(headId);
        var sourceTable = byId(viewId).getElementsByTagName("table")[0];
        var copy;
        var cells;
        var sourceCells;
        var i;
        if (!overlay || !source.firstChild) { return; }
        overlay.innerHTML = "";
        copy = document.createElement("table");
        copy.className = sourceTable.className;
        copy.style.width = sourceTable.offsetWidth + "px";
        copy.style.minWidth = "0";
        copy.appendChild(source.cloneNode(true));
        cells = copy.getElementsByTagName("th");
        sourceCells = source.getElementsByTagName("th");
        if (cells.length === sourceCells.length) {
            for (i = 0; i < cells.length; i += 1) {
                cells[i].style.width = sourceCells[i].offsetWidth + "px";
            }
        }
        overlay.appendChild(copy);
    }

    function updateFixedHeader() {
        var header;
        var toolbar;
        var shell;
        var shellRect;
        var scheduleView;
        var scheduleTable;
        var tableRect;
        var axis;
        var axisRect;
        var dailyRect;
        var scale = currentDisplayZoom / 100;
        var toolbarBottom;
        if (!fixedHeader.headerPlaceholder) { return; }
        header = document.getElementsByClassName("app-header")[0];
        toolbar = document.getElementsByClassName("toolbar")[0];
        shell = byId("app");
        scheduleView = byId(state.viewMode + "-view");
        scheduleTable = scheduleView.getElementsByTagName("table")[0];
        if (!fixedHeader.active && header.getBoundingClientRect().top > 0) { return; }
        if (fixedHeader.active && fixedHeader.headerPlaceholder.getBoundingClientRect().top > 0) {
            releaseFixedHeader();
            return;
        }
        if (!fixedHeader.active) {
            fixedHeader.headerPlaceholder.style.height = header.offsetHeight + "px";
            fixedHeader.toolbarPlaceholder.style.height = toolbar.offsetHeight + "px";
            fixedHeader.toolbarPlaceholder.style.marginTop = fixedHeader.toolbarGap + "px";
            header.style.position = "fixed";
            header.style.top = "0";
            header.style.zIndex = "1052";
            toolbar.style.position = "fixed";
            toolbar.style.marginTop = "0";
            toolbar.style.zIndex = "1051";
            fixedHeader.active = true;
            refreshFixedTimeAxis();
        }
        shellRect = shell.getBoundingClientRect();
        header.style.left = shellRect.left / scale + "px";
        header.style.width = shell.offsetWidth + "px";
        fixedHeader.headerPlaceholder.style.height = header.offsetHeight + "px";
        toolbar.style.left = shellRect.left / scale + "px";
        toolbar.style.width = shell.offsetWidth + "px";
        toolbar.style.top = header.offsetHeight + fixedHeader.toolbarGap + "px";
        fixedHeader.toolbarPlaceholder.style.height = toolbar.offsetHeight + "px";
        toolbarBottom = (header.offsetHeight + fixedHeader.toolbarGap + toolbar.offsetHeight) * scale;
        axis = byId(state.viewMode === "daily" ? "daily-head" :
            state.viewMode === "weekly" ? "weekly-head" : "monthly-head");
        axisRect = axis.getBoundingClientRect();
        dailyRect = scheduleView.getBoundingClientRect();
        tableRect = scheduleTable.getBoundingClientRect();
        if (axisRect.top <= toolbarBottom && dailyRect.bottom > toolbarBottom + axis.offsetHeight * scale) {
            fixedHeader.axisOverlay.style.display = "block";
            fixedHeader.axisOverlay.style.top = toolbarBottom / scale + "px";
            fixedHeader.axisOverlay.style.left = dailyRect.left / scale + "px";
            fixedHeader.axisOverlay.style.width = Math.min(scheduleView.offsetWidth, shell.offsetWidth) + "px";
            if (fixedHeader.axisOverlay.firstChild) {
                fixedHeader.axisOverlay.firstChild.style.marginLeft =
                    (tableRect.left - dailyRect.left) / scale - 1 + "px";
            }
        } else {
            fixedHeader.axisOverlay.style.display = "none";
        }
    }

    function updateSourceControls() {
        var mode = service.getMode();
        var editable = !service.isReadOnly();
        byId("source-csv").checked = mode === "csv";
        byId("source-sharepoint").checked = mode === "sharepoint";
        byId("new-event").disabled = !editable;
        byId("edit-mode-status").className = "edit-mode-status" + (editable ? " editing" : "");
        byId("edit-mode-status").innerHTML = editable ? "更新可能" : "読取専用";
        byId("toggle-schedule-edit").innerHTML = editable ? "更新終了" : "予定表更新";
        byId("toggle-schedule-edit").setAttribute("aria-pressed", editable ? "true" : "false");
    }

    function toggleScheduleEdit() {
        service.setEditingEnabled(!service.isEditingEnabled());
        if (byId("event-editor").style.display !== "none") { closeEditor(); }
        if (settingsController && byId("settings-panel").style.display !== "none") { settingsController.close(); }
        updateSourceControls();
        setMessage(service.isReadOnly() ? "予定表を読取専用にしました。" :
            "予定表を更新できる状態にしました。作業後は「更新終了」を押してください。", false);
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
        var daysFromMonday = (day.getDay() + 6) % 7;
        return new Date(day.getFullYear(), day.getMonth(), day.getDate() - daysFromMonday);
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

    function readDisplayPreference(name) {
        var value = getCookieValue(PREFERENCE_COOKIE_PREFIX + name);
        var legacyValue = "";
        if (value) {
            return value;
        }
        if (name !== "darkMode" && name !== "zoom") {
            return "";
        }
        try {
            legacyValue = window.localStorage.getItem(DISPLAY_PREFERENCE_KEY + "." + name) || "";
        } catch (ignore) {
            legacyValue = "";
        }
        if (legacyValue) {
            storeDisplayPreference(name, legacyValue);
        }
        return legacyValue;
    }

    function storeDisplayPreference(name, value) {
        var expires = new Date();
        var cookieText = PREFERENCE_COOKIE_PREFIX + name + "=" + encodeURIComponent(String(value)) +
            "; expires=" + new Date(expires.getTime() + 365 * 24 * 60 * 60 * 1000).toUTCString() + "; path=/";
        try {
            if (window.location.protocol === "https:") {
                cookieText += "; secure";
            }
            document.cookie = cookieText;
        } catch (ignore) {
            /* Cookieを利用できない環境でも画面操作は継続します。 */
        }
    }

    function getCookieValue(name) {
        var cookies;
        var cookie;
        var separator;
        var i;
        try {
            cookies = document.cookie ? document.cookie.split(";") : [];
            for (i = 0; i < cookies.length; i += 1) {
                cookie = cookies[i].replace(/^\s+/, "");
                separator = cookie.indexOf("=");
                if (separator >= 0 && cookie.substring(0, separator) === name) {
                    return decodeURIComponent(cookie.substring(separator + 1));
                }
            }
        } catch (ignore) {
            return "";
        }
        return "";
    }

    function deletePreferenceCookie(name) {
        var cookieText = PREFERENCE_COOKIE_PREFIX + name + "=; expires=Thu, 01 Jan 1970 00:00:00 GMT; path=/";
        try {
            if (window.location.protocol === "https:") {
                cookieText += "; secure";
            }
            document.cookie = cookieText;
        } catch (ignore) {
            /* Cookieを利用できない環境でも画面操作は継続します。 */
        }
    }

    function removeLegacyDisplayPreference(name) {
        try {
            window.localStorage.removeItem(DISPLAY_PREFERENCE_KEY + "." + name);
        } catch (ignore) {
            /* 旧形式の設定がなくても初期化を続けます。 */
        }
    }

    function copyCollapsedFlags(target, source, keyPrefix) {
        var key;
        if (!source || typeof source !== "object") {
            return;
        }
        for (key in source) {
            if (Object.prototype.hasOwnProperty.call(source, key) &&
                    source[key] === true && key.indexOf(keyPrefix) === 0) {
                target[key] = true;
            }
        }
    }

    function getCollapsedFlags(source, keyPrefix) {
        var result = {};
        var key;
        for (key in source) {
            if (Object.prototype.hasOwnProperty.call(source, key) &&
                    source[key] === true && key.indexOf(keyPrefix) === 0) {
                result[key] = true;
            }
        }
        return result;
    }

    function loadOrganizationCollapseState() {
        var savedState = readDisplayPreference("groups");
        var parsed;
        organizationCollapseState = {sections: {}, blocks: {}};
        if (!savedState) {
            return;
        }
        try {
            parsed = JSON.parse(savedState);
            copyCollapsedFlags(organizationCollapseState.sections, parsed.sections, "section:");
            copyCollapsedFlags(organizationCollapseState.blocks, parsed.blocks, "block:");
        } catch (ignore) {
            organizationCollapseState = {sections: {}, blocks: {}};
        }
    }

    function saveOrganizationCollapseState() {
        var savedState = {
            sections: getCollapsedFlags(organizationCollapseState.sections, "section:"),
            blocks: getCollapsedFlags(organizationCollapseState.blocks, "block:")
        };
        storeDisplayPreference("groups", JSON.stringify(savedState));
    }

    function clearDisplayPreferenceCookies() {
        var preferenceNames = ["viewMode", "darkMode", "zoom", "groups"];
        var modes = ["daily", "weekly", "monthly"];
        var printNames = ["paper", "orientation", "scaleX", "scaleY", "groups"];
        var i;
        var j;
        for (i = 0; i < preferenceNames.length; i += 1) {
            deletePreferenceCookie(preferenceNames[i]);
        }
        for (i = 0; i < modes.length; i += 1) {
            for (j = 0; j < printNames.length; j += 1) {
                deletePreferenceCookie("print_" + modes[i] + "_" + printNames[j]);
            }
        }
        removeLegacyDisplayPreference("darkMode");
        removeLegacyDisplayPreference("zoom");
    }

    function showCookieSettingsStatus(text, isError) {
        var element = byId("cookie-settings-status");
        element.className = isError ? "settings-status cookie-settings-status error" :
            "settings-status cookie-settings-status";
        element.innerHTML = util.escapeHtml(text || "");
        element.style.display = text ? "block" : "none";
    }

    function resetDisplayPreferences() {
        if (!window.confirm("このブラウザーに保存された表示設定と印刷設定を初期化しますか？")) {
            return;
        }
        clearDisplayPreferenceCookies();
        organizationCollapseState = {sections: {}, blocks: {}};
        state.viewMode = "monthly";
        state.displayMonth = new Date(state.displayDate.getFullYear(), state.displayDate.getMonth(), 1);
        dailyInteraction.target = null;
        dailyInteraction.selectedItemId = "";
        hideDailyContextMenu();
        setDarkMode(false, false);
        setDisplayZoom(100, false);
        renderCurrentView();
        showCookieSettingsStatus("表示・印刷設定Cookieを初期化しました。グループは展開、月間表示、100％に戻しました。", false);
    }

    function setDarkMode(enabled, shouldStore) {
        var body = document.body;
        var className = body.className.replace(/(^|\s)dark-mode(?=\s|$)/g, " ");
        var button = byId("toggle-dark-mode");
        className = className.replace(/^\s+|\s+$/g, "").replace(/\s+/g, " ");
        if (enabled) {
            className = className ? className + " dark-mode" : "dark-mode";
        }
        body.className = className;
        isDarkModeEnabled = enabled;
        button.className = "button appearance-button" + (enabled ? " active" : "");
        button.setAttribute("aria-pressed", enabled ? "true" : "false");
        if (shouldStore) {
            storeDisplayPreference("darkMode", enabled ? "1" : "0");
        }
    }

    function setDisplayZoom(zoom, shouldStore) {
        zoom = Math.max(MIN_DISPLAY_ZOOM, Math.min(MAX_DISPLAY_ZOOM, zoom));
        zoom = MIN_DISPLAY_ZOOM + Math.round((zoom - MIN_DISPLAY_ZOOM) / DISPLAY_ZOOM_STEP) * DISPLAY_ZOOM_STEP;
        currentDisplayZoom = zoom;
        document.body.style.zoom = String(zoom / 100);
        byId("zoom-value").innerHTML = zoom + "%";
        byId("zoom-out").disabled = zoom <= MIN_DISPLAY_ZOOM;
        byId("zoom-in").disabled = zoom >= MAX_DISPLAY_ZOOM;
        if (shouldStore) {
            storeDisplayPreference("zoom", String(zoom));
        }
        updateWeeklyColumnWidths();
        refreshFixedTimeAxis();
        updateFixedHeader();
    }

    function initializeDisplayPreferences() {
        var storedViewMode = readDisplayPreference("viewMode");
        var storedZoom = parseInt(readDisplayPreference("zoom"), 10);
        if (storedViewMode === "daily" || storedViewMode === "weekly" || storedViewMode === "monthly") {
            state.viewMode = storedViewMode;
        }
        setDarkMode(readDisplayPreference("darkMode") === "1", false);
        if (isNaN(storedZoom)) {
            storedZoom = 100;
        }
        setDisplayZoom(storedZoom, false);
        loadOrganizationCollapseState();
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
            lineStyle: util.normalizeLineStyle(item.lineStyle),
            lineColor: util.normalizeLineColor(item.lineColor),
            textColor: util.normalizeTextColor(item.textColor),
            createdBy: item.createdBy || "",
            createdAt: item.createdAt || null,
            modifiedBy: item.modifiedBy || "",
            modifiedAt: item.modifiedAt || null,
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
            dailyInteraction.target = getItemTarget(item,
                findParentByClass(button, "daily-timeline-cell") ||
                findParentByClass(button, "organization-schedule-cell"));
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

    function getPeriodTarget(cell, item) {
        var meta = cell && cell._periodMeta;
        var hour = parseInt(dailyViewConfig.defaultStartHour, 10);
        var minute = isNaN(hour) ? 9 * 60 : hour * 60;
        if (!meta) { return null; }
        if (item && !item.allDay) {
            minute = item.startDate.getHours() * 60 + item.startDate.getMinutes();
        }
        return {
            cell: cell,
            timeline: null,
            section: meta.section,
            team: meta.team,
            day: new Date(meta.day.getTime()),
            rangeStart: 0,
            rangeEnd: 24 * 60,
            minute: minute
        };
    }

    function getItemTarget(item, cell) {
        var purpose = splitPurpose(item.purpose || "");
        if (cell && cell._periodMeta) { return getPeriodTarget(cell, item); }
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
            selectDailyItem(item,
                findParentByClass(event.srcElement || event.target, "daily-event-bar") ||
                findParentByClass(event.srcElement || event.target, "event-item"));
        }
        if (cell) {
            dailyInteraction.target = cell._periodMeta ?
                getPeriodTarget(cell, item || (dailyInteraction.clipboard && dailyInteraction.clipboard.item)) :
                getDailyTarget(cell, event.clientX, false);
        } else if (item) {
            dailyInteraction.target = getItemTarget(item,
                findParentByClass(event.srcElement || event.target, "daily-timeline-cell") ||
                findParentByClass(event.srcElement || event.target, "organization-schedule-cell"));
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
        var before = createNew ? null : findItemById(item.id);
        if (service.isReadOnly()) {
            setMessage("現在の接続先では予定を変更できません。", true);
            return;
        }
        setConnectionStatus(sourceLabel + " 保存中", "");
        (createNew ? service.create : service.update).call(service, item, function (savedItem, items, warning) {
            if (completed) {
                completed(savedItem);
            }
            if (mode === "csv") {
                state.items = items;
                renderCurrentView();
                setConnectionStatus("試験用CSV 編集中（未保存）", "connected");
                setMessage(successMessage + " 再読込すると変更は消えます。", false);
            } else {
                reloadData(successMessage + (warning ? " " + warning : ""), !!warning);
            }
        }, function (message) {
            setConnectionStatus(sourceLabel + " 保存失敗", "error");
            setMessage(message, true);
        }, before);
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
            changed.createdBy = "";
            changed.createdAt = null;
            changed.modifiedBy = "";
            changed.modifiedAt = null;
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

    function getResizedPeriodItem(item, edge, target) {
        var changed = cloneScheduleItem(item);
        var original = edge === "start" ? item.startDate : (item.endDate || item.startDate);
        var candidate = item.allDay ?
            new Date(target.day.getFullYear(), target.day.getMonth(), target.day.getDate(),
                edge === "start" ? 0 : 23, edge === "start" ? 0 : 59) :
            new Date(target.day.getFullYear(), target.day.getMonth(), target.day.getDate(),
                original.getHours(), original.getMinutes());
        if (edge === "start") {
            if (candidate.getTime() > changed.endDate.getTime() - DAILY_SNAP_MINUTES * 60000) {
                throw new Error("開始日時は終了日時より前にしてください。");
            }
            changed.startDate = candidate;
        } else {
            if (candidate.getTime() < changed.startDate.getTime() + DAILY_SNAP_MINUTES * 60000) {
                throw new Error("終了日時は開始日時より後にしてください。");
            }
            changed.endDate = candidate;
        }
        return changed;
    }

    function clearPeriodDropTarget() {
        if (dailyInteraction.targetCell) {
            removeClass(dailyInteraction.targetCell, "period-drop-target");
            dailyInteraction.targetCell = null;
        }
    }

    function beginPeriodDrag(event, item, button, mode) {
        if (service.isReadOnly()) { return; }
        selectDailyItem(item, button);
        hideDailyContextMenu();
        dailyInteraction.drag = {
            item: item,
            button: button,
            cell: findParentByClass(button, "organization-schedule-cell"),
            mode: mode,
            period: true,
            startX: event.clientX,
            startY: event.clientY,
            moved: false,
            target: null
        };
        return preventEvent(event);
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
        var origin;
        if (!drag) {
            return;
        }
        event = event || window.event;
        if (!drag.moved && Math.abs(event.clientX - drag.startX) + Math.abs(event.clientY - drag.startY) < 4) {
            return;
        }
        drag.moved = true;
        addClass(drag.button, "daily-event-dragging");
        if (drag.period) {
            cell = findParentByClass(document.elementFromPoint(event.clientX, event.clientY),
                "organization-schedule-cell");
            target = getPeriodTarget(cell, drag.item);
            if (target && drag.mode !== "move") {
                origin = splitPurpose(drag.item.purpose || "");
                if (target.section !== origin.section || target.team !== origin.team) {
                    target = null;
                }
            }
            clearPeriodDropTarget();
            drag.target = target;
            if (target) {
                dailyInteraction.targetCell = cell;
                addClass(cell, "period-drop-target");
            }
            return preventEvent(event);
        }
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
        clearPeriodDropTarget();
        removeClass(drag.button, "daily-event-dragging");
        dailyInteraction.drag = null;
        if (!drag.moved || !drag.target) {
            return;
        }
        dailyInteraction.suppressItemId = String(drag.item.id);
        if (drag.mode === "move") {
            changed = moveItemToTarget(drag.item, drag.target);
            persistDailyItem(changed, false, "予定を移動しました。", null);
        } else if (drag.period) {
            try {
                changed = getResizedPeriodItem(drag.item, drag.mode, drag.target);
                persistDailyItem(changed, false,
                    drag.mode === "start" ? "開始日を変更しました。" : "終了日を変更しました。", null);
            } catch (error) {
                setMessage(error.message, true);
            }
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

    function createEventButton(item, day, viewMode, showCaption) {
        var button = document.createElement("button");
        var times = document.createElement("span");
        var line = document.createElement("span");
        var caption = document.createElement("span");
        var startHandle;
        var endHandle;
        var startHere = sameDate(item.startDate, day);
        var endHere = sameDate(item.endDate || item.startDate, day);
        var startTime = formatDailyTime(item.startDate.getHours() * 60 + item.startDate.getMinutes());
        var endDate = item.endDate || item.startDate;
        var endTime = formatDailyTime(endDate.getHours() * 60 + endDate.getMinutes());
        button.type = "button";
        button.className = "event-item period-event";
        if (viewMode === "monthly") {
            if (!startHere) { button.className += " period-continues-before"; }
            if (!endHere) { button.className += " period-continues-after"; }
        }
        if (dailyInteraction.selectedItemId === String(item.id)) {
            button.className += " daily-event-selected";
        }
        if (dailyInteraction.clipboard && dailyInteraction.clipboard.cut &&
                dailyInteraction.clipboard.sourceId === String(item.id)) {
            button.className += " daily-event-cut";
        }
        button.title = item.title + (item.purpose ? " / " + item.purpose : "") +
            (item.location ? " / " + item.location : "");
        times.className = "period-event-times";
        if (util.normalizeTextColor(item.textColor) !== "default") {
            times.className += " event-text-color-" + util.normalizeTextColor(item.textColor);
        }
        if (!item.allDay && viewMode !== "monthly") {
            times.appendChild(document.createTextNode(startHere && endHere ?
                startTime + "–" + endTime : startHere ? startTime + "→" :
                endHere ? "→" + endTime : "継続"));
            button.appendChild(times);
        }
        line.className = "daily-event-line" +
            (util.normalizeLineStyle(item.lineStyle) === "dotted" ? " line-dotted" : "") +
            (util.normalizeLineColor(item.lineColor) === "default" ? "" :
                " line-color-" + util.normalizeLineColor(item.lineColor));
        button.appendChild(line);
        caption.className = "period-event-caption";
        if (util.normalizeTextColor(item.textColor) !== "default") {
            caption.className += " event-text-color-" + util.normalizeTextColor(item.textColor);
        }
        caption.appendChild(document.createTextNode(showCaption === false ? "\u00a0" :
            item.title + (item.location ? "（" + item.location + "）" : "")));
        button.appendChild(caption);
        if (startHere) {
            startHandle = document.createElement("span");
            startHandle.className = "period-resize-handle period-resize-start screen-only";
            startHandle.title = "ドラッグして開始日を変更";
            button.appendChild(startHandle);
        }
        if (endHere) {
            endHandle = document.createElement("span");
            endHandle.className = "period-resize-handle period-resize-end screen-only";
            endHandle.title = "ドラッグして終了日を変更";
            button.appendChild(endHandle);
        }
        button.onclick = function (event) {
            stopEvent(event);
            if (dailyInteraction.suppressItemId === String(item.id)) {
                dailyInteraction.suppressItemId = "";
                return false;
            }
            selectDailyItem(item, button);
            openEditor(item);
            return false;
        };
        button.onfocus = function () { selectDailyItem(item, button); };
        button.onmousedown = function (event) {
            var source = (event || window.event).srcElement || (event || window.event).target;
            if (hasClass(source, "period-resize-start")) {
                return beginPeriodDrag(event || window.event, item, button, "start");
            }
            if (hasClass(source, "period-resize-end")) {
                return beginPeriodDrag(event || window.event, item, button, "end");
            }
            return beginPeriodDrag(event || window.event, item, button, "move");
        };
        button.oncontextmenu = function (event) {
            return showDailyContextMenu(event || window.event, item,
                findParentByClass(button, "organization-schedule-cell"));
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

    function getRenderedRowCount(block, requiredRows) {
        var configuredRows = block && block.rowCount ? block.rowCount : 1;
        var minimumRows = block && block.autoRows === false ? configuredRows : 1;
        return Math.max(minimumRows, requiredRows || 0, 1);
    }

    function getOrganizationKey(section, team) {
        return String(section || "") + "\u001f" + String(team || "");
    }

    function getOrganizationSectionsFromBlocks(blocks) {
        var sections = [];
        var byName = {};
        var key;
        var section;
        var i;
        for (i = 0; i < blocks.length; i += 1) {
            key = "section:" + blocks[i].section;
            section = byName[key];
            if (!section) {
                section = {
                    name: blocks[i].section,
                    blocks: [],
                    hasTeams: false
                };
                byName[key] = section;
                sections.push(section);
            }
            section.blocks.push(blocks[i]);
            if (blocks[i].team) {
                section.hasTeams = true;
            }
        }
        return sections;
    }

    function isOrganizationSectionCollapsed(section) {
        return !isCapturingPrint && organizationCollapseState.sections["section:" + section] === true;
    }

    function isOrganizationBlockCollapsed(section, team) {
        return !isCapturingPrint && organizationCollapseState.blocks["block:" + getOrganizationKey(section, team)] === true;
    }

    function toggleOrganizationSection(section) {
        var key = "section:" + section;
        organizationCollapseState.sections[key] = !organizationCollapseState.sections[key];
        saveOrganizationCollapseState();
        renderCurrentView();
    }

    function toggleOrganizationBlock(section, team) {
        var key = "block:" + getOrganizationKey(section, team);
        organizationCollapseState.blocks[key] = !organizationCollapseState.blocks[key];
        saveOrganizationCollapseState();
        renderCurrentView();
    }

    function appendOrganizationName(cell, text, collapsed, section, team, sectionToggle) {
        var button = document.createElement("button");
        var label = document.createElement("span");
        button.type = "button";
        button.className = "organization-collapse-button screen-only";
        button.setAttribute("aria-expanded", collapsed ? "false" : "true");
        button.title = text + (collapsed ? "を展開" : "を最小化");
        button.appendChild(document.createTextNode(collapsed ? "＋" : "－"));
        button.onclick = function (event) {
            stopEvent(event);
            if (sectionToggle) {
                toggleOrganizationSection(section);
            } else {
                toggleOrganizationBlock(section, team);
            }
            return false;
        };
        label.className = "organization-name-label";
        label.appendChild(document.createTextNode(text));
        cell.appendChild(button);
        cell.appendChild(label);
    }

    function appendOrganizationSectionRow(body, section, columnCount) {
        var collapsed = isOrganizationSectionCollapsed(section.name);
        var row = document.createElement("tr");
        var nameCell = document.createElement("th");
        var scheduleCell = document.createElement("td");
        row.className = "organization-section-row" + (collapsed ? " organization-collapsed" : "");
        row.setAttribute("data-print-section", section.name);
        nameCell.className = "organization-name organization-section-name";
        appendOrganizationName(nameCell, section.name, collapsed, section.name, "", true);
        scheduleCell.className = "organization-section-cell";
        scheduleCell.colSpan = columnCount;
        scheduleCell.appendChild(document.createTextNode("\u00a0"));
        row.appendChild(nameCell);
        row.appendChild(scheduleCell);
        body.appendChild(row);
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
                    rowCount: getConfiguredRowCount(groups[i], viewMode),
                    autoRows: groups[i].autoRows !== false
                });
            } else {
                for (j = 0; j < teams.length; j += 1) {
                    blocks.push({
                        section: groups[i].name,
                        team: teams[j].name,
                        label: joinOrganization(groups[i].name, teams[j].name),
                        rowCount: getConfiguredRowCount(teams[j], viewMode),
                        autoRows: teams[j].autoRows !== false
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
                    rowCount: 1,
                    autoRows: true
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

    function alignMonthlyItemsByLane(itemsByDate, dayWidth) {
        var uniqueItems = [];
        var aligned = [];
        var lanes = [];
        var item;
        var lane;
        var occupied;
        var firstIndex;
        var lastIndex;
        var reservedFirstIndex;
        var reservedLastIndex;
        var captionDays;
        var i;
        var j;
        for (j = 0; j < itemsByDate.length; j += 1) {
            aligned[j] = [];
            for (i = 0; i < itemsByDate[j].length; i += 1) {
                if (uniqueItems.indexOf(itemsByDate[j][i]) < 0) {
                    uniqueItems.push(itemsByDate[j][i]);
                }
            }
        }
        uniqueItems.sort(compareItems);
        for (i = 0; i < uniqueItems.length; i += 1) {
            item = uniqueItems[i];
            firstIndex = -1;
            lastIndex = -1;
            for (j = 0; j < itemsByDate.length; j += 1) {
                if (itemsByDate[j].indexOf(item) >= 0) {
                    if (firstIndex < 0) { firstIndex = j; }
                    lastIndex = j;
                }
            }
            captionDays = Math.ceil((estimateDailyCaptionPixels(item) + 12) / (dayWidth || 110));
            reservedFirstIndex = Math.max(0, Math.min(firstIndex, itemsByDate.length - captionDays));
            reservedLastIndex = Math.min(itemsByDate.length - 1,
                Math.max(lastIndex, reservedFirstIndex + captionDays - 1));
            for (lane = 0; lane < lanes.length; lane += 1) {
                occupied = false;
                for (j = reservedFirstIndex; j <= reservedLastIndex; j += 1) {
                    if (lanes[lane][j]) {
                        occupied = true;
                        break;
                    }
                }
                if (!occupied) { break; }
            }
            if (lane === lanes.length) { lanes.push([]); }
            for (j = reservedFirstIndex; j <= reservedLastIndex; j += 1) {
                lanes[lane][j] = true;
                if (itemsByDate[j].indexOf(item) >= 0) {
                    aligned[j][lane] = item;
                }
            }
        }
        return {itemsByDate: aligned, requiredRows: lanes.length};
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

    function constrainMonthlyCaptions(table) {
        var captions = table.querySelectorAll(".period-event-caption");
        var tableRight = table.getBoundingClientRect().right;
        var caption;
        var rect;
        var zoomRatio;
        var naturalWidth;
        var rightLimit = tableRight - 2;
        var i;
        for (i = 0; i < captions.length; i += 1) {
            caption = captions[i];
            caption.style.width = "";
            caption.style.left = "";
            caption.style.textAlign = "";
            rect = caption.getBoundingClientRect();
            zoomRatio = rect.width ? caption.offsetWidth / rect.width : 1;
            naturalWidth = caption.scrollWidth + 2;
            if (rect.left + naturalWidth / zoomRatio > rightLimit) {
                caption.style.width = naturalWidth + "px";
                caption.style.left = ((rightLimit - naturalWidth / zoomRatio - rect.left) * zoomRatio) + "px";
                caption.style.textAlign = "right";
            } else {
                caption.style.width = naturalWidth + "px";
            }
        }
    }

    function renderOrganizationSchedule(viewMode, dates, head, body) {
        var blocks = getOrganizationBlocks(viewMode, dates);
        var sections = getOrganizationSectionsFromBlocks(blocks);
        var section;
        var block;
        var collapsed;
        var itemsByDate;
        var rowCount;
        var row;
        var groupCell;
        var scheduleCell;
        var date;
        var item;
        var j;
        var rowIndex;
        var requiredRows;
        var monthlyLayout;
        var monthlyDayWidth = 110;
        var sectionIndex;
        var blockIndex;
        renderOrganizationHeader(head, dates, viewMode);
        if (viewMode === "monthly" && head.getElementsByTagName("th")[1]) {
            monthlyDayWidth = head.getElementsByTagName("th")[1].offsetWidth || monthlyDayWidth;
        }
        while (body.firstChild) {
            body.removeChild(body.firstChild);
        }
        for (sectionIndex = 0; sectionIndex < sections.length; sectionIndex += 1) {
            section = sections[sectionIndex];
            if (section.hasTeams) {
                appendOrganizationSectionRow(body, section, dates.length);
                if (isOrganizationSectionCollapsed(section.name)) {
                    continue;
                }
            }
            for (blockIndex = 0; blockIndex < section.blocks.length; blockIndex += 1) {
                block = section.blocks[blockIndex];
                collapsed = isOrganizationBlockCollapsed(block.section, block.team);
                itemsByDate = [];
                requiredRows = 0;
                for (j = 0; j < dates.length; j += 1) {
                    itemsByDate[j] = collapsed ? [] :
                        getItemsForOrganizationDay(block.section, block.team, dates[j], viewMode);
                    if (!collapsed) {
                        requiredRows = Math.max(requiredRows, itemsByDate[j].length);
                    }
                }
                if (viewMode === "monthly" && !collapsed) {
                    monthlyLayout = alignMonthlyItemsByLane(itemsByDate, monthlyDayWidth);
                    itemsByDate = monthlyLayout.itemsByDate;
                    requiredRows = monthlyLayout.requiredRows;
                }
                rowCount = collapsed ? 1 : getRenderedRowCount(block, requiredRows);
                for (rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
                    row = document.createElement("tr");
                    row.setAttribute("data-print-section", block.section);
                    row.setAttribute("data-print-team", block.team);
                    row.setAttribute("data-print-block", "1");
                    if (collapsed) {
                        row.className = "organization-collapsed";
                    }
                    if (viewMode === "weekly") {
                        row.className += " weekly-group-row" +
                            (rowIndex === 0 ? " weekly-group-first" : "") +
                            (rowIndex === rowCount - 1 ? " weekly-group-last" : "");
                    }
                    if (viewMode === "monthly") {
                        row.className += " monthly-group-row" +
                            (rowIndex === 0 ? " monthly-group-first" : "") +
                            (rowIndex === rowCount - 1 ? " monthly-group-last" : "");
                    }
                    if (rowIndex === 0) {
                        groupCell = document.createElement("th");
                        groupCell.className = "organization-name" + (block.team ? " organization-team-name" : "");
                        groupCell.rowSpan = rowCount;
                        appendOrganizationName(
                            groupCell,
                            block.team || block.section,
                            collapsed,
                            block.section,
                            block.team,
                            false
                        );
                        row.appendChild(groupCell);
                    }
                    for (j = 0; j < dates.length; j += 1) {
                        date = new Date(dates[j].getTime());
                        item = itemsByDate[j][rowIndex] || null;
                        scheduleCell = document.createElement("td");
                        scheduleCell.className = "organization-schedule-cell" +
                            (collapsed ? " organization-collapsed-cell" : " clickable-date");
                        scheduleCell.title = collapsed ? "" :
                            formatJapaneseDate(date, true) + "の" + block.label + "に予定を追加";
                        if (!collapsed) {
                            scheduleCell._periodMeta = {
                                day: new Date(date.getTime()), section: block.section, team: block.team
                            };
                        }
                        if (item) {
                            scheduleCell.appendChild(createEventButton(item, date, viewMode,
                                viewMode !== "monthly" || j === 0 || !itemOccursOn(item, dates[j - 1])));
                        } else {
                            scheduleCell.appendChild(document.createTextNode("\u00a0"));
                        }
                        if (!collapsed) {
                            (function (cell, targetDate, sectionName, teamName) {
                                cell.onclick = function () {
                                    openEditor(null, targetDate, false, {section: sectionName, team: teamName});
                                };
                                cell.oncontextmenu = function (event) {
                                    selectDailyItem(null, null);
                                    return showDailyContextMenu(event || window.event, null, cell);
                                };
                            }(scheduleCell, date, block.section, block.team));
                        }
                        row.appendChild(scheduleCell);
                    }
                    body.appendChild(row);
                }
            }
        }
        if (viewMode === "monthly") {
            constrainMonthlyCaptions(head.parentNode);
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

    function beginMonthlyPan(event) {
        var view = byId("monthly-view");
        var source;
        var tagName;
        event = event || window.event;
        if (event.button !== 0 || event.ctrlKey || event.shiftKey || event.altKey) {
            return;
        }
        source = event.target || event.srcElement;
        while (source && source !== view) {
            tagName = source.tagName;
            if (tagName === "BUTTON" || tagName === "INPUT" || tagName === "SELECT" ||
                    tagName === "TEXTAREA" || tagName === "A" || hasClass(source, "event-item")) {
                return;
            }
            source = source.parentNode;
        }
        monthlyPan.active = true;
        monthlyPan.moved = false;
        monthlyPan.startX = event.clientX;
        monthlyPan.startY = event.clientY;
        monthlyPan.startScrollLeft = view.scrollLeft;
        monthlyPan.startScrollTop = window.pageYOffset || document.documentElement.scrollTop ||
            document.body.scrollTop || 0;
    }

    function moveMonthlyPan(event) {
        var deltaX;
        var deltaY;
        if (!monthlyPan.active) { return; }
        event = event || window.event;
        deltaX = event.clientX - monthlyPan.startX;
        deltaY = event.clientY - monthlyPan.startY;
        if (!monthlyPan.moved) {
            if (Math.abs(deltaX) < 5 && Math.abs(deltaY) < 5) { return; }
            monthlyPan.moved = true;
            addClass(byId("monthly-view"), "monthly-panning");
        }
        byId("monthly-view").scrollLeft = monthlyPan.startScrollLeft - deltaX;
        window.scrollTo(window.pageXOffset || 0, monthlyPan.startScrollTop - deltaY);
        return preventEvent(event);
    }

    function endMonthlyPan() {
        if (!monthlyPan.active) { return; }
        monthlyPan.active = false;
        removeClass(byId("monthly-view"), "monthly-panning");
        if (monthlyPan.moved) {
            monthlyPan.suppressClick = true;
            window.setTimeout(function () { monthlyPan.suppressClick = false; }, 0);
        }
    }

    function suppressMonthlyPanClick(event) {
        if (!monthlyPan.suppressClick) { return; }
        monthlyPan.suppressClick = false;
        return preventEvent(event);
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

    function isDailyRangeTrigger(item, day) {
        var endDate;
        if (!item || item.allDay || !item.startDate) {
            return false;
        }
        endDate = item.endDate || item.startDate;
        return sameDate(item.startDate, day) && sameDate(endDate, day);
    }

    function getDailyDisplayRange(items, day, minimumStart, maximumEnd, standardStart, standardEnd, slotMinutes) {
        var rangeStart = standardStart;
        var rangeEnd = standardEnd;
        var itemMinutes;
        var endDate;
        var i;

        for (i = 0; i < items.length; i += 1) {
            if (!isDailyRangeTrigger(items[i], day)) {
                continue;
            }
            itemMinutes = items[i].startDate.getHours() * 60 + items[i].startDate.getMinutes();
            if (itemMinutes < standardStart) {
                rangeStart = Math.min(rangeStart, Math.floor(itemMinutes / slotMinutes) * slotMinutes);
            }
            endDate = items[i].endDate || items[i].startDate;
            itemMinutes = endDate.getHours() * 60 + endDate.getMinutes();
            if (itemMinutes > standardEnd) {
                rangeEnd = Math.max(rangeEnd, Math.ceil(itemMinutes / slotMinutes) * slotMinutes);
            }
        }

        return {
            start: Math.max(minimumStart, rangeStart),
            end: Math.min(maximumEnd, Math.max(rangeStart + slotMinutes, rangeEnd))
        };
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

    function getDailyLaneMetrics() {
        var contentHeight = 37;
        var bodyVerticalPadding = 1;
        var eventHeight = contentHeight + bodyVerticalPadding * 2;
        var laneGap = 2;
        return {
            contentHeight: contentHeight,
            bodyVerticalPadding: bodyVerticalPadding,
            eventHeight: eventHeight,
            laneGap: laneGap,
            laneHeight: eventHeight + laneGap
        };
    }

    function getDailyGroupLayout(rowCount) {
        var laneMetrics = getDailyLaneMetrics();
        var groupBorderMargin = 2;
        var scheduledAreaHeight;
        rowCount = Math.max(rowCount, 1);
        scheduledAreaHeight = rowCount * laneMetrics.eventHeight + (rowCount - 1) * laneMetrics.laneGap;
        return {
            scheduledAreaHeight: scheduledAreaHeight,
            groupBorderMargin: groupBorderMargin,
            timelineHeight: scheduledAreaHeight + groupBorderMargin * 2
        };
    }

    function createDailyEventBar(item, itemRange, displayRange, rangeStart, rangeEnd, laneIndex, groupBorderMargin) {
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
        var captionText = document.createElement("span");
        var laneMetrics = getDailyLaneMetrics();
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
        if (typeof groupBorderMargin !== "number") {
            groupBorderMargin = getDailyGroupLayout(1).groupBorderMargin;
        }
        button.style.top = (groupBorderMargin + laneIndex * laneMetrics.laneHeight) + "px";
        button.title = item.title + " / " + util.formatDateTime(item.startDate) + "～" +
            util.formatDateTime(item.endDate || item.startDate) + (item.location ? " / " + item.location : "");

        times.className = "daily-event-times";
        start.className = "daily-event-start";
        if (util.normalizeTextColor(item.textColor) !== "default") {
            start.className += " event-text-color-" + util.normalizeTextColor(item.textColor);
        }
        start.style.left = lineLeft + "%";
        start.appendChild(document.createTextNode(formatDailyTime(itemRange.start)));
        end.className = "daily-event-end";
        if (util.normalizeTextColor(item.textColor) !== "default") {
            end.className += " event-text-color-" + util.normalizeTextColor(item.textColor);
        }
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

        line.className = "daily-event-line" +
            (util.normalizeLineStyle(item.lineStyle) === "dotted" ? " line-dotted" : "") +
            (util.normalizeLineColor(item.lineColor) === "default" ? "" :
                " line-color-" + util.normalizeLineColor(item.lineColor));
        // Text may need extra room, but the line must follow the actual time axis.
        line.style.left = lineLeft + "%";
        line.style.width = lineWidth + "%";
        caption.className = "daily-event-caption";
        captionText.className = "daily-event-caption-text";
        if (util.normalizeTextColor(item.textColor) !== "default") {
            captionText.className += " event-text-color-" + util.normalizeTextColor(item.textColor);
        }
        captionText.appendChild(document.createTextNode(getDailyCaptionText(item)));
        caption.appendChild(captionText);

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
            if (findParentByClass(source, "daily-event-caption") || source === button) {
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
        var standardStartHour = parseInt(dailyViewConfig.standardStartHour, 10);
        var standardEndHour = parseInt(dailyViewConfig.standardEndHour, 10);
        var displayRange;
        var rangeStart;
        var rangeEnd;
        var totalMinutes;
        var blocks;
        var sections;
        var section;
        var block;
        var collapsed;
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
        var laneIndex;
        var groupLayout;
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
        var sectionIndex;
        var blockIndex;

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
        if (isNaN(standardStartHour) || standardStartHour < startHour || standardStartHour >= endHour) {
            standardStartHour = Math.max(startHour, 7);
        }
        if (isNaN(standardEndHour) || standardEndHour <= standardStartHour || standardEndHour > endHour) {
            standardEndHour = Math.min(endHour, 18);
        }
        if (standardEndHour <= standardStartHour) {
            standardStartHour = startHour;
            standardEndHour = endHour;
        }
        displayRange = getDailyDisplayRange(
            items,
            day,
            startHour * 60,
            endHour * 60,
            standardStartHour * 60,
            standardEndHour * 60,
            slotMinutes
        );
        rangeStart = displayRange.start;
        rangeEnd = displayRange.end;

        while (body.firstChild) {
            body.removeChild(body.firstChild);
        }
        byId("month-title").innerHTML = formatJapaneseDate(day, true);
        byId("print-heading").innerHTML = formatJapaneseDate(day, true) + "　日々予定表";

        for (i = 0; i < items.length; i += 1) {
            maximumCaptionPixels = Math.max(maximumCaptionPixels, estimateDailyCaptionPixels(items[i]));
        }
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
        sections = getOrganizationSectionsFromBlocks(blocks);
        for (sectionIndex = 0; sectionIndex < sections.length; sectionIndex += 1) {
            section = sections[sectionIndex];
            if (section.hasTeams) {
                appendOrganizationSectionRow(body, section, 1);
                if (isOrganizationSectionCollapsed(section.name)) {
                    continue;
                }
            }
            for (blockIndex = 0; blockIndex < section.blocks.length; blockIndex += 1) {
                block = section.blocks[blockIndex];
                collapsed = isOrganizationBlockCollapsed(block.section, block.team);
                row = document.createElement("tr");
                row.setAttribute("data-print-section", block.section);
                row.setAttribute("data-print-team", block.team);
                row.setAttribute("data-print-block", "1");
                if (collapsed) {
                    row.className = "organization-collapsed";
                    groupCell = document.createElement("th");
                    groupCell.className = "organization-name daily-organization-name" +
                        (block.team ? " organization-team-name" : "");
                    appendOrganizationName(
                        groupCell,
                        block.team || block.section,
                        true,
                        block.section,
                        block.team,
                        false
                    );
                    timelineCell = document.createElement("td");
                    timelineCell.className = "daily-timeline-cell organization-collapsed-cell";
                    timelineCell.appendChild(document.createTextNode("\u00a0"));
                    row.appendChild(groupCell);
                    row.appendChild(timelineCell);
                    body.appendChild(row);
                    continue;
                }

                blockItems = getItemsForOrganizationDay(block.section, block.team, day, "daily");
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
                rowCount = getRenderedRowCount(block, laneEnds.length);

                groupCell = document.createElement("th");
                groupCell.className = "organization-name daily-organization-name" +
                    (block.team ? " organization-team-name" : "");
                appendOrganizationName(
                    groupCell,
                    block.team || block.section,
                    false,
                    block.section,
                    block.team,
                    false
                );
                row.appendChild(groupCell);

                timelineCell = document.createElement("td");
                timelineCell.className = "daily-timeline-cell clickable-slot";
                timelineCell.title = block.label + "の時間帯をクリックして予定を追加。右クリックで貼り付け";
                timeline = document.createElement("div");
                timeline.className = "daily-timeline";
                groupLayout = getDailyGroupLayout(rowCount);
                timeline.style.height = groupLayout.timelineHeight + "px";
                timelineCell._dailyMeta = {
                    timeline: timeline,
                    section: block.section,
                    team: block.team,
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
                        placedItems[j].lane,
                        groupLayout.groupBorderMargin
                    ));
                }
                timelineCell.appendChild(timeline);
                (function (cell, sectionName, teamName, targetDay) {
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
                        openEditor(null, targetDate, true, {section: sectionName, team: teamName});
                    };
                    cell.oncontextmenu = function (event) {
                        selectDailyItem(null, null);
                        return showDailyContextMenu(event || window.event, null, cell);
                    };
                }(timelineCell, block.section, block.team, day));
                row.appendChild(timelineCell);
                body.appendChild(row);
            }
        }
    }

    function updateWeeklyColumnWidths() {
        var columns = byId("weekly-columns").getElementsByTagName("col");
        var tableWidth = byId("weekly-table").offsetWidth;
        var weightTotal = 0;
        var dayWidth;
        var i;
        if (!tableWidth || weeklyColumnWeights.length !== 7) {
            return;
        }
        for (i = 0; i < 7; i += 1) {
            weightTotal += weeklyColumnWeights[i];
        }
        dayWidth = Math.max(0, tableWidth - 155) / weightTotal;
        columns[0].style.width = "155px";
        for (i = 0; i < 7; i += 1) {
            columns[i + 1].style.width = (dayWidth * weeklyColumnWeights[i]) + "px";
        }
    }

    function renderWeekly() {
        var firstDay = startOfWeek(state.displayDate);
        var lastDay = new Date(firstDay.getFullYear(), firstDay.getMonth(), firstDay.getDate() + 6);
        var dates = [];
        var weight;
        var i;
        byId("month-title").innerHTML = (firstDay.getMonth() + 1) + "月" + firstDay.getDate() + "日～" +
            (lastDay.getMonth() + 1) + "月" + lastDay.getDate() + "日";
        byId("print-heading").innerHTML = firstDay.getFullYear() + "年" +
            (firstDay.getMonth() + 1) + "月" + firstDay.getDate() + "日～" +
            (lastDay.getMonth() + 1) + "月" + lastDay.getDate() + "日　週間予定表";
        weeklyColumnWeights = [];
        for (i = 0; i < 7; i += 1) {
            dates.push(new Date(firstDay.getFullYear(), firstDay.getMonth(), firstDay.getDate() + i));
            weight = dates[i].getDay() === 0 || dates[i].getDay() === 6 ?
                (getItemsForDay(dates[i], "weekly").length ? 1 : 0.5) : 1;
            weeklyColumnWeights.push(weight);
        }
        renderOrganizationSchedule("weekly", dates, byId("weekly-head"), byId("weekly-body"));
        updateWeeklyColumnWidths();
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

    function setLineChoice(kind, value) {
        var field = byId(kind === "style" ? "line-style" : "line-color");
        var group = byId(kind === "style" ? "line-style-buttons" : "line-color-buttons");
        var buttons = group.getElementsByTagName("button");
        var preview = byId("line-preview-line");
        var i;
        value = kind === "style" ? util.normalizeLineStyle(value) : util.normalizeLineColor(value);
        field.value = value;
        for (i = 0; i < buttons.length; i += 1) {
            buttons[i].className = "button line-choice-button" +
                (buttons[i].getAttribute("data-value") === value ? " active" : "");
            buttons[i].setAttribute("aria-pressed", buttons[i].getAttribute("data-value") === value ? "true" : "false");
        }
        preview.className = "daily-event-line" +
            (byId("line-style").value === "dotted" ? " line-dotted" : "") +
            (byId("line-color").value === "default" ? "" : " line-color-" + byId("line-color").value);
    }

    function setTextColorChoice(value) {
        var field = byId("text-color");
        var buttons = byId("text-color-buttons").getElementsByTagName("button");
        var i;
        value = util.normalizeTextColor(value);
        field.value = value;
        for (i = 0; i < buttons.length; i += 1) {
            buttons[i].className = "button line-choice-button" +
                (buttons[i].getAttribute("data-value") === value ? " active" : "");
            buttons[i].setAttribute("aria-pressed", buttons[i].getAttribute("data-value") === value ? "true" : "false");
        }
        byId("text-preview-text").className = value === "default" ? "" : "event-text-color-" + value;
    }

    function setAuditInfo(item) {
        var missing = item ? "情報なし" : "未登録";
        byId("event-created-by").innerHTML = util.escapeHtml(item && item.createdBy ? item.createdBy : missing);
        byId("event-created-at").innerHTML = util.escapeHtml(item && item.createdAt ? util.formatDateTime(item.createdAt) : missing);
        byId("event-modified-by").innerHTML = util.escapeHtml(item && item.modifiedBy ? item.modifiedBy : missing);
        byId("event-modified-at").innerHTML = util.escapeHtml(item && item.modifiedAt ? util.formatDateTime(item.modifiedAt) : missing);
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
        setLineChoice("style", "solid");
        setLineChoice("color", "default");
        setTextColorChoice("default");
        setAuditInfo(null);
    }

    function openEditor(item, selectedDate, useSelectedTime, selectedOrganization) {
        var readOnly = service.isReadOnly();
        var purpose;
        if (!item && readOnly) {
            setMessage("現在の接続先は閲覧専用です。", true);
            return;
        }
        state.editingItem = item || null;
        closeHistory();
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
            setLineChoice("style", item.lineStyle);
            setLineChoice("color", item.lineColor);
            setTextColorChoice(item.textColor);
            setAuditInfo(item);
        } else {
            byId("editor-title").innerHTML = "予定を追加";
        }
        setEditorReadOnly(readOnly);
        window.YoteihyouDateTimeEditor.sync();
        byId("delete-event").style.display = item && !readOnly ? "inline-block" : "none";
        byId("delete-event-top").style.display = item && !readOnly ? "inline-block" : "none";
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

    function closeHistory() {
        historyRequestId += 1;
        byId("history-panel").style.display = "none";
        if (byId("event-editor").style.display === "none") { setEditorLayoutOpen(false); }
    }

    function historyDate(value) {
        var date = value ? new Date(value) : null;
        return date && !isNaN(date.getTime()) ? util.formatDateTime(date) : "情報なし";
    }

    function historyActionLabel(action) {
        return action === "create" ? "新規作成" : action === "update" ? "更新" : "削除";
    }

    function renderHistoryDetail(entry) {
        var detail = byId("history-detail");
        var fields = [
            {key: "title", label: "件名"}, {key: "purpose", label: "グループ・反映先"},
            {key: "startDate", label: "開始日時"}, {key: "endDate", label: "終了日時"},
            {key: "allDay", label: "終日"}, {key: "category", label: "区分"},
            {key: "location", label: "場所"}, {key: "description", label: "内容"},
            {key: "lineStyle", label: "線種"}, {key: "lineColor", label: "線色"},
            {key: "textColor", label: "文字色"}
        ];
        var html = "<h3>詳細</h3><p>" + util.escapeHtml(historyActionLabel(entry.action)) + " / " +
            util.escapeHtml(historyDate(entry.at)) + "<br>操作者: " + util.escapeHtml(entry.actor || "情報なし") +
            "<br>予定リスト: " + util.escapeHtml(entry.scheduleList || config.sharePoint.listTitle) +
            "<br>予定ID: " + util.escapeHtml((entry.after || entry.before || {}).id || entry.itemId || "情報なし") + "</p>";
        var i;
        var before;
        var after;
        function display(snapshot, key) {
            var value = snapshot ? snapshot[key] : "";
            if (key === "startDate" || key === "endDate") { return value ? historyDate(value) : ""; }
            if (key === "allDay") { return value ? "はい" : "いいえ"; }
            if (key === "lineStyle") { return value === "dotted" ? "点線" : "実線"; }
            if (key === "lineColor" || key === "textColor") {
                return value === "red" ? "赤" : value === "green" ? "緑" : value === "brown" ? "茶" : "標準（青）";
            }
            return String(value || "");
        }
        html += "<table class=\"history-detail-table\"><thead><tr><th>項目</th><th>変更前</th><th>変更後</th></tr></thead><tbody>";
        for (i = 0; i < fields.length; i += 1) {
            before = entry.before ? display(entry.before, fields[i].key) : "－";
            after = entry.after ? display(entry.after, fields[i].key) : "－";
            html += "<tr><th>" + util.escapeHtml(fields[i].label) + "</th><td>" + util.escapeHtml(before) +
                "</td><td>" + util.escapeHtml(after) + "</td></tr>";
        }
        detail.innerHTML = html + "</tbody></table>";
    }

    function renderHistory(entries) {
        var list = byId("history-list");
        var i;
        var entry;
        var button;
        list.innerHTML = "";
        byId("history-detail").innerHTML = "";
        if (!entries.length) {
            list.innerHTML = "<p>履歴はありません。</p>";
            return;
        }
        for (i = 0; i < entries.length; i += 1) {
            entry = entries[i];
            button = document.createElement("button");
            button.type = "button";
            button.className = "history-entry";
            button.innerHTML = "<strong>" + util.escapeHtml(historyActionLabel(entry.action)) + "</strong> " +
                util.escapeHtml((entry.after || entry.before || {}).title || entry.title || "（件名なし）") +
                "<br><span>" + util.escapeHtml(historyDate(entry.at)) + " / " +
                util.escapeHtml(entry.actor || "情報なし") + "</span>";
            (function (selected) {
                util.addEvent(button, "click", function () { renderHistoryDetail(selected); });
            }(entry));
            list.appendChild(button);
        }
        renderHistoryDetail(entries[0]);
    }

    function openHistory() {
        var requestId;
        closeEditor();
        historyRequestId += 1;
        requestId = historyRequestId;
        byId("history-panel").style.display = "block";
        setEditorLayoutOpen(true);
        byId("history-status").innerHTML = service.getMode() === "csv" ?
            "試験用CSVのサンプル履歴です。今回の画面操作も表示しますが、保存されません。" : "SharePointから読込中…";
        byId("history-list").innerHTML = "";
        byId("history-detail").innerHTML = "";
        service.loadHistory(function (entries) {
            if (requestId !== historyRequestId) { return; }
            byId("history-status").innerHTML = service.getMode() === "csv" ?
                "試験用CSVのサンプル履歴です。今回の画面操作も表示しますが、保存されません。" :
                util.escapeHtml("最新" + entries.length + "件を表示（上限300件）");
            renderHistory(entries);
        }, function (message) {
            if (requestId !== historyRequestId) { return; }
            byId("history-status").innerHTML = util.escapeHtml("履歴の読込に失敗しました。" + message);
        });
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
            lineStyle: util.normalizeLineStyle(byId("line-style").value),
            lineColor: util.normalizeLineColor(byId("line-color").value),
            textColor: util.normalizeTextColor(byId("text-color").value),
            createdBy: state.editingItem ? state.editingItem.createdBy : "",
            createdAt: state.editingItem ? state.editingItem.createdAt : null,
            modifiedBy: state.editingItem ? state.editingItem.modifiedBy : "",
            modifiedAt: state.editingItem ? state.editingItem.modifiedAt : null,
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

    function reloadData(successMessage, isWarning) {
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
                setMessage(successMessage, !!isWarning);
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
        closeHistory();
        service.setEditingEnabled(false);
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
        (current ? service.update : service.create).call(service, item, function (savedItem, items, warning) {
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
                reloadData((current ? "予定を更新しました。" : "予定を登録しました。") +
                    (warning ? " " + warning : ""), !!warning);
            }
        }, function (message) {
            setConnectionStatus(sourceLabel + " 保存失敗", "error");
            setMessage(message, true);
        }, current);
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
        service.remove(item, function (removedItem, items, warning) {
            closeEditor();
            if (mode === "csv") {
                state.items = items;
                renderCurrentView();
                setConnectionStatus("試験用CSV 編集中（未保存）", "connected");
                setMessage("予定を画面上で削除しました。再読込すると元に戻ります。", false);
            } else {
                reloadData("予定を削除しました。" + (warning ? " " + warning : ""), !!warning);
            }
        }, function (message) {
            setConnectionStatus(sourceLabel + " 削除失敗", "error");
            setMessage(message, true);
        });
    }

    function updatePrintPageStyle() {
        var style = byId("print-page-style");
        var preferences = printPreferences || readPrintPreferences(state.viewMode);
        var cssText = "@media print { @page { size: " + preferences.paper + " " +
            preferences.orientation + "; margin: 0; } }";
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

    function readPrintPreferences(mode) {
        var paper = readDisplayPreference("print_" + mode + "_paper");
        var orientation = readDisplayPreference("print_" + mode + "_orientation");
        var scaleX = parseInt(readDisplayPreference("print_" + mode + "_scaleX"), 10);
        var scaleY = parseInt(readDisplayPreference("print_" + mode + "_scaleY"), 10);
        var groups = readDisplayPreference("print_" + mode + "_groups");
        var excluded = [];
        try {
            excluded = groups ? JSON.parse(groups) : [];
            if (!excluded || Object.prototype.toString.call(excluded) !== "[object Array]") {
                excluded = [];
            }
        } catch (ignore) {
            excluded = [];
        }
        return {
            paper: paper === "A4" || paper === "A3" ? paper : (mode === "weekly" ? "A4" : "A3"),
            orientation: orientation === "portrait" || orientation === "landscape" ? orientation :
                (mode === "weekly" ? "portrait" : "landscape"),
            scaleX: isNaN(scaleX) ? 100 : Math.max(25, Math.min(200, scaleX)),
            scaleY: isNaN(scaleY) ? 100 : Math.max(25, Math.min(200, scaleY)),
            excluded: excluded
        };
    }

    function createPrintSource() {
        var source = document.createElement("div");
        var activeView;
        var viewCopy;
        var table;
        var scrollLeft = getActiveViewElement().scrollLeft;
        source.className = "print-area";
        isCapturingPrint = true;
        try {
            renderCurrentView();
            activeView = getActiveViewElement();
            table = activeView.getElementsByTagName("table")[0];
            source._printBaseWidth = Math.max(activeView.scrollWidth, table ? table.scrollWidth : 0, 1);
            source.appendChild(byId("print-heading").cloneNode(true));
            viewCopy = activeView.cloneNode(true);
            viewCopy.style.display = "block";
            source.appendChild(viewCopy);
            stripPrintIds(source);
        } finally {
            isCapturingPrint = false;
            renderCurrentView();
            getActiveViewElement().scrollLeft = scrollLeft;
        }
        return source;
    }

    function stripPrintIds(source) {
        var descendants = source.getElementsByTagName("*");
        var i;
        for (i = 0; i < descendants.length; i += 1) {
            descendants[i].removeAttribute("id");
        }
    }

    function fillPrintGroups() {
        var list = byId("print-groups");
        var rows = printState.source.getElementsByTagName("tr");
        var known = {};
        var section;
        var team;
        var key;
        var label;
        var input;
        var i;
        list.innerHTML = "";
        for (i = 0; i < rows.length; i += 1) {
            if (rows[i].getAttribute("data-print-block") !== "1") { continue; }
            section = rows[i].getAttribute("data-print-section") || "";
            team = rows[i].getAttribute("data-print-team") || "";
            key = getOrganizationKey(section, team);
            if (known[key]) { continue; }
            known[key] = true;
            label = document.createElement("label");
            input = document.createElement("input");
            input.type = "checkbox";
            input.checked = printPreferences.excluded.indexOf(key) < 0;
            input._printKey = key;
            label.appendChild(input);
            label.appendChild(document.createTextNode(team ? joinOrganization(section, team) : section));
            list.appendChild(label);
            util.addEvent(input, "change", savePrintGroupSelection);
        }
    }

    function savePrintGroupSelection() {
        var inputs = byId("print-groups").getElementsByTagName("input");
        var excluded = [];
        var i;
        for (i = 0; i < inputs.length; i += 1) {
            if (!inputs[i].checked) { excluded.push(inputs[i]._printKey); }
        }
        printPreferences.excluded = excluded;
        storeDisplayPreference("print_" + state.viewMode + "_groups", JSON.stringify(excluded));
        updatePrintPreview();
    }

    function selectAllPrintGroups(checked) {
        var inputs = byId("print-groups").getElementsByTagName("input");
        var i;
        for (i = 0; i < inputs.length; i += 1) { inputs[i].checked = checked; }
        savePrintGroupSelection();
    }

    function filterPrintGroups(source) {
        var rows = source.getElementsByTagName("tr");
        var selectedSections = {};
        var section;
        var team;
        var key;
        var i;
        for (i = rows.length - 1; i >= 0; i -= 1) {
            if (rows[i].getAttribute("data-print-block") !== "1") { continue; }
            section = rows[i].getAttribute("data-print-section") || "";
            team = rows[i].getAttribute("data-print-team") || "";
            key = getOrganizationKey(section, team);
            if (printPreferences.excluded.indexOf(key) >= 0) {
                rows[i].parentNode.removeChild(rows[i]);
            } else {
                selectedSections["section:" + section] = true;
            }
        }
        for (i = rows.length - 1; i >= 0; i -= 1) {
            if (rows[i].parentNode && hasClass(rows[i], "organization-section-row") &&
                    !selectedSections["section:" + rows[i].getAttribute("data-print-section")]) {
                rows[i].parentNode.removeChild(rows[i]);
            }
        }
    }

    function applyPrintDimensions(source, widthRatio, heightRatio) {
        var fontRatio = Math.min(widthRatio, heightRatio);
        var table = source.getElementsByTagName("table")[0];
        var elements = source.querySelectorAll("tr, th, td, col, .print-heading, .event-item, " +
            ".period-event-times, .period-event-caption, .daily-event-times, .daily-event-start, " +
            ".daily-event-end, .daily-event-caption, .daily-event-caption-text, .daily-event-line, " +
            ".daily-axis-label, .daily-timeline-axis, .daily-timeline, .daily-event-bar");
        var metrics = [];
        var element;
        var computed;
        var i;
        for (i = 0; i < elements.length; i += 1) {
            element = elements[i];
            computed = window.getComputedStyle(element);
            metrics.push({
                element: element,
                fontSize: parseFloat(computed.fontSize),
                lineHeight: parseFloat(computed.lineHeight),
                paddingTop: parseFloat(computed.paddingTop),
                paddingBottom: parseFloat(computed.paddingBottom),
                paddingLeft: parseFloat(computed.paddingLeft),
                paddingRight: parseFloat(computed.paddingRight),
                marginTop: parseFloat(computed.marginTop),
                marginBottom: parseFloat(computed.marginBottom),
                minHeight: parseFloat(computed.minHeight),
                height: element.offsetHeight,
                top: parseFloat(computed.top),
                width: parseFloat(element.style.width)
            });
        }
        source.style.width = Math.round(printState.source._printBaseWidth * widthRatio) + "px";
        if (table) {
            table.style.width = "100%";
            table.style.minWidth = "0";
        }
        for (i = 0; i < metrics.length; i += 1) {
            element = metrics[i].element;
            if (!isNaN(metrics[i].fontSize)) {
                element.style.fontSize = (metrics[i].fontSize * fontRatio) + "px";
            }
            if (!isNaN(metrics[i].lineHeight)) {
                element.style.lineHeight = (metrics[i].lineHeight * fontRatio) + "px";
            }
            element.style.paddingTop = (metrics[i].paddingTop * heightRatio) + "px";
            element.style.paddingBottom = (metrics[i].paddingBottom * heightRatio) + "px";
            element.style.paddingLeft = (metrics[i].paddingLeft * widthRatio) + "px";
            element.style.paddingRight = (metrics[i].paddingRight * widthRatio) + "px";
            if (!isNaN(metrics[i].marginTop)) {
                element.style.marginTop = (metrics[i].marginTop * heightRatio) + "px";
            }
            if (!isNaN(metrics[i].marginBottom)) {
                element.style.marginBottom = (metrics[i].marginBottom * heightRatio) + "px";
            }
            if (!isNaN(metrics[i].minHeight)) {
                element.style.minHeight = (metrics[i].minHeight * heightRatio) + "px";
            }
            if (element.tagName === "TR" || element.tagName === "TH" || element.tagName === "TD" ||
                    hasClass(element, "daily-timeline") || hasClass(element, "daily-event-bar") ||
                    hasClass(element, "daily-timeline-axis") || hasClass(element, "daily-event-times")) {
                element.style.height = Math.max(1, metrics[i].height * heightRatio) + "px";
            }
            if (element.tagName === "COL" && !isNaN(metrics[i].width)) {
                element.style.width = (metrics[i].width * widthRatio) + "px";
            }
            if (hasClass(element, "organization-column") || hasClass(element, "organization-name")) {
                element.style.width = (155 * widthRatio) + "px";
            }
            if ((hasClass(element, "daily-event-bar") || hasClass(element, "daily-axis-label")) &&
                    !isNaN(metrics[i].top)) {
                element.style.top = (metrics[i].top * heightRatio) + "px";
            }
        }
        if (table && hasClass(table, "monthly-schedule")) {
            constrainMonthlyCaptions(table);
        }
    }

    function updatePrintPreview() {
        var page = byId("print-preview-page");
        var content = byId("print-preview-content");
        var source = printState.source.cloneNode(true);
        var paperWidth = printPreferences.paper === "A3" ? 297 : 210;
        var paperHeight = printPreferences.paper === "A3" ? 420 : 297;
        var swappedWidth;
        var marginMm = parseFloat(printConfig.marginMm) || 10;
        var leftMarginMm = Math.max(20, marginMm);
        var pixelsPerMm = 96 / 25.4;
        var availableWidth;
        var availableHeight;
        var contentWidth;
        var contentHeight;
        var fitX;
        var fitY;
        var widthRatio;
        var heightRatio;
        var widthCorrection;
        var heightCorrection;
        var selectedCount;
        var inputs;
        var i;
        if (printPreferences.orientation === "landscape") {
            swappedWidth = paperWidth;
            paperWidth = paperHeight;
            paperHeight = swappedWidth;
        }
        filterPrintGroups(source);
        page.style.width = paperWidth + "mm";
        page.style.height = paperHeight + "mm";
        page.style.zoom = String(Math.min(1, Math.max(0.25,
            (window.innerWidth - 340) / (paperWidth * pixelsPerMm))));
        content.style.left = leftMarginMm + "mm";
        content.style.top = marginMm + "mm";
        content.innerHTML = "";
        content.appendChild(source);
        contentWidth = printState.source._printBaseWidth;
        content.style.width = contentWidth + "px";
        contentHeight = Math.max(content.scrollHeight, 1);
        availableWidth = (paperWidth - leftMarginMm - marginMm) * pixelsPerMm;
        availableHeight = (paperHeight - 2 * marginMm) * pixelsPerMm;
        fitX = Math.min(1, availableWidth / contentWidth);
        fitY = Math.min(1, availableHeight / contentHeight);
        for (i = 0; i < 3; i += 1) {
            if (i > 0) {
                source = printState.source.cloneNode(true);
                filterPrintGroups(source);
                content.innerHTML = "";
                content.appendChild(source);
            }
            applyPrintDimensions(source, fitX, fitY);
            content.style.width = source.style.width;
            widthCorrection = Math.min(1, (availableWidth - 2) / Math.max(content.scrollWidth, 1));
            heightCorrection = Math.min(1, (availableHeight - 2) / Math.max(content.scrollHeight, 1));
            if (widthCorrection === 1 && heightCorrection === 1) { break; }
            fitX *= widthCorrection;
            fitY *= heightCorrection;
        }
        widthRatio = fitX * printPreferences.scaleX / 100;
        heightRatio = fitY * printPreferences.scaleY / 100;
        source = printState.source.cloneNode(true);
        filterPrintGroups(source);
        content.innerHTML = "";
        content.appendChild(source);
        applyPrintDimensions(source, widthRatio, heightRatio);
        content.style.width = source.style.width;
        inputs = byId("print-groups").getElementsByTagName("input");
        selectedCount = 0;
        for (i = 0; i < inputs.length; i += 1) {
            if (inputs[i].checked) { selectedCount += 1; }
        }
        byId("execute-print").disabled = selectedCount === 0;
        byId("print-preview-status").innerHTML = selectedCount === 0 ? "印刷するグループを選択してください。" :
            (content.scrollWidth > availableWidth + 1 || content.scrollHeight > availableHeight + 1 ?
                "指定した縮尺では用紙からはみ出す部分があります。" :
                "用紙に合わせて配置しています（基準：横" + Math.round(fitX * 100) +
                "％・縦" + Math.round(fitY * 100) + "％）。");
        updatePrintPageStyle();
    }

    function openPrintPreview() {
        closeEditor();
        printPreferences = readPrintPreferences(state.viewMode);
        printState = {source: createPrintSource(), bodyZoom: document.body.style.zoom};
        document.body.style.zoom = "1";
        byId("print-paper").value = printPreferences.paper;
        byId("print-orientation").value = printPreferences.orientation;
        byId("print-scale-x").value = printPreferences.scaleX;
        byId("print-scale-y").value = printPreferences.scaleY;
        fillPrintGroups();
        byId("print-preview").style.display = "block";
        updatePrintPreview();
        byId("close-print-preview").focus();
    }

    function closePrintPreview() {
        if (printState) { document.body.style.zoom = printState.bodyZoom; }
        byId("print-preview").style.display = "none";
        byId("print-preview-content").innerHTML = "";
        printState = null;
        printPreferences = null;
        updatePrintPageStyle();
        byId("print-view").focus();
    }

    function updatePrintOption(name, value) {
        if (name === "scaleX" || name === "scaleY") {
            value = parseInt(value, 10);
            value = isNaN(value) ? 100 : Math.max(25, Math.min(200, value));
            byId(name === "scaleX" ? "print-scale-x" : "print-scale-y").value = value;
        }
        printPreferences[name] = value;
        storeDisplayPreference("print_" + state.viewMode + "_" + name, value);
        updatePrintPreview();
    }

    function printCurrentView() {
        if (byId("print-preview").style.display === "none") {
            openPrintPreview();
        } else if (!byId("execute-print").disabled) {
            window.print();
        }
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
        refreshFixedTimeAxis();
        updateFixedHeader();
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
        storeDisplayPreference("viewMode", mode);
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
        util.addEvent(byId("toggle-schedule-edit"), "click", toggleScheduleEdit);
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
        util.addEvent(byId("toggle-dark-mode"), "click", function () {
            setDarkMode(!isDarkModeEnabled, true);
        });
        util.addEvent(byId("zoom-out"), "click", function () {
            setDisplayZoom(currentDisplayZoom - DISPLAY_ZOOM_STEP, true);
        });
        util.addEvent(byId("zoom-in"), "click", function () {
            setDisplayZoom(currentDisplayZoom + DISPLAY_ZOOM_STEP, true);
        });
        util.addEvent(byId("reset-display-cookies"), "click", resetDisplayPreferences);
        util.addEvent(byId("print-view"), "click", function () {
            printCurrentView();
        });
        util.addEvent(byId("close-print-preview"), "click", closePrintPreview);
        util.addEvent(byId("execute-print"), "click", function () { window.print(); });
        util.addEvent(byId("print-paper"), "change", function () {
            updatePrintOption("paper", this.value);
        });
        util.addEvent(byId("print-orientation"), "change", function () {
            updatePrintOption("orientation", this.value);
        });
        util.addEvent(byId("print-scale-x"), "change", function () {
            updatePrintOption("scaleX", this.value);
        });
        util.addEvent(byId("print-scale-x"), "input", function () {
            if (this.value !== "") { updatePrintOption("scaleX", this.value); }
        });
        util.addEvent(byId("print-scale-y"), "change", function () {
            updatePrintOption("scaleY", this.value);
        });
        util.addEvent(byId("print-scale-y"), "input", function () {
            if (this.value !== "") { updatePrintOption("scaleY", this.value); }
        });
        util.addEvent(byId("print-select-all"), "click", function () { selectAllPrintGroups(true); });
        util.addEvent(byId("print-select-none"), "click", function () { selectAllPrintGroups(false); });
        util.addEvent(byId("open-history"), "click", openHistory);
        util.addEvent(byId("close-history"), "click", closeHistory);
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
        util.addEvent(byId("monthly-view"), "mousedown", beginMonthlyPan);
        util.addEvent(document, "mousemove", moveMonthlyPan);
        util.addEvent(document, "mouseup", endMonthlyPan);
        util.addEvent(window, "blur", endMonthlyPan);
        byId("monthly-view").addEventListener("click", suppressMonthlyPanClick, true);
        util.addEvent(document, "click", function (event) {
            var target = (event || window.event).srcElement || (event || window.event).target;
            if (!findParentByClass(target, "daily-context-menu")) {
                hideDailyContextMenu();
            }
        });
        util.addEvent(document, "keydown", function (event) {
            var handled = false;
            event = event || window.event;
            if (event.ctrlKey && !event.altKey && !isFormInput(event.srcElement || event.target)) {
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
                if (byId("print-preview").style.display !== "none") {
                    closePrintPreview();
                    return;
                }
                hideDailyContextMenu();
                if (dailyInteraction.drag) {
                    removeClass(dailyInteraction.drag.button, "daily-event-dragging");
                    dailyInteraction.drag = null;
                    if (dailyInteraction.indicator && dailyInteraction.indicator.parentNode) {
                        dailyInteraction.indicator.parentNode.removeChild(dailyInteraction.indicator);
                    }
                    dailyInteraction.indicator = null;
                    clearPeriodDropTarget();
                }
                if (byId("event-editor").style.display !== "none") {
                    closeEditor();
                }
                if (byId("history-panel").style.display !== "none") {
                    closeHistory();
                }
            }
        });
        util.addEvent(window, "scroll", hideDailyContextMenu);
        util.addEvent(window, "scroll", updateFixedHeader);
        util.addEvent(byId("monthly-view"), "scroll", updateFixedHeader);
        util.addEvent(window, "resize", function () {
            updateWeeklyColumnWidths();
            refreshFixedTimeAxis();
            updateFixedHeader();
            if (printState) { updatePrintPreview(); }
        });
        util.addEvent(byId("cancel-edit"), "click", closeEditor);
        util.addEvent(byId("cancel-edit-top"), "click", closeEditor);
        util.addEvent(byId("delete-event"), "click", deleteEvent);
        util.addEvent(byId("delete-event-top"), "click", deleteEvent);
        util.addEvent(byId("save-event-top"), "click", saveEvent);
        util.addEvent(byId("event-form"), "submit", saveEvent);
        function bindLineChoices(kind) {
            var group = byId(kind === "style" ? "line-style-buttons" : "line-color-buttons");
            var buttons = group.getElementsByTagName("button");
            var i;
            for (i = 0; i < buttons.length; i += 1) {
                util.addEvent(buttons[i], "click", function () {
                    setLineChoice(kind, this.getAttribute("data-value"));
                });
            }
        }
        bindLineChoices("style");
        bindLineChoices("color");
        (function () {
            var buttons = byId("text-color-buttons").getElementsByTagName("button");
            var i;
            for (i = 0; i < buttons.length; i += 1) {
                util.addEvent(buttons[i], "click", function () {
                    setTextColorChoice(this.getAttribute("data-value"));
                });
            }
        }());
    }

    function initialize() {
        window.YoteihyouDateTimeEditor.initialize();
        var mode = getStoredMode() || config.defaultDataSource;
        byId("app-title").innerHTML = util.escapeHtml(config.appTitle);
        document.title = config.appTitle;
        initializeDisplayPreferences();
        loadOrganizationSections("", "");
        initializeFixedHeader();
        service.setMode(mode);
        updateSourceControls();
        accessCounter.increment(function (count) {
            byId("access-counter").innerHTML = "アクセス " + String(count);
            byId("access-counter").title = "全利用者の累計アクセス数";
        }, function (message) {
            byId("access-counter").innerHTML = "アクセス 123";
            byId("access-counter").title = "暫定表示（集計できません: " + message + "）";
        });
        bindEvents();
        settingsController = new window.YoteihyouSettingsController({
            source: organizationSettingsSource,
            baseConfig: organizationConfig,
            canEdit: function () { return !service.isReadOnly(); },
            onOpen: function () {
                closeEditor();
                closeHistory();
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
