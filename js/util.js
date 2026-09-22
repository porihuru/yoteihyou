(function (window) {
    "use strict";

    function pad2(value) {
        value = String(value);
        return value.length < 2 ? "0" + value : value;
    }

    function trim(value) {
        return String(value === null || typeof value === "undefined" ? "" : value).replace(/^\s+|\s+$/g, "");
    }

    function escapeHtml(value) {
        return String(value === null || typeof value === "undefined" ? "" : value)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#39;");
    }

    function formatDateKey(date) {
        return date.getFullYear() + "-" + pad2(date.getMonth() + 1) + "-" + pad2(date.getDate());
    }

    function formatDateTime(date) {
        if (!date || isNaN(date.getTime())) {
            return "";
        }
        return formatDateKey(date) + " " + pad2(date.getHours()) + ":" + pad2(date.getMinutes());
    }

    function parseDate(value) {
        var text;
        var match;
        var date;
        var hour;
        var minute;

        if (Object.prototype.toString.call(value) === "[object Date]") {
            return new Date(value.getTime());
        }

        text = trim(value);
        if (!text) {
            return null;
        }

        match = /^(\d{4})[-\/]?(\d{1,2})[-\/]?(\d{1,2})(?:[ T](\d{1,2})(?::?(\d{1,2}))?)?$/.exec(text);
        if (!match) {
            return null;
        }

        hour = parseInt(match[4] || "0", 10);
        minute = parseInt(match[5] || "0", 10);
        if (hour < 0 || hour > 23 || minute < 0 || minute > 59) {
            return null;
        }

        date = new Date(
            parseInt(match[1], 10),
            parseInt(match[2], 10) - 1,
            parseInt(match[3], 10),
            hour,
            minute,
            0,
            0
        );

        if (date.getFullYear() !== parseInt(match[1], 10) ||
                date.getMonth() !== parseInt(match[2], 10) - 1 ||
                date.getDate() !== parseInt(match[3], 10)) {
            return null;
        }
        return date;
    }

    function toIsoString(date) {
        if (!date || isNaN(date.getTime())) {
            return null;
        }
        return date.getUTCFullYear() + "-" + pad2(date.getUTCMonth() + 1) + "-" + pad2(date.getUTCDate()) +
            "T" + pad2(date.getUTCHours()) + ":" + pad2(date.getUTCMinutes()) + ":" + pad2(date.getUTCSeconds()) + "Z";
    }

    function isTrue(value) {
        var text;
        if (value === true || value === 1) {
            return true;
        }
        text = trim(value).toLowerCase();
        return text === "true" || text === "1" || text === "yes" || text === "はい" || text === "○";
    }

    function addEvent(element, eventName, handler) {
        if (element.addEventListener) {
            element.addEventListener(eventName, handler, false);
        } else if (element.attachEvent) {
            element.attachEvent("on" + eventName, handler);
        }
    }

    function getJson(xhr) {
        if (!xhr.responseText) {
            return {};
        }
        return JSON.parse(xhr.responseText);
    }

    function getErrorMessage(xhr, fallback) {
        var data;
        var message;
        try {
            data = getJson(xhr);
            message = data && data.error && data.error.message && (data.error.message.value || data.error.message);
        } catch (ignore) {
            message = "";
        }
        return message || (fallback + "（HTTP " + xhr.status + "）");
    }

    window.YoteihyouUtil = {
        pad2: pad2,
        trim: trim,
        escapeHtml: escapeHtml,
        formatDateKey: formatDateKey,
        formatDateTime: formatDateTime,
        parseDate: parseDate,
        toIsoString: toIsoString,
        isTrue: isTrue,
        addEvent: addEvent,
        getJson: getJson,
        getErrorMessage: getErrorMessage
    };
}(window));
