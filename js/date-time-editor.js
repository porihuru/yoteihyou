(function (window, document) {
    "use strict";
    var util = window.YoteihyouUtil;
    function el(id) { return document.getElementById(id); }
    function button(text, action) {
        var node = document.createElement("button");
        node.type = "button";
        node.appendChild(document.createTextNode(text));
        node.onclick = action;
        return node;
    }
    function calendar(prefix, month) {
        var panel = el(prefix + "-calendar");
        var year = month.getFullYear();
        var m = month.getMonth();
        var table = document.createElement("table");
        var row;
        var cell;
        var i;
        var offset = new Date(year, m, 1).getDay();
        var days = new Date(year, m + 1, 0).getDate();
        panel.innerHTML = "";
        panel.appendChild(button("前月", function () { calendar(prefix, new Date(year, m - 1, 1)); }));
        panel.appendChild(document.createTextNode(year + "年" + (m + 1) + "月"));
        panel.appendChild(button("翌月", function () { calendar(prefix, new Date(year, m + 1, 1)); }));
        row = document.createElement("tr");
        for (i = 0; i < 7; i += 1) {
            cell = document.createElement("th");
            cell.appendChild(document.createTextNode("日月火水木金土".charAt(i)));
            row.appendChild(cell);
        }
        table.appendChild(row);
        for (i = 0; i < Math.ceil((offset + days) / 7) * 7; i += 1) {
            if (i % 7 === 0) { row = document.createElement("tr"); table.appendChild(row); }
            cell = document.createElement("td");
            if (i >= offset && i < offset + days) {
                (function (day) {
                    cell.appendChild(button(String(day), function () {
                        el(prefix + "-date").value = util.formatDateKey(new Date(year, m, day));
                        panel.hidden = true;
                        el(prefix + "-date").focus();
                    }));
                }(i - offset + 1));
            }
            row.appendChild(cell);
        }
        panel.appendChild(table);
        panel.appendChild(button("閉じる", function () { panel.hidden = true; }));
    }
    function sync() {
        var disabled = el("all-day").checked || el("all-day").disabled;
        ["start", "end"].forEach(function (prefix) {
            el(prefix + "-time").disabled = disabled;
            el(prefix + "-time-toggle").disabled = disabled;
            showTimes(prefix, false);
        });
    }
    function set(prefix, date) {
        el(prefix + "-date").value = date ? util.formatDateKey(date) : "";
        el(prefix + "-time").value = date ? util.pad2(date.getHours()) + util.pad2(date.getMinutes()) : "";
        el(prefix + "-calendar").hidden = true;
        showTimes(prefix, false);
    }
    function read(prefix, allDay) {
        var date = util.trim(el(prefix + "-date").value);
        var time = util.trim(el(prefix + "-time").value);
        var match;
        if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) { return null; }
        if (allDay) { return util.parseDate(date + (prefix === "end" ? " 23:59" : " 00:00")); }
        match = /^(\d{2}):?(\d{2})$/.exec(time);
        return match ? util.parseDate(date + " " + match[1] + ":" + match[2]) : null;
    }
    function showTimes(prefix, visible) {
        el(prefix + "-time-options").hidden = !visible;
        el(prefix + "-time-toggle").setAttribute("aria-expanded", visible ? "true" : "false");
    }
    function initialize() {
        ["start", "end"].forEach(function (prefix) {
            var select = el(prefix + "-time-options");
            var option;
            var minute;
            for (minute = 480; minute <= 1020; minute += 15) {
                (function (value) {
                    option = button(value, function () {
                        el(prefix + "-time").value = value;
                        showTimes(prefix, false);
                        el(prefix + "-time").focus();
                    });
                }(util.pad2(Math.floor(minute / 60)) + util.pad2(minute % 60)));
                select.appendChild(option);
            }
            el(prefix + "-time-toggle").onclick = function () {
                var open = select.hidden;
                showTimes("start", false);
                showTimes("end", false);
                showTimes(prefix, open);
                if (open) { select.scrollTop = 0; }
            };
            el(prefix + "-time-picker").onkeydown = function (event) {
                event = event || window.event;
                if (event.keyCode === 27) { showTimes(prefix, false); el(prefix + "-time").focus(); }
                if (event.keyCode === 40 && event.target === el(prefix + "-time")) {
                    event.preventDefault();
                    showTimes(prefix, true);
                    select.firstChild.focus();
                }
            };
            el(prefix + "-calendar-button").onclick = function () {
                var panel = el(prefix + "-calendar");
                panel.hidden = !panel.hidden;
                if (!panel.hidden) { calendar(prefix, util.parseDate(el(prefix + "-date").value) || new Date()); }
            };
        });
        el("all-day").onchange = sync;
        util.addEvent(document, "click", function (event) {
            var target = (event || window.event).target;
            ["start", "end"].forEach(function (prefix) {
                if (!el(prefix + "-time-picker").contains(target)) { showTimes(prefix, false); }
            });
        });
    }
    window.YoteihyouDateTimeEditor = {initialize: initialize, set: set, read: read, sync: sync};
}(window, document));
