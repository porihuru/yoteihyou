"use strict";

var fs = require("fs");
var path = require("path");

var GROUPS = [
    "GP1", "GP2", "GP3", "GP4", "GP5", "GP6", "GP7", "GP8",
    "GP1科／GS班", "GP1科／JN班", "GP1科／JH班",
    "GP2科／UK班", "GP2科／KR班",
    "GP3科／KY班", "GP3科／KK班", "GP3科／SS班",
    "GP4班／KY班", "GP4班／KY2班", "GP4班／QY班",
    "GP01", "GP02", "GP03", "GP04", "GP05"
];

var BUSY_GROUPS = ["GP2", "GP3", "GP4", "GP5"];
var CATEGORIES = ["会議", "確認", "作業", "連絡", "点検", "調整"];
var LOCATIONS = ["第1会議室", "第2会議室", "作業室", "事務室", "オンライン", "現地"];
var EXTRA_TITLES = {
    GP2: ["受入確認", "資料レビュー", "進捗会議"],
    GP3: ["作業調整", "品質確認", "連絡会"],
    GP4: ["進捗整理", "課題確認", "定例打合せ"],
    GP5: ["設備確認", "運用レビュー", "引継ぎ"]
};
var EXTRA_TIMES = [
    {start: "08:15", end: "09:00"},
    {start: "09:00", end: "10:00"},
    {start: "09:15", end: "10:30"}
];

function pad2(value) {
    return value < 10 ? "0" + value : String(value);
}

function formatDate(date) {
    return date.getUTCFullYear() + "-" + pad2(date.getUTCMonth() + 1) + "-" + pad2(date.getUTCDate());
}

function formatTime(totalMinutes) {
    return pad2(Math.floor(totalMinutes / 60)) + ":" + pad2(totalMinutes % 60);
}

function addDays(date, days) {
    return new Date(date.getTime() + days * 24 * 60 * 60 * 1000);
}

function csvField(value) {
    var text = String(value == null ? "" : value);
    if (/[",\r\n]/.test(text)) {
        return '"' + text.replace(/"/g, '""') + '"';
    }
    return text;
}

function createRow(id, title, start, end, category, location, description, group) {
    return {
        ID: id,
        Title: title,
        EventDate: start,
        EndDate: end,
        Category: category,
        Location: location,
        Description: description,
        Purpose: group + "｜日々・週間・月間"
    };
}

function isEvenlySelected(index, total, selectedCount) {
    return Math.floor((index + 1) * selectedCount / total) >
        Math.floor(index * selectedCount / total);
}

function reducePeriodTargets(rows, specialCount) {
    var regularCount = rows.length - specialCount;
    var weeklyCount = Math.round(rows.length * 0.5) - specialCount;
    var monthlyCount = Math.round(rows.length * 0.2) - specialCount;
    var weeklyIndex = 0;
    var targets;
    var i;
    for (i = 0; i < regularCount; i += 1) {
        targets = ["日々"];
        if (isEvenlySelected(i, regularCount, weeklyCount)) {
            targets.push("週間");
            if (isEvenlySelected(weeklyIndex, weeklyCount, monthlyCount)) {
                targets.push("月間");
            }
            weeklyIndex += 1;
        }
        rows[i].Purpose = rows[i].Purpose.replace(/｜.*$/, "｜" + targets.join("・"));
    }
}

function buildRows() {
    var rows = [];
    var firstDay = new Date(Date.UTC(2026, 8, 1));
    var lastDay = new Date(Date.UTC(2026, 9, 31));
    var date = firstDay;
    var dayIndex = 0;
    var id = 1;
    var groupIndex;
    var group;
    var startMinutes;
    var durationMinutes;
    var start;
    var end;
    var extraIndex;

    while (date.getTime() <= lastDay.getTime()) {
        if (date.getUTCDay() !== 0 && date.getUTCDay() !== 6) {
            for (groupIndex = 0; groupIndex < GROUPS.length; groupIndex += 1) {
            group = GROUPS[groupIndex];
            startMinutes = 8 * 60 + ((dayIndex * 2 + groupIndex * 3) % 28) * 15;
            durationMinutes = 45 + ((dayIndex + groupIndex) % 3) * 15;
            start = formatDate(date) + " " + formatTime(startMinutes);
            end = formatDate(date) + " " + formatTime(startMinutes + durationMinutes);
            rows.push(createRow(
                id,
                group + " 日次確認",
                start,
                end,
                CATEGORIES[(dayIndex + groupIndex) % CATEGORIES.length],
                LOCATIONS[(dayIndex * 2 + groupIndex) % LOCATIONS.length],
                formatDate(date) + "のサンプル予定です。反映先は目的列の設定に従います。",
                group
            ));
            id += 1;

            if (BUSY_GROUPS.indexOf(group) >= 0) {
                for (extraIndex = 0; extraIndex < EXTRA_TIMES.length; extraIndex += 1) {
                    rows.push(createRow(
                        id,
                        group + " " + EXTRA_TITLES[group][extraIndex],
                        formatDate(date) + " " + EXTRA_TIMES[extraIndex].start,
                        formatDate(date) + " " + EXTRA_TIMES[extraIndex].end,
                        CATEGORIES[(groupIndex + extraIndex + 1) % CATEGORIES.length],
                        LOCATIONS[(dayIndex + groupIndex + extraIndex + 1) % LOCATIONS.length],
                        group + "の表示件数と重なり回避を確認するサンプル予定です。",
                        group
                    ));
                    id += 1;
                }
            }
            }
        }
        date = addDays(date, 1);
        dayIndex += 1;
    }

    rows.push(createRow(
        id,
        "GP1 日またぎ予定",
        "2026-09-08 22:00",
        "2026-09-09 02:00",
        "夜間対応",
        "オンライン",
        "開始日から翌日にまたがる予定の表示確認用です。",
        "GP1"
    ));
    id += 1;
    rows.push(createRow(
        id,
        "GP1 月またぎ計画",
        "2026-09-28 09:00",
        "2026-10-02 17:00",
        "計画",
        "本部会議室",
        "9月から10月へまたがる予定の表示確認用です。",
        "GP1"
    ));
    id += 1;
    rows.push(createRow(
        id,
        "GP1 週またぎ対応",
        "2026-10-09 13:00",
        "2026-10-13 12:00",
        "対応",
        "オンライン",
        "金曜日から翌週火曜日へまたがる予定の表示確認用です。",
        "GP1"
    ));

    reducePeriodTargets(rows, 3);
    return rows;
}

function toCsv(rows) {
    var headers = ["ID", "Title", "EventDate", "EndDate", "Category", "Location", "Description", "Purpose"];
    var lines = [headers.join(",")];
    rows.forEach(function (row) {
        lines.push(headers.map(function (header) {
            return csvField(row[header]);
        }).join(","));
    });
    return lines.join("\r\n") + "\r\n";
}

function writeSampleCsv() {
    var outputPath = path.resolve(__dirname, "..", "data", "schedule.csv");
    var rows = buildRows();
    fs.writeFileSync(outputPath, toCsv(rows), "utf8");
    process.stdout.write(outputPath + ": " + rows.length + " schedules\n");
}

if (require.main === module) {
    writeSampleCsv();
}

module.exports = {
    groups: GROUPS.slice(0),
    busyGroups: BUSY_GROUPS.slice(0),
    buildRows: buildRows,
    toCsv: toCsv,
    writeSampleCsv: writeSampleCsv
};
