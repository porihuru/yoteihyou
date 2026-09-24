(function (window) {
    "use strict";

    window.YOTEIHYOU_CONFIG = {
        appTitle: "予定表",
        defaultDataSource: "csv",
        rememberDataSource: true,
        storageKey: "yoteihyou.dataSource",

        dailyView: {
            startHour: 6,
            endHour: 22,
            standardStartHour: 7,
            standardEndHour: 18,
            slotMinutes: 60,
            defaultStartHour: 9,
            defaultDurationMinutes: 60
        },

        print: {
            marginMm: 10,
            minimumScale: 0.65
        },

        csv: {
            url: "data/schedule.csv"
        },

        sharePoint: {
            /* 空欄の場合は、現在のSharePointサイトを使用します。 */
            siteUrl: "",
            listTitle: "予定表",
            pageSize: 500,
            organizationSettings: {
                enabled: true,
                listTitle: "予定表組織設定",
                pageSize: 500,
                fields: {
                    id: "ID",
                    groupName: "Title",
                    teamName: "TeamName",
                    monthlyRows: "MonthlyRows",
                    weeklyRows: "WeeklyRows",
                    dailyRows: "DailyRows",
                    autoRows: "AutoRows",
                    sortOrder: "SortOrder",
                    isActive: "IsActive"
                }
            },
            fields: {
                id: "ID",
                title: "Title",
                startDate: "EventDate",
                endDate: "EndDate",
                allDay: "AllDay",
                category: "Category",
                location: "Location",
                description: "Description",
                purpose: "Purpose"
            }
        }
    };
}(window));
