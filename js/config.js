(function (window) {
    "use strict";

    window.YOTEIHYOU_CONFIG = {
        appTitle: "予定表",
        defaultDataSource: "csv",
        rememberDataSource: true,
        storageKey: "yoteihyou.dataSource",

        csv: {
            url: "data/schedule.csv"
        },

        sharePoint: {
            /* 空欄の場合は、現在のSharePointサイトを使用します。 */
            siteUrl: "",
            listTitle: "予定表",
            pageSize: 500,
            fields: {
                id: "ID",
                title: "Title",
                startDate: "EventDate",
                endDate: "EndDate",
                category: "Category",
                location: "Location",
                description: "Description",
                purpose: "Purpose"
            }
        }
    };
}(window));
