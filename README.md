# 予定表

Edge 95 と Internet Explorer 11（IEモードを含む）で動作する、SharePointリスト連携の月間予定表です。

## 表示と印刷

画面上部で次の表示を切り替えます。

- 日々: 1日分を表示。印刷はA4縦
- 週間: 日曜日から土曜日までを表示。印刷はA4縦
- 月間: 月間カレンダーを表示。印刷はA3横

印刷直前に表示内容を測定し、用紙1枚へ収まる範囲で自動縮小します。初期設定では65%未満まで縮小しなければ収まらない場合、警告を表示して印刷を中止します。余白と縮小下限は `js/config.js` の `print` で変更できます。画面の印刷ボタンと `Ctrl+P` のどちらも同じ事前確認を行います。

予定の入力時に「日々」「週間」「月間」から反映先を複数選択できます。反映先が保存されていない従来データは、3つすべてに表示します。

## 入力と修正

- 表示中の予定をクリックすると、内容を直接修正できます。
- 月間画面は日付枠の空白、週間画面は各日の予定欄の空白をクリックして新規入力します。クリックした日付が開始日時・終了日時へ自動設定されます。
- 日々画面は時間枠の空白をクリックして新規入力します。クリックした日付と時刻が開始日時へ入り、終了日時は設定された標準時間後になります。
- 編集中は画面を分割し、左側4分の3に予定表、右側4分の1に入力・修正画面を表示します。
- 入力・修正画面は取消ボタンまたはEscキーで閉じられます。
- 試験用CSVは閲覧専用です。新規登録・更新・削除はSharePoint接続時だけ使用できます。

日々画面の時間枠は `js/config.js` の `dailyView` で変更できます。初期設定は6時から22時、1時間単位、標準開始時刻9時、標準予定時間1時間です。

## データ接続先

画面上部の「データ接続」で切り替えます。

- **試験用CSV**: `data/schedule.csv` を `XMLHttpRequest` で読み込みます。閲覧専用です。
- **SharePoint**: SharePoint REST APIを使用し、予定の読込・登録・更新・削除を行います。

最後に選択した接続先は `localStorage` に保存されます。保存しない場合は `js/config.js` の `rememberDataSource` を `false` にしてください。

## SharePointリスト

既存のSharePoint標準「イベント」リストを使用します。アプリが読み書きする列は次のとおりです。

| 表示名 | 内部名 |
|---|---|
| タイトル | `Title` |
| 開始時刻 | `EventDate` |
| 終了時刻 | `EndDate` |
| 分類 | `Category` |
| 場所 | `Location` |
| 説明 | `Description` |
| 目的（入力グループ保存先） | `Purpose` |

`ID` は既存の標準列を使用します。入力グループと反映先は `総務科／人事班｜日々・週間・月間` の形式で「目的」列へ保存します。バナーURL、位置情報、空き時間情報、重複予約のチェック、出席者、設備、登録・更新情報など、アプリで使用しない列は変更しません。

## 組織設定

組織は `config/organizations.js` から読み込みます。画面では「科」を選ぶと、設定された「班」だけが次の選択欄へ表示されます。

```javascript
window.YOTEIHYOU_ORGANIZATIONS = {
    separator: "／",
    sections: [
        {
            name: "総務科",
            teams: ["総務班", "人事班", "情報班"]
        }
    ]
};
```

科を追加する場合は、`sections` 内へ同じ形式で追記します。既存予定に設定ファイル外の組織名が保存されている場合は、「設定外」と表示して値を保持します。

## 設定

`js/config.js` の次の項目を環境に合わせて変更します。

```javascript
sharePoint: {
    siteUrl: "https://サーバー/サイト/サブサイト",
    listTitle: "予定表"
}
```

`siteUrl` が空欄の場合は `_spPageContextInfo.webAbsoluteUrl` を使用します。それも取得できない場合は、配置URL中の `SiteAssets`、`Documents`、`DocLib` などからサイトURLを推定します。確実に接続するには `siteUrl` を明示してください。

リスト名や内部名を変更した場合は、同じファイルの `listTitle` と `fields` を変更します。

## 試験用CSV

`data/schedule.csv` の列構成は次のとおりです。

```text
ID,Title,EventDate,EndDate,Category,Location,Description,Purpose
```

- 文字コード: UTF-8（BOMあり・なしの両方に対応）
- 日時: `YYYY-MM-DD HH:mm`
- カンマや改行を含む値: ダブルクォートで囲む

ブラウザの制約上、CSVをサーバー上へ直接上書きしません。試験中の編集結果を保存する場合もSharePointへ切り替えます。

## 配置

リポジトリ一式を同じ構成のまま、SharePointのドキュメントライブラリまたはIISへ配置します。CSV読込はHTTP/HTTPS経由で行うため、`index.html` をローカルファイルとして直接開くより、Webサーバー上で確認してください。

## IE11互換方針

- JavaScriptはES5構文のみ
- `XMLHttpRequest` を使用
- `let`、`const`、アロー関数、`async/await`、`fetch`、`Promise`、`padStart` は不使用
- 外部ライブラリ、ES Modules、CSS Grid、Flexboxの `gap`、`position: sticky` は不使用
