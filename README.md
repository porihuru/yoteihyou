# 予定表

Edge 95 と Internet Explorer 11（IEモードを含む）で動作する、SharePointリスト連携の月間予定表です。

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

`ID` は既存の標準列を使用します。バナーURL、位置情報、空き時間情報、重複予約のチェック、出席者、設備、目的、登録・更新情報など、アプリで使用しない列は変更しません。

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
ID,Title,EventDate,EndDate,Category,Location,Description
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
