# 予定表

Edge 95 と Internet Explorer 11（IEモードを含む）で動作する、SharePointリスト連携の月間予定表です。

## データ接続先

画面上部の「データ接続」で切り替えます。

- **試験用CSV**: `data/schedule.csv` を `XMLHttpRequest` で読み込みます。閲覧専用です。
- **SharePoint**: SharePoint REST APIを使用し、予定の読込・登録・更新・削除を行います。

最後に選択した接続先は `localStorage` に保存されます。保存しない場合は `js/config.js` の `rememberDataSource` を `false` にしてください。

## SharePointリスト

カスタムリスト「予定表」を作成し、次の列を用意します。内部名が一致するよう、列は英字の内部名で作成してから表示名を日本語に変更してください。

| 表示名 | 内部名 | 種類 | 必須 | 初期値 |
|---|---|---|---|---|
| 件名 | `Title` | 1行テキスト | はい | － |
| 開始日時 | `StartDate` | 日付と時刻 | はい | － |
| 終了日時 | `EndDate` | 日付と時刻 | いいえ | － |
| 終日 | `AllDay` | はい／いいえ | いいえ | いいえ |
| 区分 | `Category` | 1行テキスト | いいえ | － |
| 場所 | `Location` | 1行テキスト | いいえ | － |
| 内容 | `Description` | 複数行テキスト（プレーンテキスト） | いいえ | － |
| 表示順 | `SortOrder` | 数値 | いいえ | 0 |
| 表示する | `IsActive` | はい／いいえ | いいえ | はい |

既存列の `ID` はそのまま使用します。

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
ID,Title,StartDate,EndDate,AllDay,Category,Location,Description,SortOrder,IsActive
```

- 文字コード: UTF-8（BOMあり・なしの両方に対応）
- 日時: `YYYY-MM-DD HH:mm`
- 真偽値: `true/false`、`1/0`、`yes`、`はい`、`○` に対応
- カンマや改行を含む値: ダブルクォートで囲む

ブラウザの制約上、CSVをサーバー上へ直接上書きしません。試験中の編集結果を保存する場合もSharePointへ切り替えます。

## 配置

リポジトリ一式を同じ構成のまま、SharePointのドキュメントライブラリまたはIISへ配置します。CSV読込はHTTP/HTTPS経由で行うため、`index.html` をローカルファイルとして直接開くより、Webサーバー上で確認してください。

## IE11互換方針

- JavaScriptはES5構文のみ
- `XMLHttpRequest` を使用
- `let`、`const`、アロー関数、`async/await`、`fetch`、`Promise`、`padStart` は不使用
- 外部ライブラリ、ES Modules、CSS Grid、Flexboxの `gap`、`position: sticky` は不使用
