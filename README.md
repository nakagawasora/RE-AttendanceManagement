# AM - 勤怠管理・請求書発行 GAS

Google Apps Script で動作する勤怠入力、月次実績確認、請求書発行、回収、管理画面をまとめたシステムです。Web アプリはデプロイユーザー権限で実行し、一般ユーザーはデータベーススプレッドシートへの権限を持たずに利用します。

## 現在の主な機能

- 勤怠入力フォーム
  - メールアドレスとパスワードで認証
  - PJ、作業内容、開始時刻、終了時刻を入力
  - 一時保存と呼び出し
  - 二重勤怠チェック
  - 入力メールアドレスの SHA-256 ハッシュと PJ 割当表をフロントで突合し、PJ 選択肢をリアルタイムに絞り込み
- 実績確認
  - フォームから「実績」画面を表示
  - 対象年月を指定可能
  - 表示できる年月は当月まで
  - PJ と役職手当を同じ粒度で、単価、勤務時間、税抜金額、税込金額、税込合計として表示
  - 現時点の時給・手当で計算するため参考値として扱う
- 請求書発行
  - 通常請求書とインボイス請求書を自動判定
  - 請求書テンプレートを一時コピーし、そのコピーに書き込んで PDF 化
  - PDF は対象ユーザーの Drive フォルダに保存
  - 発行後に Slack DM と請求書チャンネルへ通知
  - 役職手当は専用行として請求書に追加
- 請求書回収
  - 対象月の PDF を回収用 Drive フォルダへコピー
  - 一括発行、一括回収、個別再発行は先月以前のみ実行可能
- 管理画面
  - PJ 追加、有効化、無効化、割当ユーザー編集
  - 時給増分管理
  - 役職手当 CRUD
  - ユーザー一覧、名前編集、論理削除
  - 管理者ホワイトリスト管理
  - 一括発行、一括回収、個別再発行
  - 請求書チャンネル最新まとめ投稿の未リアクションユーザー確認と催促 DM
  - 管理操作ログ表示

## ファイル構成

```text
AM/
├── .clasp.dev.json
├── .clasp.prod.json
├── package.json
├── README.md
└── src/
    ├── appsscript.json
    ├── BillUtil.js
    ├── ManagementUtil.js
    ├── index.html
    ├── worktime.html
    ├── admin.html
    └── submitted.html
```

## デプロイ設定

`src/appsscript.json` の Web アプリ設定は以下です。

```json
{
  "webapp": {
    "executeAs": "USER_DEPLOYING",
    "access": "ANYONE"
  }
}
```

この構成では、フォーム利用者や管理画面利用者がデータベースシートへ直接アクセスする必要はありません。DB、Drive、Gmail、Slack などへのアクセスはデプロイユーザーの承認で実行されます。

`executeAs` を `USER_ACCESSING` に変更すると、一般ユーザーが DB や Drive にアクセスできず、フォーム送信、実績取得、管理操作が失敗します。

## スクリプトプロパティ

GAS のスクリプトプロパティに以下を設定します。

| プロパティ | 用途 |
|---|---|
| `SLACK_BOT_TOKEN` | Slack Bot OAuth トークン |
| `BILL_CHANNEL_ID` | 請求書まとめ通知を送る Slack チャンネル ID |
| `SLACK_ID_OF_ADMIN` | 管理者の Slack ユーザー ID |
| `MAIL_ADDRESS` | エラー通知先メールアドレス |
| `CC_RECIPIENT` | エラー通知メールの CC |
| `BILL_FOLDER_ID` | 回収済み請求書を格納する Drive フォルダ ID |
| `DATABASE_SHEET_ID` | データベーススプレッドシート ID |
| `NORMAL_BILL_SHEET_ID` | 通常請求書テンプレートのスプレッドシート ID |
| `INVOICE_BILL_SHEET_ID` | インボイス請求書テンプレートのスプレッドシート ID |
| `TAX_PERCENTAGE` | 消費税率。未設定時は `10` |

以下は手動初期化関数 `resetBill()` 用に残っているプロパティです。現在の通常発行はテンプレートの一時コピー方式なので、通常運用ではテンプレート本体を初期化しません。

| プロパティ | 用途 |
|---|---|
| `RESET_NORMAL_BILL_SHEET_ID` | 通常請求書の初期化用シート ID |
| `RESET_INVOICE_BILL_SHEET_ID` | インボイス請求書の初期化用シート ID |

## データベースシート

データベーススプレッドシートは、業務データ 4 シートとメタシートで構成します。

業務データは `Users`、`Projects`、`WageUpgrades`、`RoleAllowances` のシート名で判定します。シート順を並び替えても、この名前が合っていれば同じシートとして扱います。

メタシートの `AdminWhitelist` と `AdminExecutionLogs` は名前で除外されるため、業務データの 1〜4 枚目としては扱われません。メタシート作成でシート順が崩れた場合は `repairDatabaseSheetOrder()` を実行すると、業務シート、ホワイトリスト、ログの順に整えます。

### 1枚目: Users

| 列 | 内容 |
|---|---|
| A | ユーザー ID |
| B | 勤怠シート ID |
| C | Slack ユーザー ID |
| D | 請求書出力先 Drive フォルダ ID |
| E | 個人情報シート ID |
| F | メールアドレス |
| G | パスワードハッシュ SHA-256 |
| H | 一時保存データ JSON |
| I | 名前 |
| J | 論理削除日時 |

J 列に値があるユーザーは論理削除済みです。フォーム、実績、管理画面の通常一覧、請求書発行対象から除外されます。

### 2枚目: Projects

| 列 | 内容 |
|---|---|
| A | PJ ID |
| B | PJ 名 |
| C | 基本時給 |
| D | 状態 |
| E | 対象ユーザー ID |

D 列は `TRUE`、`有効`、`active`、`1`、空欄を有効として扱います。それ以外は無効です。

E 列が空欄なら全ユーザーに割当、ユーザー ID のカンマ区切りなら対象ユーザーだけに割当です。管理画面で割当ユーザーが 0 人になる操作をすると、その PJ は自動的に無効化されます。

### 3枚目: 時給増分

| 列 | 内容 |
|---|---|
| A | ID |
| B | ユーザー ID |
| C | 対象 PJ ID |
| D | 時給増分 |

管理画面ではユーザーごとにまとめて表示し、対象 PJ は `id_name` 形式で表示します。

### 4枚目: 役職手当

| 列 | 内容 |
|---|---|
| A | ID |
| B | 手当名 |
| C | 対象ユーザー ID |
| D | 単価 |
| E | 論理削除日時 |

E 列に値がある手当は無効です。管理画面の削除は物理削除ではなく論理削除です。ユーザー選択は `id_name` 形式で表示します。

### AdminWhitelist

管理画面を利用できる Google アカウントを管理するメタシートです。

| 列 | 内容 |
|---|---|
| A | メールアドレス |
| B | active |
| C | memo |

管理画面は `Session.getActiveUser().getEmail()` で取得したログイン中メールアドレスを、このシートと照合します。初回は DB 権限を持つユーザーが `setupAdminWhitelistSheet()` を実行するか、同名シートを手動作成して管理者メールアドレスを登録します。

### AdminExecutionLogs

管理画面から実行した一括処理などのログを記録するメタシートです。管理画面上の表示件数は制限していますが、シート自体は自動削除しないため、長期運用では必要に応じてアーカイブまたは削除してください。

## 既存DBからの移行

既存のデータベーススプレッドシートを直接加工せず、新しいスプレッドシートへコピーして正規化できます。

GAS エディタから以下を実行します。

```javascript
migrateDatabaseToNamedSheets()
```

この関数は `DATABASE_SHEET_ID` のスプレッドシートを移行元として、次の処理を行います。

- 業務シートを `Users`、`Projects`、`WageUpgrades`、`RoleAllowances` の名前に統一
- 旧シート名や日本語名があれば名前で判定
- 名前だけで判定できない場合は、ヘッダー/ラベルと先頭データ行の型・形式で静的検証
- それでも判定しきれない場合のみ、メタシートを除外した従来の並び順で補完
- 必要な列数とヘッダー行を整備
- 既存 PJ の状態列が空欄なら有効として補完
- `AdminWhitelist` と `AdminExecutionLogs` をコピーまたは新規作成
- `AdminWhitelist` が空の場合、実行中の Google アカウントを初期管理者として追加

静的検証では実データを読みますが、戻り値やログにメールアドレス、Slack ID、シート ID、名前などの実値は出しません。返すのはシート名、判定方法、ヘッダー整合、先頭データ行の整合、不一致だった期待列名だけです。

旧DBの代表的なシート名として、Users 相当の `ID`、時給増分相当の `upgrade` も判定対象に含めます。旧ヘッダーの `outputFolderId`、`store`、`時給（税抜き）`、`増分（税抜き）` も現在の列へ対応付けます。Users の `name` は運用上正確でない場合があるため、移行判定では必須の強い根拠にせず、管理画面から後で編集できます。

戻り値の `targetSpreadsheetId` を確認し、問題なければスクリプトプロパティ `DATABASE_SHEET_ID` を移行後 ID に変更します。

確認と切り替えを分けずに行う場合は以下を実行します。

```javascript
migrateDatabaseToNamedSheetsAndUseIt()
```

この関数はコピー後に `DATABASE_SHEET_ID` も移行先へ更新します。実行前に必ずコピー先を作る権限と、移行元 DB を読む権限があるアカウントで実行してください。

移行後のシート認識を確認するには以下を実行します。

```javascript
validateNamedDatabaseSheets()
```

Apps Script エディタの実行ボタンから選ぶ場合は、引数なしの `run...` 関数を使います。

```javascript
runValidateCurrentDatabase()
runValidateMigrationSourceDatabase()
runMigrateDatabaseToNamedSheets()
```

`runValidateMigrationSourceDatabase()` と `runMigrateDatabaseToNamedSheets()` は、スクリプトプロパティ `MIGRATION_SOURCE_SHEET_ID` があればそれを移行元にします。未設定なら `DATABASE_SHEET_ID` を移行元にします。コピー先名を固定したい場合は `MIGRATION_TARGET_SHEET_NAME` を設定します。

フェイクの移行前テストDBに対しては以下を実行します。

```javascript
runValidateFakePreMigrationDatabase()
runMigrateFakePreMigrationDatabase()
```

検証結果では、各シートについて以下を確認できます。

- `matchedBy`: `name_and_data`、`name`、`data`、`legacy_index` などの判定方法
- `headerOk`: ヘッダー/ラベルが想定通りか
- `firstDataRowOk`: 先頭データ行の型・形式が想定通りか
- `mismatchedHeaderColumns`: ヘッダー/ラベルが合わない期待列
- `invalidColumns`: 先頭データ行の形式が合わない期待列

## 画面 URL

| 画面 | URL |
|---|---|
| 勤怠入力フォーム | `/exec` |
| 管理画面 | `/exec?mode=admin` |
| 実績画面 | `/exec?view=worktime` |

実績画面は通常フォームから開きます。URLを直接開く場合は、メールアドレスとパスワードを指定します。

```text
/exec?view=worktime&email={email}&password={password}&year={year}&month={month}
```

## 実績画面

勤務時間、PJ 別金額、役職手当、税込合計をHTML画面で表示します。JSON APIとしての利用は廃止しています。

```text
/exec?view=worktime&email={email}&password={password}&year={year}&month={month}
```

| パラメータ | 必須 | 内容 |
|---|---|---|
| `email` | 必須 | 登録メールアドレス |
| `password` | 必須 | 登録パスワード |
| `year` | 任意 | 対象年。未指定時は当年 |
| `month` | 任意 | 対象月。未指定時は当月 |

対象年月は当月までです。未来月は取得できません。

画面には PJ 明細、役職手当明細、税抜合計、税込合計、対象年月、参考値である旨の注記が表示されます。

## 管理画面

```text
/exec?mode=admin
```

管理画面は `AdminWhitelist` に登録された Google ログイン中メールアドレスだけが利用できます。スクリプトプロパティではなく DB シートで管理します。

主な操作は以下です。

- PJ 追加
- 有効 PJ の割当編集、無効化
- 無効 PJ の展開表示、有効化
- 時給増分のユーザー単位編集
- 役職手当の追加、編集、論理削除
- ユーザー一覧確認、名前編集、論理削除
- 管理者ホワイトリストの追加、有効化、無効化、削除
- 一括発行、一括回収、個別再発行
- 最新まとめ投稿の未リアクションユーザー確認
- 未リアクションユーザーへの確認催促 DM
- 管理操作ログ確認

一括発行、一括回収、個別再発行は先月以前のみ対象です。管理操作は確認モーダルを経由し、実行後は管理画面上部に結果メッセージを表示します。

## 論理削除の扱い

- ユーザー
  - Users シート J 列に削除日時を入れます。
  - フォームログイン、PJ 割当候補、管理画面の通常ユーザー一覧、請求書発行対象から除外されます。
  - 過去の勤怠シートや発行済み PDF は削除しません。
  - 論理削除済みユーザーの請求書再発行はできません。
- 役職手当
  - 役職手当シート E 列に削除日時を入れます。
  - 実績表示と請求書計算から除外されます。
  - 過去に発行済みの請求書 PDF は変更しません。
- PJ
  - 無効化するとフォームの選択肢と管理画面の有効 PJ 一覧から外れます。
  - 過去の勤怠行に記録済みの PJ 名は勤怠シート上には残ります。

## Slack

`src/appsscript.json` は SlackApp ライブラリを利用します。

必要な Bot Token Scopes は以下です。

| Scope | 用途 |
|---|---|
| `chat:write` | チャンネルと DM への投稿 |
| `im:write` | DM チャンネルの作成・取得 |
| `channels:read` | 公開チャンネル情報の取得 |
| `channels:history` | 請求書チャンネルの最新まとめ投稿とリアクション確認 |

請求書チャンネルがプライベートチャンネルの場合は、追加で以下が必要です。

| Scope | 用途 |
|---|---|
| `groups:read` | プライベートチャンネル情報の取得 |
| `groups:history` | プライベートチャンネルの投稿とリアクション確認 |

## clasp 操作

```bash
npm run push:dev
npm run push:prod
npm run pull:dev
npm run pull:prod
```

## 運用メモ

- 管理画面や実績画面への遷移は、現在のデプロイ URL をベースにした `/exec?...` で行います。GAS の `userCodeAppPanel?createOAuthDialog=true` に飛ぶ場合は、フォーム action、リンク生成、リダイレクト先 URL を確認してください。
- 一括系処理は GAS の同期実行です。画面側では待機アニメーションを表示しますが、サーバー側の細かな進捗ストリーミングはできません。
- PJ 選択肢のフロント絞り込みは UX と情報露出の抑制が目的です。最終的な認証、論理削除、PJ 有効状態、対象年月制限はサーバー側でも確認します。
- 管理操作でシート構成を変更した後に業務データ参照がずれた場合は `repairDatabaseSheetOrder()` を実行してください。
