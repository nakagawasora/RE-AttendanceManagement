# AM - 勤怠管理 & 請求書自動発行システム

Google Apps Script (GAS) ベースの勤怠管理・請求書自動発行システムです。  
インターン生の勤怠入力をWebフォームで受け付け、月次の請求書を自動生成してSlack/メールで通知します。

## 主な機能

### 勤怠入力 (Webフォーム)
- メール・パスワード認証付きの勤怠入力フォーム (`doGet` / `doPost`)
- プロジェクト選択、作業内容、開始/終了時刻の入力
- 二重勤怠チェック（時間帯の重複検知）
- 一時保存・呼び出し機能（セキュアな保存形式に対応）
- 入力バリデーション・サニタイズ処理

### 請求書自動発行
- 月次の勤怠データを集計し、プロジェクト別・時給別に請求書を自動生成 (`setBill`)
- 通常請求書 / インボイス請求書の自動判別
- 請求書のPDF出力・Google Drive への保存
- 通し番号の自動採番
- 2ヶ月分合算請求書の生成 (`generate2MonthBill`)
- 特定ユーザーの請求書再発行 (`reSetBill`)

### 請求書回収
- 全ユーザーの請求書PDFを月別フォルダに一括コピー (`gatherBills`)

### 通知
- Slack Bot による個別DM・チャンネル通知
- エラー発生時のメールフォールバック

### 月次勤務実績API
- GETリクエストでプロジェクト別勤務時間・合計金額（税込）を取得 (`handleMonthlyWorkTimeApi`)
- メール・パスワード認証付き

## ディレクトリ構成

```
AM/
├── .clasp.dev.json       # clasp設定（開発環境）
├── .clasp.prod.json      # clasp設定（本番環境）
├── .gitignore
├── package.json           # npm scripts（push/pull）
├── README.md
└── src/
    ├── appsscript.json    # GASマニフェスト（SlackAppライブラリ依存）
    ├── BillUtil.js        # 請求書生成・勤務時間計算ユーティリティ
    ├── ManagementUtil.js  # メイン処理（Webフォーム・通知・API）
    ├── index.html         # 勤怠入力フォームテンプレート
    └── submitted.html     # 提出完了ページテンプレート
```

## セットアップ

### 前提条件

- [clasp](https://github.com/google/clasp) がインストール済みであること
- `clasp login` で認証済みであること

### スクリプトプロパティ

GAS のスクリプトプロパティに以下の値を設定してください。

| プロパティ名 | 説明 |
|---|---|
| `SLACK_BOT_TOKEN` | Slack Bot のOAuthトークン |
| `BILL_CHANNEL_ID` | 請求書通知先のSlackチャンネルID |
| `SLACK_ID_OF_ADMIN` | 管理者のSlackユーザーID |
| `MAIL_ADDRESS` | エラー通知先メールアドレス |
| `CC_RECIPIENT` | メール通知のCCアドレス |
| `BILL_FOLDER_ID` | 請求書回収先のGoogle DriveフォルダID |
| `DATABASE_SHEET_ID` | データベースシートのスプレッドシートID |
| `NORMAL_BILL_SHEET_ID` | 通常請求書テンプレートのスプレッドシートID |
| `INVOICE_BILL_SHEET_ID` | インボイス請求書テンプレートのスプレッドシートID |
| `TAX_PERCENTAGE` | 消費税率（デフォルト: `10`） |
| `API_TOKEN` | プログラムから勤務実績APIを取得するための共有トークン |

### データベースシート構成

データベーススプレッドシートは以下のシートで構成されています。

- **Usersシート（1枚目）**: ユーザー情報
  - `A列`: ユーザーID
  - `B列`: 勤怠シートID
  - `C列`: SlackユーザーID
  - `D列`: 請求書出力先フォルダID
  - `E列`: 個人情報シートID
  - `F列`: メールアドレス
  - `G列`: パスワードハッシュ（SHA-256）
  - `H列`: 一時保存データ（JSON）
  - `I列`: 名前
  - `J列`: 論理削除日時（値があれば無効ユーザー）
- **Projectsシート（2枚目）**: プロジェクト情報
  - `A列`: プロジェクトID
  - `B列`: プロジェクト名
  - `C列`: 時給
  - `D列`: アクティブ状態（空欄/TRUE/有効/active/1 を有効として扱う）
  - `E列`: フォーム表示対象user_id（カンマ区切り、空欄なら全ユーザー）
- **時給アップシート（3枚目）**: ユーザー別時給加算情報
  - `A列`: ID
  - `B列`: ユーザーID
  - `C列`: プロジェクトID
  - `D列`: 時給加算額
- **役職手当シート（4枚目）**: ユーザー別役職手当
  - `A列`: ID
  - `B列`: 手当名
  - `C列`: 対象user_id
  - `D列`: 単価
  - `E列`: 論理削除日時（値があれば無効な手当）
- **AdminWhitelistシート（名前付きメタシート）**: 管理画面を許可するGoogleアカウント
  - `A列`: email
  - `B列`: active（空欄/TRUE/有効/active/1 を有効として扱う）
  - `C列`: memo
  - 通常の1〜4枚目シート参照からは除外されます。
- **AdminExecutionLogsシート（名前付きメタシート）**: 管理画面の一括処理ログ
  - 通常の1〜4枚目シート参照からは除外されます。

## 管理画面

```
GET /exec?mode=admin
```

管理画面は `Session.getActiveUser().getEmail()` で取得したGoogleログイン中のメールアドレスが、データベーススプレッドシートの `AdminWhitelist` シートに含まれる場合のみ表示・操作できます。初回はDB権限を持つユーザーが `setupAdminWhitelistSheet()` を実行するか、同名シートを作成して許可メールアドレスを登録してください。

請求書作成時は `NORMAL_BILL_SHEET_ID` / `INVOICE_BILL_SHEET_ID` の雛形を一時コピーし、そのコピーへ書き込んでPDF化した後に一時コピーをゴミ箱へ移動します。雛形本体は書き換えません。

管理画面では以下を操作できます。

- PJ追加、アクティブ/非アクティブ切り替え、フォーム表示対象ユーザーの割当
- 時給増分追加
- ユーザー一覧確認、論理削除
- 役職手当追加
- 管理者ホワイトリスト追加、有効/無効切り替え、削除
- 一括再発行、一括回収、個人再発行
- 請求書チャンネル最新おまとめ投稿のSlack未リアクションユーザー確認

## npm scripts

```bash
# 開発環境にプッシュ
npm run push:dev

# 本番環境にプッシュ
npm run push:prod

# 開発環境からプル
npm run pull:dev

# 本番環境からプル
npm run pull:prod
```

## GASライブラリ依存

| ライブラリ | 用途 |
|---|---|
| [SlackApp](https://github.com/nickovs/SlackApp) | Slack API連携（メッセージ送信） |

### Slack OAuth Scopes

現在の実装で利用するSlack APIに必要なBot Token Scopesは以下です。

| Scope | 用途 |
|---|---|
| `chat:write` | チャンネル/DMへのメッセージ送信 |
| `im:write` | `conversations.open` によるDMチャンネル作成/取得 |
| `channels:read` | 公開チャンネルのメンバー確認 |
| `channels:history` | 請求書チャンネルのおまとめ投稿とリアクション確認 |

請求書チャンネルがプライベートチャンネルの場合は、上記に加えて `groups:read` / `groups:history` が必要です。

## Webアプリのデプロイ

`appsscript.json` の設定に基づき、以下の条件でWebアプリとしてデプロイされます。

- **実行ユーザー**: デプロイしたユーザー (`USER_DEPLOYING`)
- **アクセス範囲**: 全員 (`ANYONE`)

DBスプレッドシート、請求書雛形、Drive、Gmail、Slack連携に必要なGoogle OAuth承認はデプロイしたユーザー側で行います。一般ユーザーにDBスプレッドシート権限は不要です。`executeAs` を `USER_ACCESSING` に変更すると、一般ユーザーがDBやDriveへアクセスできずフォーム/APIが失敗します。

## API仕様

### 月次勤務実績API

勤務時間と合計金額を取得するGET APIです。

**エンドポイント**:  
```
GET /exec?email={email}&password={password}&month={month}&year={year}
```

サーバー間連携などパスワードを使わない取得では、`API_TOKEN` を設定したうえで以下も利用できます。

```
GET /exec?email={email}&apiToken={API_TOKEN}&month={month}&year={year}
```

| パラメータ | 必須 | 説明 |
|---|---|---|
| `email` | ✅ | 登録メールアドレス |
| `password` | ✅ | パスワード（平文） |
| `month` | - | 対象月（省略時は当月） |
| `year` | - | 対象年（省略時は当年） |

**レスポンス例（成功時）**:
```json
{
  "status": "success",
  "data": {
    "workTimeByProject": {
      "プロジェクトA": "10時間30分",
      "プロジェクトB": "5時間0分"
    },
    "roleAllowances": [
      {
        "name": "役職手当",
        "price": 100,
        "workTime": "15時間30分",
        "amount": 1550,
        "amountWithTax": 1705
      }
    ],
    "totalAmount": 17050,
    "month": 6,
    "year": 2026,
    "summary": "2026年6月の勤務実績"
  }
}
```
