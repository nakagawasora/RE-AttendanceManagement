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
| `RESET_NORMAL_BILL_SHEET_ID` | 通常請求書初期化用シートID |
| `RESET_INVOICE_BILL_SHEET_ID` | インボイス請求書初期化用シートID |
| `TAX_PERCENTAGE` | 消費税率（デフォルト: `10`） |

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
- **Projectsシート（2枚目）**: プロジェクト情報
  - `A列`: プロジェクトID
  - `B列`: プロジェクト名
  - `C列`: 時給
- **時給アップシート（3枚目）**: ユーザー別時給加算情報
  - `A列`: ID
  - `B列`: ユーザーID
  - `C列`: プロジェクトID
  - `D列`: 時給加算額

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

## Webアプリのデプロイ

`appsscript.json` の設定に基づき、以下の条件でWebアプリとしてデプロイされます。

- **実行ユーザー**: デプロイしたユーザー (`USER_DEPLOYING`)
- **アクセス範囲**: 全員 (`ANYONE`)

## API仕様

### 月次勤務実績API

勤務時間と合計金額を取得するGET APIです。

**エンドポイント**:  
```
GET /exec?email={email}&password={password}&month={month}&year={year}
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
    "totalAmount": 17050,
    "month": 6,
    "year": 2026,
    "summary": "2026年6月の勤務実績"
  }
}
```
