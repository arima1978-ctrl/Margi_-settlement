# Margi_-settlement

名大SKYの月次マージン清算書を自動生成するツール。

対象サービス: プログラミング / 将棋 / 文理ヴィクトリー / 速読

## 使い方

### プログラミング清算書の生成

```bash
python margin_settlement.py programming \
  --source  "Y:\_★20170701作業用\【エデュプラス請求書】\【業者請求書】エクセルbackup\2026年3月21日送信分\2026.3.23送信.xlsm" \
  --template "Y:\_★20170701作業用\【エデュプラス請求書】\プログラミング清算書\プログラミング売上管理表_202603月分.xlsx" \
  --output   "Y:\_★20170701作業用\【エデュプラス請求書】\プログラミング清算書\プログラミング売上管理表_202604月分.xlsx" \
  --month    2026-04
```

### 対応表（source → settlement）

source .xlsm は **1ヶ月前の送信分** を使う。

| 清算書 | source folder |
|---|---|
| 202603月分（3月分） | 2026年2月17日送信分 |
| 202604月分（4月分） | 2026年3月21日送信分 |
| 202605月分（5月分） | 2026年4月18日送信分 |

### 引数

| 引数 | 説明 |
|---|---|
| `--source`   | 当月の業者請求書 `.xlsm` |
| `--template` | 前月の清算書 `.xlsx`（雛形として使用） |
| `--output`   | 出力先のパス |
| `--month`    | 対象月 `YYYY-MM` |

## 自動処理内容（プログラミング）

1. 前月清算書をコピーして新ファイル作成
2. 7シートを source から上書き:
   - `らくらく ユーザー基本情報貼り付ける`
   - `保護者情報DL貼付⑩AKへ`
   - `④_2プロ_管理者ＩＤ` ← source `④プロ_管理者ＩＤ`
   - `④_3プロ_生徒ＩＤ` ← source `④プロ_生徒ＩＤ`
   - `④ゲームクリエイター生徒ID`
   - `④_4カルチャ加盟金`
   - `④カルチャー_基本料金`
3. `報告書!D1` に対象月を設定
4. 新規家族IDを検出 → [Googleスプレッドシート](https://docs.google.com/spreadsheets/d/1fT5niRMqfvdIMm0GuW0l1_fetHS8eoEfdU84BJLOk6k) から情報取得
5. 新規塾行を `プログラミング営業管理` と `報告書` に追加

## 手動作業として残る項目

- `報告書!AJ3`（名大SKY直営分マイクラID数）は手入力

## セットアップ

```bash
pip install -r requirements.txt
```

### Google Sheets API認証（新規塾自動追加を使う場合）

サービスアカウント JSON を `credentials.json` として配置するか、環境変数 `GOOGLE_APPLICATION_CREDENTIALS` でパス指定。

## 検証

```bash
# 3月分を再現して既存ファイルと比較
python scripts/verify_march.py
```
