# 192.168.1.16 への本番デプロイ

Ubuntu Server 25 (skyuser@ubuntuserver25-2) に margin-settlement を
デプロイする手順。ブラウザ UI と毎月1日の自動実行を両方有効化する。

Y:\ は `/mnt/nas_share/` に既にマウントされている前提。

## 1. セットアップ

```bash
# ログイン
ssh skyuser@192.168.1.16

# リポジトリをクローン
cd ~
git clone https://github.com/arima1978-ctrl/Margi_-settlement.git margin-settlement
cd margin-settlement

# venv 作成
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# .env を配置 (パスワード・トークンは Windows の .env からコピー)
cp .env.example .env
nano .env
# → MARGIN_BASE_DIR=/mnt/nas_share/_★20170701作業用/【エデュプラス請求書】
# → TELEGRAM_BOT_TOKEN=...
# → TELEGRAM_CHAT_ID=8620367664
# → WEB_BASIC_AUTH_USERS=arima:強いパスワード,miura:強いパスワード

# 動作確認 (dry-run)
python scripts/run_monthly.py --month 2026-05 --only programming --dry-run
```

## 2. Web UI を systemd サービス化

```bash
# systemd サービスファイルを配置
sudo cp deploy/margin-settlement-web.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now margin-settlement-web.service

# 状態確認
sudo systemctl status margin-settlement-web.service
journalctl -u margin-settlement-web.service -f   # ライブログ

# ブラウザから確認
# http://192.168.1.16:8081/  (Basic 認証: arima / miura)
```

ポート 8081 をファイアウォールで開放する必要があるかも:

```bash
sudo ufw allow 8081/tcp
```

## 3. 毎月1日 09:00 に自動実行する cron

```bash
# skyuser の crontab に追加
crontab -e
# → deploy/margin-settlement.crontab の内容をコピペ

# または直接反映
crontab deploy/margin-settlement.crontab

# 確認
crontab -l

# ログ出力先を用意
mkdir -p ~/margin-settlement/logs
```

## 4. 動作テスト

### Web UI テスト
ブラウザで `http://192.168.1.16:8081/` を開き、以下を確認:

- [ ] Basic 認証が効く (arima / miura でログイン可)
- [ ] 対象月を選んで「生成実行」を押すとログが流れる
- [ ] 完了後に「最近生成されたファイル」にリストアップされる
- [ ] [ダウンロード] リンクからファイルを落とせる
- [ ] 実行後に Telegram に通知が届く

### cron テスト (1日を待たずに即実行)

```bash
# 今月分をその場で実行して通知確認
cd ~/margin-settlement
./.venv/bin/python scripts/run_monthly.py --yes --notify --month $(date +%Y-%m)
```

## 5. アップデート手順

```bash
ssh skyuser@192.168.1.16
cd ~/margin-settlement
git pull
source .venv/bin/activate
pip install -r requirements.txt   # 依存更新があれば
sudo systemctl restart margin-settlement-web.service
```

## トラブルシューティング

### Web UI にアクセスできない
- ファイアウォール: `sudo ufw status`
- サービス動作: `sudo systemctl status margin-settlement-web.service`
- ログ: `journalctl -u margin-settlement-web.service -n 50`

### Telegram に届かない
- `.env` の `TELEGRAM_BOT_TOKEN` / `TELEGRAM_CHAT_ID` を確認
- 手動テスト: `python -c "from src.notifier import load_dotenv, send_telegram; load_dotenv('.env', override=True); print(send_telegram('test'))"`

### cron が動いていない
- `grep CRON /var/log/syslog | tail -20`
- ログファイル: `~/margin-settlement/logs/cron-YYYYMM.log`

### Y:\ のマウントが外れている
- `ls /mnt/nas_share/_★20170701作業用/` でファイルが見えるか
- 見えない場合: `mount /mnt/nas_share` または `/etc/fstab` を確認
