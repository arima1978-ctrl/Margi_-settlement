# eteacher の refresh + check だけ実行 (apply なし)
$target = "Y:\_★20170701作業用\9三浦\eteacher売上管理表2026年4月.xlsx"
$source = "Y:\_★20170701作業用\【エデュプラス請求書】\【業者請求書】エクセルbackup\2026年4月18日送信分\2026年4月17日送信.xlsm"

Write-Host "=== 1/2: 売上を家族IDベースで再反映 ===" -ForegroundColor Cyan
python scripts/refresh_eteacher.py --target $target --source $source
if ($LASTEXITCODE -ne 0) { Write-Host "refresh 失敗" -ForegroundColor Red; exit 1 }

Write-Host ""
Write-Host "=== 2/2: 取りこぼし最終チェック ===" -ForegroundColor Cyan
python scripts/check_eteacher_missing.py --target $target --source $source
