# eteacher 修正の 3 連コマンド (apply → refresh → check)
# Usage: .\eteacher_apply.ps1

$corrections = "C:\tmp\margin_inspect\id_name_mismatch_full.xlsx"
$target      = "Y:\_★20170701作業用\9三浦\eteacher売上管理表2026年4月.xlsx"
$source      = "Y:\_★20170701作業用\【エデュプラス請求書】\【業者請求書】エクセルbackup\2026年4月18日送信分\2026年4月17日送信.xlsm"

Write-Host "=== 1/3: 修正家族IDを eteacher に反映 ===" -ForegroundColor Cyan
python scripts/apply_eteacher_corrections.py --corrections $corrections --target $target
if ($LASTEXITCODE -ne 0) { Write-Host "apply 失敗" -ForegroundColor Red; exit 1 }

Write-Host ""
Write-Host "=== 2/3: 売上を家族IDベースで再反映 ===" -ForegroundColor Cyan
python scripts/refresh_eteacher.py --target $target --source $source
if ($LASTEXITCODE -ne 0) { Write-Host "refresh 失敗" -ForegroundColor Red; exit 1 }

Write-Host ""
Write-Host "=== 3/3: 取りこぼし最終チェック ===" -ForegroundColor Cyan
python scripts/check_eteacher_missing.py --target $target --source $source
