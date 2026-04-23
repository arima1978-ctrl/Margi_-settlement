# eteacher 月次タスクを Windows Task Scheduler に登録 (schtasks.exe 経由)
#
# 実行方法:
#   cd C:\Users\USER\projects\Margi_-settlement
#   .\register_eteacher_task.ps1
#
# 毎月 1 日 09:00 に eteacher_monthly.bat を実行

$TaskName = "EteacherMonthly"
$BatPath  = "$PSScriptRoot\eteacher_monthly.bat"

if (-not (Test-Path $BatPath)) {
    Write-Error "$BatPath が見つかりません"
    exit 1
}

# 既存タスクを削除 (存在しなければ無視)
schtasks.exe /Delete /TN $TaskName /F 2>$null | Out-Null

# 引数を配列で構築し、& で呼び出す (バックチック行継続を避けるため)
$schtasksArgs = @(
    "/Create",
    "/TN", $TaskName,
    "/TR", "`"$BatPath`"",
    "/SC", "MONTHLY",
    "/D",  "1",
    "/ST", "09:00",
    "/RL", "LIMITED",
    "/F"
)

& schtasks.exe @schtasksArgs

if ($LASTEXITCODE -ne 0) {
    Write-Error "schtasks.exe 登録に失敗: ExitCode=$LASTEXITCODE"
    exit 1
}

Write-Host ""
Write-Host "登録完了: $TaskName"
Write-Host "次回実行: 次の月 1 日 09:00 (以降 毎月 1 日 09:00)"
Write-Host ""
Write-Host "確認:     schtasks /Query /TN $TaskName /V /FO LIST"
Write-Host "手動実行: schtasks /Run /TN $TaskName"
Write-Host "削除:     schtasks /Delete /TN $TaskName /F"
Write-Host "ログ:     $PSScriptRoot\logs\eteacher_monthly.log"
