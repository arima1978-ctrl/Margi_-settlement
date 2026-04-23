# eteacher 月次タスクを Windows Task Scheduler に登録
#
# 実行方法: PowerShell を管理者権限で開き、以下を実行
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

# 毎月 1 日 09:00 実行トリガ
$Trigger = New-CimInstance -CimClass (Get-CimClass -ClassName MSFT_TaskMonthlyTrigger -Namespace Root\Microsoft\Windows\TaskScheduler) -ClientOnly
$Trigger.DaysOfMonth   = 1
$Trigger.MonthsOfYear  = 4095   # 全月 (ビットマスク: 2^12 - 1)
$Trigger.StartBoundary = (Get-Date "2026-05-01 09:00:00" -Format "yyyy-MM-ddTHH:mm:ss")
$Trigger.Enabled       = $true

$Action    = New-ScheduledTaskAction -Execute $BatPath
$Principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType S4U -RunLevel Limited
$Settings  = New-ScheduledTaskSettingsSet `
                -AllowStartIfOnBatteries `
                -DontStopIfGoingOnBatteries `
                -StartWhenAvailable `
                -ExecutionTimeLimit (New-TimeSpan -Hours 2)

Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Principal $Principal `
    -Settings $Settings `
    -Description "eteacher 売上管理表 月次自動生成 (毎月1日 09:00)"

Write-Host ""
Write-Host "登録完了: $TaskName"
Write-Host "次回実行: 2026-05-01 09:00 (以降 毎月 1 日 09:00)"
Write-Host ""
Write-Host "確認:     Get-ScheduledTask -TaskName $TaskName"
Write-Host "手動実行: Start-ScheduledTask -TaskName $TaskName"
Write-Host "ログ:     $PSScriptRoot\logs\eteacher_monthly.log"
