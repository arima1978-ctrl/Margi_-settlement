@echo off
REM eteacher 月次自動生成 (Windows Task Scheduler 用)
REM 毎月 1 日 09:00 に実行 → 前月分を生成 → Telegram 通知
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set PYTHONUNBUFFERED=1
python -u scripts\eteacher_monthly.py --notify >> logs\eteacher_monthly.log 2>&1
exit /b %ERRORLEVEL%
