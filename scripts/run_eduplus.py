"""Eduplus 月次集計 CLI.

Usage
-----
    # 最新の送信分 .xlsm を自動検出して処理
    python scripts/run_eduplus.py

    # 特定の .xlsm を明示
    python scripts/run_eduplus.py --source "Y:\\...\\2026年4月18日送信分\\2026年4月17日送信.xlsm"

    # バックアップ無しで処理
    python scripts/run_eduplus.py --no-backup

    # Telegram 通知を有効化
    python scripts/run_eduplus.py --notify
"""
from __future__ import annotations

import argparse
import os
import sys
from datetime import date
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from src.eduplus_processor import format_summary, process_eduplus
from src.notifier import load_dotenv, send_telegram

load_dotenv(REPO_ROOT / ".env", override=True)

_DEFAULT_BASE = r"Y:\_★20170701作業用\【エデュプラス請求書】"
BASE_DIR = Path(os.environ.get("MARGIN_BASE_DIR") or _DEFAULT_BASE)
SOURCE_BACKUP_DIR = BASE_DIR / "【業者請求書】エクセルbackup"


def prev_month(d: date) -> date:
    if d.month == 1:
        return date(d.year - 1, 12, 1)
    return date(d.year, d.month - 1, 1)


def find_latest_source_xlsm(target_month: date) -> Path:
    """Locate the source .xlsm for the target month's previous-month 送信分."""
    src_month = prev_month(target_month)
    pattern = f"{src_month.year}年{src_month.month}月*日送信分"
    folders = sorted(SOURCE_BACKUP_DIR.glob(pattern))
    if not folders:
        raise FileNotFoundError(f"No source folder found: {SOURCE_BACKUP_DIR}\\{pattern}")
    src_folder = folders[-1]

    # Exclude 入金チェック variants and Office lock files
    candidates: list[Path] = []
    for f in src_folder.glob("*.xlsm"):
        if f.name.startswith("~$") or "入金チェック" in f.name:
            continue
        candidates.append(f)
    if not candidates:
        raise FileNotFoundError(f"No usable .xlsm in {src_folder}")
    # Prefer the shortest filename (canonical "送信.xlsm"), tie-break by mtime
    candidates.sort(key=lambda p: (len(p.name), -p.stat().st_mtime))
    return candidates[0]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--source", help="対象の .xlsm パス (省略時は今月の前月送信分を自動検出)")
    parser.add_argument("--month", help="対象月 YYYY-MM (--source 省略時のみ使用、既定=今月)")
    parser.add_argument("--no-backup", action="store_true", help=".xlsm.bak を作らない")
    parser.add_argument("--notify", action="store_true",
                        help="Telegram に実行結果を通知 (env: TELEGRAM_BOT_TOKEN/TELEGRAM_CHAT_ID)")
    args = parser.parse_args()

    if args.source:
        source_path = Path(args.source)
    else:
        if args.month:
            y, m = args.month.split("-")
            target = date(int(y), int(m), 1)
        else:
            today = date.today()
            target = date(today.year, today.month, 1)
        source_path = find_latest_source_xlsm(target)
        print(f"[info] source auto-detected: {source_path}")

    print(f"Processing: {source_path}")
    try:
        result = process_eduplus(source_path, backup=not args.no_backup)
    except Exception as exc:
        err_msg = f"【eduplus 集計 FAILED】{type(exc).__name__}: {exc}"
        print(err_msg, file=sys.stderr)
        if args.notify:
            send_telegram(err_msg)
        return 1

    summary = format_summary(result)
    print()
    print(summary)

    if args.notify:
        sent = send_telegram(summary)
        if sent:
            print("\nTelegram 通知送信しました。")
        else:
            print("\nWARN: Telegram 通知送信に失敗しました。", file=sys.stderr)

    return 0


if __name__ == "__main__":
    sys.exit(main())
