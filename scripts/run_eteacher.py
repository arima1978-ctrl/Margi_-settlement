"""eteacher売上管理表 の月次更新 CLI.

Usage
-----
    # 4月17日送信.xlsm から 2026-04 月分を生成 (3月テンプレ使用)
    python scripts/run_eteacher.py --month 2026-04

    # 明示的にファイル指定
    python scripts/run_eteacher.py --month 2026-04 \
      --source "Y:\\...\\2026年4月17日送信.xlsm" \
      --template "Y:\\...\\eteacher売上管理表2026年3月.xls" \
      --output "Y:\\...\\eteacher売上管理表2026年4月.xlsx"

    # Telegram 通知
    python scripts/run_eteacher.py --month 2026-04 --notify
"""
from __future__ import annotations

import argparse
import os
import sys
from datetime import date
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from src.eteacher_updater import (
    EteacherUpdateResult,
    compute_sales_by_shop,
    convert_xls_to_xlsx,
    format_eteacher_summary,
    update_eteacher,
    write_unmatched_report,
)
from src.notifier import load_dotenv, send_telegram

load_dotenv(REPO_ROOT / ".env", override=True)

_DEFAULT_BASE = r"Y:\_★20170701作業用\【エデュプラス請求書】"
BASE_DIR = Path(os.environ.get("MARGIN_BASE_DIR") or _DEFAULT_BASE)
SOURCE_BACKUP_DIR = BASE_DIR / "【業者請求書】エクセルbackup"
# eteacher files live under 9三浦, which sits next to 【エデュプラス請求書】 on the share
ETEACHER_DIR = BASE_DIR.parent / "9三浦"


def prev_month(d: date) -> date:
    if d.month == 1:
        return date(d.year - 1, 12, 1)
    return date(d.year, d.month - 1, 1)


def find_source_xlsm(month_date: date) -> Path:
    src_month = prev_month(month_date)
    pattern = f"{src_month.year}年{src_month.month}月*日送信分"
    folders = sorted(SOURCE_BACKUP_DIR.glob(pattern))
    if not folders:
        raise FileNotFoundError(f"源泉フォルダ未検出: {SOURCE_BACKUP_DIR}\\{pattern}")
    candidates: list[Path] = []
    for f in folders[-1].glob("*.xlsm"):
        if f.name.startswith("~$") or "入金チェック" in f.name:
            continue
        candidates.append(f)
    if not candidates:
        raise FileNotFoundError(f".xlsm 未検出: {folders[-1]}")
    candidates.sort(key=lambda p: (len(p.name), -p.stat().st_mtime))
    return candidates[0]


def find_template_xls(month_date: date) -> Path:
    """eteacher売上管理表{YYYY}年{M}月.xls の前月分を探す (MDPS 無しの本ファイル)."""
    prev = prev_month(month_date)
    # 既存ファイルの命名は 2026年3月.xls のように 月は1〜2桁 (ゼロ埋めなし)
    candidate = ETEACHER_DIR / f"eteacher売上管理表{prev.year}年{prev.month}月.xls"
    if candidate.exists():
        return candidate
    # フォールバックで ゼロ埋め
    alt = ETEACHER_DIR / f"eteacher売上管理表{prev.year}年{prev.month:02d}月.xls"
    if alt.exists():
        return alt
    raise FileNotFoundError(f"前月テンプレ未検出: {candidate}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--month", required=False, help="対象月 YYYY-MM (省略時は今月)")
    parser.add_argument("--source", help="源泉 .xlsm のパス (省略時は前月送信分を自動検出)")
    parser.add_argument("--template", help="前月の eteacher .xls テンプレのパス (省略時は自動)")
    parser.add_argument("--output", help="出力 .xlsx のパス (省略時は 9三浦 フォルダに自動命名)")
    parser.add_argument("--no-insert", action="store_true",
                        help="家族ID 列を挿入しない (デバッグ用)")
    parser.add_argument("--notify", action="store_true", help="Telegram に結果通知")
    args = parser.parse_args()

    if args.month:
        y, m = args.month.split("-")
        target = date(int(y), int(m), 1)
    else:
        t = date.today()
        target = date(t.year, t.month, 1)

    try:
        source_xlsm = Path(args.source) if args.source else find_source_xlsm(target)
        template_xls = Path(args.template) if args.template else find_template_xls(target)
    except FileNotFoundError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    if args.output:
        output_xlsx = Path(args.output)
    else:
        output_xlsx = ETEACHER_DIR / f"eteacher売上管理表{target.year}年{target.month}月.xlsx"

    print(f"対象月:    {target.year}-{target.month:02d}")
    print(f"source:    {source_xlsm}")
    print(f"template:  {template_xls}")
    print(f"output:    {output_xlsx}")
    print()

    # Step 1: compute sales per shop from processed xlsm
    print("[1/3] マージン計算用 から 塾別売上を算出中...")
    sales = compute_sales_by_shop(source_xlsm)
    print(f"     {len(sales)} 塾の売上データを読み込み")

    # Step 2: convert .xls template -> .xlsx output (preserving formatting via Excel COM)
    print("[2/3] .xls テンプレを .xlsx に変換中 (Excel COM)...")
    convert_xls_to_xlsx(template_xls, output_xlsx)

    # Step 3: insert 家族ID column and write this month's data
    print("[3/3] 家族ID列挿入 + 売上反映...")
    result = update_eteacher(
        xlsx_path=output_xlsx,
        sales=sales,
        insert_family_id_col=not args.no_insert,
    )
    # Keep the original template path for display
    result.template_path = template_xls

    summary = format_eteacher_summary(result)
    print()
    print(summary)

    # Write full unmatched list for operator review (side-by-side with output)
    report_path = output_xlsx.with_suffix(".unmatched.txt")
    write_unmatched_report(result, report_path)
    print(f"\n未マッチ塾リスト: {report_path}")

    if args.notify:
        sent = send_telegram(summary)
        print(f"\nTelegram: {'送信OK' if sent else '送信失敗'}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
