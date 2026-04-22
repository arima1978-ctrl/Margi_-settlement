r"""One-shot monthly settlement generator for all 4 services.

Usage
-----
    python scripts/run_monthly.py --month 2026-05
    python scripts/run_monthly.py                 # 引数なし → 次の月を対象
    python scripts/run_monthly.py --yes           # 確認プロンプトをスキップ
    python scripts/run_monthly.py --dry-run       # 実行せず計画のみ表示

何をするか
---------
1. 対象月 (YYYY-MM) の **前月** の源泉ファイル (.xlsm) を自動で探す
2. 4サービス (programming / shogi / bunri / sokudoku) の清算書を順番に生成
3. Y:\ の定位置に ``..._YYYYMM月分.xlsx`` で保存

前月送信分のフォルダや ``.xlsm`` はファイル名がブレるので glob で柔軟に探す。
2026 年以降はファイル名が ``YYYYMM月分`` (ゼロ埋め) の形式だが、古い 2025 年の
ファイルは ``YYYY{M}月分`` (ゼロ埋め無し) のこともあるので両方試す。
"""
from __future__ import annotations

import argparse
import calendar
import os
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path

# Allow running from repo root
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from margin_settlement import run as run_service
from src.notifier import format_run_summary, load_dotenv, send_telegram

# Load .env early so MARGIN_BASE_DIR is respected when the module-level
# BASE_DIR constant is resolved.
load_dotenv(Path(__file__).resolve().parent.parent / ".env", override=True)

# Windows default points to Y:\, Linux deploy overrides via MARGIN_BASE_DIR
# (e.g. /mnt/nas_share/_★20170701作業用/【エデュプラス請求書】).
_DEFAULT_BASE = r"Y:\_★20170701作業用\【エデュプラス請求書】"
BASE_DIR = Path(os.environ.get("MARGIN_BASE_DIR") or _DEFAULT_BASE)
SOURCE_BACKUP_DIR = BASE_DIR / "【業者請求書】エクセルbackup"


@dataclass(frozen=True)
class ServiceSpec:
    service: str          # CLI id (programming / shogi / bunri / sokudoku)
    folder: str           # 清算書フォルダ名
    prefix: str           # ファイル名の先頭 (例: "プログラミング売上管理表_")
    suffix: str = "月分.xlsx"


SERVICES: list[ServiceSpec] = [
    ServiceSpec("programming", "プログラミング清算書",        "プログラミング売上管理表_"),
    ServiceSpec("shogi",       "スマイル将棋清算書",          "スマイル将棋売上管理表_"),
    ServiceSpec("bunri",       "文理ヴィクトリー清算書",      "文理ヴィクトリー売上管理表_"),
    ServiceSpec("sokudoku",    "１００万人の速読　清算書",    "速読_売上管理表_"),
]


def parse_month(s: str) -> date:
    year, month = s.split("-")
    return date(int(year), int(month), 1)


def next_month(d: date) -> date:
    if d.month == 12:
        return date(d.year + 1, 1, 1)
    return date(d.year, d.month + 1, 1)


def prev_month(d: date) -> date:
    if d.month == 1:
        return date(d.year - 1, 12, 1)
    return date(d.year, d.month - 1, 1)


def month_suffixes(year: int, month: int) -> list[str]:
    """Try both '202605' (zero-padded) and '20265' (unpadded) for month lookup."""
    padded = f"{year}{month:02d}"
    unpadded = f"{year}{month}"
    # Deduplicate while preserving order
    return list(dict.fromkeys([padded, unpadded]))


def find_file(folder: Path, prefix: str, year: int, month: int, suffix: str) -> Path | None:
    for mstr in month_suffixes(year, month):
        candidate = folder / f"{prefix}{mstr}{suffix}"
        if candidate.exists():
            return candidate
    return None


def find_source_xlsm(month_date: date) -> tuple[Path, Path]:
    """Return (source_folder, source_xlsm). Source is the PREVIOUS month's 送信分."""
    src_month = prev_month(month_date)
    # Folder names look like "2026年3月21日送信分"
    pattern = f"{src_month.year}年{src_month.month}月*日送信分"
    folders = sorted(SOURCE_BACKUP_DIR.glob(pattern))
    if not folders:
        raise FileNotFoundError(
            f"Source folder not found: {SOURCE_BACKUP_DIR}\\{pattern}"
        )
    src_folder = folders[-1]  # Most recent if multiple

    # .xlsm inside, excluding '入金チェック' variants and '~$' temp lock files
    candidates: list[Path] = []
    for f in src_folder.glob("*.xlsm"):
        if f.name.startswith("~$"):
            continue
        if "入金チェック" in f.name:
            continue
        candidates.append(f)
    if not candidates:
        raise FileNotFoundError(f"No usable .xlsm in {src_folder}")
    # Prefer shortest filename (usually the canonical "送信.xlsm"), else latest mtime
    candidates.sort(key=lambda p: (len(p.name), -p.stat().st_mtime))
    return src_folder, candidates[0]


@dataclass
class Plan:
    service: str
    source_xlsm: Path
    template: Path | None  # None = template missing, service will be skipped
    output: Path
    month: str  # YYYY-MM


def build_plans(month_date: date, source_xlsm: Path) -> list[Plan]:
    plans: list[Plan] = []
    template_month = prev_month(month_date)
    for spec in SERVICES:
        folder = BASE_DIR / spec.folder
        template = find_file(folder, spec.prefix, template_month.year, template_month.month, spec.suffix)
        output = folder / f"{spec.prefix}{month_date.year}{month_date.month:02d}{spec.suffix}"
        plans.append(Plan(
            service=spec.service,
            source_xlsm=source_xlsm,
            template=template,
            output=output,
            month=f"{month_date.year}-{month_date.month:02d}",
        ))
    return plans


def print_plan(month_date: date, source_folder: Path, source_xlsm: Path, plans: list[Plan]) -> None:
    print("=" * 80)
    print(f"対象月:        {month_date.year}年{month_date.month}月分")
    print(f"源泉フォルダ:  {source_folder}")
    print(f"源泉 xlsm:     {source_xlsm.name}")
    print("=" * 80)
    print("\n生成するファイル:")
    for p in plans:
        if p.template is None:
            print(f"  [{p.service:11s}] テンプレ未検出 → SKIP (前月分ファイルが Y:\\ に無い)")
            continue
        exists = " ← 既存を上書き" if p.output.exists() else ""
        print(f"  [{p.service:11s}] {p.output}{exists}")
        print(f"  {'':13s}   template: {p.template.name}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--month", help="対象月 YYYY-MM (省略時は今月の翌月)")
    parser.add_argument("--yes", "-y", action="store_true", help="確認プロンプトをスキップ")
    parser.add_argument("--dry-run", action="store_true", help="計画表示のみで実行しない")
    parser.add_argument("--only", nargs="+", choices=[s.service for s in SERVICES],
                        help="指定サービスだけ実行")
    parser.add_argument("--skip-google-sheet", action="store_true",
                        help="Google Sheet 照会をスキップ (programming の新規塾検出)")
    parser.add_argument("--notify", action="store_true",
                        help="Telegram に実行結果を通知 (env: TELEGRAM_BOT_TOKEN/TELEGRAM_CHAT_ID)")
    args = parser.parse_args()


    if args.month:
        month_date = parse_month(args.month)
    else:
        month_date = next_month(date.today().replace(day=1))
        print(f"[info] --month 未指定 → 翌月 {month_date.year}-{month_date.month:02d} を対象にします")

    try:
        source_folder, source_xlsm = find_source_xlsm(month_date)
    except FileNotFoundError as e:
        print(f"\nERROR: {e}", file=sys.stderr)
        return 2

    plans = build_plans(month_date, source_xlsm)
    if args.only:
        plans = [p for p in plans if p.service in args.only]

    print_plan(month_date, source_folder, source_xlsm, plans)

    if args.dry_run:
        print("\n[--dry-run] 何も実行しませんでした。")
        return 0

    if not args.yes:
        reply = input("\n続行しますか? [y/N]: ").strip().lower()
        if reply not in ("y", "yes"):
            print("キャンセルしました。")
            return 1

    # Execute
    results: list[tuple[str, bool, str]] = []
    for plan in plans:
        if plan.template is None:
            results.append((plan.service, False, "template missing → SKIP"))
            continue
        print("\n" + "=" * 80)
        print(f"[{plan.service}] 生成開始...")
        print("=" * 80)
        try:
            run_service(
                service=plan.service,
                source=str(plan.source_xlsm),
                template=str(plan.template),
                output=str(plan.output),
                month=plan.month,
                skip_google_sheet=args.skip_google_sheet,
            )
            results.append((plan.service, True, str(plan.output)))
        except Exception as e:
            results.append((plan.service, False, f"{type(e).__name__}: {e}"))
            print(f"\n[{plan.service}] FAILED: {e}", file=sys.stderr)

    # Final summary
    print("\n" + "=" * 80)
    print("実行サマリ")
    print("=" * 80)
    ok_count = sum(1 for _, ok, _ in results if ok)
    fail_count = len(results) - ok_count
    for svc, ok, info in results:
        mark = "OK " if ok else "NG "
        print(f"  {mark} {svc:11s}  {info}")
    print(f"\n成功 {ok_count}/{len(results)}, 失敗 {fail_count}")

    if args.notify:
        month_str = f"{month_date.year}-{month_date.month:02d}"
        summary = format_run_summary(month_str, results)
        sent = send_telegram(summary)
        if sent:
            print("\nTelegram 通知送信しました。")
        else:
            print("\nWARN: Telegram 通知送信に失敗しました (認証情報未設定 or ネットワーク障害)。", file=sys.stderr)

    return 0 if fail_count == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
