"""eteacher売上管理表 の月次更新ツール.

Workflow
--------
1. マージン計算用 の各行から (塾名, 家族ID, T=売上) を算出
   - T = ROUNDDOWN((初期費用 + 基本料金 + 利用料金 + 解約分) * 1.1, 0)
   - 各成分は eduplus_processor が書き込んだ 4 シートの N/O を Python で
     VLOOKUP 相当の処理で取得する
2. 3月.xls テンプレ を .xlsx にコピー/変換
3. openpyxl で D列の前に「家族ID」列を新規挿入(既存データ 1列ずつ右シフト)
4. 各データ行で塾名照合 → 新 D 列に家族ID、新 AO 列に売上を書く
5. 新 r7 に「入金日」「売上額」ヘッダーを追加
"""
from __future__ import annotations

import math
import shutil
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import openpyxl


# ---- lookup from processed xlsm -----------------------------------------

_SHEET_START_ROWS = {
    "⑤Edu　ID利用料":          5,
    "⑥Edu　基本料金DLデータ":  4,
    "⑥-2 Edu初期費用　":       4,
    "⑥-3 Edu解約塾":            4,
}


def _build_family_amount_lookup(wb: openpyxl.Workbook, sheet_name: str) -> dict[int, float]:
    """Read N (family_id) / O (amount) columns from a source sheet.

    Our eduplus_processor writes unique (family_id, sum_amount) pairs
    starting at that sheet's data_start_row, so no VLOOKUP is needed.
    """
    if sheet_name not in wb.sheetnames:
        return {}
    start_row = _SHEET_START_ROWS.get(sheet_name, 4)
    ws = wb[sheet_name]
    out: dict[int, float] = {}
    for row in range(start_row, ws.max_row + 1):
        key = ws.cell(row=row, column=14).value  # N
        val = ws.cell(row=row, column=15).value  # O
        if key is None or key == "":
            continue
        try:
            fid = int(key)
        except (TypeError, ValueError):
            continue
        if isinstance(val, (int, float)):
            out[fid] = float(val)
        else:
            out[fid] = 0.0
    return out


def _roundown_toward_zero(x: float) -> int:
    """Excel-style ROUNDDOWN(x, 0) — truncate toward zero."""
    return math.trunc(x)


@dataclass
class ShopSales:
    family_id: int
    sales: int     # ROUNDDOWN(base * 1.1, 0)
    base: float    # 税抜 = O + P + Q + R


def _build_hogosha_name_lookup(wb: openpyxl.Workbook) -> dict[int, str]:
    """Build a 家族ID -> 塾名 lookup from 保護者情報DL貼付⑩AKへ.

    Column A holds family ID, column B holds the shop (customer) name.
    Used as a fallback when ``マージン計算用`` D column is a formula whose
    cached value is not available.
    """
    sheet_name = "保護者情報DL貼付⑩AKへ"
    if sheet_name not in wb.sheetnames:
        return {}
    ws = wb[sheet_name]
    out: dict[int, str] = {}
    for row in range(2, ws.max_row + 1):
        k = ws.cell(row=row, column=1).value
        v = ws.cell(row=row, column=2).value
        if k is None or v is None:
            continue
        try:
            fid = int(k)
        except (TypeError, ValueError):
            continue
        if isinstance(v, str):
            out[fid] = v.strip()
    return out


def compute_sales_by_shop(xlsm_path: str | Path) -> dict[str, ShopSales]:
    """Return a dict mapping 塾名 -> ShopSales computed from the processed xlsm.

    Reads with ``data_only=True`` so cached formula values from マージン計算用
    D列 (VLOOKUP で 塾名 を導出している行) も取れる。キャッシュが空の場合は
    保護者情報DL貼付⑩AKへ を直接引いて 塾名 を解決する。

    塾名 is trimmed of surrounding whitespace for stable matching.
    """
    # data_only=True: cached formula values. Any NOT-yet-calculated cells
    # will fall through to the 保護者情報 fallback below.
    wb = openpyxl.load_workbook(xlsm_path, data_only=True, keep_vba=True)

    shoki = _build_family_amount_lookup(wb, "⑥-2 Edu初期費用　")
    kihon = _build_family_amount_lookup(wb, "⑥Edu　基本料金DLデータ")
    riyo = _build_family_amount_lookup(wb, "⑤Edu　ID利用料")
    kaiyaku = _build_family_amount_lookup(wb, "⑥-3 Edu解約塾")
    hogosha_names = _build_hogosha_name_lookup(wb)

    ws = wb["マージン計算用"]
    out: dict[str, ShopSales] = {}
    for row in range(12, ws.max_row + 1):
        raw_fid = ws.cell(row=row, column=5).value  # E = 家族ID (usually literal int)
        try:
            fid = int(raw_fid) if raw_fid is not None else None
        except (TypeError, ValueError):
            fid = None
        if fid is None:
            continue

        # Resolve 塾名: prefer cached value from D, fall back to 保護者情報 lookup.
        name_raw = ws.cell(row=row, column=4).value
        name: str | None = None
        if isinstance(name_raw, str) and name_raw.strip() and not name_raw.startswith("="):
            name = name_raw.strip()
        else:
            name = hogosha_names.get(fid)
        if not name:
            continue

        base = (
            shoki.get(fid, 0.0)
            + kihon.get(fid, 0.0)
            + riyo.get(fid, 0.0)
            + kaiyaku.get(fid, 0.0)
        )
        t = _roundown_toward_zero(base * 1.1)
        out[name] = ShopSales(family_id=fid, sales=t, base=base)

    wb.close()
    return out


# ---- .xls -> .xlsx conversion via Excel COM -----------------------------

def convert_xls_to_xlsx(src_xls: str | Path, dst_xlsx: str | Path) -> None:
    """Convert .xls -> .xlsx using Excel via win32com.

    Requires Microsoft Excel installed on the machine. This is run on the
    user's Windows PC, not on the Linux cron server.
    """
    import win32com.client  # type: ignore[import]

    src_abs = str(Path(src_xls).resolve())
    dst_abs = str(Path(dst_xlsx).resolve())
    dst_parent = Path(dst_xlsx).resolve().parent
    dst_parent.mkdir(parents=True, exist_ok=True)

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(src_abs)
        # FileFormat 51 = xlOpenXMLWorkbook (.xlsx)
        wb.SaveAs(dst_abs, FileFormat=51)
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()


# ---- eteacher .xlsx modification ----------------------------------------

@dataclass
class EteacherUpdateResult:
    template_path: Path
    output_path: Path
    matched: list[tuple[str, int, int]] = field(default_factory=list)   # (塾名, 家族ID, 売上)
    unmatched_in_eteacher: list[str] = field(default_factory=list)       # 塾名 in eteacher but not in source
    unmatched_in_source: list[str] = field(default_factory=list)         # 塾名 in source but not in eteacher
    skipped_no_name: int = 0


# Column indices (1-based) in the 3月 eteacher template (before shift).
_COL_SHOP_NAME = 4   # D = 塾名
_COL_HEADER_ROW = 6  # r6 has '塾名' header
_COL_MONTH_HEADER_ROW = 7  # r7 has 入金日/売上額 pattern


def update_eteacher(
    xlsx_path: str | Path,
    sales: dict[str, ShopSales],
    insert_family_id_col: bool = True,
) -> EteacherUpdateResult:
    """Insert 家族ID column before D, then write this month's 売上 to AO.

    Modifies ``xlsx_path`` in place — the caller is responsible for making
    a copy of the template first (via ``convert_xls_to_xlsx``).

    Layout after insertion (1-based columns):
      A, B, C: unchanged
      D:  家族ID (new, populated by shop-name lookup)
      E:  塾名 (was D in template)
      F.. onwards: shifted from old E onwards
      old AL -> new AM (last historical month moves to AM)
      new AN: this month's 入金日 (left blank)
      new AO: this month's 売上 (from sales dict)
    """
    output_xlsx = Path(xlsx_path)

    wb = openpyxl.load_workbook(output_xlsx)
    ws = wb.active  # Sheet2 in the template

    result = EteacherUpdateResult(template_path=output_xlsx, output_path=output_xlsx)

    # Snapshot the old 塾名 column before insertion.
    # Row 7 and below: r8+ are data; r7 has month headers; r6 has col headers.
    old_shop_names: dict[int, str] = {}
    max_row = ws.max_row
    for row in range(7, max_row + 1):
        v = ws.cell(row=row, column=_COL_SHOP_NAME).value
        if isinstance(v, str) and v.strip():
            old_shop_names[row] = v.strip()

    if insert_family_id_col:
        ws.insert_cols(_COL_SHOP_NAME, 1)   # insert before old D; old D becomes E
        family_id_col = _COL_SHOP_NAME       # new D
        # r6 header
        ws.cell(row=_COL_HEADER_ROW, column=family_id_col).value = "家族ID"
    else:
        family_id_col = None

    # After insertion (if done), 塾名 column is at old D + 1 = 5. If not inserted, it's still 4.
    shop_name_col = _COL_SHOP_NAME + (1 if insert_family_id_col else 0)

    # New month columns: pushed right by 1 if we inserted.
    # Template ends at AL (col 38). After insert, last existing data col is AM (col 39).
    # We write:  AN (col 40) = 入金日 blank,  AO (col 41) = 売上
    if insert_family_id_col:
        new_month_date_col = 40  # AN
        new_month_sales_col = 41  # AO
    else:
        new_month_date_col = 39  # AM
        new_month_sales_col = 40  # AN

    # Set month headers (r7)
    ws.cell(row=_COL_MONTH_HEADER_ROW, column=new_month_date_col).value = "入金日"
    ws.cell(row=_COL_MONTH_HEADER_ROW, column=new_month_sales_col).value = "売上額"

    # Fill data rows
    matched_names: set[str] = set()
    for row_idx, shop_name in old_shop_names.items():
        if row_idx < 8:
            # r7 or higher header; skip data-fill
            continue
        info = sales.get(shop_name)
        if info is None:
            # No match — log for user review
            result.unmatched_in_eteacher.append(shop_name)
            continue
        # Write 家族ID to new D
        if family_id_col is not None:
            ws.cell(row=row_idx, column=family_id_col).value = info.family_id
        # Write 売上 to new AO
        if info.sales != 0:
            ws.cell(row=row_idx, column=new_month_sales_col).value = info.sales
        # 入金日 (AN) left blank intentionally
        matched_names.add(shop_name)
        result.matched.append((shop_name, info.family_id, info.sales))

    # Shops in source but not matched in eteacher
    result.unmatched_in_source = sorted(set(sales.keys()) - matched_names)

    wb.save(output_xlsx)
    wb.close()
    return result


def format_eteacher_summary(result: EteacherUpdateResult, detail_preview: int = 5) -> str:
    """Short summary used in CLI stdout and Telegram. See ``write_unmatched_report``
    for the full list intended for operator review."""
    lines = [
        f"【eteacher 売上反映】{result.output_path.name}",
        f"  テンプレ: {result.template_path.name}",
        f"  照合成功: {len(result.matched)}塾 (うち売上>0 は Excel で確認)",
        f"  eteacher 側で未マッチ: {len(result.unmatched_in_eteacher)}塾",
        f"  source 側で未マッチ: {len(result.unmatched_in_source)}塾",
    ]
    if result.unmatched_in_eteacher:
        lines.append("  先頭サンプル (eteacher 側):")
        for name in result.unmatched_in_eteacher[:detail_preview]:
            lines.append(f"    - {name}")
    if result.unmatched_in_source:
        lines.append("  先頭サンプル (source 側):")
        for name in result.unmatched_in_source[:detail_preview]:
            lines.append(f"    - {name}")
    return "\n".join(lines)


def write_unmatched_report(result: EteacherUpdateResult, report_path: str | Path) -> Path:
    """Write the full list of unmatched shops to a text file for manual review.

    目視で似た塾名を統一するときに使う。
    """
    report = Path(report_path)
    report.parent.mkdir(parents=True, exist_ok=True)
    lines = [
        f"# eteacher 更新 未マッチ塾リスト",
        f"出力ファイル: {result.output_path}",
        f"テンプレ:     {result.template_path}",
        "",
        f"## eteacher 側にあるが source にない塾 ({len(result.unmatched_in_eteacher)}件)",
        "これらは eteacher の行は残るが、AO 列(売上)が空のままになります。",
        "source 側の塾名の表記揺れが原因なら手動で統一してください。",
        "",
    ]
    lines.extend(f"- {n}" for n in result.unmatched_in_eteacher)
    lines.append("")
    lines.append(f"## source 側にあるが eteacher にない塾 ({len(result.unmatched_in_source)}件)")
    lines.append("これらの売上は eteacher に載らないので、新規塾なら手動で行追加が必要です。")
    lines.append("")
    lines.extend(f"- {n}" for n in result.unmatched_in_source)
    report.write_text("\n".join(lines), encoding="utf-8")
    return report
