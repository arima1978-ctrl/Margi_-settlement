"""Fetch shop info from the master Google Spreadsheet."""
from __future__ import annotations

import os
from pathlib import Path


def load_google_sheet_rows(sheet_id: str, tab_name: str) -> list[list[str]]:
    """Load all rows from a Google Sheet tab as a 2D list.

    Requires gspread + service-account credentials at credentials.json
    (or path set via GOOGLE_APPLICATION_CREDENTIALS env var).
    """
    try:
        import gspread
    except ImportError as e:
        raise RuntimeError(
            "gspread is required for Google Sheet access. "
            "Install with: pip install gspread google-auth"
        ) from e

    creds_path = os.environ.get(
        "GOOGLE_APPLICATION_CREDENTIALS", "credentials.json"
    )
    if not Path(creds_path).exists():
        raise FileNotFoundError(
            f"Google Sheets credentials not found at {creds_path}. "
            "Create a service-account JSON and set GOOGLE_APPLICATION_CREDENTIALS."
        )

    gc = gspread.service_account(filename=creds_path)
    sh = gc.open_by_key(sheet_id)
    worksheet = sh.worksheet(tab_name)
    return worksheet.get_all_values()


def find_shop_rows_by_family_ids(
    rows: list[list[str]], family_ids: list[int], id_column_index: int = 1
) -> dict[int, list[str]]:
    """Look up rows by family ID (first column assumed to be 塾ID by default).

    Returns mapping: family_id -> row values.
    """
    result: dict[int, list[str]] = {}
    for row in rows:
        if len(row) <= id_column_index:
            continue
        cell = row[id_column_index]
        if not cell:
            continue
        try:
            fid = int(str(cell).strip())
        except ValueError:
            continue
        if fid in family_ids:
            result[fid] = row
    return result
