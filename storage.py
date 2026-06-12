"""Excel storage for receipts (openpyxl) — one isolated folder per Telegram user.

Layout on disk:
  data/users/<telegram_user_id>/receipts.xlsx   one workbook per user
  data/users/<telegram_user_id>/photos/         that user's receipt photos

Workbook layout:
  Receipts sheet — one row per receipt.
  Summary sheet  — per-year totals via live Excel formulas, plus grand total.
"""

import os
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")

HEADERS = ["ID", "Added", "Receipt Date", "Year", "Store",
           "Subtotal", "Sales Tax", "Total", "Photo", "Hash", "User"]
MONEY_COLS = (6, 7, 8)          # Subtotal, Sales Tax, Total
COL_ID, COL_YEAR, COL_STORE = 1, 4, 5
COL_SUBTOTAL, COL_TAX, COL_TOTAL = 6, 7, 8
COL_DATE, COL_HASH = 3, 10
MONEY_FMT = '"$"#,##0.00'


def user_dir(user_id) -> str:
    return os.path.join(DATA_DIR, "users", str(user_id))


def excel_path(user_id) -> str:
    return os.path.join(user_dir(user_id), "receipts.xlsx")


def photos_dir(user_id) -> str:
    return os.path.join(user_dir(user_id), "photos")


def _open(user_id) -> Workbook:
    os.makedirs(photos_dir(user_id), exist_ok=True)
    path = excel_path(user_id)
    if os.path.exists(path):
        return load_workbook(path)
    wb = Workbook()
    ws = wb.active
    ws.title = "Receipts"
    ws.append(HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    widths = [6, 18, 12, 8, 24, 11, 11, 11, 28, 16, 14]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    wb.create_sheet("Summary")
    return wb


def _rows(ws):
    """Data rows as tuples (skips header)."""
    return list(ws.iter_rows(min_row=2, values_only=True))


def _rebuild_summary(wb: Workbook) -> None:
    ws = wb["Receipts"]
    sm = wb["Summary"]
    sm.delete_rows(1, sm.max_row)

    years = sorted({row[COL_YEAR - 1] for row in _rows(ws) if row[COL_YEAR - 1]})
    sm.append(["Year", "Receipts", "Total Spent", "Total Sales Tax"])
    for cell in sm[1]:
        cell.font = Font(bold=True)
    for w, col in zip((10, 10, 14, 16), "ABCD"):
        sm.column_dimensions[col].width = w

    yc = get_column_letter(COL_YEAR)
    tc = get_column_letter(COL_TOTAL)
    xc = get_column_letter(COL_TAX)
    for y in years:
        sm.append([
            y,
            f'=COUNTIF(Receipts!{yc}:{yc},A{sm.max_row + 1})',
            f'=SUMIF(Receipts!{yc}:{yc},A{sm.max_row + 1},Receipts!{tc}:{tc})',
            f'=SUMIF(Receipts!{yc}:{yc},A{sm.max_row + 1},Receipts!{xc}:{xc})',
        ])
    last = sm.max_row
    sm.append(["TOTAL",
               f"=SUM(B2:B{last})", f"=SUM(C2:C{last})", f"=SUM(D2:D{last})"])
    for cell in sm[sm.max_row]:
        cell.font = Font(bold=True)
    for row in sm.iter_rows(min_row=2, min_col=3, max_col=4):
        for cell in row:
            cell.number_format = MONEY_FMT


def _save(wb: Workbook, user_id) -> None:
    try:
        wb.save(excel_path(user_id))
    except PermissionError as e:
        raise RuntimeError(
            "Can't write receipts.xlsx — close it in Excel and try again.") from e


def find_duplicate(user_id, *, file_hash: str = "", store: str = "",
                   date: str = "", tax=None, total=None) -> tuple[int, str] | None:
    """Return (receipt_id, reason) if this receipt is already stored.

    Matches either the exact same photo (file hash) or the same parsed
    receipt (store + date + tax + total), which catches re-photographed
    duplicates too. Only searches this user's own workbook.
    """
    store_key = (store or "").strip().lower()
    for row in _rows(_open(user_id)["Receipts"]):
        if file_hash and row[COL_HASH - 1] == file_hash:
            return row[0], "same photo"
        if (total is not None and store_key
                and (row[COL_STORE - 1] or "").strip().lower() == store_key
                and (row[COL_DATE - 1] or "") == (date or "")
                and row[COL_TOTAL - 1] == total
                and row[COL_TAX - 1] == tax):
            return row[0], "same store, date and amounts"
    return None


def add_receipt(user_id, *, store: str, date: str, subtotal, tax, total,
                photo: str, user: str, file_hash: str = "") -> int:
    """Append a receipt row; returns its ID."""
    wb = _open(user_id)
    ws = wb["Receipts"]
    rows = _rows(ws)
    rid = (rows[-1][0] + 1) if rows else 1

    year = None
    if date:
        year = int(date[-4:])
    if not year:
        year = datetime.now().year

    ws.append([rid, datetime.now().strftime("%m/%d/%Y %H:%M"), date or "",
               year, store or "Unknown", subtotal, tax, total, photo,
               file_hash, user])
    for col in MONEY_COLS:
        ws.cell(row=ws.max_row, column=col).number_format = MONEY_FMT

    _rebuild_summary(wb)
    _save(wb, user_id)
    return rid


def update_receipt(user_id, rid: int, field: str, value) -> bool:
    """field: 'tax' | 'total' | 'subtotal' | 'store' | 'date'."""
    col = {"tax": COL_TAX, "total": COL_TOTAL, "subtotal": COL_SUBTOTAL,
           "store": COL_STORE, "date": COL_DATE}[field]
    wb = _open(user_id)
    ws = wb["Receipts"]
    for row in ws.iter_rows(min_row=2):
        if row[0].value == rid:
            row[col - 1].value = value
            if col in MONEY_COLS:
                row[col - 1].number_format = MONEY_FMT
            if field == "date" and value:
                row[COL_YEAR - 1].value = int(str(value)[-4:])
            _rebuild_summary(wb)
            _save(wb, user_id)
            return True
    return False


def delete_receipt(user_id, rid: int) -> bool:
    wb = _open(user_id)
    ws = wb["Receipts"]
    for row in ws.iter_rows(min_row=2):
        if row[0].value == rid:
            ws.delete_rows(row[0].row, 1)
            _rebuild_summary(wb)
            _save(wb, user_id)
            return True
    return False


def last_receipt_id(user_id) -> int | None:
    rows = _rows(_open(user_id)["Receipts"])
    return rows[-1][0] if rows else None


def get_receipt(user_id, rid: int) -> dict | None:
    for row in _rows(_open(user_id)["Receipts"]):
        if row[0] == rid:
            return dict(zip(HEADERS, row))
    return None


def get_summary(user_id, year: int | None = None) -> dict:
    """Totals overall or for one year, plus per-year breakdown."""
    rows = _rows(_open(user_id)["Receipts"])
    if year:
        rows = [r for r in rows if r[COL_YEAR - 1] == year]
    per_year: dict[int, dict] = {}
    for r in rows:
        y = r[COL_YEAR - 1]
        d = per_year.setdefault(y, {"count": 0, "spent": 0.0, "tax": 0.0})
        d["count"] += 1
        d["spent"] += r[COL_TOTAL - 1] or 0
        d["tax"] += r[COL_TAX - 1] or 0
    return {
        "count": len(rows),
        "spent": round(sum(r[COL_TOTAL - 1] or 0 for r in rows), 2),
        "tax": round(sum(r[COL_TAX - 1] or 0 for r in rows), 2),
        "per_year": {y: {k: round(v, 2) if isinstance(v, float) else v
                         for k, v in d.items()}
                     for y, d in sorted(per_year.items())},
    }


def get_last(user_id, n: int = 5) -> list[dict]:
    rows = _rows(_open(user_id)["Receipts"])
    return [dict(zip(HEADERS, r)) for r in rows[-n:]]


def excel_exists(user_id) -> bool:
    return os.path.exists(excel_path(user_id))
