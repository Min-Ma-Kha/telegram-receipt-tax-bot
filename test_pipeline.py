"""Offline end-to-end test: synthetic receipt image -> OCR -> parse -> Excel.
Covers per-user isolation, per-year files, duplicates, year-end report
markers, and legacy migration."""

import os
import shutil
import sys
import tempfile

from openpyxl import Workbook
from PIL import Image, ImageDraw, ImageFont

import storage
from ocr import extract_text
from receipt_parser import parse_receipt

RECEIPT_TEXT = [
    "WALMART SUPERCENTER",
    "123 MAIN ST, DALLAS TX",
    "06/08/2026  14:32",
    "",
    "GV MILK 1GAL        3.48",
    "BREAD WHEAT         2.18",
    "EGGS LARGE 12CT     4.92",
    "PAPER TOWELS        8.97",
    "BATTERIES AA        9.97",
    "",
    "SUBTOTAL           29.52",
    "TAX 8.250%          2.44",
    "TOTAL              31.96",
    "CASH               40.00",
    "CHANGE              8.04",
]


def make_receipt_image(path: str) -> None:
    try:
        font = ImageFont.truetype("consola.ttf", 28)
    except OSError:
        font = ImageFont.load_default()
    img = Image.new("L", (520, 40 + 40 * len(RECEIPT_TEXT)), 255)
    d = ImageDraw.Draw(img)
    for i, line in enumerate(RECEIPT_TEXT):
        d.text((20, 20 + 40 * i), line, fill=0, font=font)
    img.save(path)


def main() -> int:
    failures = []

    def check(cond, msg):
        if not cond:
            failures.append(msg)

    img_path = os.path.join(os.path.dirname(__file__), "test_receipt.png")
    make_receipt_image(img_path)

    # 1. OCR + parse
    r = parse_receipt(extract_text(img_path))
    print("--- Parsed ---")
    print(r)
    check(r.tax == 2.44, f"tax: expected 2.44, got {r.tax}")
    check(r.total == 31.96, f"total: expected 31.96, got {r.total}")
    check(r.subtotal == 29.52, f"subtotal: expected 29.52, got {r.subtotal}")
    check(r.date == "06/08/2026", f"date: expected 06/08/2026, got {r.date!r}")
    check("walmart" in r.store.lower(), f"store wrong: {r.store!r}")

    # 2. Per-year files + global IDs in a scratch data dir
    storage.DATA_DIR = tempfile.mkdtemp(prefix="receipt_test_")
    U1, U2 = 111, 222
    id1 = storage.add_receipt(U1, store=r.store, date=r.date, subtotal=r.subtotal,
                              tax=r.tax, total=r.total, photo="x.jpg", user="u1")
    id2 = storage.add_receipt(U1, store="Target", date="01/15/2025", subtotal=10.00,
                              tax=0.83, total=10.83, photo="", user="u1")
    id3 = storage.add_receipt(U1, store="Costco", date="03/02/2026", subtotal=50.0,
                              tax=4.13, total=54.13, photo="", user="u1")
    check([id1, id2, id3] == [1, 2, 3], f"global ids wrong: {[id1, id2, id3]}")
    check(os.path.exists(storage.excel_path(U1, 2025)), "2025 file missing")
    check(os.path.exists(storage.excel_path(U1, 2026)), "2026 file missing")
    check(storage.years_with_data(U1) == [2025, 2026],
          f"years_with_data wrong: {storage.years_with_data(U1)}")

    # 3. Summaries: per-year file vs all-time aggregate across files
    s26 = storage.get_summary(U1, 2026)
    check(s26["count"] == 2, f"2026 count: {s26['count']}")
    check(s26["tax"] == round(2.44 + 4.13, 2), f"2026 tax: {s26['tax']}")
    s25 = storage.get_summary(U1, 2025)
    check(s25["tax"] == 0.83, f"2025 tax: {s25['tax']}")
    allt = storage.get_summary(U1)
    check(allt["count"] == 3, f"all-time count: {allt['count']}")
    check(allt["tax"] == round(2.44 + 4.13 + 0.83, 2), f"all-time tax: {allt['tax']}")
    check(set(allt["per_year"]) == {2025, 2026}, f"per_year keys: {allt['per_year']}")

    # 4. Per-user isolation
    check(storage.get_summary(U2)["count"] == 0, "user 2 should start empty")
    storage.add_receipt(U2, store="CVS", date="02/01/2026", subtotal=5.00,
                        tax=0.41, total=5.41, photo="p.jpg", user="u2",
                        file_hash="abc123")
    check(storage.get_summary(U1)["count"] == 3, "user 2 leaked into user 1")
    check(storage.last_receipt_id(U2) == 1, "user 2 ids should start at 1")

    # 5. Duplicate detection across a user's files
    check(storage.find_duplicate(U2, file_hash="abc123") is not None,
          "hash dup not found")
    check(storage.find_duplicate(U1, file_hash="abc123") is None,
          "dup crossed user boundary")
    check(storage.find_duplicate(U1, store="Target", date="01/15/2025",
                                 tax=0.83, total=10.83) is not None,
          "field dup not found across years")

    # 6. fix/delete by global ID find the right file; date edit moves year
    check(storage.update_receipt(U1, id3, "tax", 5.00), "update failed")
    check(storage.get_receipt(U1, id3)["Sales Tax"] == 5.00, "tax not updated")
    storage.update_receipt(U1, id2, "date", "12/20/2024")   # 2025 -> 2024 move
    check(storage.get_receipt(U1, id2)["Year"] == 2024, "year not moved")
    check(2024 in storage.years_with_data(U1), "2024 file not created on move")
    check(2025 not in storage.years_with_data(U1), "empty 2025 file not removed")
    check(storage.delete_receipt(U1, id1), "delete failed")
    check(storage.get_summary(U1)["count"] == 2, "count after delete wrong")

    # 7. Year-end report markers
    check(not storage.is_reported(U1, 2024), "should not be reported yet")
    storage.mark_reported(U1, 2024)
    check(storage.is_reported(U1, 2024), "mark_reported didn't stick")

    # 8. Legacy migration: build an old single-file workbook and split it
    U3 = 333
    os.makedirs(storage.user_dir(U3), exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Receipts"
    ws.append(storage.HEADERS)
    ws.append([1, "x", "05/01/2025", 2025, "OldStore", 9.0, 0.74, 9.74, "", "h1", "u3"])
    ws.append([2, "x", "06/01/2026", 2026, "NewStore", 19.0, 1.57, 20.57, "", "h2", "u3"])
    wb.create_sheet("Summary")
    wb.save(os.path.join(storage.user_dir(U3), "receipts.xlsx"))
    check(storage.migrate_legacy(U3), "migrate returned False")
    check(storage.years_with_data(U3) == [2025, 2026], "migration years wrong")
    check(storage.get_summary(U3)["count"] == 2, "migration lost rows")
    check(storage.is_reported(U3, 2025), "completed year not auto-marked")
    check(not storage.is_reported(U3, 2026), "current year wrongly marked")
    check(not os.path.exists(os.path.join(storage.user_dir(U3), "receipts.xlsx")),
          "legacy file not renamed")

    shutil.rmtree(storage.DATA_DIR, ignore_errors=True)
    os.remove(img_path)

    if failures:
        print("\nFAILURES:")
        for f in failures:
            print(" -", f)
        return 1
    print("\nALL TESTS PASSED")
    return 0


if __name__ == "__main__":
    sys.exit(main())
