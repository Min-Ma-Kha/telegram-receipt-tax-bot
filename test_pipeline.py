"""Offline end-to-end test: synthetic receipt image -> OCR -> parse -> Excel.
Also covers per-user isolation and duplicate detection."""

import os
import shutil
import sys
import tempfile

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
    img_path = os.path.join(os.path.dirname(__file__), "test_receipt.png")
    make_receipt_image(img_path)

    # 1. OCR
    text = extract_text(img_path)
    print("--- OCR text ---")
    print(text)

    # 2. Parse
    r = parse_receipt(text)
    print("--- Parsed ---")
    print(r)
    if r.tax != 2.44:
        failures.append(f"tax: expected 2.44, got {r.tax}")
    if r.total != 31.96:
        failures.append(f"total: expected 31.96, got {r.total}")
    if r.subtotal != 29.52:
        failures.append(f"subtotal: expected 29.52, got {r.subtotal}")
    if r.date != "06/08/2026":
        failures.append(f"date: expected 06/08/2026, got {r.date!r}")
    if "walmart" not in r.store.lower():
        failures.append(f"store: expected Walmart, got {r.store!r}")

    # 3. Storage in a scratch data dir, two separate users
    storage.DATA_DIR = tempfile.mkdtemp(prefix="receipt_test_")
    U1, U2 = 111, 222
    rid1 = storage.add_receipt(U1, store=r.store, date=r.date, subtotal=r.subtotal,
                               tax=r.tax, total=r.total, photo="x.jpg", user="u1")
    rid2 = storage.add_receipt(U1, store="Target", date="01/15/2025", subtotal=10.00,
                               tax=0.83, total=10.83, photo="", user="u1")
    storage.update_receipt(U1, rid2, "tax", 0.85)
    s = storage.get_summary(U1)
    print("--- Summary (user 1) ---")
    print(s)
    if s["count"] != 2:
        failures.append(f"summary count: expected 2, got {s['count']}")
    if s["tax"] != round(2.44 + 0.85, 2):
        failures.append(f"summary tax: expected 3.29, got {s['tax']}")
    if 2026 not in s["per_year"] or 2025 not in s["per_year"]:
        failures.append(f"per-year keys wrong: {list(s['per_year'])}")
    s25 = storage.get_summary(U1, 2025)
    if s25["tax"] != 0.85:
        failures.append(f"2025 tax: expected 0.85, got {s25['tax']}")

    # 4. Per-user isolation: user 2 sees nothing of user 1
    if storage.get_summary(U2)["count"] != 0:
        failures.append("user isolation broken: user 2 sees user 1's receipts")
    storage.add_receipt(U2, store="CVS", date="02/01/2026", subtotal=5.00,
                        tax=0.41, total=5.41, photo="p.jpg", user="u2",
                        file_hash="abc123")
    if storage.get_summary(U2)["count"] != 1:
        failures.append("user 2 add failed")
    if storage.get_summary(U1)["count"] != 2:
        failures.append("user 2's receipt leaked into user 1's storage")
    if not os.path.exists(storage.excel_path(U1)) or not os.path.exists(storage.excel_path(U2)):
        failures.append("per-user excel files missing")

    # 5. Duplicate detection (scoped to one user)
    dup = storage.find_duplicate(U2, file_hash="abc123")
    if not dup or dup[1] != "same photo":
        failures.append(f"hash duplicate not found: {dup}")
    dup = storage.find_duplicate(U1, file_hash="abc123")
    if dup:
        failures.append("duplicate check crossed user boundary")
    dup = storage.find_duplicate(U2, store="cvs ", date="02/01/2026",
                                 tax=0.41, total=5.41)
    if not dup:
        failures.append("field duplicate not found")
    dup = storage.find_duplicate(U2, store="CVS", date="02/01/2026",
                                 tax=0.41, total=99.99)
    if dup:
        failures.append(f"false field duplicate (different total): {dup}")

    storage.delete_receipt(U1, rid1)
    if storage.get_summary(U1)["count"] != 1:
        failures.append("delete failed")

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
