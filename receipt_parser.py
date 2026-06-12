"""Parse OCR'd receipt text into store / date / subtotal / sales tax / total."""

import re
from dataclasses import dataclass, field


@dataclass
class Receipt:
    store: str = ""
    date: str = ""          # MM/DD/YYYY
    subtotal: float | None = None
    tax: float | None = None
    total: float | None = None
    warnings: list[str] = field(default_factory=list)


# An amount like 12.34 / 12,34 / $ 12.34 — OCR often mangles the separator.
_MONEY = re.compile(r"\$?\s*(\d{1,5})\s*[.,]\s*(\d{2})\b")
_DATE = re.compile(r"\b(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})\b")

# Words that disqualify a "total" line from being the grand total.
_NOT_TOTAL = ("subtotal", "sub total", "sub-total", "savings", "discount",
              "items", "item count", "points", "tax", "you saved")
_TOTAL_WORDS = ("total", "balance due", "amount due", "balance")

# Lines we skip when guessing the store name.
_NOT_STORE = re.compile(
    r"receipt|welcome|thank|invoice|order|cashier|register|store\s*#|tel|phone|"
    r"www\.|\.com|^\W*$", re.IGNORECASE)


def _amounts(line: str) -> list[float]:
    return [float(f"{d}.{c}") for d, c in _MONEY.findall(line)]


def _last_amount(line: str) -> float | None:
    amts = _amounts(line)
    return amts[-1] if amts else None


def parse_receipt(text: str) -> Receipt:
    r = Receipt()
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    tax_lines: list[float] = []
    total_tax: float | None = None
    total_candidates: list[float] = []

    for line in lines:
        low = line.lower()

        # --- sales tax ---
        if "tax" in low and "taxable" not in low:
            amt = _last_amount(line)
            if amt is not None:
                if "total tax" in low:
                    total_tax = amt
                else:
                    tax_lines.append(amt)

        # --- subtotal ---
        if r.subtotal is None and any(w in low for w in ("subtotal", "sub total", "sub-total")):
            r.subtotal = _last_amount(line)

        # --- total ---
        if any(w in low for w in _TOTAL_WORDS) and not any(w in low for w in _NOT_TOTAL):
            amt = _last_amount(line)
            if amt is not None:
                total_candidates.append(amt)

        # --- date ---
        if not r.date:
            m = _DATE.search(line)
            if m:
                mo, dy, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
                if yr < 100:
                    yr += 2000
                if 1 <= mo <= 12 and 1 <= dy <= 31 and 2000 <= yr <= 2099:
                    r.date = f"{mo:02d}/{dy:02d}/{yr}"

    # Prefer an explicit "total tax" line; otherwise sum individual tax lines
    # (handles receipts with Tax 1 / Tax 2).
    if total_tax is not None:
        r.tax = total_tax
    elif tax_lines:
        r.tax = round(sum(tax_lines), 2)

    if total_candidates:
        # The grand total is usually the last "total"-ish line, but if
        # subtotal + tax matches another candidate better, prefer that.
        r.total = total_candidates[-1]
        if r.subtotal is not None and r.tax is not None:
            expected = round(r.subtotal + r.tax, 2)
            if abs(r.total - expected) > 0.02:
                close = [t for t in total_candidates if abs(t - expected) <= 0.02]
                if close:
                    r.total = close[0]

    # Derive missing pieces when two of three are known.
    if r.total is None and r.subtotal is not None and r.tax is not None:
        r.total = round(r.subtotal + r.tax, 2)
    if r.subtotal is None and r.total is not None and r.tax is not None:
        r.subtotal = round(r.total - r.tax, 2)

    # --- store name: first plausible line near the top ---
    for line in lines[:6]:
        letters = sum(ch.isalpha() for ch in line)
        if letters >= 3 and not _NOT_STORE.search(line) and not _MONEY.search(line):
            r.store = line.title() if line.isupper() else line
            break

    # --- sanity warnings ---
    if r.tax is None:
        r.warnings.append("Couldn't find a tax line — fix it with /fix tax <amount>")
    if r.total is None:
        r.warnings.append("Couldn't find the total — fix it with /fix total <amount>")
    if (r.subtotal is not None and r.tax is not None and r.total is not None
            and abs(round(r.subtotal + r.tax, 2) - r.total) > 0.02):
        r.warnings.append("Subtotal + tax doesn't add up to total — double-check the numbers")

    return r
