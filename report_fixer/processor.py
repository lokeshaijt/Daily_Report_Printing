"""Correction rules for the downloaded stock report.

Rules implemented (numbers match the requirement document):

1. Remove rows where every cell is empty / NULL (partially empty rows are kept).
2. Day Consumption: tons -> kg (x1000) and No. of Days Consumption recomputed
   from the kg value.
3. Raw Aging Stock sheet: drop Lot Code, Lot No, Supplier Lot No, Category,
   GSM, Qty, Machine Line.
4. All TagEnv Aging Stock sheet: drop Lot Code, Lot No, Supplier Lot No,
   Width, GSM, Source Type; group repeated Item Names with the average QTN and
   Holding Days on the group (summary) row.
5/7. Zero values -> "-" (all sheets, which includes the 4th sheet).
6. Decimal values rounded to the nearest whole number (half-up).

Rounding and the zero -> "-" swap are applied by :func:`display_value` at
output time, so averages and the days calculation use the full-precision
numbers.
"""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal
from typing import Dict, List, Optional

from .reader import SheetRows

# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
NULL_TOKENS = {"", "null", "none", "nan", "nat", "#n/a", "n/a"}


def norm(text) -> str:
    """'No. of Days  Consumption' -> 'noofdaysconsumption'."""
    return re.sub(r"[^a-z0-9]", "", str(text or "").lower())


def is_blank(v) -> bool:
    return v is None or (isinstance(v, str) and v.strip().lower() in NULL_TOKENS)


_NUM_RE = re.compile(r"^-?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$")


def to_number(v):
    """Return v as int/float if it is numeric (or a clean numeric string), else None."""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        s = v.strip()
        if _NUM_RE.match(s):
            digits = s.lstrip("-").replace(",", "")
            # keep codes such as "00123" or very long ids as text
            if (len(digits.split(".")[0]) > 1 and digits.startswith("0")) or len(digits) > 15:
                return None
            f = float(s.replace(",", ""))
            return int(f) if "." not in s else f
    return None


def round_half_up(x: float) -> int:
    return int(Decimal(str(x)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def display_value(v):
    """Final cell value: decimals rounded, zeros shown as '-'.

    Values like 0.4 (between -1 and 1, but not zero) are kept exactly as they are.
    """
    if isinstance(v, bool) or isinstance(v, (dt.date, dt.datetime, dt.time)):
        return v
    if isinstance(v, (int, float)):
        if v != 0 and -1 < v < 1:
            return v
        r = round_half_up(v)
        return "-" if r == 0 else r
    return v


# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #
RAW_AGING_DROP = ["Lot Code", "Lot No", "Supplier Lot No", "Category", "GSM", "Qty", "Machine Line"]
TAGENV_DROP = ["Lot Code", "Lot No", "Supplier Lot No", "Width", "GSM", "Source Type"]

# alternative spellings that also count as the same column
_ALIASES = {
    "lotno": {"lotno", "lotnumber"},
    "supplierlotno": {"supplierlotno", "supplierlotnumber", "suplotno", "vendorlotno"},
    "machineline": {"machineline", "machinelines"},
    "qty": {"qty", "quantity"},
    "category": {"category", "itemcategory"},
}


def _drop_keys(names: List[str]) -> set:
    keys = set()
    for n in names:
        k = norm(n)
        keys |= _ALIASES.get(k, {k})
    return keys


def is_raw_aging_sheet(name: str) -> bool:
    n = norm(name)
    return "raw" in n and ("aging" in n or "ageing" in n)


def is_tagenv_sheet(name: str) -> bool:
    n = norm(name)
    return "tagenv" in n or ("alltag" in n and ("aging" in n or "ageing" in n))


@dataclass
class Options:
    raw_aging_sheet: Optional[str] = None  # None -> auto-detect, "" -> skip
    tagenv_sheet: Optional[str] = None
    convert_day_consumption: bool = True
    recalc_days: bool = True
    stock_column: Optional[str] = None  # header used as stock qty for days calc; None -> auto


# --------------------------------------------------------------------------- #
# Data model
# --------------------------------------------------------------------------- #
@dataclass
class Group:
    item: str
    summary: list            # one row with averages
    details: List[list]      # the individual records


@dataclass
class Sheet:
    name: str
    title_rows: List[list] = field(default_factory=list)
    header: Optional[list] = None
    rows: List[list] = field(default_factory=list)
    groups: Optional[List[Group]] = None   # only for the grouped TagEnv sheet
    group_value_cols: List[int] = field(default_factory=list)
    item_col: Optional[int] = None
    notes: List[str] = field(default_factory=list)

    @property
    def width(self) -> int:
        widths = [len(self.header or [])] + [len(r) for r in self.rows + self.title_rows]
        return max(widths) if widths else 0


# --------------------------------------------------------------------------- #
# Header detection
# --------------------------------------------------------------------------- #
_HEADER_HINTS = ("item", "name", "code", "qty", "qtn", "stock", "days", "date", "uom", "lot", "consumption", "description")


def find_header_row(rows: List[list], scan: int = 30) -> Optional[int]:
    best, best_score = None, 0.0
    for i, row in enumerate(rows[:scan]):
        filled = [c for c in row if not is_blank(c)]
        if len(filled) < 2:
            continue
        texts = [c for c in filled if isinstance(c, str) and to_number(c) is None]
        score = len(texts) * (1.0 if len(texts) == len(filled) else 0.5)
        score += sum(1 for t in texts if any(h in norm(t) for h in _HEADER_HINTS))
        if score > best_score:
            best, best_score = i, score
    return best


def _find_col(header: list, predicate) -> Optional[int]:
    for i, h in enumerate(header):
        if h is not None and predicate(norm(h)):
            return i
    return None


def is_day_consumption(n: str) -> bool:
    return ("dayconsumption" in n or "dailyconsumption" in n or "perdaycons" in n) and "noof" not in n and "days" not in n


def is_days_consumption(n: str) -> bool:
    return ("noofday" in n or "daysconsumption" in n or "daysofconsumption" in n or "daysstock" in n
            or "daysinventory" in n) and not is_day_consumption(n)


_STOCK_PRIORITY = ("closingstock", "currentstock", "stockqty", "stockinhand", "balanceqty",
                   "availableqty", "stock", "qtn", "qty", "quantity")


def find_stock_col(header: list, exclude: set) -> Optional[int]:
    keys = [(i, norm(h)) for i, h in enumerate(header) if h is not None and i not in exclude]
    for wanted in _STOCK_PRIORITY:
        for i, n in keys:
            if n == wanted or n.startswith(wanted):
                return i
    for wanted in _STOCK_PRIORITY[:7]:
        for i, n in keys:
            if wanted in n:
                return i
    return None


# --------------------------------------------------------------------------- #
# Rules
# --------------------------------------------------------------------------- #
def drop_columns(sheet: Sheet, names: List[str]) -> None:
    keys = _drop_keys(names)
    idx = [i for i, h in enumerate(sheet.header) if h is not None and norm(h) in keys]
    if not idx:
        sheet.notes.append("None of the columns to remove were found: " + ", ".join(names))
        return
    removed = [str(sheet.header[i]) for i in idx]
    keep = [i for i in range(len(sheet.header)) if i not in idx]
    sheet.header = [sheet.header[i] for i in keep]
    sheet.rows = [[r[i] if i < len(r) else None for i in keep] for r in sheet.rows]
    sheet.notes.append("Removed columns: " + ", ".join(removed))
    removed_keys = {norm(r) for r in removed}
    missing = [n for n in names if not (_drop_keys([n]) & removed_keys)]
    if missing:
        sheet.notes.append("Not present (nothing to remove): " + ", ".join(missing))


def convert_day_consumption(sheet: Sheet, opts: Options) -> None:
    dc = _find_col(sheet.header, is_day_consumption)
    if dc is None:
        return
    days = _find_col(sheet.header, is_days_consumption)
    stock = None
    if opts.stock_column:
        stock = _find_col(sheet.header, lambda n: n == norm(opts.stock_column))
    if stock is None:
        stock = find_stock_col(sheet.header, exclude={dc, days} - {None})

    old_header = str(sheet.header[dc])
    new_header = re.sub(r"\(?\s*(m\.?t\.?|tons?|tonnes?)\s*\)?", "", old_header, flags=re.I).strip(" -_")
    sheet.header[dc] = f"{new_header or 'Day Consumption'} (Kg)"
    converted = 0
    recalculated = 0
    for r in sheet.rows:
        ton = to_number(r[dc])
        if ton is None:
            continue
        kg = ton * 1000
        r[dc] = kg
        converted += 1
        if days is None or not opts.recalc_days:
            continue
        if stock is not None:
            qty = to_number(r[stock])
            if qty is None:
                continue
            r[days] = qty / kg if kg else 0
        else:
            old_days = to_number(r[days])
            if old_days is None:
                continue
            # days were stock / tons -> stock / kg is 1000x smaller
            r[days] = old_days / 1000
        recalculated += 1
    sheet.notes.append(f"'{old_header}' converted from tons to kg on {converted} rows (renamed to '{sheet.header[dc]}').")
    if days is not None and opts.recalc_days:
        basis = (f"'{sheet.header[stock]}' / '{sheet.header[dc]}'" if stock is not None
                 else "previous value / 1000 (no stock column found)")
        sheet.notes.append(f"'{sheet.header[days]}' recalculated on {recalculated} rows as {basis}.")


def _item_col(header: list) -> Optional[int]:
    for pred in (lambda n: n == "itemname", lambda n: "itemname" in n, lambda n: n in ("itemdescription", "description"),
                 lambda n: n == "item", lambda n: n.startswith("item") and "code" not in n):
        c = _find_col(header, pred)
        if c is not None:
            return c
    return None


def _qtn_col(header: list) -> Optional[int]:
    for pred in (lambda n: n.startswith("qtn"), lambda n: n.startswith("qty") or n.startswith("quantity"),
                 lambda n: "qtn" in n or "qty" in n):
        c = _find_col(header, pred)
        if c is not None:
            return c
    return None


def _holding_col(header: list) -> Optional[int]:
    c = _find_col(header, lambda n: "holding" in n)
    if c is None:
        c = _find_col(header, lambda n: "agingdays" in n or "ageingdays" in n)
    return c


def group_items(sheet: Sheet) -> None:
    ic = _item_col(sheet.header)
    if ic is None:
        sheet.notes.append("Item Name column not found - rows were not grouped.")
        return
    value_cols = [c for c in (_qtn_col(sheet.header), _holding_col(sheet.header)) if c is not None]
    sheet.group_value_cols = value_cols
    sheet.item_col = ic

    order: List[str] = []
    buckets: Dict[str, List[list]] = {}
    trailing: List[list] = []
    for r in sheet.rows:
        item = r[ic]
        if is_blank(item) or "total" in norm(item):
            trailing.append(r)
            continue
        key = str(item).strip()
        if key not in buckets:
            buckets[key] = []
            order.append(key)
        buckets[key].append(r)

    groups = []
    width = len(sheet.header)
    for key in order:
        recs = buckets[key]
        summary = [None] * width
        summary[ic] = key
        # carry over any value that is identical for every record (e.g. Item Code, UOM)
        for c in range(width):
            if c in value_cols or c == ic:
                continue
            vals = {repr(r[c]) for r in recs}
            if len(vals) == 1:
                summary[c] = recs[0][c]
        for c in value_cols:
            nums = [to_number(r[c]) for r in recs]
            nums = [n for n in nums if n is not None]
            summary[c] = sum(nums) / len(nums) if nums else None
        groups.append(Group(item=key, summary=summary, details=recs))
    sheet.groups = groups
    sheet.rows = trailing  # rows that are not part of an item group (e.g. totals)
    repeated = sum(1 for g in groups if len(g.details) > 1)
    names = ", ".join(f"'{sheet.header[c]}'" for c in value_cols) or "(no QTN / Holding Days column found)"
    sheet.notes.append(
        f"Grouped {sum(len(g.details) for g in groups)} records into {len(groups)} unique items "
        f"({repeated} repeated). Group rows show the average of {names}.")


# --------------------------------------------------------------------------- #
# Pipeline
# --------------------------------------------------------------------------- #
def build_sheet(name: str, rows: List[list]) -> Sheet:
    sheet = Sheet(name=name)
    h = find_header_row(rows)
    if h is None:
        sheet.rows = [r for r in rows if not all(is_blank(c) for c in r)]
        return sheet
    sheet.title_rows = [r for r in rows[:h] if not all(is_blank(c) for c in r)]
    sheet.header = list(rows[h])
    body = rows[h + 1:]
    kept = [list(r) for r in body if not all(is_blank(c) for c in r)]
    removed = len(body) - len(kept)
    # numeric text -> numbers so they can be rounded / formatted
    sheet.rows = [[(to_number(c) if to_number(c) is not None else c) for c in r] for r in kept]
    if removed:
        sheet.notes.append(f"Removed {removed} completely empty rows.")
    return sheet


def drop_item_only_rows(sheet: Sheet) -> None:
    """Remove rows that have an Item Name but no value in any other column."""
    ic = _item_col(sheet.header)
    if ic is None:
        return
    before = len(sheet.rows)
    sheet.rows = [r for r in sheet.rows
                  if is_blank(r[ic]) or not all(is_blank(c) for i, c in enumerate(r) if i != ic)]
    removed = before - len(sheet.rows)
    if removed:
        sheet.notes.append(f"Removed {removed} rows where '{sheet.header[ic]}' had no values in any other column.")


def resolve_sheets(names: List[str], opts: Options):
    """None -> auto-detect by sheet name, "" -> rule disabled, otherwise the named sheet."""
    def pick(choice, detect):
        if choice == "":
            return None
        if choice in names:
            return choice
        return next((n for n in names if detect(n)), None)
    return pick(opts.raw_aging_sheet, is_raw_aging_sheet), pick(opts.tagenv_sheet, is_tagenv_sheet)


def process(sheets: List[SheetRows], opts: Optional[Options] = None) -> List[Sheet]:
    opts = opts or Options()
    raw_name, tag_name = resolve_sheets([n for n, _ in sheets], opts)
    out = []
    for name, rows in sheets:
        sheet = build_sheet(name, rows)
        if sheet.header is not None:
            if name == raw_name:
                drop_columns(sheet, RAW_AGING_DROP)
            if name == tag_name:
                drop_columns(sheet, TAGENV_DROP)
            drop_item_only_rows(sheet)
            if opts.convert_day_consumption:
                convert_day_consumption(sheet, opts)
            if name == tag_name:
                group_items(sheet)
        out.append(sheet)
    return out
