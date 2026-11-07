"""Read an uploaded report into a simple, engine-independent structure.

Every sheet becomes ``(sheet_name, rows)`` where ``rows`` is a list of lists of
plain Python values (str / int / float / datetime / None). Supports .xlsx /
.xlsm (openpyxl), legacy .xls (xlrd) and the HTML-table-disguised-as-.xls
files that many web portals produce.
"""
from __future__ import annotations

import datetime as dt
import io
import math
from typing import List, Tuple

import pandas as pd
from openpyxl import load_workbook

SheetRows = Tuple[str, List[list]]


def _clean_value(v):
    if v is None:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    if isinstance(v, pd.Timestamp):
        return None if pd.isna(v) else v.to_pydatetime()
    if v is pd.NaT:
        return None
    if isinstance(v, str):
        return v.strip() if v.strip() else None
    return v


def _trim(rows: List[list]) -> List[list]:
    """Drop trailing empty cells so every sheet has a tight used range."""
    width = 0
    for r in rows:
        for i in range(len(r) - 1, -1, -1):
            if r[i] is not None:
                width = max(width, i + 1)
                break
    return [list(r[:width]) + [None] * (width - len(r[:width])) for r in rows]


def _read_openpyxl(data: bytes) -> List[SheetRows]:
    wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    out = []
    for ws in wb.worksheets:
        rows = [[_clean_value(c) for c in row] for row in ws.iter_rows(values_only=True)]
        out.append((ws.title, _trim(rows)))
    wb.close()
    return out


def _read_pandas(frames: dict) -> List[SheetRows]:
    out = []
    for name, df in frames.items():
        rows = [[_clean_value(v) for v in rec] for rec in df.itertuples(index=False, name=None)]
        out.append((str(name), _trim(rows)))
    return out


def read_workbook(data: bytes, filename: str = "") -> List[SheetRows]:
    name = filename.lower()
    errors = []
    if not name.endswith(".xls"):
        try:
            return _read_openpyxl(data)
        except Exception as e:  # noqa: BLE001
            errors.append(f"openpyxl: {e}")
    try:
        frames = pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, engine="xlrd")
        return _read_pandas(frames)
    except Exception as e:  # noqa: BLE001
        errors.append(f"xlrd: {e}")
    try:
        tables = pd.read_html(io.BytesIO(data), header=None)
        return _read_pandas({f"Sheet{i + 1}": t for i, t in enumerate(tables)})
    except Exception as e:  # noqa: BLE001
        errors.append(f"html: {e}")
    raise ValueError("Could not read the uploaded file as an Excel report.\n" + "\n".join(errors))


__all__ = ["read_workbook", "SheetRows", "dt"]
