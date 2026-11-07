"""Write processed sheets to a clean, consistently formatted, print-ready .xlsx."""
from __future__ import annotations

import datetime as dt
import io
import re
from typing import List

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import Outline, PageSetupProperties

from .processor import Sheet, display_value

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
TITLE_FONT = Font(name="Calibri", size=13, bold=True, color="1F4E78")
BODY_FONT = Font(name="Calibri", size=10)
GROUP_FILL = PatternFill("solid", fgColor="DDEBF7")
GROUP_FONT = Font(name="Calibri", size=10, bold=True)
ZEBRA_FILL = PatternFill("solid", fgColor="F7F9FC")
THIN = Side(style="thin", color="B7C3D0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MAX_WIDTH = 45
MIN_WIDTH = 8
_CODE_WORDS = {"code", "id", "no", "number", "sl", "sno", "srno", "slno", "year", "pin"}


def _is_code_column(header) -> bool:
    words = set(re.findall(r"[a-z]+", str(header or "").lower()))
    return bool(words & _CODE_WORDS) and not words & {"days", "day", "qty", "qtn", "stock"}


def _text_len(v) -> int:
    if v is None:
        return 0
    if isinstance(v, (dt.datetime, dt.date)):
        return 11
    if isinstance(v, int):
        return len(f"{v:,}")
    return max(len(line) for line in str(v).split("\n"))


def _style_cell(cell, value, code_col: bool, bold=False, fill=None, indent=0):
    cell.font = GROUP_FONT if bold else BODY_FONT
    cell.border = BORDER
    if fill is not None:
        cell.fill = fill
    if isinstance(value, (dt.datetime, dt.date)):
        cell.number_format = "dd-mmm-yyyy"
        cell.alignment = Alignment(horizontal="center", vertical="center")
    elif isinstance(value, (int, float)) and not isinstance(value, bool):
        cell.number_format = "0" if code_col else "#,##0"
        cell.alignment = Alignment(horizontal="right", vertical="center")
    elif value == "-":
        cell.alignment = Alignment(horizontal="center", vertical="center")
    else:
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=indent)


def _write_sheet(ws, sheet: Sheet) -> None:
    width = max(sheet.width, 1)
    header = sheet.header or []
    code_cols = {i for i, h in enumerate(header) if _is_code_column(h)}
    col_len = [0] * width
    r = 1

    # Title rows (report name, date, etc. above the table)
    for t in sheet.title_rows:
        vals = [display_value(v) for v in t]
        filled = [i for i, v in enumerate(vals) if v is not None]
        for i, v in enumerate(vals):
            if v is not None:
                c = ws.cell(row=r, column=i + 1, value=v)
                c.font = TITLE_FONT
                c.alignment = Alignment(horizontal="left", vertical="center")
        if len(filled) == 1 and width > 1:
            ws.merge_cells(start_row=r, start_column=filled[0] + 1, end_row=r, end_column=width)
        ws.row_dimensions[r].height = 22
        r += 1
    if sheet.title_rows:
        r += 1  # spacer

    header_row = None
    if sheet.header is not None:
        header_row = r
        for i, h in enumerate(header):
            c = ws.cell(row=r, column=i + 1, value=h)
            c.font, c.fill, c.border = HEADER_FONT, HEADER_FILL, BORDER
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            col_len[i] = max(col_len[i], min(_text_len(h), 18))
        ws.row_dimensions[r].height = 32
        if sheet.groups is not None:
            for ci in sheet.group_value_cols:
                ws.cell(row=r, column=ci + 1).comment = Comment(
                    "Item group rows show the AVERAGE of the item's records. "
                    "Click the + on the left to expand an item.", "Report Fixer")
        r += 1

    def put_row(values, bold=False, fill=None, indent_col=None):
        nonlocal r
        for i in range(width):
            v = display_value(values[i]) if i < len(values) else None
            c = ws.cell(row=r, column=i + 1, value=v)
            _style_cell(c, v, i in code_cols, bold=bold, fill=fill, indent=2 if i == indent_col else 0)
            col_len[i] = max(col_len[i], _text_len(v) + (2 if i == indent_col else 0))
        r += 1

    if sheet.groups is not None:
        ws.sheet_properties.outlinePr = Outline(summaryBelow=False, summaryRight=False)
        item_col = sheet.item_col
        for g in sheet.groups:
            if len(g.details) == 1:
                put_row(g.details[0])
                continue
            label = list(g.summary)
            if item_col is not None:
                label[item_col] = f"{g.item}  ({len(g.details)} records)"
            summary_row = r
            put_row(label, bold=True, fill=GROUP_FILL)
            for d in g.details:
                ws.row_dimensions[r].outline_level = 1
                ws.row_dimensions[r].hidden = True
                put_row(d, indent_col=item_col)
            ws.row_dimensions[summary_row].collapsed = True
        for row in sheet.rows:
            put_row(row, bold=True)
    else:
        for n, row in enumerate(sheet.rows):
            put_row(row, fill=ZEBRA_FILL if n % 2 else None)

    for i, ln in enumerate(col_len):
        ws.column_dimensions[get_column_letter(i + 1)].width = max(MIN_WIDTH, min(MAX_WIDTH, ln + 3))

    # Freeze header, filter, and print setup (landscape, fit to page width, repeat header)
    if header_row is not None:
        ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
        last = r - 1
        if sheet.groups is None and last > header_row:
            ws.auto_filter.ref = f"A{header_row}:{get_column_letter(width)}{last}"
        ws.print_title_rows = f"{header_row}:{header_row}"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.oddFooter.center.text = "Page &P of &N"
    ws.sheet_view.showGridLines = False


def write_workbook(sheets: List[Sheet]) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)
    for s in sheets:
        ws = wb.create_sheet(title=s.name[:31])
        _write_sheet(ws, s)
    if not wb.worksheets:
        wb.create_sheet("Sheet1")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
