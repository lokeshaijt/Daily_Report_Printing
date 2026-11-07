import datetime as dt
import io

from openpyxl import Workbook, load_workbook

from report_fixer import Options, fix_report


def make_sample() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Stock Summary"
    ws.append(["Daily Stock Report"])
    ws.append([])
    ws.append(["Item Code", "Item Name", "Stock Qty", "Day Consumption", "No. of Days Consumption"])
    ws.append(["1001", "Kraft Paper", 5000, 0.25, 20000])      # 5000 kg / 0.25 t
    ws.append([None, None, None, None, None])                   # fully empty -> removed
    ws.append(["1002", "Board", 0, 0.5, 0])
    ws.append(["NULL", None, "", None, None])                   # NULL/empty -> removed
    ws.append(["1003", None, 1234.567, None, None])            # partially empty -> kept
    ws.append([None, "Empty Item", None, "NULL", None])         # item name only -> removed

    raw = wb.create_sheet("Raw Aging Stock")
    raw.append(["Item Name", "Lot Code", "Lot No", "Supplier Lot No", "Category", "GSM", "Qty",
                "Machine Line", "Stock Qty", "Aging Days", "GRN Date"])
    raw.append(["Paper A", "LC1", "L1", "S1", "Cat", 80, 10, "M1", 100.4, 12.5, dt.datetime(2026, 1, 5)])
    raw.append(["Paper B", "LC2", "L2", "S2", "Cat", 90, 0, "M2", 0, 0.4, dt.datetime(2026, 2, 5)])

    tag = wb.create_sheet("All TagEnv Aging Stock")
    tag.append(["Item Name", "Lot Code", "Lot No", "Supplier Lot No", "Width", "GSM", "Source Type",
                "UOM", "QTN", "Holding Days", "Location"])
    tag.append(["Film X", "a", "1", "s", 100, 50, "PO", "KG", 10, 5, "R1"])
    tag.append(["Film Y", "b", "2", "s", 100, 50, "PO", "KG", 7, 3, "R2"])
    tag.append(["Film X", "c", "3", "s", 100, 50, "PO", "KG", 21, 8, "R3"])
    tag.append([None] * 11)
    tag.append(["Film Z", "d", "4", "s", 100, 50, "PO", None, None, None, None])  # only removed cols -> dropped
    tag.append(["Film X", None, None, None, None, None, None, None, None, "NULL", None])  # item only -> dropped

    fourth = wb.create_sheet("Fourth")
    fourth.append(["Name", "Jan", "Feb"])
    fourth.append(["A", 0, 2.6])
    fourth.append(["B", 0.0, "0"])
    fourth.append(["C", 0.75, -0.2])
    fourth.append(["D", 0.3456, 0.00456])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def values(ws):
    return [list(r) for r in ws.iter_rows(values_only=True)]


def test_full_pipeline():
    sheets, xlsx = fix_report(make_sample(), "sample.xlsx", Options())
    wb = load_workbook(io.BytesIO(xlsx))
    assert wb.sheetnames == ["Stock Summary", "Raw Aging Stock", "All TagEnv Aging Stock", "Fourth"]

    # 1 + 2: empty rows removed, tons -> kg, days recalculated = stock / kg
    s = values(wb["Stock Summary"])
    hdr = s[2]
    assert hdr == ["Item Code", "Item Name", "Stock Qty", "Day Consumption (Kg)", "No. of Days Consumption"]
    body = s[3:]
    assert len(body) == 3
    assert body[0] == [1001, "Kraft Paper", 5000, 250, 20]
    assert body[1] == [1002, "Board", "-", 500, "-"]
    assert body[2] == [1003, None, 1235, None, None]

    # 3: columns removed, rounding, zeros -> "-"
    r = values(wb["Raw Aging Stock"])
    assert r[0] == ["Item Name", "Stock Qty", "Aging Days", "GRN Date"]
    assert r[1][:3] == ["Paper A", 100, 13]
    assert r[2][:3] == ["Paper B", "-", 0.4]   # 0.x values are kept as they are

    # 4: columns removed, grouped with averages, details hidden in an outline
    ws = wb["All TagEnv Aging Stock"]
    t = values(ws)
    assert t[0] == ["Item Name", "UOM", "QTN", "Holding Days", "Location"]
    assert t[1] == ["Film X  (2 records)", "KG", 16, 7, None]   # avg(10,21)=15.5->16, avg(5,8)=6.5->7
    assert t[2][2:4] == [10, 5] and t[3][2:4] == [21, 8]
    assert ws.row_dimensions[3].outline_level == 1 and ws.row_dimensions[3].hidden
    assert t[4] == ["Film Y", "KG", 7, 3, "R2"]
    assert len(t) == 5

    # 5/7: zeros on 4th sheet
    f = values(wb["Fourth"])
    assert f[1] == ["A", "-", 3]
    assert f[2] == ["B", "-", "-"]
    assert f[3] == ["C", 0.75, -0.2]
    assert f[4] == ["D", 0.35, 0.0046]   # long decimals shortened


def test_days_fallback_without_stock_column():
    wb = Workbook()
    ws = wb.active
    ws.append(["Item", "Day Consumption (MT)", "No of Days Consumption"])
    ws.append(["X", 2, 3000])
    buf = io.BytesIO()
    wb.save(buf)
    _, xlsx = fix_report(buf.getvalue(), "x.xlsx")
    rows = values(load_workbook(io.BytesIO(xlsx)).active)
    assert rows[0][1] == "Day Consumption (Kg)"
    assert rows[1] == ["X", 2000, 3]


def test_item_rows_without_details_removed_spec_example():
    wb = Workbook()
    ws = wb.active
    ws.append(["Item Name", "QTN", "Stock", "Holding Days"])
    ws.append(["Paper A", 100, 50, 10])
    ws.append(["Paper B", "NULL", 20, "NULL"])
    ws.append(["Paper C", "NULL", "NULL", "NULL"])
    buf = io.BytesIO()
    wb.save(buf)
    _, xlsx = fix_report(buf.getvalue(), "x.xlsx")
    out = load_workbook(io.BytesIO(xlsx)).active
    rows = values(out)
    assert [r[0] for r in rows[1:]] == ["Paper A", "Paper B"]
    assert all(out.row_dimensions[i].height for i in range(1, out.max_row + 1))


def test_printing_report_two_tables_and_zero_only_items():
    """Mirrors the real 'PRINTING DAILY REPORT' layout: Master Roll table, then a Slitting Process table."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Printing"
    ws.append(["PRINTING DAILY REPORT - 05 Nov 2026"])
    ws.append(["Master Roll"])
    ws.append([])
    ws.append(["Item Name", "Master Reel Opening in Kgs", "Master Reel Inward in Kgs", "Transfer to Slit in Kgs",
               "Master Reel Closing in Kgs", "Pending PO Quantity", "Day Consumption", "No of days Consumption"])
    ws.append(["PAPER 70 GSM", 76960, 0, 0, 76960, 0, 5.77, 13338])
    ws.append(["PAPER 100 GSM", 0, 0, 0, 0, 0, 0, 0])          # all zeros -> removed
    ws.append(["PAPER 170 GSM (UC)", 0, None, "NULL", 0, None, None, 0])  # zeros / blanks -> removed
    ws.append(["PAPER 90 GSM FOIL", 0, 0, 0, 0, 0, 0.07, 0])   # has a day consumption -> kept
    ws.append(["Grand Total", 76960, 0, 0, 76960, 0, None, None])
    ws.append(["Slitting Process"])
    ws.append(["Item Name", "Slit Opening in Kgs", "Slit Inward in Kgs", "Receive from other branch in Kgs",
               "Transfer to other branch in Kgs", "Transfer to Printing (Consumbtion)", "Wastage",
               "Wastage Percentage", "Closing stock in Kgs"])
    ws.append(["PAPER 70 GSM (PIGMENT)", 510, 715, 0, 0, 510, 11.02, 2.16, 704])
    ws.append(["PAPER 80 GSM", 0, 0, 0, 0, 0, 0, 0, 0])         # all zeros -> removed
    buf = io.BytesIO()
    wb.save(buf)
    sheets, xlsx = fix_report(buf.getvalue(), "p.xlsx")
    rows = values(load_workbook(io.BytesIO(xlsx)).active)
    items = [r[0] for r in rows]
    assert "PAPER 100 GSM" not in items and "PAPER 170 GSM (UC)" not in items and "PAPER 80 GSM" not in items
    assert "PAPER 90 GSM FOIL" in items and "Grand Total" in items and "Slitting Process" in items
    main = rows[items.index("PAPER 70 GSM")]
    assert main[6:8] == [5770, 13]                       # tons -> kg, days = closing / kg
    assert rows[items.index("PAPER 90 GSM FOIL")][6] == 70
    slit = rows[items.index("PAPER 70 GSM (PIGMENT)")]
    assert slit[6:9] == [11, 2, 704]                      # second table NOT treated as Day Consumption
    assert rows[items.index("Slitting Process") + 1][0] == "Item Name"
