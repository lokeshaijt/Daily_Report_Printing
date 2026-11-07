# Excel Report Corrector (Streamlit)

Upload the stock report downloaded from the portal and get back a corrected, consistently formatted, print-ready Excel file.

## What it does

| # | Rule |
|---|------|
| 1 | Removes rows where the **Item Name has no value in any other column** (all empty, `NULL` or zero), plus fully empty rows. Grand Total rows and section titles are kept. Sheets with several tables (e.g. Master Roll and Slitting Process) are handled table by table, each with its own header. Rows with at least one other value are kept unchanged. |
| 2 | **Day Consumption** tons → kg (×1000, header renamed `… (Kg)`); **No. of Days Consumption** recalculated as `Stock ÷ Day Consumption (Kg)`. If no stock column exists, the old value is divided by 1000. |
| 3 | **Raw Aging Stock**: removes Lot Code, Lot No, Supplier Lot No, Category, GSM, Qty, Machine Line. |
| 4 | **All TagEnv Aging Stock**: removes Lot Code, Lot No, Supplier Lot No, Width, GSM, Source Type. Repeated Item Names appear once, with the **average QTN and Holding Days**. In Excel, click **+** beside the item to expand its records. On the website, click the item row. |
| 5/7 | Zero values are shown as `-` on every sheet, including the 4th sheet. |
| 6 | Decimals are rounded half-up to whole numbers, except values like 0.4 (between -1 and 1, not zero), which are not rounded to a whole number; long decimals are shortened to 2 decimal places (0.3456 → 0.35). Averages and calculations use full precision first. |
| 8 | Uniform fonts, styled headers, borders, alignment, column widths and row heights sized to the content (no cut-off text), frozen header, filters, and A4 landscape fit-to-width printing with a repeated header row. |

The uploaded file is validated: the website warns if it doesn't have 4 sheets, or if an expected sheet or column (Day Consumption, Item Name, QTN, Holding Days) isn't found. Sheets, header rows and column names are detected automatically, and title rows above the table are kept. If a sheet is detected wrongly, change it in the sidebar.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Push this repo to GitHub.
2. Go to https://share.streamlit.io and choose **New app**.
3. Pick this repository and branch, and set **Main file path** to `app.py`.
4. Click **Deploy**. Dependencies install from `requirements.txt`.

## Project layout

```
app.py                  Streamlit UI (upload, settings, preview, download)
report_fixer/reader.py  Reads .xlsx / .xlsm / .xls (and HTML-based .xls exports)
report_fixer/processor.py  All correction rules
report_fixer/writer.py  Formatted Excel output
tests/                  pytest tests built on a synthetic sample report
```

Run the tests with `python -m pytest -q`.
