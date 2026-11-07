# Excel Report Corrector (Streamlit)

Upload the stock report downloaded from the portal and get back a corrected, consistently formatted, print-ready Excel file.

## What it does

| # | Rule |
|---|------|
| 1 | Removes rows where **every** cell is empty / `NULL`. Rows with only some blank cells are kept unchanged. |
| 2 | **Day Consumption** tons → kg (×1000, header renamed `… (Kg)`); **No. of Days Consumption** recalculated as `Stock ÷ Day Consumption (Kg)`. If no stock column exists, the old value is divided by 1000. |
| 3 | **Raw Aging Stock**: removes Lot Code, Lot No, Supplier Lot No, Category, GSM, Qty, Machine Line. |
| 4 | **All TagEnv Aging Stock**: removes Lot Code, Lot No, Supplier Lot No, Width, GSM, Source Type. Repeated Item Names appear once, with the **average QTN and Holding Days**. In Excel, click **+** beside the item to expand its records. On the website, click the item row. |
| 5/7 | Zero values are shown as `-` on every sheet, including the 4th sheet. |
| 6 | Decimals are rounded half-up to whole numbers. Averages and calculations use full precision first. |
| 8 | Styled headers, borders, alignment, auto column widths, frozen header, filters, and A4 landscape fit-to-width printing with a repeated header row. |

Sheets, header rows and column names are detected automatically, and title rows above the table are kept. If a sheet is detected wrongly, change it in the sidebar.

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
