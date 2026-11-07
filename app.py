"""Streamlit front-end: upload the downloaded report, download the corrected one.

Run locally:  streamlit run app.py
"""
from __future__ import annotations

import os

import pandas as pd
import streamlit as st

from report_fixer import Options, display_value, fix_report, read_workbook
from report_fixer.processor import find_header_row, is_day_consumption, norm, resolve_sheets

st.set_page_config(page_title="Excel Report Corrector", page_icon="📊", layout="wide")

st.markdown(
    """
    <style>
      .block-container {padding-top: 2rem; max-width: 1300px;}
      div[data-testid="stMetricValue"] {font-size: 1.6rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📊 Excel Report Corrector")
st.caption("Upload the report downloaded from the portal — get back a cleaned, correctly calculated, print-ready Excel file.")

with st.expander("What gets corrected?", expanded=False):
    st.markdown(
        """
1. **Completely empty / NULL rows removed** (rows with only some blank cells are kept as-is).
2. **Day Consumption converted from tons to kg**, and **No. of Days Consumption** recalculated from the kg value.
3. **Raw Aging Stock** – removes Lot Code, Lot No, Supplier Lot No, Category, GSM, Qty, Machine Line.
4. **All TagEnv Aging Stock** – removes Lot Code, Lot No, Supplier Lot No, Width, GSM, Source Type, and groups
   repeated Item Names into one row showing the **average QTN and Holding Days**. In Excel, click the **+** next to
   an item to expand its individual records; on this page, click an item row to see them.
5. **Zero values shown as "-"** on every sheet (including the 4th sheet).
6. **Decimals rounded** to the nearest whole number.
7. **Professional formatting** – styled headers, borders, alignment, auto column widths, frozen header,
   filters, and landscape fit-to-width print setup.
        """
    )

uploaded = st.file_uploader("Upload the Excel report", type=["xlsx", "xlsm", "xls"])

if not uploaded:
    st.info("⬆️ Upload the downloaded report to begin.")
    st.stop()

data = uploaded.getvalue()


@st.cache_data(show_spinner=False)
def _read(raw: bytes, name: str):
    return read_workbook(raw, name)


try:
    original = _read(data, uploaded.name)
except Exception as e:  # noqa: BLE001
    st.error(f"Could not read this file: {e}")
    st.stop()

names = [n for n, _ in original]
auto_raw, auto_tag = resolve_sheets(names, Options())

# ----------------------------------------------------------------- settings
with st.sidebar:
    st.header("⚙️ Settings")
    st.caption("Sheets are detected automatically. Change only if detection is wrong.")
    none_opt = "— none —"
    raw_sel = st.selectbox("Raw Aging Stock sheet", [none_opt] + names,
                           index=(names.index(auto_raw) + 1) if auto_raw else 0)
    tag_sel = st.selectbox("All TagEnv Aging Stock sheet", [none_opt] + names,
                           index=(names.index(auto_tag) + 1) if auto_tag else 0)
    st.divider()
    convert = st.checkbox("Convert Day Consumption tons → kg", value=True)
    recalc = st.checkbox("Recalculate No. of Days Consumption", value=True, disabled=not convert)

    # every header that appears in a sheet with a Day Consumption column
    stock_choices = []
    for _, rows in original:
        h = find_header_row(rows)
        if h is None:
            continue
        hdr = [str(x) for x in rows[h] if x is not None]
        if any("consumption" in x.lower() and "day" in x.lower() and "no" not in x.lower() for x in hdr):
            stock_choices += [x for x in hdr if x not in stock_choices]
    stock_sel = st.selectbox("Stock quantity column used for No. of Days", ["Auto-detect"] + stock_choices,
                             help="No. of Days Consumption = Stock ÷ Day Consumption (Kg)", disabled=not recalc)

opts = Options(
    raw_aging_sheet="" if raw_sel == none_opt else raw_sel,
    tagenv_sheet="" if tag_sel == none_opt else tag_sel,
    convert_day_consumption=convert,
    recalc_days=recalc,
    stock_column=None if stock_sel == "Auto-detect" else stock_sel,
)
with st.spinner("Correcting the report…"):
    try:
        sheets, xlsx = fix_report(data, uploaded.name, opts)
    except Exception as e:  # noqa: BLE001
        st.exception(e)
        st.stop()

# ----------------------------------------------------------------- validation
problems = []
if len(sheets) != 4:
    problems.append(f"Expected the 4-sheet report, but this file has {len(sheets)} sheet(s).")
res_raw, res_tag = resolve_sheets(names, opts)
if raw_sel != none_opt and not res_raw:
    problems.append("Raw Aging Stock sheet not found — its columns were not removed. Pick it in the sidebar.")
if tag_sel != none_opt and not res_tag:
    problems.append("All TagEnv Aging Stock sheet not found — items were not grouped. Pick it in the sidebar.")
for s in sheets:
    if s.name == res_tag and s.groups is None:
        problems.append(f"'{s.name}': Item Name column not found, so items were not grouped.")
    elif s.name == res_tag and len(s.group_value_cols) < 2:
        problems.append(f"'{s.name}': QTN or Holding Days column not found, so its average could not be shown.")
if convert and not any(s.header and any(h is not None and is_day_consumption(norm(h)) for h in s.header)
                       for s in sheets):
    problems.append("No Day Consumption column found — nothing was converted to kg.")
for s in sheets:
    if s.header is None and s.rows:
        problems.append(f"'{s.name}': header row could not be identified; only cleaning and formatting were applied.")

# ----------------------------------------------------------------- result
base = os.path.splitext(uploaded.name)[0]
if problems:
    st.warning("**Please check:**\n\n" + "\n".join(f"- {p}" for p in problems))
st.success("✅ Report corrected.")
c1, c2, c3 = st.columns([2, 1, 1])
with c1:
    st.download_button("⬇️ Download corrected report", xlsx, file_name=f"{base}_Corrected.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary", width="stretch")
c2.metric("Sheets", len(sheets))
c3.metric("Data rows", sum(len(s.rows) + sum(len(g.details) for g in (s.groups or []))
                           + sum(len(sec.rows) for sec in s.sections) for s in sheets))


def _frame(header, rows):
    width = len(header) if header else max((len(r) for r in rows), default=0)
    cols = []
    for i in range(width):
        h = str(header[i]) if header and header[i] is not None else f"Column {i + 1}"
        while h in cols:
            h += " "
        cols.append(h)

    def fmt(v):
        v = display_value(v)
        if v is None:
            return ""
        if hasattr(v, "strftime"):
            return v.strftime("%d-%b-%Y")
        if isinstance(v, int):
            return f"{v:,}"
        return str(v)

    return pd.DataFrame([[fmt(r[i]) if i < len(r) else "" for i in range(width)] for r in rows], columns=cols)


st.subheader("Preview")
tabs = st.tabs([s.name for s in sheets])
for tab, s in zip(tabs, sheets):
    with tab:
        for note in s.notes:
            st.markdown(f"- {note}")
        if s.header is None and not s.rows:
            st.write("_Empty sheet_")
            continue
        if s.groups is None:
            st.dataframe(_frame(s.header, s.rows), width="stretch", hide_index=True,
                         height=480 if not s.sections else "auto")
            for sec in s.sections:
                title = next((str(c) for t in sec.title_rows for c in t if c is not None), "")
                if title:
                    st.markdown(f"**{title}**")
                st.dataframe(_frame(sec.header, sec.rows), width="stretch", hide_index=True)
            continue

        # Grouped TagEnv view: one row per item; click a row to see its records.
        query = st.text_input("Search item", key=f"q_{s.name}", placeholder="Type part of an item name…")
        groups = [g for g in s.groups if query.lower() in g.item.lower()]
        summary_rows = []
        for g in groups:
            row = list(g.summary)
            row.insert(0, len(g.details))
            summary_rows.append(row)
        summary_header = ["Records"] + [
            f"Average {h}" if i in s.group_value_cols else h for i, h in enumerate(s.header)]
        summary_df = _frame(summary_header, summary_rows)
        st.caption("Each item appears once, with the average QTN and Holding Days of its records. "
                   "**Click an item row** to expand all of its individual records.")
        event = st.dataframe(summary_df, width="stretch", hide_index=True, height=420,
                             on_select="rerun", selection_mode="single-row", key=f"grid_{s.name}")
        picked = event.selection.rows if event and hasattr(event, "selection") else []
        if picked:
            g = groups[picked[0]]
            with st.expander(f"{g.item} — {len(g.details)} record(s)", expanded=True):
                st.dataframe(_frame(s.header, g.details), width="stretch", hide_index=True)
        if s.rows:
            st.markdown("**Other rows (totals / unnamed)**")
            st.dataframe(_frame(s.header, s.rows), width="stretch", hide_index=True)
