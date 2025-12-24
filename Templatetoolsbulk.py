import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# =====================================================
# PAGE CONFIG
# =====================================================
st.set_page_config(
    page_title="Bulk Profitability Checker – Multi Coverage",
    layout="wide"
)

st.title("📊 Bulk Profitability Checker – Multi Coverage")
st.caption("Profitability checking tool (PAR, EQVET, MB, PL, FG)")

# =====================================================
# USER ASSUMPTIONS
# =====================================================
st.sidebar.header("Asumsi Profitability")

loss_ratio = st.sidebar.number_input("Asumsi Loss Ratio", 0.0, 1.0, 0.45, 0.01)
premi_xol = st.sidebar.number_input("Asumsi Premi XOL", 0.0, 1.0, 0.12, 0.01)
expense_ratio = st.sidebar.number_input("Asumsi Expense", 0.0, 1.0, 0.20, 0.01)

# =====================================================
# CONSTANTS
# =====================================================
KOMISI_BPPDAN = 0.35
KOMISI_MAIPARK = 0.30

RATE_MB_INDUSTRIAL = 0.0015
RATE_MB_NON_INDUSTRIAL = 0.0001
RATE_PL = 0.0005
RATE_FG = 0.0010

COVERAGE_ORDER = [
    "PAR",
    "EQVET",
    "MACHINERY",
    "PUBLIC LIABILITY",
    "FIDELITY GUARANTEE"
]

# =====================================================
# DISPLAY CONFIG
# =====================================================
AMOUNT_COLS = [
    "Kurs",
    "TSI Full Value original currency",
    "Limit of Liability original currency",
    "Top Risk original currency",
    "TSI_IDR", "Limit_IDR", "TopRisk_IDR",
    "ExposureBasis", "Exposure_Loss",
    "S_Askrindo", "Pool_amt", "Fac_amt", "OR_amt",
    "Prem100", "Prem_Askrindo", "Prem_POOL", "Prem_Fac", "Prem_OR",
    "Acq_amt", "Komisi_POOL", "Komisi_Fakultatif",
    "EL_100", "EL_Askrindo", "EL_POOL", "EL_Fac",
    "XL_cost", "Expense", "Result"
]

PERCENT_COLS = [
    "Rate",
    "% Askrindo Share",
    "% Fakultatif Share",
    "% Komisi Fakultatif",
    "% LOL Premi",
    "%POOL",
    "%OR",
    "%Result"
]

INT_COLS = ["Kode Okupasi"]

# =====================================================
# FILE UPLOAD
# =====================================================
uploaded_file = st.file_uploader(
    "📁 Upload Excel (PAR, EQVET, MACHINERY, PUBLIC LIABILITY, FIDELITY GUARANTEE)",
    type=["xlsx"]
)

process_btn = st.button("🚀 Proses Profitability")

# =====================================================
# DISPLAY FORMATTER (STREAMLIT)
# =====================================================
def format_display(df):
    fmt = {}
    for c in AMOUNT_COLS:
        if c in df.columns:
            fmt[c] = "{:,.0f}"
    for c in PERCENT_COLS:
        if c in df.columns:
            fmt[c] = "{:.2%}"
    for c in INT_COLS:
        if c in df.columns:
            fmt[c] = "{:.0f}"
    return df.style.format(fmt)

# =====================================================
# CORE ENGINE
# =====================================================
def run_profitability(df, coverage):

    df = df.copy()
    df.columns = [c.strip() for c in df.columns]

    df["TSI_IDR"] = df["TSI Full Value original currency"] * df["Kurs"]
    df["Limit_IDR"] = df["Limit of Liability original currency"] * df["Kurs"]
    df["TopRisk_IDR"] = df["Top Risk original currency"] * df["Kurs"]

    df["ExposureBasis"] = df[["Limit_IDR", "TopRisk_IDR"]].max(axis=1)
    df["Exposure_Loss"] = np.where(df["Limit_IDR"] > 0, df["Limit_IDR"], df["TSI_IDR"])

    df["S_Askrindo"] = df["% Askrindo Share"] * df["ExposureBasis"]

    if coverage == "PAR":
        df["Pool_amt"] = np.minimum(0.025 * df["S_Askrindo"], 500_000_000 * df["% Askrindo Share"])
        komisi_pool = KOMISI_BPPDAN
    elif coverage == "EQVET":
        rate = np.where(df["Wilayah Gempa Prioritas"] == "DKI-JABAR-BANTEN", 0.10, 0.25)
        df["Pool_amt"] = np.minimum(rate * df["S_Askrindo"], 10_000_000_000 * df["% Askrindo Share"])
        komisi_pool = KOMISI_MAIPARK
    else:
        df["Pool_amt"] = 0
        komisi_pool = 0

    df["%POOL"] = np.where(df["ExposureBasis"] > 0, df["Pool_amt"] / df["ExposureBasis"], 0)

    df["Fac_amt"] = df["% Fakultatif Share"] * df["ExposureBasis"]
    df["OR_amt"] = df["S_Askrindo"] - df["Pool_amt"] - df["Fac_amt"]
    df["%OR"] = df["OR_amt"] / df["ExposureBasis"]

    df["Prem100"] = df["Rate"] * df["% LOL Premi"] * df["TSI_IDR"]
    df["Prem_Askrindo"] = df["Prem100"] * df["% Askrindo Share"]
    df["Prem_POOL"] = df["Prem100"] * df["%POOL"]
    df["Prem_Fac"] = df["Prem100"] * df["% Fakultatif Share"]
    df["Prem_OR"] = df["Prem100"] * df["%OR"]

    df["Acq_amt"] = df["% Akuisisi"] * df["Prem_Askrindo"]
    df["Komisi_POOL"] = komisi_pool * df["Prem_POOL"]
    df["Komisi_Fakultatif"] = df["% Komisi Fakultatif"].fillna(0) * df["Prem_Fac"]

    if coverage == "MACHINERY":
        rate_acuan = np.where(df["Occupancy"].str.lower() == "industrial", RATE_MB_INDUSTRIAL, RATE_MB_NON_INDUSTRIAL)
        df["EL_100"] = rate_acuan * df["Exposure_Loss"] * loss_ratio
    elif coverage == "PUBLIC LIABILITY":
        df["EL_100"] = RATE_PL * df["Exposure_Loss"] * loss_ratio
    elif coverage == "FIDELITY GUARANTEE":
        df["EL_100"] = RATE_FG * df["Exposure_Loss"] * loss_ratio
    else:
        df["EL_100"] = loss_ratio * df["Prem100"]

    df["EL_Askrindo"] = df["EL_100"] * df["% Askrindo Share"]
    df["EL_POOL"] = df["EL_100"] * df["%POOL"]
    df["EL_Fac"] = df["EL_100"] * df["% Fakultatif Share"]

    df["XL_cost"] = premi_xol * df["Prem_OR"]
    df["Expense"] = expense_ratio * df["Prem_Askrindo"]

    df["Result"] = (
        df["Prem_Askrindo"]
        - df["Prem_POOL"]
        - df["Prem_Fac"]
        - df["Acq_amt"]
        + df["Komisi_POOL"]
        + df["Komisi_Fakultatif"]
        - df["EL_Askrindo"]
        + df["EL_POOL"]
        + df["EL_Fac"]
        - df["XL_cost"]
        - df["Expense"]
    )

    df["%Result"] = np.where(df["Prem_Askrindo"] != 0, df["Result"] / df["Prem_Askrindo"], 0)

    return df

# =====================================================
# EXCEL HELPERS
# =====================================================
def add_total_row(df):
    total = {}
    for col in df.columns:
        if col in AMOUNT_COLS:
            total[col] = df[col].sum()
        elif col == "%Result":
            total[col] = df["Result"].sum() / df["Prem_Askrindo"].sum() if df["Prem_Askrindo"].sum() != 0 else 0
        else:
            total[col] = np.nan
    return pd.concat([df, pd.DataFrame([total], index=["JUMLAH"])])

def write_formatted_sheet(writer, df, sheet_name):
    wb = writer.book
    ws = wb.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = ws

    df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)

    fmt_header = wb.add_format({"bold": True, "align": "center", "border": 1})
    fmt_amt = wb.add_format({"num_format": "#,##0"})
    fmt_pct = wb.add_format({"num_format": "0.00%"})
    fmt_int = wb.add_format({"num_format": "0"})
    fmt_red = wb.add_format({"bg_color": "#F8CBAD", "num_format": "0.00%"})
    fmt_green = wb.add_format({"bg_color": "#C6EFCE", "num_format": "0.00%"})

    for c, col in enumerate(df.columns):
        ws.write(0, c, col, fmt_header)
        if col in AMOUNT_COLS:
            ws.set_column(c, c, 18, fmt_amt)
        elif col in PERCENT_COLS:
            ws.set_column(c, c, 14, fmt_pct)
        elif col in INT_COLS:
            ws.set_column(c, c, 12, fmt_int)
        else:
            ws.set_column(c, c, 20)

    ws.freeze_panes(1, 0)

    if "%Result" in df.columns:
        idx = df.columns.get_loc("%Result")
        ws.conditional_format(1, idx, len(df), idx, {
            "type": "cell", "criteria": "<", "value": 0.05, "format": fmt_red
        })
        ws.conditional_format(1, idx, len(df), idx, {
            "type": "cell", "criteria": ">=", "value": 0.05, "format": fmt_green
        })

# =====================================================
# RUN
# =====================================================
if process_btn and uploaded_file:

    xls = pd.ExcelFile(uploaded_file)
    results = {c: run_profitability(pd.read_excel(xls, c), c) for c in COVERAGE_ORDER}

    rows, tp, tr = [], 0, 0
    for c in COVERAGE_ORDER:
        p = results[c]["Prem_Askrindo"].sum()
        r = results[c]["Result"].sum()
        rows.append([c, p, r, r / p if p else 0])
        tp += p; tr += r

    rows.append(["JUMLAH", tp, tr, tr / tp if tp else 0])
    summary_df = pd.DataFrame(rows, columns=["Coverage", "Jumlah Premi Ourshare", "Result", "%Result"])

    st.subheader("📊 Summary Profitability")
    st.dataframe(format_display(summary_df), use_container_width=True)

    for c in COVERAGE_ORDER:
        dfc = add_total_row(results[c])
        st.subheader(f"📋 Detail {c}")
        st.dataframe(format_display(dfc), use_container_width=True)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        write_formatted_sheet(writer, summary_df, "SUMMARY")
        for c in COVERAGE_ORDER:
            write_formatted_sheet(writer, add_total_row(results[c]), c)

    st.download_button(
        "📥 Download Excel Output",
        data=output.getvalue(),
        file_name="Profitability_Output_All_Coverages.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
