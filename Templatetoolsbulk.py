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
# COLUMN RULES
# =====================================================
NON_SUM_COVERAGE_COLS = [
    "Kode Okupasi",
    "Kurs",
    "TSI Full Value original currency",
    "Limit of Liability original currency",
    "Top Risk original currency",
    "Rate",
    "% Askrindo Share",
    "% Fakultatif Share",
    "% Komisi Fakultatif",
    "% LOL Premi",
    "%POOL",
    "%OR"
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
# STREAMLIT DISPLAY FORMAT
# =====================================================
def format_display(df):
    fmt = {}
    for col in df.columns:
        if col in PERCENT_COLS:
            fmt[col] = "{:.2%}"
        elif col in INT_COLS:
            fmt[col] = "{:.0f}"
        elif pd.api.types.is_numeric_dtype(df[col]):
            fmt[col] = "{:,.0f}"
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
        df["Pool_amt"] = np.minimum(
            0.025 * df["S_Askrindo"],
            500_000_000 * df["% Askrindo Share"]
        )
        komisi_pool = KOMISI_BPPDAN

    elif coverage == "EQVET":
        rate = np.where(
            df["Wilayah Gempa Prioritas"] == "DKI-JABAR-BANTEN",
            0.10,
            0.25
        )
        df["Pool_amt"] = np.minimum(
            rate * df["S_Askrindo"],
            10_000_000_000 * df["% Askrindo Share"]
        )
        komisi_pool = KOMISI_MAIPARK

    else:
        df["Pool_amt"] = 0
        komisi_pool = 0

    df["%POOL"] = np.where(df["ExposureBasis"] > 0, df["Pool_amt"] / df["ExposureBasis"], 0)

    df["Fac_amt"] = df["% Fakultatif Share"] * df["ExposureBasis"]
    df["OR_amt"] = df["S_Askrindo"] - df["Pool_amt"] - df["Fac_amt"]
    df["%OR"] = np.where(df["ExposureBasis"] > 0, df["OR_amt"] / df["ExposureBasis"], 0)

    df["Prem100"] = df["Rate"] * df["% LOL Premi"] * df["TSI_IDR"]
    df["Prem_Askrindo"] = df["Prem100"] * df["% Askrindo Share"]
    df["Prem_POOL"] = df["Prem100"] * df["%POOL"]
    df["Prem_Fac"] = df["Prem100"] * df["% Fakultatif Share"]
    df["Prem_OR"] = df["Prem100"] * df["%OR"]

    df["Acq_amt"] = df["% Akuisisi"] * df["Prem_Askrindo"]
    df["Komisi_POOL"] = komisi_pool * df["Prem_POOL"]
    df["Komisi_Fakultatif"] = df["% Komisi Fakultatif"].fillna(0) * df["Prem_Fac"]

    if coverage == "MACHINERY":
        rate_acuan = np.where(
            df["Occupancy"].str.lower() == "industrial",
            RATE_MB_INDUSTRIAL,
            RATE_MB_NON_INDUSTRIAL
        )
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

    df["%Result"] = np.where(
        df["Prem_Askrindo"] != 0,
        df["Result"] / df["Prem_Askrindo"],
        0
    )

    return df

# =====================================================
# TOTAL ROW HANDLER
# =====================================================
def add_total_row(df, is_summary=False):

    total = {}

    denom_col = "Jumlah Premi Ourshare" if is_summary else "Prem_Askrindo"

    for col in df.columns:

        if col == "%Result":
            if denom_col in df.columns and df[denom_col].sum() != 0:
                total[col] = df["Result"].sum() / df[denom_col].sum()
            else:
                total[col] = 0

        elif is_summary:
            if pd.api.types.is_numeric_dtype(df[col]):
                total[col] = df[col].sum()
            else:
                total[col] = np.nan

        else:
            if col in NON_SUM_COVERAGE_COLS:
                total[col] = np.nan
            elif pd.api.types.is_numeric_dtype(df[col]):
                total[col] = df[col].sum()
            else:
                total[col] = np.nan

    return pd.concat([df, pd.DataFrame([total], index=["JUMLAH"])])

# =====================================================
# EXCEL WRITER
# =====================================================
def write_formatted_sheet(writer, df, sheet_name):

    wb = writer.book
    ws = wb.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = ws

    df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1, header=False)

    fmt_header = wb.add_format({"bold": True, "align": "center", "border": 1})
    fmt_amt = wb.add_format({"num_format": "#,##0"})
    fmt_pct = wb.add_format({"num_format": "0.00%"})
    fmt_int = wb.add_format({"num_format": "0"})

    fmt_amt_bold = wb.add_format({"num_format": "#,##0", "bold": True})
    fmt_pct_bold = wb.add_format({"num_format": "0.00%", "bold": True})
    fmt_txt_bold = wb.add_format({"bold": True})

    fmt_red = wb.add_format({"bg_color": "#F8CBAD", "num_format": "0.00%"})
    fmt_green = wb.add_format({"bg_color": "#C6EFCE", "num_format": "0.00%"})

    for c, col in enumerate(df.columns):
        ws.write(0, c, col, fmt_header)

        if col in PERCENT_COLS:
            ws.set_column(c, c, 14, fmt_pct)
        elif col in INT_COLS:
            ws.set_column(c, c, 12, fmt_int)
        else:
            ws.set_column(c, c, 18, fmt_amt)

    jumlah_row = len(df)

    for c, col in enumerate(df.columns):
        val = df.iloc[-1, c]

        if pd.isna(val):
            ws.write_blank(jumlah_row, c, None, fmt_txt_bold)

        elif col in PERCENT_COLS:
            ws.write_number(jumlah_row, c, float(val), fmt_pct_bold)

        elif isinstance(val, (int, float, np.integer, np.floating)):
            ws.write_number(jumlah_row, c, float(val), fmt_amt_bold)

        else:
            ws.write(jumlah_row, c, val, fmt_txt_bold)

    if "%Result" in df.columns:
        idx = df.columns.get_loc("%Result")
        ws.conditional_format(
            1, idx, jumlah_row, idx,
            {"type": "cell", "criteria": "<", "value": 0.05, "format": fmt_red}
        )
        ws.conditional_format(
            1, idx, jumlah_row, idx,
            {"type": "cell", "criteria": ">=", "value": 0.05, "format": fmt_green}
        )

# =====================================================
# RUN
# =====================================================
if process_btn and uploaded_file:

    xls = pd.ExcelFile(uploaded_file)
    results = {
        c: run_profitability(pd.read_excel(xls, c), c)
        for c in COVERAGE_ORDER
    }

    # SUMMARY DATAFRAME
    rows, tp, tr = [], 0, 0
    for c in COVERAGE_ORDER:
        p = results[c]["Prem_Askrindo"].sum()
        r = results[c]["Result"].sum()
        rows.append([c, p, r, r / p if p else 0])
        tp += p
        tr += r

    rows.append(["JUMLAH", tp, tr, tr / tp if tp else 0])
    summary_df = pd.DataFrame(
        rows,
        columns=["Coverage", "Jumlah Premi Ourshare", "Result", "%Result"]
    )

    st.subheader("📊 Summary Profitability")
    st.dataframe(format_display(summary_df), use_container_width=True)

    for c in COVERAGE_ORDER:
        st.subheader(f"📋 Detail {c}")
        st.dataframe(format_display(add_total_row(results[c])), use_container_width=True)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        write_formatted_sheet(
            writer,
            add_total_row(summary_df, is_summary=True),
            "SUMMARY"
        )
        for c in COVERAGE_ORDER:
            write_formatted_sheet(
                writer,
                add_total_row(results[c]),
                c
            )

    st.download_button(
        "📥 Download Excel Output",
        data=output.getvalue(),
        file_name="Profitability_Output_All_Coverages.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
