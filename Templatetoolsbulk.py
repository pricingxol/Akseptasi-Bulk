import streamlit as st
import pandas as pd
import numpy as np

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

loss_ratio = st.sidebar.number_input(
    "Asumsi Loss Ratio", 0.0, 1.0, 0.4500, 0.0001, format="%.4f"
)
premi_xol = st.sidebar.number_input(
    "Asumsi Premi XOL", 0.0, 1.0, 0.1200, 0.0001, format="%.4f"
)
expense_ratio = st.sidebar.number_input(
    "Asumsi Expense", 0.0, 1.0, 0.2000, 0.0001, format="%.4f"
)

# =====================================================
# CONSTANTS
# =====================================================
KOMISI_BPPDAN = 0.35
KOMISI_MAIPARK = 0.30

RATE_MB_INDUSTRIAL = 0.0015
RATE_MB_NON_INDUSTRIAL = 0.0001
RATE_PL = 0.0005
RATE_FG = 0.0010

OR_CAP = {
    "PAR": 350_000_000_000,
    "EQVET": 350_000_000_000,
    "MACHINERY": 300_000_000_000,
    "PUBLIC LIABILITY": 80_000_000_000,
    "FIDELITY GUARANTEE": 80_000_000_000,
}

COVERAGE_ORDER = [
    "PAR",
    "EQVET",
    "MACHINERY",
    "PUBLIC LIABILITY",
    "FIDELITY GUARANTEE"
]

# =====================================================
# DISPLAY FORMAT
# =====================================================
PERCENT_COLS = [
    "Rate", "% Askrindo Share", "% Fakultatif Share",
    "% Komisi Fakultatif", "% LOL Premi",
    "%POOL", "%OR", "%Shortfall", "%Result"
]

INT_COLS = ["Kode Okupasi"]

def format_display(df):
    fmt = {}
    for c in df.columns:
        if c in PERCENT_COLS:
            fmt[c] = "{:.2%}"
        elif c in INT_COLS:
            fmt[c] = "{:.0f}"
        elif pd.api.types.is_numeric_dtype(df[c]):
            fmt[c] = "{:,.0f}"
    return df.style.format(fmt)

# =====================================================
# CORE ENGINE
# =====================================================
def run_profitability(df, coverage):

    df = df.copy()
    df.columns = [c.strip() for c in df.columns]

    # =============================
    # EXPOSURE
    # =============================
    df["TSI_IDR"] = df["TSI Full Value original currency"] * df["Kurs"]
    df["Limit_IDR"] = df["Limit of Liability original currency"].fillna(0) * df["Kurs"]
    df["TopRisk_IDR"] = df["Top Risk original currency"].fillna(0) * df["Kurs"]

    df["ExposureBasis"] = df[["Limit_IDR", "TopRisk_IDR"]].max(axis=1)

    # =============================
    # SHARE ASKRINDO
    # =============================
    df["S_Askrindo"] = df["% Askrindo Share"] * df["ExposureBasis"]

    # =============================
    # POOL
    # =============================
    if coverage == "PAR":
        df["Pool_amt"] = np.minimum(
            0.025 * df["S_Askrindo"],
            500_000_000 * df["% Askrindo Share"]
        )
        komisi_pool = KOMISI_BPPDAN

    elif coverage == "EQVET":
        rate_eq = np.where(
            df["Wilayah Gempa Prioritas"] == "DKI-JABAR-BANTEN",
            0.10, 0.25
        )
        df["Pool_amt"] = np.minimum(
            rate_eq * df["S_Askrindo"],
            100_000_000_000 * df["% Askrindo Share"]
        )
        komisi_pool = KOMISI_MAIPARK

    else:
        df["Pool_amt"] = 0
        komisi_pool = 0

    # =============================
    # FAC & OR
    # =============================
    df["Fac_amt"] = df["% Fakultatif Share"] * df["ExposureBasis"]

    df["OR_amt_raw"] = df["S_Askrindo"] - df["Pool_amt"] - df["Fac_amt"]
    df["OR_amt"] = np.minimum(
        np.maximum(df["OR_amt_raw"], 0),
        OR_CAP[coverage]
    )

    # =============================
    # SHORTFALL
    # =============================
    df["Shortfall_amt"] = np.maximum(
        df["S_Askrindo"] - (df["Pool_amt"] + df["Fac_amt"] + df["OR_amt"]),
        0
    )

    # =============================
    # % SPREADING (Exposure Basis)
    # =============================
    df["%POOL"] = np.where(df["ExposureBasis"] > 0, df["Pool_amt"] / df["ExposureBasis"], 0)
    df["%OR"] = np.where(df["ExposureBasis"] > 0, df["OR_amt"] / df["ExposureBasis"], 0)
    df["%Shortfall"] = np.where(df["ExposureBasis"] > 0, df["Shortfall_amt"] / df["ExposureBasis"], 0)

    # =============================
    # PREMIUM
    # =============================
    df["Prem100"] = df["Rate"] * df["TSI_IDR"] * df["% LOL Premi"]

    df["Prem_Askrindo"] = df["% Askrindo Share"] * df["Prem100"]
    df["Prem_POOL"] = df["%POOL"] * df["Prem100"]
    df["Prem_Fac"] = df["% Fakultatif Share"] * df["Prem100"]
    df["Prem_OR"] = df["%OR"] * df["Prem100"]
    df["Prem_Shortfall"] = df["%Shortfall"] * df["Prem100"]

    # =============================
    # COMMISSION
    # =============================
    df["Akuisisi"] = df["% Akuisisi"] * df["Prem_Askrindo"]
    df["Komisi_POOL"] = komisi_pool * df["Prem_POOL"]
    df["Komisi_Fakultatif"] = df["% Komisi Fakultatif"].fillna(0) * df["Prem_Fac"]

    # =============================
    # LOSS (FINAL & CORRECT)
    # =============================
    df["EL_BASIS"] = np.where(
        df["% LOL Premi"].notna(),
        df["TSI_IDR"] * df["% LOL Premi"],
        df["TSI_IDR"]
    )

    if coverage in ["PAR", "EQVET"]:
        df["EL_100"] = loss_ratio * df["Prem100"]

    elif coverage == "MACHINERY":
        rate_min = np.where(
            df["Occupancy"].str.lower() == "industrial",
            RATE_MB_INDUSTRIAL,
            RATE_MB_NON_INDUSTRIAL
        )
        df["EL_100"] = rate_min * loss_ratio * df["EL_BASIS"]

    elif coverage == "PUBLIC LIABILITY":
        df["EL_100"] = RATE_PL * loss_ratio * df["EL_BASIS"]

    elif coverage == "FIDELITY GUARANTEE":
        df["EL_100"] = RATE_FG * loss_ratio * df["EL_BASIS"]

    df["EL_Askrindo"] = df["% Askrindo Share"] * df["EL_100"]
    df["EL_POOL"] = df["%POOL"] * df["EL_100"]
    df["EL_Fac"] = df["% Fakultatif Share"] * df["EL_100"]

    # =============================
    # COST
    # =============================
    df["XOL"] = premi_xol * df["Prem_OR"]
    df["Expense"] = expense_ratio * df["Prem_Askrindo"]

    # =============================
    # RESULT
    # =============================
    df["Result"] = (
        df["Prem_Askrindo"]
        - df["Prem_POOL"]
        - df["Prem_Fac"]
        - df["Akuisisi"]
        + df["Komisi_POOL"]
        + df["Komisi_Fakultatif"]
        - df["EL_Askrindo"]
        + df["EL_POOL"]
        + df["EL_Fac"]
        - df["XOL"]
        - df["Expense"]
    )

    df["%Result"] = np.where(
        df["Prem_Askrindo"] != 0,
        df["Result"] / df["Prem_Askrindo"],
        0
    )

    return df

# =====================================================
# TOTAL ROW
# =====================================================
def add_total_row(df):
    total = {}
    for c in df.columns:
        if pd.api.types.is_numeric_dtype(df[c]):
            total[c] = df[c].sum()
        else:
            total[c] = ""
    total["%Result"] = (
        total["Result"] / total["Prem_Askrindo"]
        if total["Prem_Askrindo"] != 0 else 0
    )
    total_df = pd.DataFrame([total], index=["JUMLAH"])
    return pd.concat([df, total_df])

# =====================================================
# RUN APP
# =====================================================
uploaded_file = st.file_uploader("📁 Upload Excel", type=["xlsx"])
process_btn = st.button("🚀 Proses Profitability")

if process_btn and uploaded_file:
    xls = pd.ExcelFile(uploaded_file)

    for cov in COVERAGE_ORDER:
        df_raw = pd.read_excel(xls, cov)
        df_res = run_profitability(df_raw, cov)
        df_res = add_total_row(df_res)

        st.subheader(f"📋 Detail {cov}")

        if df_res["Shortfall_amt"].sum() > 0:
            st.warning(f"⚠️ Shortfall detected | {cov}")

        st.dataframe(format_display(df_res), use_container_width=True)
