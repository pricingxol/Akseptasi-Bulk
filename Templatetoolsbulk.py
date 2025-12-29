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
st.caption("Profitability checking tool (PAR, EQVET, MB, PL, FG) by Divisi Aktuaria Askrindo")

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
# DISPLAY FORMAT RULES
# =====================================================
PERCENT_COLS = [
    "Rate", "% Askrindo Share", "% Fakultatif Share",
    "% Komisi Fakultatif", "% LOL Premi",
    "%POOL", "%OR", "%Result"
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

    # ===== IDR Conversion =====
    df["TSI_IDR"] = df["TSI Full Value original currency"] * df["Kurs"]
    df["Limit_IDR"] = df["Limit of Liability original currency"] * df["Kurs"]
    df["TopRisk_IDR"] = df["Top Risk original currency"] * df["Kurs"]

    df["ExposureBasis"] = df[["Limit_IDR", "TopRisk_IDR"]].max(axis=1)
    df["Exposure_OR"] = np.minimum(df["ExposureBasis"], OR_CAP[coverage])

    # ===== Spreading =====
    df["S_Askrindo"] = df["% Askrindo Share"] * df["Exposure_OR"]

    if coverage == "PAR":
        df["Pool_amt"] = np.minimum(
            0.025 * df["S_Askrindo"],
            500_000_000 * df["% Askrindo Share"]
        )
        komisi_pool = KOMISI_BPPDAN

    elif coverage == "EQVET":
        rate = np.where(
            df["Wilayah Gempa Prioritas"] == "DKI-JABAR-BANTEN",
            0.10, 0.25
        )
        df["Pool_amt"] = np.minimum(
            rate * df["S_Askrindo"],
            10_000_000_000 * df["% Askrindo Share"]
        )
        komisi_pool = KOMISI_MAIPARK

    else:
        df["Pool_amt"] = 0
        komisi_pool = 0

    df["Fac_amt"] = df["% Fakultatif Share"] * df["Exposure_OR"]

    df["OR_amt"] = np.maximum(
        df["S_Askrindo"] - df["Pool_amt"] - df["Fac_amt"], 0
    )

    # ===== SHORTFALL =====
    df["Shortfall_amt"] = np.maximum(
        df["Exposure_OR"] - (df["Pool_amt"] + df["Fac_amt"] + df["OR_amt"]),
        0
    )

    df["%POOL"] = np.where(df["Exposure_OR"] > 0, df["Pool_amt"] / df["Exposure_OR"], 0)
    df["%OR"] = np.where(df["Exposure_OR"] > 0, df["OR_amt"] / df["Exposure_OR"], 0)

    # ===== PREMIUM =====
    df["Prem100"] = df["Rate"] * df["% LOL Premi"] * df["TSI_IDR"]

    df["Prem_Askrindo_Normal"] = df["Prem100"] * df["% Askrindo Share"]
    df["Prem_Shortfall"] = df["Rate"] * df["% LOL Premi"] * df["Shortfall_amt"]
    df["Prem_Askrindo"] = df["Prem_Askrindo_Normal"] + df["Prem_Shortfall"]

    df["Prem_POOL"] = df["Prem100"] * df["%POOL"]
    df["Prem_Fac"] = df["Prem100"] * df["% Fakultatif Share"]

    # ===== COMMISSION =====
    df["Acq_amt"] = df["% Akuisisi"] * df["Prem_Askrindo"]
    df["Komisi_POOL"] = komisi_pool * df["Prem_POOL"]
    df["Komisi_Fakultatif"] = df["% Komisi Fakultatif"].fillna(0) * df["Prem_Fac"]

    # ===== LOSS =====
    if coverage == "MACHINERY":
        rate_acuan = np.where(
            df["Occupancy"].str.lower() == "industrial",
            RATE_MB_INDUSTRIAL,
            RATE_MB_NON_INDUSTRIAL
        )
        df["EL_100"] = rate_acuan * df["ExposureBasis"] * loss_ratio

    elif coverage == "PUBLIC LIABILITY":
        df["EL_100"] = RATE_PL * df["ExposureBasis"] * loss_ratio

    elif coverage == "FIDELITY GUARANTEE":
        df["EL_100"] = RATE_FG * df["ExposureBasis"] * loss_ratio

    else:
        df["EL_100"] = loss_ratio * df["Prem100"]

    df["EL_Askrindo_Normal"] = df["EL_100"] * df["% Askrindo Share"]
    df["EL_Shortfall"] = np.where(
        df["ExposureBasis"] > 0,
        df["EL_100"] * (df["Shortfall_amt"] / df["ExposureBasis"]),
        0
    )
    df["EL_Askrindo"] = df["EL_Askrindo_Normal"] + df["EL_Shortfall"]

    df["EL_POOL"] = df["EL_100"] * df["%POOL"]
    df["EL_Fac"] = df["EL_100"] * df["% Fakultatif Share"]

    # ===== COST =====
    df["XL_cost"] = premi_xol * df["OR_amt"]
    df["Expense"] = expense_ratio * df["Prem_Askrindo"]

    # ===== RESULT =====
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
# RUN APP
# =====================================================
uploaded_file = st.file_uploader("📁 Upload Excel", type=["xlsx"])
process_btn = st.button("🚀 Proses Profitability")

if process_btn and uploaded_file:

    xls = pd.ExcelFile(uploaded_file)
    results = {c: run_profitability(pd.read_excel(xls, c), c) for c in COVERAGE_ORDER}

    for c in COVERAGE_ORDER:
        df = results[c]
        st.subheader(f"📋 Detail {c}")

        shortfall_total = df["Shortfall_amt"].sum()
        if shortfall_total > 0:
            st.warning(
                f"⚠️ Shortfall detected | {c} | "
                f"{(df['Shortfall_amt'] > 0).sum()} policy | "
                f"Total Shortfall = {shortfall_total:,.0f}"
            )

        st.dataframe(format_display(df), use_container_width=True)
