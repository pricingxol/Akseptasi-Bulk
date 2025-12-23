import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# =========================
# PAGE CONFIG
# =========================
st.set_page_config(
    page_title="Bulk Profitability Checker – PAR",
    layout="wide"
)

st.title("📊 Bulk Profitability Checker – PAR")
st.caption("Tool untuk mengecek profitability portofolio PAR (bukan pricing)")

# =========================
# SIDEBAR – ASSUMPTIONS
# =========================
st.sidebar.header("Asumsi Profitability")

komisi_bppdan = st.sidebar.number_input(
    "% Komisi BPPDAN",
    min_value=0.0, max_value=1.0, value=0.10, step=0.01
)

loss_ratio = st.sidebar.number_input(
    "% Asumsi Loss Ratio (AL9)",
    min_value=0.0, max_value=1.0, value=0.45, step=0.01
)

premi_xol = st.sidebar.number_input(
    "% Premi XOL (AQ9)",
    min_value=0.0, max_value=1.0, value=0.12, step=0.01
)

expense_ratio = st.sidebar.number_input(
    "% Expense (AR9)",
    min_value=0.0, max_value=1.0, value=0.20, step=0.01
)

# =========================
# FILE UPLOAD
# =========================
uploaded_file = st.file_uploader(
    "📁 Upload Excel Bulk Input (Template)",
    type=["xlsx"]
)

process_btn = st.button("🚀 Proses Profitability")

# =========================
# CORE ENGINE
# =========================
def process_profitability(df: pd.DataFrame) -> pd.DataFrame:

    df.columns = [c.strip() for c in df.columns]

    required_cols = [
        "TSI Full Value original currency",
        "Limit of Liability original currency",
        "Top Risk original currency",
        "Kurs",
        "% Askrindo Share",
        "% Fakultatif Share",
        "Rate",
        "% LOL Premi",
        "% Akuisisi",
        "% Komisi Fakultatif"
    ]

    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"❌ Kolom berikut tidak ditemukan di Excel: {missing}")
        st.stop()

    # =========================
    # CURRENCY & EXPOSURE
    # =========================
    df["TSI_IDR"] = df["TSI Full Value original currency"] * df["Kurs"]
    df["Limit_IDR"] = df["Limit of Liability original currency"] * df["Kurs"]
    df["TopRisk_IDR"] = df["Top Risk original currency"] * df["Kurs"]

    df["ExposureBasis"] = df[["Limit_IDR", "TopRisk_IDR"]].max(axis=1)

    # =========================
    # SHARE STRUCTURE
    # =========================
    df["S_Askrindo"] = df["% Askrindo Share"] * df["ExposureBasis"]

    df["BPPDAN_amt"] = np.minimum(
        0.025 * df["S_Askrindo"],
        500_000_000 * df["% Askrindo Share"]
    )

    df["%BPPDAN"] = df["BPPDAN_amt"] / df["ExposureBasis"]

    df["Fac_amt"] = df["% Fakultatif Share"] * df["ExposureBasis"]

    df["OR_amt"] = df["S_Askrindo"] - df["BPPDAN_amt"] - df["Fac_amt"]
    df["%OR"] = df["OR_amt"] / df["ExposureBasis"]

    # =========================
    # PREMIUM
    # =========================
    df["Prem100"] = (
        df["Rate"]
        * df["% LOL Premi"]
        * df["TSI_IDR"]
    )

    df["Prem_Askrindo"] = df["Prem100"] * df["% Askrindo Share"]
    df["Prem_BPPDAN"] = df["Prem100"] * df["%BPPDAN"]
    df["Prem_OR"] = df["Prem100"] * df["%OR"]
    df["Prem_Fac"] = df["Prem100"] * df["% Fakultatif Share"]

    # =========================
    # COMMISSION & ACQUISITION
    # =========================
    df["Acq_amt"] = df["% Akuisisi"] * df["Prem_Askrindo"]
    df["Komisi_BPPDAN"] = komisi_bppdan * df["Prem_BPPDAN"]
    df["Komisi_Fac"] = df["% Komisi Fakultatif"] * df["Prem_Fac"]

    # =========================
    # EXPECTED LOSS
    # =========================
    df["EL_100"] = loss_ratio * df["Prem100"]
    df["EL_Askrindo"] = df["EL_100"] * df["% Askrindo Share"]
    df["EL_BPPDAN"] = df["EL_100"] * df["%BPPDAN"]
    df["EL_OR"] = df["EL_100"] * df["%OR"]
    df["EL_Fac"] = df["EL_100"] * df["% Fakultatif Share"]

    # =========================
    # XL & EXPENSE
    # =========================
    df["XL_cost"] = premi_xol * df["Prem_OR"]
    df["Expense"] = expense_ratio * df["Prem_Askrindo"]

    # =========================
    # FINAL RESULT (AS10)
    # =========================
    df["Result"] = (
        df["Prem_Askrindo"]
        - df["Prem_BPPDAN"]
        - df["Prem_Fac"]
        - df["Acq_amt"]
        + df["Komisi_BPPDAN"]
        + df["Komisi_Fac"]
        - df["EL_Askrindo"]
        + df["EL_BPPDAN"]
        + df["EL_Fac"]
        - df["XL_cost"]
        - df["Expense"]
    )

    return df


# =========================
# RUN PROCESS
# =========================
if process_btn:

    if uploaded_file is None:
        st.error("❗ Silakan upload file Excel terlebih dahulu.")
        st.stop()

    raw_df = pd.read_excel(uploaded_file)
    result_df = process_profitability(raw_df.copy())

    st.success("✅ Profitability berhasil dihitung")

    st.subheader("📋 Hasil Profitability per Polis")
    st.dataframe(result_df, use_container_width=True)

    st.subheader("📈 Summary Portofolio")
    st.metric("Total Premi Askrindo", f"{result_df['Prem_Askrindo'].sum():,.0f}")
    st.metric("Total Result", f"{result_df['Result'].sum():,.0f}")

    # =========================
    # DOWNLOAD EXCEL
    # =========================
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        result_df.to_excel(writer, index=False, sheet_name="Profitability_PAR")

    st.download_button(
        label="📥 Download Excel Output",
        data=output.getvalue(),
        file_name="Profitability_PAR_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
