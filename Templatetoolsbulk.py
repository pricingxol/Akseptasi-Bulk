import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Bulk Profitability Checker - Askrindo", layout="wide")
st.title("📊 Bulk Profitability Checker – Askrindo")

# =========================
# INPUT EXCEL
# =========================
uploaded = st.file_uploader("📁 Upload Excel Bulk Input Template", type=["xlsx"])

# =========================
# INPUT ASSUMPTIONS
# =========================
st.sidebar.header("Asumsi Profitability")
komisi_bppdan = st.sidebar.number_input("% Komisi BPPDAN", 0.0, 1.0, 0.10)
loss_ratio = st.sidebar.number_input("% Loss Ratio (AL9)", 0.0, 1.0, 0.45)
premi_xol = st.sidebar.number_input("% Premi XOL (AQ9)", 0.0, 1.0, 0.12)
expense_ratio = st.sidebar.number_input("% Expense (AR9)", 0.0, 1.0, 0.20)

process_btn = st.button("🚀 Proses Profitability")


# =========================
# ENGINE PROCESS
# =========================
def process_profitability(df):

    # Rename for easy reference (optional)
    df.columns = [c.strip() for c in df.columns]

    # Currency conversion
    df["TSI_IDR"] = df["TSI Full Value original currency"] * df["Kurs"]
    df["Limit_IDR"] = df["Limit of Liability original currency"] * df["Kurs"]
    df["TopRisk_IDR"] = df["Top Risk original currency"] * df["Kurs"]

    # Exposure Basis = basis pembagian share
    df["ExposureBasis"] = df[["Limit_IDR", "TopRisk_IDR"]].max(axis=1)

    # Askrindo retained
    df["S_Askrindo"] = df["% Askrindo Share"] * df["ExposureBasis"]

    # BPPDAN capped
    df["BPPDAN_amt"] = np.minimum(
        0.025 * df["S_Askrindo"],
        500_000_000 * df["% Askrindo Share"]
    )
    df["%BPPDAN"] = df["BPPDAN_amt"] / df["ExposureBasis"]

    # Facultative
    df["Fac_amt"] = df["% Fakultatif Share"] * df["ExposureBasis"]

    # OR = residual
    df["OR_amt"] = df["S_Askrindo"] - df["BPPDAN_amt"] - df["Fac_amt"]
    df["%OR"] = df["OR_amt"] / df["ExposureBasis"]

    # Premium 100% Share
    df["Prem100"] = (
        df["Rate"] *
        df["% LOL Premi"] *
        df["TSI_IDR"]
    )

    # Premium split
    df["Prem_Askrindo"] = df["Prem100"] * df["% Askrindo Share"]
    df["Prem_BPPDAN"] = df["Prem100"] * df["%BPPDAN"]
    df["Prem_OR"] = df["Prem100"] * df["%OR"]
    df["Prem_Fac"] = df["Prem100"] * df["% Fakultatif Share"]

    # Acquisition (Askrindo)
    df["Acq_amt"] = df["% Akuisisi"] * df["Prem_Askrindo"]

    # Komisi BPPDAN
    df["Komisi_BPPDAN"] = komisi_bppdan * df["Prem_BPPDAN"]

    # Komisi Facultative
    df["Komisi_Fac"] = df["% Komisi Fakultatif"] * df["Prem_Fac"]

    # Expected Loss
    df["EL_100"] = loss_ratio * df["Prem100"]
    df["EL_Askrindo"] = df["EL_100"] * df["% Askrindo Share"]
    df["EL_BPPDAN"] = df["EL_100"] * df["%BPPDAN"]
    df["EL_OR"] = df["EL_100"] * df["%OR"]
    df["EL_Fac"] = df["EL_100"] * df["% Fakultatif Share"]

    # XL & Expense
    df["XL_cost"] = premi_xol * df["Prem_OR"]
    df["Expense"] = expense_ratio * df["Prem_Askrindo"]

    # Final UW Result
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
# DISPLAY RESULT
# =========================
if process_btn and uploaded:
    raw_df = pd.read_excel(uploaded)
    result_df = process_profitability(raw_df.copy())

    st.success("Selesai diproses! 🎉")
    st.subheader("📊 Hasil Profitability")
    st.dataframe(result_df)

    # Summary
    st.subheader("📈 Summary Portfolio")
    st.write(f"Total Premi Askrindo: {result_df['Prem_Askrindo'].sum():,.0f}")
    st.write(f"Total Result: {result_df['Result'].sum():,.0f}")

    # Download Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        result_df.to_excel(writer, index=False, sheet_name="Profitability")
    st.download_button(
        label="📥 Download Excel Hasil",
        data=output.getvalue(),
        file_name="Profitability_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

elif process_btn and not uploaded:
    st.error("❗ Harap upload file input terlebih dahulu.")
