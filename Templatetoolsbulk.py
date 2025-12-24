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
# FILE UPLOAD
# =====================================================
uploaded_file = st.file_uploader(
    "📁 Upload Excel (PAR, EQVET, MACHINERY, PUBLIC LIABILITY, FIDELITY GUARANTEE)",
    type=["xlsx"]
)

process_btn = st.button("🚀 Proses Profitability")

# =====================================================
# CORE ENGINE
# =====================================================
def run_profitability(df, coverage):

    df = df.copy()
    df.columns = [c.strip() for c in df.columns]

    # Currency & exposure
    df["TSI_IDR"] = df["TSI Full Value original currency"] * df["Kurs"]
    df["Limit_IDR"] = df["Limit of Liability original currency"] * df["Kurs"]
    df["TopRisk_IDR"] = df["Top Risk original currency"] * df["Kurs"]

    df["ExposureBasis"] = df[["Limit_IDR", "TopRisk_IDR"]].max(axis=1)
    df["Exposure_Loss"] = np.where(df["Limit_IDR"] > 0, df["Limit_IDR"], df["TSI_IDR"])

    # Retention
    df["S_Askrindo"] = df["% Askrindo Share"] * df["ExposureBasis"]

    # Pool logic
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

    # Facultative & OR
    df["Fac_amt"] = df["% Fakultatif Share"] * df["ExposureBasis"]
    df["OR_amt"] = df["S_Askrindo"] - df["Pool_amt"] - df["Fac_amt"]
    df["%OR"] = df["OR_amt"] / df["ExposureBasis"]

    # Premium
    df["Prem100"] = df["Rate"] * df["% LOL Premi"] * df["TSI_IDR"]
    df["Prem_Askrindo"] = df["Prem100"] * df["% Askrindo Share"]
    df["Prem_POOL"] = df["Prem100"] * df["%POOL"]
    df["Prem_Fac"] = df["Prem100"] * df["% Fakultatif Share"]
    df["Prem_OR"] = df["Prem100"] * df["%OR"]

    # Commission & acquisition
    df["Acq_amt"] = df["% Akuisisi"] * df["Prem_Askrindo"]
    df["Komisi_POOL"] = komisi_pool * df["Prem_POOL"]
    df["Komisi_Fakultatif"] = df["% Komisi Fakultatif"].fillna(0) * df["Prem_Fac"]

    # Expected loss
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

    # XL & expense
    df["XL_cost"] = premi_xol * df["Prem_OR"]
    df["Expense"] = expense_ratio * df["Prem_Askrindo"]

    # Result
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
# RUN PROCESS
# =====================================================
if process_btn:

    if uploaded_file is None:
        st.error("Upload file Excel terlebih dahulu.")
        st.stop()

    xls = pd.ExcelFile(uploaded_file)
    results = {}

    for cov in COVERAGE_ORDER:
        if cov not in xls.sheet_names:
            st.error(f"Sheet {cov} tidak ditemukan di Excel.")
            st.stop()
        results[cov] = run_profitability(pd.read_excel(xls, sheet_name=cov), cov)

    # ========================= SUMMARY =========================
    summary_rows = []
    total_prem = total_res = 0

    for cov in COVERAGE_ORDER:
        df = results[cov]
        prem = df["Prem_Askrindo"].sum()
        res = df["Result"].sum()
        pct = res / prem if prem != 0 else 0

        total_prem += prem
        total_res += res

        summary_rows.append([cov, prem, res, pct])

    summary_rows.append(["JUMLAH", total_prem, total_res, total_res / total_prem if total_prem != 0 else 0])

    summary_df = pd.DataFrame(summary_rows, columns=["Coverage", "Jumlah Premi Ourshare", "Result", "%Result"])

    st.subheader("📊 Summary Profitability")
    st.dataframe(
        summary_df.style.format({
            "Jumlah Premi Ourshare": "{:,.0f}",
            "Result": "{:,.0f}",
            "%Result": "{:.2%}"
        }),
        use_container_width=True
    )

    # ========================= DETAILS =========================
    for cov in COVERAGE_ORDER:
        df = results[cov]

        total_row = pd.DataFrame([{
            "Prem_Askrindo": df["Prem_Askrindo"].sum(),
            "Result": df["Result"].sum(),
            "%Result": df["Result"].sum() / df["Prem_Askrindo"].sum() if df["Prem_Askrindo"].sum() != 0 else 0
        }], index=["JUMLAH"])

        display_df = pd.concat([df, total_row], axis=0)

        st.subheader(f"📋 Detail {cov}")
        st.dataframe(
            display_df.style.format({
                "%Result": "{:.2%}",
                "Prem_Askrindo": "{:,.0f}",
                "Prem_POOL": "{:,.0f}",
                "Prem_Fac": "{:,.0f}",
                "Prem_OR": "{:,.0f}",
                "Acq_amt": "{:,.0f}",
                "Komisi_POOL": "{:,.0f}",
                "Komisi_Fakultatif": "{:,.0f}",
                "EL_100": "{:,.0f}",
                "EL_Askrindo": "{:,.0f}",
                "EL_POOL": "{:,.0f}",
                "EL_Fac": "{:,.0f}",
                "XL_cost": "{:,.0f}",
                "Expense": "{:,.0f}",
                "Result": "{:,.0f}"
            }),
            use_container_width=True
        )

    # ========================= DOWNLOAD =========================
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for cov in COVERAGE_ORDER:
            results[cov].to_excel(writer, index=False, sheet_name=cov)

    st.download_button(
        "📥 Download Excel Output",
        data=output.getvalue(),
        file_name="Profitability_Output_All_Coverages.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
