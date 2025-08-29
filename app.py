import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Loan Summary App", layout="wide")

st.title("📊 Loan Summary Dashboard")

# --- Upload Excel file ---
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    merged_data = []

    # Merge all sheets
    for sheet_name in xls.sheet_names:
        df = xls.parse(sheet_name)
        df['SheetDate'] = sheet_name
        merged_data.append(df)

    df_combined = pd.concat(merged_data, ignore_index=True)

    # Convert to numeric
    df_combined['LOAN AMOUNT'] = pd.to_numeric(df_combined['LOAN AMOUNT'], errors='coerce')
    df_combined['DISBURSEMENT AMOUNT'] = pd.to_numeric(df_combined['DISBURSEMENT AMOUNT'], errors='coerce')

    # --- Step 1: Aggregate ---
    summary = df_combined.groupby(
        ['SheetDate', 'NAME OF EMPLOYER', 'LOAN TYPE']
    ).agg(
        total_loan_amount=('LOAN AMOUNT', 'sum'),
        total_disbursed_amount=('DISBURSEMENT AMOUNT', 'sum'),
        loan_count=('LOAN AMOUNT', 'count')
    ).reset_index()

    # --- Step 2: Clean LOAN TYPE ---
    def clean_loan_type(value):
        if pd.isna(value): return value
        value = str(value).upper()
        if "NEW" in value: return "New"
        elif "TOP UP" in value: return "Top up"
        elif "RETURNING" in value: return "Returning"
        else: return "Other"

    summary["LOAN TYPE"] = summary["LOAN TYPE"].apply(clean_loan_type)

    # --- Step 3: Clean EMPLOYER ---
    def clean_employer(value):
        if pd.isna(value): return value
        value = str(value)
        value = re.sub(r'\bTOP UP\b', '', value, flags=re.IGNORECASE)
        value = re.sub(r'\bNEW\b', '', value, flags=re.IGNORECASE)
        value = re.sub(r'\bRETURNING\b', '', value, flags=re.IGNORECASE)
        return value.strip()

    summary["NAME OF EMPLOYER"] = summary["NAME OF EMPLOYER"].apply(clean_employer)

    # --- Step 4: Pivot for final summary ---
    overall = summary.pivot_table(
        index="SheetDate",
        columns="LOAN TYPE",
        values="total_disbursed_amount",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    overall = overall.rename(columns={
        "New": "NEW LOAN",
        "Returning": "RETURNING",
        "Top up": "TOP UP"
    })

    employer = summary.pivot_table(
        index="SheetDate",
        columns=["NAME OF EMPLOYER", "LOAN TYPE"],
        values="total_disbursed_amount",
        aggfunc="sum",
        fill_value=0
    )
    employer.columns = [f"{emp} {loan}".strip() for emp, loan in employer.columns]
    employer = employer.reset_index()

    final = pd.merge(overall, employer, on="SheetDate", how="outer")
    final = final.reindex(columns=sorted(final.columns, key=lambda x: (x != "SheetDate", x)))

    # --- Display ---
    st.subheader("📌 Cleaned Data")
    st.dataframe(summary, use_container_width=True)

    st.subheader("📌 Final Summary Report")
    st.dataframe(final, use_container_width=True)

    # --- Download Excel ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        summary.to_excel(writer, sheet_name="Cleaned Data", index=False)
        final.to_excel(writer, sheet_name="Summary Report", index=False)
    st.download_button(
        label="📥 Download Summary Excel",
        data=output.getvalue(),
        file_name="loan_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
