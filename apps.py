import pandas as pd
import streamlit as st
from io import BytesIO

#Cleansing Data
def clean_data(df, source="claim"):
    df.columns = df.columns.str.strip()
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).str.strip()

    if source == "claim" and "ClaimStatus" in df.columns:
        df = df[df["ClaimStatus"] == "R"]
    elif source == "benefit" and "Status_Claim" in df.columns:
        df = df[df["Status_Claim"] == "R"]

    df = df.drop(columns=["Status_Claim", "BAmount"], errors='ignore')
    df = df.drop(columns=["ClaimStatus"], errors='ignore')
    return df

#Save to Excel
def save_to_excel(transformed_df, benefit_df, summary_stats, filtered_cr_df, filename):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Sheet 1: Summary
        summary_df = pd.DataFrame({
            "Metric": ["Total Claims", "Total Billed", "Total Accepted", "Total Excess", "Total Unpaid"],
            "Value": summary_stats
        })
        summary_df.to_excel(writer, index=False, sheet_name='Summary', startrow=0)

        if filtered_cr_df is not None:
            filtered_cr_df.to_excel(writer, index=False, sheet_name='Summary', startrow=8)

        # Sheet 2: Transformed Claim Data
        transformed_df.to_excel(writer, index=False, sheet_name='SC')

        # Sheet 3: Benefit Data
        if benefit_df is not None:
            benefit_df.to_excel(writer, index=False, sheet_name='Benefit')

    output.seek(0)
    return output, filename

# Streamlit UI
st.title("Claim & Benefit Excel Template Generator")

# Upload file
st.subheader("Upload Required Files")
uploaded_claim = st.file_uploader("1. Upload Claim Data (.csv)", type=["csv"])
uploaded_cr = st.file_uploader("2. Upload Claim Ratio File (.xlsx)", type=["xlsx"])
uploaded_benefit = st.file_uploader("3. Upload Benefit Data (.csv)", type=["csv"])

if uploaded_claim and uploaded_cr and uploaded_benefit:
    # Proses Claim Data
    claim_df = pd.read_csv(uploaded_claim)
    transformed_data = clean_data(claim_df, source="claim")

    #Statistics Summary
    total_claims = len(transformed_data)
    total_billed = int(transformed_data["Billed"].sum())
    total_accepted = int(transformed_data["Accepted"].sum())
    total_excess = int(transformed_data["ExcessTotal"].sum())
    total_unpaid = int(transformed_data["Unpaid"].sum())
    summary_stats = [total_claims, total_billed, total_accepted, total_excess, total_unpaid]

    #Claim Ratio
    cr_df = pd.read_excel(uploaded_cr)
    cr_df.columns = cr_df.columns.str.strip()
    policy_nos = transformed_data["PolicyNo"].unique().tolist()
    filtered_cr_df = cr_df[cr_df["PolicyNo"].isin(policy_nos)]

    required_cols = ["Company", "Net Premi", "Billed", "Unpaid", "Excess Total",
                     "Excess Coy", "Excess Emp", "Claim", "CR", "Est Claim"]
    existing_cols = [col for col in required_cols if col in filtered_cr_df.columns]
    filtered_cr_df = filtered_cr_df[existing_cols]

    #Benefit
    benefit_df = pd.read_csv(uploaded_benefit)
    benefit_df = clean_data(benefit_df, source="benefit")

    #Preview
    st.subheader("Data Preview")
    st.write("Transformed Claim Data:")
    st.dataframe(transformed_data.head())

    st.write("Filtered Claim Ratio Data:")
    st.dataframe(filtered_cr_df.head())

    st.write("Filtered Benefit Data:")
    st.dataframe(benefit_df.head())

    # Input file name and download
    st.subheader("Export to Excel")
    filename = st.text_input("Enter Excel file name (without extension):", "Transformed_Claim_Data")

    if filename:
        excel_file, final_filename = save_to_excel(
            transformed_df=transformed_data,
            benefit_df=benefit_df,
            summary_stats=summary_stats,
            filtered_cr_df=filtered_cr_df,
            filename=filename + ".xlsx"
        )

        st.download_button(
            label="Download Final Excel File",
            data=excel_file,
            file_name=final_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Please upload all three files to continue.")
