import pandas as pd
import streamlit as st
from io import BytesIO

# Function to filter and clean claim data
def filter_data(df):
    if 'Status Claim' in df.columns:
        df = df[df['Status_Claim'] == 'R']
    return df

def move_to_template(df):
    df.columns = df.columns.str.strip()
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).str.strip()
    df = df.drop(columns=["Status_Claim", "BAmount"], errors='ignore')
    return filter_data(df)

# Save all data into Excel file with 3 sheets
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

# --- Streamlit UI ---
st.title("Claim & Benefit Excel Template Generator")

# Upload all files first
st.subheader("Upload Required Files")
uploaded_claim = st.file_uploader("Claim Data (.csv)", type=["csv"])
uploaded_cr = st.file_uploader("Claim Ratio File (.xlsx)", type=["xlsx"])
uploaded_benefit = st.file_uploader("Benefit Data (.csv)", type=["csv"])

if uploaded_claim and uploaded_cr and uploaded_benefit:
    # Process Claim Data
    claim_df = pd.read_csv(uploaded_claim)
    transformed_data = move_to_template(claim_df)

    # Summary stats
    total_claims = len(transformed_data)
    total_billed = int(transformed_data["Billed"].sum())
    total_accepted = int(transformed_data["Accepted"].sum())
    total_excess = int(transformed_data["ExcessTotal"].sum())
    total_unpaid = int(transformed_data["Unpaid"].sum())
    summary_stats = [total_claims, total_billed, total_accepted, total_excess, total_unpaid]

    # Process Claim Ratio
    cr_df = pd.read_excel(uploaded_cr)
    cr_df.columns = cr_df.columns.str.strip()
    policy_nos = transformed_data["Policy No"].unique().tolist()
    filtered_cr_df = cr_df[cr_df["Policy No"].isin(policy_nos)]
    required_cols = ["Company", "Net Premi", "Billed", "Unpaid", "Excess Total",
                     "Excess Coy", "Excess Emp", "Claim", "CR", "Est Claim"]
    existing_cols = [col for col in required_cols if col in filtered_cr_df.columns]
    filtered_cr_df = filtered_cr_df[existing_cols]

    # Load Benefit File
    benefit_df = pd.read_csv(uploaded_benefit)

    # Show Previews
    st.subheader("Data Preview")
    st.write("Transformed Claim Data:")
    st.dataframe(transformed_data.head())

    st.write("Filtered Claim Ratio (based on Policy No):")
    st.dataframe(filtered_cr_df.head())

    st.write("Raw Benefit Data:")
    st.dataframe(benefit_df.head())

    # File name + download
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
    st.info("Please upload all three files above to continue.")
