import pandas as pd
import streamlit as st
from io import BytesIO

# Claim data functions
def filter_claim_data(df):
    df = df[df['ClaimStatus'] == 'R']
    return df

def keep_last_duplicate_claim(df):
    duplicate_claims = df[df.duplicated(subset='ClaimNo', keep=False)]
    if not duplicate_claims.empty:
        st.write("Duplicated ClaimNo values:")
        st.dataframe(duplicate_claims[['ClaimNo']].drop_duplicates())
    df = df.drop_duplicates(subset='ClaimNo', keep='last')
    return df

def move_to_claim_template(df):
    # Step 1: Filter the data
    new_df = filter_claim_data(df)
    # Step 2: Handle duplicates
    new_df = keep_last_duplicate_claim(new_df)
    # Step 3: Convert date columns to datetime
    date_columns = ["TreatmentStart", "TreatmentFinish", "Date"]
    for col in date_columns:
        new_df[col] = pd.to_datetime(new_df[col], errors='coerce')
        if new_df[col].isnull().any():
            st.warning(f"Invalid date values detected in column '{col}'. Coerced to NaT.")
    # Step 4: Build the new template
    df_transformed = pd.DataFrame({
        "No": range(1, len(new_df) + 1),
        "Policy No": new_df["PolicyNo"],
        "Client Name": new_df["ClientName"],
        "Claim No": new_df["ClaimNo"],
        "Member No": new_df["MemberNo"],
        "Emp ID": new_df["EmpID"],
        "Emp Name": new_df["EmpName"],
        "Patient Name": new_df["PatientName"],
        "Membership": new_df["Membership"],
        "Product Type": new_df["ProductType"],
        "Claim Type": new_df["ClaimType"],
        "Room Option": new_df["RoomOption"].fillna('').astype(str).str.upper().str.replace(r"\s+", "", regex=True),
        "Area": new_df["Area"],
        "Plan": new_df["PPlan"],
        "Diagnosis": new_df["PrimaryDiagnosis"].str.upper(),
        "Treatment Place": new_df["TreatmentPlace"].str.upper(),
        "Treatment Start": new_df["TreatmentStart"],
        "Treatment Finish": new_df["TreatmentFinish"],
        "Settled Date": new_df["Date"],
        "Year": new_df["Date"].dt.year,
        "Month": new_df["Date"].dt.month,
        "Length of Stay": new_df["LOS"],
        "Sum of Billed": new_df["Billed"],
        "Sum of Accepted": new_df["Accepted"],
        "Sum of Excess Coy": new_df["ExcessCoy"],
        "Sum of Excess Emp": new_df["ExcessEmp"],
        "Sum of Excess Total": new_df["ExcessTotal"],
        "Sum of Unpaid": new_df["Unpaid"]
    })
    return df_transformed

# Benefit data functions:
def filter_benefit_data(df):
    if 'Status Claim' in df.columns or 'Status_Claim' in df.columns:
        if 'Status_Claim' in df.columns:
            df = df[df['Status_Claim'] == 'R']
        else:
            df = df[df['Status Claim'] == 'R']
    else:
        st.warning("Column 'Status Claim' not found. Data not filtered.")
    return df

def move_to_benefit_template(df):
    df = filter_benefit_data(df)
    df.columns = df.columns.str.strip()
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).str.strip()
    # Rename Benefit columns
    rename_mapping = {
        'ClientName': 'Client Name',
        'PolicyNo': 'Policy No',
        'ClaimNo': 'Claim No',
        'MemberNo': 'Member No',
        'PatientName': 'Patient Name',
        'EmpID': 'Emp ID',
        'EmpName': 'Emp Name',
        'ClaimType': 'Claim Type',
        'TreatmentPlace': 'Treatment Place',
        'RoomOption': 'Room Option',
        'TreatmentRoomClass': 'Treatment Room Class',
        'TreatmentStart': 'Treatment Start',
        'TreatmentFinish': 'Treatment Finish',
        'ProductType': 'Product Type',
        'BenefitName': 'Benefit Name',
        'PaymentDate': 'Payment Date',
        'ExcessTotal': 'Excess Total',
        'ExcessCoy': 'Excess Coy',
        'ExcessEmp': 'Excess Emp'
    }
    df = df.rename(columns=rename_mapping)
    # Drop unnecessary columns if available
    df = df.drop(columns=["Status_Claim", "BAmount"], errors='ignore')
    return df

# Save to excel
def save_to_excel(claim_df, benefit_df, summary_top_df, claim_ratio_df, filename):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        # Summary Sheet:
        summary_sheet = workbook.add_worksheet("Summary")
        bold_format = workbook.add_format({'bold': True})
        
        row = 0
        # Write summary top statistics (no header row)
        # Each row will simply have Metric and Value, one under the other.
        for i, data in summary_top_df.iterrows():
            summary_sheet.write(row, 0, data["Metric"], bold_format)
            summary_sheet.write(row, 1, data["Value"])
            row += 1

        # Insert a blank row (between sum stats and CR)
        row += 1

        # Write header for Claim Ratio table (horizontal layout)
        cr_columns = ["Company", "Net Premi", "Billed", "Unpaid", 
                      "Excess Total", "Excess Coy", "Excess Emp", "Claim", "CR"]
        col = 0
        for header in cr_columns:
            summary_sheet.write(row, col, header, bold_format)
            col += 1
        row += 1

        # Write each row of the Claim Ratio table
        for i, data in claim_ratio_df.iterrows():
            col = 0
            for header in cr_columns:
                summary_sheet.write(row, col, data.get(header, ""))
                col += 1
            row += 1

        # SC & Benefit sheets naming
        claim_df.to_excel(writer, index=False, sheet_name='SC')
        benefit_df.to_excel(writer, index=False, sheet_name='Benefit')
        writer.close()
    output.seek(0)
    return output, filename

# Streamlit app UI
st.title("Template - Standardisasi Report")
st.header("Upload Files")

uploaded_claim = st.file_uploader("Upload Claim Data (.csv)", type=["csv"], key="claim")
uploaded_claim_ratio = st.file_uploader("Upload Claim Ratio Data (.xlsx)", type=["xlsx"], key="claim_ratio")
uploaded_benefit = st.file_uploader("Upload Benefit Data (.csv)", type=["csv"], key="benefit")

if uploaded_claim and uploaded_claim_ratio and uploaded_benefit:
    # Process claim data
    raw_claim = pd.read_csv(uploaded_claim)
    st.write("Processing Claim Data...")
    claim_transformed = move_to_claim_template(raw_claim)
    st.subheader("Claim Data Preview:")
    st.dataframe(claim_transformed.head())

    # Process claim ratio data
    claim_ratio_raw = pd.read_excel(uploaded_claim_ratio)
    # Filter: keep only rows with 'Policy No' that appear in Claim data
    policy_list = claim_transformed["Policy No"].unique().tolist()
    claim_ratio_filtered = claim_ratio_raw[claim_ratio_raw["Policy No"].isin(policy_list)]
    # Drop duplicates based on "Policy No"
    claim_ratio_unique = claim_ratio_filtered.drop_duplicates(subset="Policy No")
    # Select desired columns (excluding "Policy No")
    desired_cols = ['Company', 'Net Premi', 'Billed', 'Unpaid', 
                    'Excess Total', 'Excess Coy', 'Excess Emp', 'Claim', 'CR', 'Est CR Total']
    missing_cols = [col for col in desired_cols if col not in claim_ratio_unique.columns]
    if missing_cols:
        st.warning(f"Missing columns in Claim Ratio Data: {missing_cols}")
    claim_ratio_unique = claim_ratio_unique[[col for col in desired_cols if col in claim_ratio_unique.columns]]
    # Rename 'Est CR Total' to 'Est Claim'
    claim_ratio_unique = claim_ratio_unique.rename(columns={'Est CR Total': 'Est Claim'})
    # For the horizontal summary table, use only the first 9 columns
    summary_cr_df = claim_ratio_unique[['Company', 'Net Premi', 'Billed', 'Unpaid', 
                                         'Excess Total', 'Excess Coy', 'Excess Emp', 'Claim', 'CR', 'Est Claim']]
    
    st.subheader("Claim Ratio Data Preview (unique by Policy No):")
    st.dataframe(summary_cr_df.head())

    # Process benefit data
    raw_benefit = pd.read_csv(uploaded_benefit)
    st.write("Processing Benefit Data...")
    benefit_transformed = move_to_benefit_template(raw_benefit)
    # Retain rows where Benefit data's 'ClaimNo' appears in Claim data "Claim No"
    claim_no_list = claim_transformed["Claim No"].unique().tolist()
    if "ClaimNo" in benefit_transformed.columns:
        benefit_transformed = benefit_transformed[benefit_transformed["ClaimNo"].isin(claim_no_list)]
    elif "Claim No" in benefit_transformed.columns:
        benefit_transformed = benefit_transformed[benefit_transformed["Claim No"].isin(claim_no_list)]
    else:
        st.warning("Column 'ClaimNo' not found in Benefit data; skipping filtering based on ClaimNo.")
    
    st.subheader("Benefit Data Preview:")
    st.dataframe(benefit_transformed.head())

    # Summary Top Section (Claim Stats + Overall Claim Ratio)
    total_claims = len(claim_transformed)
    total_billed = int(claim_transformed["Sum of Billed"].sum())
    total_accepted = int(claim_transformed["Sum of Accepted"].sum())
    total_excess = int(claim_transformed["Sum of Excess Total"].sum())
    total_unpaid = int(claim_transformed["Sum of Unpaid"].sum())
    
    claim_summary_data = {
        "Metric": ["Total Claims", "Total Billed", "Total Accepted", "Total Excess", "Total Unpaid"],
        "Value": [f"{total_claims:,}", f"{total_billed:,.2f}", f"{total_accepted:,.2f}",
                  f"{total_excess:,.2f}", f"{total_unpaid:,.2f}"]
    }
    claim_summary_df = pd.DataFrame(claim_summary_data)
    
    if "Claim" in claim_ratio_unique.columns and "Net Premi" in claim_ratio_unique.columns:
        total_claim_ratio_claim = claim_ratio_unique["Claim"].sum()
        total_net_premi = claim_ratio_unique["Net Premi"].sum()
        overall_cr = (total_claim_ratio_claim / total_net_premi) * 100 if total_net_premi != 0 else 0
        claim_ratio_overall = pd.DataFrame({"Metric": ["Claim Ratio (%)"],
                                            "Value": [f"{overall_cr:.2f}%"]})
    else:
        claim_ratio_overall = pd.DataFrame({"Metric": ["Claim Ratio (%)"], "Value": ["N/A"]})
    
    summary_top_df = pd.concat([claim_summary_df, claim_ratio_overall], ignore_index=True)
    
    # Download excel file
    filename_input = st.text_input("Enter the Excel file name (without extension):", "SC & Benefit - - YTD")
    if filename_input:
        excel_file, final_filename = save_to_excel(claim_transformed, benefit_transformed,
                                                   summary_top_df, summary_cr_df, filename_input + ".xlsx")
        st.download_button(
            label="Download Excel File",
            data=excel_file,
            file_name=final_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
