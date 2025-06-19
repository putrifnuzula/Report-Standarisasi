import pandas as pd
import streamlit as st
from io import BytesIO

# ===================================
# CLAIM DATA FUNCTIONS
# ===================================
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
    # Step 4: Transform to the new template
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

# ===================================
# BENEFIT DATA FUNCTIONS
# ===================================
def filter_benefit_data(df):
    if 'Status Claim' in df.columns or 'Status_Claim' in df.columns:
        if 'Status_Claim' in df.columns:
            df = df[df['Status_Claim'] == 'R']
        else:
            df = df[df['Status Claim'] == 'R']
    else:
        st.warning("⚠️ Column 'Status Claim' not found. Data not filtered.")
    return df

def move_to_benefit_template(df):
    df = filter_benefit_data(df)
    df.columns = df.columns.str.strip()
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).str.strip()
    # Drop unnecessary columns if available
    df = df.drop(columns=["Status_Claim", "BAmount"], errors='ignore')
    return df

# ===================================
# UTILITY: Prepare Vertical Claim Ratio Summary
# ===================================
def vertical_claim_ratio_summary(df):
    """
    For each unique policy (from claim ratio unique data), output six metrics on separate rows.
    The metrics are displayed vertically as:
      Company, Net Premi, Billed, Unpaid, Claim, CR.
    Each policy's metrics are separated by a blank row.
    """
    metrics = ["Company", "Net Premi", "Billed", "Unpaid", "Claim", "CR"]
    rows = []
    for idx, r in df.iterrows():
        for metric in metrics:
            # Bold formatting will be applied when writing the Excel file.
            rows.append({"Metric": metric, "Value": r.get(metric, "")})
        # Append a blank row separating each policy block.
        rows.append({"Metric": "", "Value": ""})
    return pd.DataFrame(rows)

# ===================================
# SAVE TO EXCEL FUNCTION
# ===================================
def save_to_excel(claim_df, benefit_df, summary_df, filename):
    output = BytesIO()
    # Create a writer using the xlsxwriter engine
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        # -------------------------------
        # Write Summary sheet manually for formatting
        # -------------------------------
        summary_sheet = workbook.add_worksheet("Summary")
        # Define a bold format for metric names
        bold_format = workbook.add_format({'bold': True})
        # Write headers for the two columns: "Metric" and "Value"
        summary_sheet.write(0, 0, "Metric", bold_format)
        summary_sheet.write(0, 1, "Value")
        
        # Write each row from summary_df (starting at row 1)
        for i, row in summary_df.iterrows():
            metric = row["Metric"]
            value = row["Value"]
            if metric:  # non blank cell => bold the metric
                summary_sheet.write(i+1, 0, metric, bold_format)
            else:
                summary_sheet.write(i+1, 0, metric)
            summary_sheet.write(i+1, 1, value)
        
        # -------------------------------
        # Write Claim (SC) and Benefit sheets using pandas output
        # -------------------------------
        claim_df.to_excel(writer, index=False, sheet_name='SC')
        benefit_df.to_excel(writer, index=False, sheet_name='Benefit')
        writer.close()
    output.seek(0)
    return output, filename

# ===================================
# STREAMLIT APP UI
# ===================================
st.title("Integrated Claims, Claim Ratio, and Benefit Data Processor")

st.header("Upload Files")
uploaded_claim = st.file_uploader("Upload Claim Data (.csv)", type=["csv"], key="claim")
uploaded_claim_ratio = st.file_uploader("Upload Claim Ratio Data (.xlsx)", type=["xlsx"], key="claim_ratio")
uploaded_benefit = st.file_uploader("Upload Benefit Data (.csv)", type=["csv"], key="benefit")

if uploaded_claim and uploaded_claim_ratio and uploaded_benefit:
    # ----- Process Claim Data -----
    raw_claim = pd.read_csv(uploaded_claim)
    st.write("Processing Claim Data...")
    claim_transformed = move_to_claim_template(raw_claim)
    st.subheader("Claim Data Preview:")
    st.dataframe(claim_transformed.head())
    
    # ----- Process Claim Ratio Data -----
    claim_ratio_raw = pd.read_excel(uploaded_claim_ratio)
    # Filter: only keep rows with Policy No values present in the claim data
    policy_list = claim_transformed["Policy No"].unique().tolist()
    claim_ratio_filtered = claim_ratio_raw[claim_ratio_raw["Policy No"].isin(policy_list)]
    # We need only these columns (Policy No will not be shown in the final summary):
    desired_cols = ['Policy No', 'Company', 'Net Premi', 'Billed', 'Unpaid', 
                    'ExcessTotal', 'ExcessCoy', 'ExcessEmp', 'Claim', 'CR', 'Est Claim']
    missing_cols = [col for col in desired_cols if col not in claim_ratio_filtered.columns]
    if missing_cols:
        st.warning(f"Missing columns in Claim Ratio Data: {missing_cols}")
    claim_ratio_filtered = claim_ratio_filtered[[col for col in desired_cols if col in claim_ratio_filtered.columns]]
    # Drop duplicates based on "Policy No" (one row per unique policy)
    claim_ratio_unique = claim_ratio_filtered.drop_duplicates(subset="Policy No")
    
    st.subheader("Claim Ratio Data Preview (unique by Policy No):")
    st.dataframe(claim_ratio_unique.head())
    
    # ----- Process Benefit Data -----
    raw_benefit = pd.read_csv(uploaded_benefit)
    st.write("Processing Benefit Data...")
    benefit_transformed = move_to_benefit_template(raw_benefit)
    # Drop rows from Benefit data where 'ClaimNo' is not in the cleaned Claim data "Claim No"
    claim_no_list = claim_transformed["Claim No"].unique().tolist()
    if "ClaimNo" in benefit_transformed.columns:
        benefit_transformed = benefit_transformed[benefit_transformed["ClaimNo"].isin(claim_no_list)]
    else:
        st.warning("Column 'ClaimNo' not found in Benefit data; skipping filtering based on ClaimNo.")
    
    st.subheader("Benefit Data Preview:")
    st.dataframe(benefit_transformed.head())
    
    # ===================================
    # PREPARE THE SUMMARY SHEET
    # ===================================
    # --- Top summary: Claim statistics ---
    total_claims = len(claim_transformed)
    total_billed = int(claim_transformed["Sum of Billed"].sum())
    total_accepted = int(claim_transformed["Sum of Accepted"].sum())
    total_excess = int(claim_transformed["Sum of Excess Total"].sum())
    total_unpaid = int(claim_transformed["Sum of Unpaid"].sum())
    
    claim_summary_data = {
        "Metric": ["Total Claims", "Total Billed", "Total Accepted", "Total Excess", "Total Unpaid"],
        "Value": [f"{total_claims:,}", f"{total_billed:,.2f}", f"{total_accepted:,.2f}", f"{total_excess:,.2f}", f"{total_unpaid:,.2f}"]
    }
    claim_summary_df = pd.DataFrame(claim_summary_data)
    
    # --- Overall Claim Ratio (from Claim Ratio data) ---
    if "Claim" in claim_ratio_unique.columns and "Net Premi" in claim_ratio_unique.columns:
        total_claim_ratio_claim = claim_ratio_unique["Claim"].sum()
        total_net_premi = claim_ratio_unique["Net Premi"].sum()
        overall_cr = (total_claim_ratio_claim / total_net_premi) * 100 if total_net_premi != 0 else 0
        claim_ratio_overall = pd.DataFrame({"Metric": ["Claim Ratio (%)"],
                                            "Value": [f"{overall_cr:.2f}%"]})
    else:
        claim_ratio_overall = pd.DataFrame({"Metric": ["Claim Ratio (%)"], "Value": ["N/A"]})
    
    summary_top = pd.concat([claim_summary_df, claim_ratio_overall], ignore_index=True)
    
    # --- Blank row ---
    blank_row = pd.DataFrame({"Metric": [""], "Value": [""]})
    
    # --- Bottom summary: Vertical layout for claim ratio details ---
    vertical_cr_summary = vertical_claim_ratio_summary(claim_ratio_unique)
    
    # Combine top summary, a blank row, and vertical claim ratio summary
    summary_df = pd.concat([summary_top, blank_row, vertical_cr_summary], ignore_index=True)
    
    st.subheader("Summary Preview:")
    st.dataframe(summary_df)
    
    # ===================================
    # DOWNLOAD THE EXCEL FILE
    # ===================================
    filename_input = st.text_input("Enter the Excel file name (without extension):", "Processed_Data")
    if filename_input:
        excel_file, final_filename = save_to_excel(claim_transformed, benefit_transformed, summary_df, filename_input + ".xlsx")
        st.download_button(
            label="Download Excel File",
            data=excel_file,
            file_name=final_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
